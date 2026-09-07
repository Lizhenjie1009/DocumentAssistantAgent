# -*- coding: utf-8 -*-
"""Feature 2: LLM-assisted drafting.

Flow: ask the user a few questions -> send the matched template's chapter
outline + user info to a DeepSeek-compatible chat API -> the model returns a
structured JSON (chapters + paragraphs/tables/lists) -> we render it into a
real .docx (with genuine Heading styles) under the input folder -> the normal
conversion pipeline turns it into the template-formatted output.
"""
import json
import os
import re
import urllib.error
import urllib.request
from dataclasses import dataclass

from docx import Document
from docx.shared import RGBColor

from .report import FileResult, write_report
from .util import strip_number, detect_heading_level, normalize_text

DEFAULT_BASE_URL = 'https://api.deepseek.com'
DEFAULT_MODEL = 'deepseek-chat'

SYSTEM_PROMPT = (
    '你是专业的军工/软件工程文档撰写助手。根据用户给出的模板章节大纲和素材要点，'
    '起草文档正文内容。\n'
    '要求：\n'
    '1) 只输出一个 JSON 对象，不要输出任何解释文字、注释或 Markdown 代码块；\n'
    '2) 每个章节的 level 和 heading 必须严格使用大纲给出的值，不得增删改章节标题；\n'
    '3) 正文使用正式中文技术文档风格，语句完整通顺；\n'
    '4) 素材不足时可写合理的、留待人工核实的占位内容，但尽量具体。\n'
)

JSON_SCHEMA_HINT = (
    '输出 JSON 结构如下（chapters 顺序与大写大纲一致）：\n'
    '{\n'
    '  "title": "文档完整标题",\n'
    '  "chapters": [\n'
    '    {\n'
    '      "level": 1,\n'
    '      "heading": "章节标题（原样来自大纲）",\n'
    '      "blocks": [\n'
    '        {"type": "para", "text": "段落文字"},\n'
    '        {"type": "list", "items": ["条目1", "条目2"]},\n'
    '        {"type": "table", "caption": "表题(可省略)", "headers": ["列1", "列2"], "rows": [["a", "b"]]},\n'
    '        {"type": "placeholder", "text": "【此处插入XXX图片】"}\n'
    '      ]\n'
    '    }\n'
    '  ]\n'
    '}\n'
    'blocks 的 type 只能是 para / list / table / placeholder 之一。\n'
)


@dataclass
class LLMConfig:
    api_key: str = ''
    base_url: str = DEFAULT_BASE_URL
    model: str = DEFAULT_MODEL
    temperature: float = 0.7
    timeout: int = 240


def load_llm_config(path=None):
    cfg = LLMConfig(api_key=os.environ.get('DEEPSEEK_API_KEY', ''))
    if not path and os.path.exists('llm.yaml'):
        path = 'llm.yaml'
    if path and os.path.exists(path):
        import yaml
        with open(path, 'r', encoding='utf-8') as f:
            data = yaml.safe_load(f) or {}
        if data.get('api_key'):
            cfg.api_key = str(data['api_key'])
        cfg.base_url = str(data.get('base_url', cfg.base_url))
        cfg.model = str(data.get('model', cfg.model))
        if data.get('temperature') is not None:
            cfg.temperature = float(data['temperature'])
    return cfg


# --------------------------------------------------------------------------- #
# LLM call
# --------------------------------------------------------------------------- #
def call_llm(messages, cfg):
    """POST to an OpenAI-compatible chat endpoint (stdlib urllib only)."""
    url = cfg.base_url.rstrip('/') + '/chat/completions'
    payload = json.dumps({
        'model': cfg.model,
        'messages': messages,
        'temperature': cfg.temperature,
        'stream': False,
        'response_format': {'type': 'json_object'},
    }).encode('utf-8')
    req = urllib.request.Request(url, data=payload, method='POST', headers={
        'Content-Type': 'application/json',
        'Authorization': 'Bearer ' + cfg.api_key,
    })
    try:
        with urllib.request.urlopen(req, timeout=cfg.timeout) as resp:
            data = json.loads(resp.read().decode('utf-8'))
    except urllib.error.HTTPError as e:
        body = e.read().decode('utf-8', 'replace')[:300]
        raise RuntimeError(f'API 请求失败 HTTP {e.code}: {body}')
    except Exception as e:
        raise RuntimeError(f'API 请求失败: {e}')
    try:
        return data['choices'][0]['message']['content']
    except (KeyError, IndexError, TypeError):
        raise RuntimeError('API 返回格式异常: ' + json.dumps(data, ensure_ascii=False)[:300])


# --------------------------------------------------------------------------- #
# Template outline -> prompt
# --------------------------------------------------------------------------- #
def template_outline(template_doc, tpl_cfg, max_level=3):
    """Return [{'level':.., 'heading':..}] for the template's body chapters."""
    out = []
    started = False
    for p in template_doc.paragraphs:
        name = (p.style.name or '') if p.style is not None else ''
        if not started:
            if name != tpl_cfg.body_start_style:
                continue
            started = True
        lvl = detect_heading_level(p)
        if lvl is not None and 1 <= lvl <= max_level:
            text = strip_number(p.text).strip()
            if text:
                out.append({'level': lvl, 'heading': text})
    return out


def build_user_message(tpl_cfg, outline, info):
    parts = [
        f'文档模板：{tpl_cfg.name}',
        f'文档标题：{info.get("title") or ""}',
    ]
    if info.get('system'):
        parts.append(f'系统/项目名称：{info["system"]}')
    parts.append('章节大纲（严格照此输出，不得增删改标题文本）：')
    parts.append(json.dumps(outline, ensure_ascii=False))
    if info.get('notes'):
        parts.append('素材要点（写作以此为准，缺细节可合理补充）：\n' + info['notes'])
    parts.append('若素材不足，可写合理、留待人工核实的占位内容，但尽量具体正式。')
    parts.append(JSON_SCHEMA_HINT)
    return '\n'.join(parts)


def parse_content_json(text):
    """Parse the model reply into a content dict; None if invalid."""
    if not text:
        return None
    t = text.strip()
    t = re.sub(r'^```(?:json)?\s*', '', t)
    t = re.sub(r'\s*```$', '', t)
    try:
        content = json.loads(t)
    except (json.JSONDecodeError, ValueError):
        # try to salvage the first {...} block
        m = re.search(r'\{.*\}', t, re.S)
        if not m:
            return None
        try:
            content = json.loads(m.group(0))
        except (json.JSONDecodeError, ValueError):
            return None
    if not isinstance(content, dict):
        return None
    chapters = content.get('chapters')
    if not isinstance(chapters, list) or not chapters:
        return None
    if not isinstance(content.get('title'), str):
        content['title'] = ''
    return content


def _missing_chapters(outline, content):
    """Outline chapters the model failed to produce (by level + normalized title)."""
    have = set()
    for ch in content.get('chapters') or []:
        try:
            lvl = int(ch.get('level', 1))
        except (TypeError, ValueError):
            lvl = 1
        have.add((lvl, normalize_text(str(ch.get('heading') or ''))))
    return [o for o in outline if (o['level'], normalize_text(o['heading'])) not in have]


def generate_content(tpl_cfg, outline, info, llm_cfg, retries=2):
    user_msg = build_user_message(tpl_cfg, outline, info)
    for attempt in range(retries):
        text = call_llm([
            {'role': 'system', 'content': SYSTEM_PROMPT},
            {'role': 'user', 'content': user_msg},
        ], llm_cfg)
        content = parse_content_json(text)
        if content:
            missing = _missing_chapters(outline, content)
            if not missing:
                return content
            if attempt < retries - 1:
                names = '、'.join(o['heading'] for o in missing)
                user_msg += (f'\n\n注意：你漏掉了以下章节，必须全部补齐且标题原文照抄大纲：{names}')
                continue
            print(f'[警告] AI 仍漏掉章节：{"、".join(o["heading"] for o in missing)}，已按现有内容继续')
            return content
        user_msg += '\n\n注意：你上一次的输出不是合法 JSON，请只输出一个 JSON 对象。'
    raise RuntimeError('AI 连续返回无法解析的内容，已放弃')


# --------------------------------------------------------------------------- #
# Render content JSON -> .docx (source for the conversion pipeline)
# --------------------------------------------------------------------------- #
def _render_block(doc, b):
    if not isinstance(b, dict):
        return
    kind = b.get('type')
    if kind == 'para':
        text = str(b.get('text') or '').strip()
        if text:
            doc.add_paragraph(text)
    elif kind == 'list':
        for it in b.get('items') or []:
            doc.add_paragraph(str(it).strip())
    elif kind == 'table':
        headers = [str(x) for x in (b.get('headers') or [])]
        rows = b.get('rows') or []
        ncol = 1
        if headers:
            ncol = max(ncol, len(headers))
        for r in rows:
            ncol = max(ncol, len(r))
        if b.get('caption'):
            doc.add_paragraph(str(b['caption']))
        tbl = doc.add_table(rows=0, cols=ncol)
        tbl.style = 'Table Grid'
        if headers:
            row = tbl.add_row()
            for i, h in enumerate(headers[:ncol]):
                cell = row.cells[i]
                cell.text = str(h)
                for p in cell.paragraphs:
                    for r in p.runs:
                        r.bold = True
        for r in rows:
            row = tbl.add_row()
            for i in range(ncol):
                row.cells[i].text = str(r[i]) if i < len(r) else ''
    elif kind == 'placeholder':
        p = doc.add_paragraph()
        run = p.add_run(str(b.get('text') or ''))
        run.bold = True
        run.font.color.rgb = RGBColor(0xC0, 0x00, 0x00)


def render_source(content, out_path):
    doc = Document()
    title = str(content.get('title') or '未命名文档').strip()
    doc.add_heading(title, level=0)  # Title style -> picked up by _extract_title

    counters = [0, 0, 0, 0, 0]

    def number_for(level):
        i = level - 1
        counters[i] += 1
        for j in range(i + 1, len(counters)):
            counters[j] = 0
        return '.'.join(str(c) for c in counters[:level])

    for ch in content.get('chapters') or []:
        try:
            lvl = int(ch.get('level', 1))
        except (TypeError, ValueError):
            lvl = 1
        lvl = max(1, min(5, lvl))
        heading = str(ch.get('heading') or '').strip() or '未命名章节'
        doc.add_heading(f'{number_for(lvl)} {heading}', level=lvl)
        for b in ch.get('blocks') or []:
            _render_block(doc, b)

    out_dir = os.path.dirname(out_path)
    if out_dir:
        os.makedirs(out_dir, exist_ok=True)
    doc.save(out_path)
    return out_path


def sanitize_filename(name):
    name = re.sub(r'[\\/:*?"<>|\r\n\t]', '_', str(name)).strip()
    return name or 'untitled'


# --------------------------------------------------------------------------- #
# Interactive flow
# --------------------------------------------------------------------------- #
def ask_template(configs):
    print('可选模板：')
    for i, c in enumerate(configs):
        print(f'  [{i + 1}] {c.name}')
    while True:
        s = input(f'选择模板编号 (1-{len(configs)})：').strip()
        if s.isdigit() and 1 <= int(s) <= len(configs):
            return configs[int(s) - 1]
        print('输入无效，请重新选择。')


def interactive_generate(templates_dir):
    """Ask the user what to write; returns (tpl_cfg, content, outline)."""
    from .config import load_configs
    configs = load_configs(templates_dir)
    if not configs:
        raise RuntimeError('模板文件夹里没有 .docx 文件')

    cfg = ask_template(configs)
    template_doc = Document(cfg.file)
    outline = template_outline(template_doc, cfg)
    if not outline:
        raise RuntimeError(f'模板「{cfg.name}」里没有可用的章节大纲')

    print(f'\n模板「{cfg.name}」的章节大纲：')
    for o in outline:
        print('  ' * (o['level'] - 1) + f'└ {o["heading"]}')

    print('\n接下来回答几个问题（直接回车可跳过）：')
    title = input('  文档完整标题：').strip()
    if not title:
        title = f'{cfg.name}草稿'
    system = input('  系统/项目名称（例如 量子控制系统V2.0）：').strip()
    notes = input('  素材要点（粘贴你的素材，可空回车 = 让 AI 起草占位内容）：').strip()

    info = {'title': title, 'system': system, 'notes': notes}
    return cfg, info, outline


def run_generate(args):
    """Full --generate flow: ask -> draft -> render into input -> convert."""
    from .config import load_configs
    from .mapper import convert as do_convert

    llm_cfg = load_llm_config(getattr(args, 'llm', None))
    if not llm_cfg.api_key:
        print('错误：未找到 API key。请任选一种方式配置：\n'
              '  1) 设置环境变量 DEEPSEEK_API_KEY=sk-xxx\n'
              '  2) 在当前目录放 llm.yaml（含 api_key 字段），本程序会自动读取\n'
              '  3) 用 --llm 指定配置文件路径')
        return 1

    cfg, info, outline = interactive_generate(args.templates)
    print('\n正在调用 AI 起草内容（可能需要一两分钟）……')
    content = generate_content(cfg, outline, info, llm_cfg)
    if not content.get('title'):
        content['title'] = info['title']

    gen_dir = args.input or 'input'
    src_path = os.path.join(gen_dir, sanitize_filename(content['title']) + '.docx')
    render_source(content, src_path)
    print('已生成源文档:', src_path)

    # ---- reuse the normal single-file conversion (template forced, no re-match)
    configs = load_configs(args.templates, args.config)
    out_dir = args.output
    os.makedirs(out_dir, exist_ok=True)
    report_path = args.report or os.path.join(out_dir, 'report.md')

    from .history import RunLog
    log = RunLog(args.db)
    log.begin(input_=gen_dir, templates=args.templates, output=out_dir,
              report=report_path, args=['--generate'])

    stem = os.path.splitext(os.path.basename(src_path))[0]
    out_path = os.path.join(out_dir, f'{stem}_{cfg.name}.docx')
    try:
        ctx = do_convert(src_path, cfg, out_path)
        result = FileResult(input=src_path, output=out_path, status='ok',
                            template=cfg.name, score=1.0, stats=ctx.stats,
                            warnings=ctx.warnings,
                            detail=json.dumps(ctx.toc_entries, ensure_ascii=False))
    except Exception as e:
        result = FileResult(input=src_path, output=out_path, status='error',
                            template=cfg.name, error=repr(e))

    log.log_file(result)
    log.finish(1, 1 if result.status == 'ok' else 0, 0 if result.status == 'ok' else 1)
    log.close()

    write_report([result], report_path)

    if result.status == 'ok':
        print('转换完成:', out_path)
        print('处理报告:', report_path)
        return 0
    print('转换失败:', result.error)
    return 1
