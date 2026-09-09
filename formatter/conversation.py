# -*- coding: utf-8 -*-
"""Conversational document drafting (GitHub issue #1).

REPL session: pick a template -> requirements -> AI draft -> keep chatting to
revise chapters -> export docx (through the normal conversion pipeline) or
Markdown. The whole conversation is persisted to SQLite.
"""
import json
import os
import sqlite3
from datetime import datetime

from docx import Document

from .config import load_configs
from .generate import (load_llm_config, template_outline, generate_content,
                       revise_content, render_source, sanitize_filename,
                       ask_template)
from .history import RunLog
from .mapper import convert as do_convert
from .report import FileResult

HELP = """输入说明：
  - 直接输入文字  = 作为“修改意见”发给 AI 修订文档（例如：第2章补一张表格 / 测试结果写详细些）
  - /目录         查看当前章节
  - /导出docx     渲染并转成模板格式的 Word（写入 output）
  - /导出md       导出 Markdown（写入 output）
  - /重来         放弃当前内容，按原需求重新起草
  - /退出         结束会话（对话自动保存）
"""


# --------------------------------------------------------------------------- #
# conversation persistence (SQLite)
# --------------------------------------------------------------------------- #
def _now():
    return datetime.now().strftime('%Y-%m-%d %H:%M:%S')


def _open_conversation_db(db_path):
    conn = sqlite3.connect(db_path)
    conn.executescript('''
    CREATE TABLE IF NOT EXISTS chat_sessions (
        id         INTEGER PRIMARY KEY AUTOINCREMENT,
        started_at TEXT NOT NULL,
        template   TEXT,
        title      TEXT,
        outline    TEXT
    );
    CREATE TABLE IF NOT EXISTS chat_messages (
        id         INTEGER PRIMARY KEY AUTOINCREMENT,
        session_id INTEGER NOT NULL,
        role       TEXT NOT NULL,
        content    TEXT,
        created_at TEXT
    );
    ''')
    return conn


# --------------------------------------------------------------------------- #
# helpers
# --------------------------------------------------------------------------- #
def _numbered_heading(counters, level):
    i = level - 1
    counters[i] += 1
    for j in range(i + 1, len(counters)):
        counters[j] = 0
    return '.'.join(str(c) for c in counters[:level])


def _md_cell(v):
    return str(v).replace('|', '\\|').replace('\n', ' ')


def content_to_markdown(content):
    """Render the content JSON as Markdown text."""
    lines = ['# ' + (content.get('title') or '未命名文档').strip(), '']
    counters = [0] * 5
    for ch in content.get('chapters') or []:
        try:
            lvl = int(ch.get('level', 1))
        except (TypeError, ValueError):
            lvl = 1
        lvl = max(1, min(5, lvl))
        heading = str(ch.get('heading') or '').strip() or '未命名章节'
        lines.append('#' * (lvl + 1) + ' ' + _numbered_heading(counters, lvl) + ' ' + heading)
        lines.append('')
        for b in ch.get('blocks') or []:
            if not isinstance(b, dict):
                continue
            kind = b.get('type')
            if kind == 'para':
                lines.append(str(b.get('text') or '').strip())
                lines.append('')
            elif kind == 'list':
                for it in b.get('items') or []:
                    lines.append('- ' + str(it).strip())
                lines.append('')
            elif kind == 'table':
                if b.get('caption'):
                    lines.append('**' + str(b['caption']) + '**')
                headers = [str(x) for x in (b.get('headers') or [])]
                rows = b.get('rows') or []
                ncol = 1
                for h in headers:
                    ncol = max(ncol, len([h]))
                for r in rows:
                    ncol = max(ncol, len(r))
                if headers:
                    lines.append('| ' + ' | '.join(_md_cell(h) for h in headers[:ncol]) + ' |')
                    lines.append('|' + '---|' * ncol)
                for r in rows:
                    lines.append('| ' + ' | '.join(_md_cell(r[i]) if i < len(r) else ''
                                  for i in range(ncol)) + ' |')
                lines.append('')
            elif kind == 'placeholder':
                lines.append('> ' + str(b.get('text') or ''))
                lines.append('')
    return '\n'.join(lines)


def _describe(content):
    """One-line summary of the current draft."""
    n = len(content.get('chapters') or [])
    blocks = sum(len(ch.get('blocks') or []) for ch in (content.get('chapters') or []))
    return f'共 {n} 章、{blocks} 个内容块'


# --------------------------------------------------------------------------- #
# the chat session
# --------------------------------------------------------------------------- #
def run_chat(args):
    llm_cfg = load_llm_config(getattr(args, 'llm', None))
    if not llm_cfg.api_key:
        print('错误：未找到 API key。请设置环境变量 DEEPSEEK_API_KEY，或放一个含 api_key 的 llm.yaml / 用 --llm 指定。')
        return 1

    configs = load_configs(args.templates, getattr(args, 'config', None))
    cfg = ask_template(configs)
    template_doc = Document(cfg.file)
    outline = template_outline(template_doc, cfg)
    if not outline:
        print(f'错误：模板「{cfg.name}」没有可用章节大纲')
        return 1

    print(f'\n模板「{cfg.name}」共 {len(outline)} 个章节标题，开始收集需求：')
    title = input('  文档完整标题：').strip() or f'{cfg.name}草稿'
    system = input('  系统/项目名称（可空）：').strip()
    notes = input('  素材要点（可空，留空由 AI 起草占位内容）：').strip()
    info = {'title': title, 'system': system, 'notes': notes}

    conn = _open_conversation_db(args.db)
    cur = conn.execute(
        'INSERT INTO chat_sessions (started_at, template, title, outline) VALUES (?,?,?,?)',
        (_now(), cfg.name, title, json.dumps(outline, ensure_ascii=False)))
    session_id = cur.lastrowid
    conn.commit()

    def save_msg(role, text):
        conn.execute('INSERT INTO chat_messages (session_id, role, content, created_at) VALUES (?,?,?,?)',
                     (session_id, role, text, _now()))
        conn.commit()

    save_msg('user', f'选择模板={cfg.name}; 标题={title}; 系统={system or "-"}; 素材={notes or "-"}')

    gen_dir = getattr(args, 'input', None) or 'input'
    out_dir = args.output
    os.makedirs(gen_dir, exist_ok=True)
    os.makedirs(out_dir, exist_ok=True)

    print('\n正在起草初稿（约 1-2 分钟）……')
    content = generate_content(cfg, outline, info, llm_cfg)
    content.setdefault('title', title)
    save_msg('assistant', '[初稿] ' + _describe(content))
    print('初稿完成：' + _describe(content))
    print('\n' + HELP)

    last_export_docx = ''
    try:
        while True:
            try:
                user_in = input('你 > ').strip()
            except EOFError:
                break
            if not user_in:
                continue
            low = user_in.lower()
            if low in ('/退出', '退出', 'exit', 'quit'):
                break
            if low.startswith('/') or low in ('目录', '导出', '导出docx', '导出md', '重来'):
                cmd = low[1:] if low.startswith('/') else low
                if cmd in ('目录',):
                    print('当前章节：')
                    counters = [0] * 5
                    for ch in content.get('chapters') or []:
                        lvl = max(1, min(5, int(ch.get('level', 1))))
                        print('  ' + '  ' * (lvl - 1) + f'{_numbered_heading(counters, lvl)} {ch.get("heading")}')
                    continue
                if cmd in ('导出docx', '导出'):
                    fname = sanitize_filename(content.get('title') or title) + '.docx'
                    src = os.path.join(gen_dir, fname)
                    render_source(content, src)
                    stem = os.path.splitext(fname)[0]
                    out_path = os.path.join(out_dir, f'{stem}_{cfg.name}.docx')
                    ctx = do_convert(src, cfg, out_path)
                    last_export_docx = out_path
                    print(f'已导出 Word（模板格式）: {out_path}')
                    print(f'  源草稿: {src}（{_describe(content)}）')
                    # record in run history
                    log = RunLog(args.db)
                    log.begin(input_=gen_dir, templates=args.templates, output=out_dir,
                              report=os.path.join(out_dir, 'report.md'), args=['--chat', cmd])
                    result = FileResult(input=src, output=out_path, status='ok', template=cfg.name,
                                        score=1.0, stats=ctx.stats, warnings=ctx.warnings,
                                        detail=json.dumps(ctx.toc_entries, ensure_ascii=False))
                    log.log_file(result)
                    log.finish(1, 1, 0)
                    log.close()
                    continue
                if cmd in ('导出md',):
                    md_path = os.path.join(out_dir, sanitize_filename(content.get('title') or title) + '.md')
                    with open(md_path, 'w', encoding='utf-8') as f:
                        f.write(content_to_markdown(content))
                    print('已导出 Markdown:', md_path)
                    continue
                if cmd in ('重来',):
                    print('正在按原需求重新起草……')
                    content = generate_content(cfg, outline, info, llm_cfg)
                    content.setdefault('title', title)
                    save_msg('assistant', '[重来] ' + _describe(content))
                    print('重写完成：' + _describe(content))
                    continue
                # unknown '/cmd'
                print('未知指令，试试：/目录 /导出docx /导出md /重来 /退出（或直接输入修改意见）')
                continue

            # treat as a revision instruction
            save_msg('user', user_in)
            print('正在按你的意见修订……')
            content = revise_content(cfg, outline, content, user_in, llm_cfg)
            content.setdefault('title', title)
            save_msg('assistant', '[修订] ' + _describe(content))
            print('修订完成：' + _describe(content))
    except KeyboardInterrupt:
        print('\n已中断')

    conn.close()
    print('\n会话结束（已保存到 ' + args.db + ' 的 chat_sessions/chat_messages 表）。')
    if last_export_docx:
        print('最近一次导出:', last_export_docx)
    return 0
