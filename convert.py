# -*- coding: utf-8 -*-
"""Command-line entry: convert source docx files to match templates."""
import argparse
import json
import os
import sys

from docx import Document

from formatter.config import load_configs
from formatter.history import RunLog, show_history
from formatter.mapper import convert
from formatter.matcher import collect_headings, match_template
from formatter.report import FileResult, write_report


def _fix_stdout():
    try:
        sys.stdout.reconfigure(encoding='utf-8')
    except Exception:
        pass


def _collect_inputs(path):
    if os.path.isdir(path):
        return [
            os.path.join(path, n)
            for n in sorted(os.listdir(path))
            if n.lower().endswith('.docx') and not n.startswith('~$')
        ]
    return [path]


def process_one(path, configs, out_dir):
    stem = os.path.splitext(os.path.basename(path))[0]
    try:
        source = Document(path)
    except Exception as e:
        return FileResult(input=path, status='error', error=str(e))

    if not collect_headings(source):
        return FileResult(input=path, status='no_headings',
                          warnings=['源文档未检测到任何标题'])

    cfg, score = match_template(source, configs, source_path=path)
    if cfg is None:
        return FileResult(input=path, status='no_match', score=score,
                          warnings=['未匹配到模板（最高相似度 %.2f）' % score])

    out_path = os.path.join(out_dir, f'{stem}_{cfg.name}.docx')
    detail = ''
    try:
        ctx = convert(path, cfg, out_path)
        detail = json.dumps(ctx.toc_entries, ensure_ascii=False)
    except PermissionError:
        return FileResult(input=path, output=out_path, status='error', template=cfg.name,
                          error='输出文件被占用（可能正被 Word/WPS 打开），请关闭后重试')
    except Exception as e:
        return FileResult(input=path, output=out_path, status='error',
                          template=cfg.name, error=repr(e))
    return FileResult(input=path, output=out_path, status='ok',
                      template=cfg.name, score=score, stats=ctx.stats,
                      warnings=ctx.warnings, detail=detail)


def main(argv=None):
    ap = argparse.ArgumentParser(description='把源 Word 文档套用为模板格式')
    ap.add_argument('--input', '-i', default=None, help='源 .docx 文件或文件夹')
    ap.add_argument('--templates', '-t', default=None, help='模板文件夹')
    ap.add_argument('--output', '-o', default='output', help='输出文件夹（默认 output）')
    ap.add_argument('--config', '-c', default=None, help='可选 config.yaml')
    ap.add_argument('--report', '-r', default=None, help='报告路径（默认 <output>/report.md）')
    ap.add_argument('--db', default='run_history.db', help='SQLite 运行历史库（默认 run_history.db）')
    ap.add_argument('--history', nargs='?', const=10, type=int, metavar='N',
                    help='查看最近 N 次运行记录后退出（默认 10）')
    ap.add_argument('--generate', action='store_true',
                    help='交互生成模式：问答 → AI 按所选模板大纲起草 → 生成 input 文档 → 自动转换')
    ap.add_argument('--llm', default=None, help='LLM 配置文件（可选；api_key 默认读 DEEPSEEK_API_KEY）')
    args = ap.parse_args(argv)

    _fix_stdout()

    if args.history is not None:
        show_history(args.db, limit=args.history)
        return 0

    if args.generate:
        if not args.templates:
            ap.error('--generate 需要提供 --templates/-t 模板文件夹')
        from formatter.generate import run_generate
        return run_generate(args)

    if not args.input or not args.templates:
        ap.error('--input/-i 和 --templates/-t 为必填参数（除非使用 --history 或 --generate）')

    configs = load_configs(args.templates, args.config)
    if not configs:
        print('错误：模板文件夹里没有 .docx 文件')
        return 1

    inputs = _collect_inputs(args.input)
    if not inputs:
        print('错误：未找到源 .docx 文件')
        return 1

    os.makedirs(args.output, exist_ok=True)
    report_path = args.report or os.path.join(args.output, 'report.md')

    log = RunLog(args.db)
    log.begin(input_=args.input, templates=args.templates, output=args.output,
              report=report_path, args=(argv if argv is not None else sys.argv[1:]))

    results = []
    for path in inputs:
        print('处理中:', os.path.basename(path))
        result = process_one(path, configs, args.output)
        results.append(result)
        log.log_file(result)

    ok = sum(1 for r in results if r.status == 'ok')
    log.finish(total=len(results), ok=ok, failed=len(results) - ok)
    log.close()

    write_report(results, report_path)
    print('报告已生成:', report_path)
    print('运行历史已记录到:', args.db)
    return 0


if __name__ == '__main__':
    sys.exit(main())
