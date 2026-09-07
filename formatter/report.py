# -*- coding: utf-8 -*-
"""Processing report."""
import os
from dataclasses import dataclass, field


@dataclass
class FileResult:
    input: str
    output: str = ''
    status: str = 'ok'
    template: str = ''
    score: float = 0.0
    stats: object = None
    warnings: list = field(default_factory=list)
    detail: str = ''
    error: str = ''


def write_report(results, path, also_print=True):
    ok = sum(1 for r in results if r.status == 'ok')
    lines = [
        '# 处理报告', '',
        f'- 总数：{len(results)}　成功：{ok}　失败：{len(results) - ok}', '',
        '| 源文件 | 状态 | 匹配模板 | 相似度 | 标题 | 段落 | 表格 | 图片 | 备注 |',
        '|---|---|---|---|---|---|---|---|---|',
    ]
    for r in results:
        st = r.stats
        hd = st.headings if st else '-'
        pa = st.paragraphs if st else '-'
        tb = st.tables if st else '-'
        im = st.images if st else '-'
        note = '; '.join(r.warnings) or (r.error or '')
        note = note.replace('|', '\\|').replace('\n', ' ')
        lines.append(
            f'| {os.path.basename(r.input)} | {r.status} | {r.template or "-"} | '
            f'{r.score:.2f} | {hd} | {pa} | {tb} | {im} | {note} |'
        )
    lines.append('')
    text = '\n'.join(lines)
    os.makedirs(os.path.dirname(path) or '.', exist_ok=True)
    with open(path, 'w', encoding='utf-8') as f:
        f.write(text)
    if also_print:
        print(text)
    return path
