# -*- coding: utf-8 -*-
"""Template matching by chapter (heading) fingerprint."""
import os
import re
from difflib import SequenceMatcher

from .util import detect_heading_level, normalize_text, find_style


def collect_headings(doc, from_body=False, body_start_style='Heading 1'):
    """Return [(level, normalized_text), ...] for headings (deduplicated).

    When from_body=True, only headings after the first paragraph whose style is
    `body_start_style` are collected (i.e. skip front matter). Duplicates
    (same level + text) are dropped so a hand-typed TOC does not double count.
    """
    items = []
    started = not from_body
    for p in doc.paragraphs:
        if not started:
            name = (p.style.name or '') if p.style is not None else ''
            if name == body_start_style or detect_heading_level(p) == 1:
                started = True
        if started:
            lvl = detect_heading_level(p)
            if lvl is not None and 1 <= lvl <= 5:
                items.append((lvl, normalize_text(p.text)))
    seen = set()
    out = []
    for item in items:
        if item not in seen:
            seen.add(item)
            out.append(item)
    return out


def _implicit_keywords(template_file):
    """Derive candidate keywords from a template filename, e.g.
    '1-【模板】XX分系统关键技术攻关报告.docx' -> ['XX分系统关键技术攻关报告', '关键技术攻关报告'].
    """
    stem = os.path.splitext(os.path.basename(template_file))[0]
    stem = re.sub(r'[【】]', '', stem)
    stem = stem.replace('模板', '')
    stem = re.sub(r'^[\d\-\.\s]*', '', stem)
    kws = [stem] if stem else []
    m = re.match(r'^X{1,3}分?系统(.*)$', stem)
    if m and m.group(1):
        kws.append(m.group(1))
    return [k for k in kws if k]


def _source_full_text(source_doc):
    return ' '.join((p.text or '') for p in source_doc.paragraphs)


def _score(src_h1, tpl_h1, src_h2, tpl_h2):
    s1 = SequenceMatcher(None, ' '.join(src_h1), ' '.join(tpl_h1)).ratio()
    s2 = SequenceMatcher(None, ' '.join(src_h2), ' '.join(tpl_h2)).ratio() if (src_h2 and tpl_h2) else s1
    return 0.7 * s1 + 0.3 * s2


def match_template(source_doc, configs, source_path=''):
    """Return (best_config, best_score) or (None, 0.0).

    `source_path` (optional) enables a filename-based boost: template keywords
    found in the source file name add a small bonus (safety net when the
    document text alone is ambiguous).
    """
    src_items = collect_headings(source_doc, from_body=False)
    src_h1 = [t for l, t in src_items if l == 1]
    src_h2 = [t for l, t in src_items if l == 2]
    full_text = _source_full_text(source_doc)
    src_stem = os.path.splitext(os.path.basename(source_path))[0] if source_path else ''

    best_cfg, best_score = None, 0.0
    for cfg in configs:
        try:
            from docx import Document
            tpl = Document(cfg.file)
        except Exception:
            continue
        tpl_items = collect_headings(tpl, from_body=True, body_start_style=cfg.body_start_style)
        tpl_h1 = [t for l, t in tpl_items if l == 1]
        tpl_h2 = [t for l, t in tpl_items if l == 2]

        if not tpl_h1:
            continue

        score = _score(src_h1, tpl_h1, src_h2, tpl_h2)

        # keyword boost: configured + derived from template filename, matched
        # against the source document text
        keywords = list(cfg.match_keywords) + _implicit_keywords(cfg.file)
        if keywords:
            hits = sum(1 for kw in keywords if kw and kw in full_text)
            if hits:
                score = min(1.0, score + 0.15 * hits)

        # filename boost: template keywords found in the source file name
        if src_stem:
            for kw in _implicit_keywords(cfg.file):
                if kw and kw in src_stem:
                    score = min(1.0, score + 0.15)
                    break

        if score > best_score:
            best_score, best_cfg = score, cfg

    if best_cfg is not None and best_score < best_cfg.min_similarity:
        return None, best_score
    return best_cfg, best_score
