# -*- coding: utf-8 -*-
"""Shared helpers for the docx formatter."""
import re
from docx.oxml.ns import qn as _qn


def qn(tag):
    return _qn(tag)


# --------------------------------------------------------------------------- #
# Styles
# --------------------------------------------------------------------------- #
def find_style(doc, *names):
    """Return the style matching `names`, honoring priority order (not doc order).

    For each name in `names`, the first style with that name (case-insensitive)
    is returned; the first name that exists wins.
    """
    by_name = {}
    for st in doc.styles:
        key = (st.name or '').lower()
        if key and key not in by_name:
            by_name[key] = st
    for n in names:
        st = by_name.get(n.lower())
        if st is not None:
            return st
    return None


def style_outline_level(style):
    """0-based outline level of a style, or None."""
    if style is None:
        return None
    ppr = style.element.find(qn('w:pPr'))
    if ppr is None:
        return None
    ol = ppr.find(qn('w:outlineLvl'))
    if ol is None:
        return None
    try:
        return int(ol.get(qn('w:val')))
    except (TypeError, ValueError):
        return None


# --------------------------------------------------------------------------- #
# Text / numbering
# --------------------------------------------------------------------------- #
_NUM_STRIP = [
    re.compile(r'^\d+(?:\.\d+)*[\s、．.．、]*'),                      # 1 / 1.1 / 1.1.1
    re.compile(r'^[一二三四五六七八九十百千]+[、.．\s]*'),            # 一、
    re.compile(r'^第[一二三四五六七八九十百千\d]+[章节部分条款][\s、.．]*'),
    re.compile(r'^[\(（【\[]?\d+[\)）】\]]?[\s、.．]*'),              # (1) / 1)
]


def strip_number(text):
    """Strip a leading manual number from heading text."""
    s = (text or '').strip()
    for pat in _NUM_STRIP:
        s2 = pat.sub('', s, count=1)
        if s2 != s:
            return s2.strip()
    return s


def normalize_text(text):
    t = strip_number(text)
    t = re.sub(r'[\s\u3000]+', '', t)
    return t.lower()


# --------------------------------------------------------------------------- #
# Heading detection
# --------------------------------------------------------------------------- #
def detect_heading_level(p):
    """Return 1-based heading level, or None if the paragraph is not a heading."""
    st = (p.style.name or '') if p.style is not None else ''

    # 'Title' is a document title, not a numbered heading (even if it has outlineLvl)
    if st.lower() == 'title':
        return None

    m = re.search(r'(?:heading|标题)\s*([1-9])', st, re.I)
    if m:
        return int(m.group(1))

    lvl = style_outline_level(p.style)
    if lvl is not None and 0 <= lvl <= 8:
        return lvl + 1

    txt = (p.text or '').strip()
    if txt:
        # manual "1.1.1 xxx" (only for reasonably short paragraphs = headings)
        m = re.match(r'^(\d+(?:\.\d+)+)(?!\d)[\s、.．]*\S', txt)
        if m and len(txt) <= 60:
            return m.group(1).count('.') + 1
        # manual single-number heading: "1 范围", "2、目的", "3. 概述"
        m = re.match(r'^(\d+)[\s、.．]+(\S.*)$', txt)
        if m and len(txt) <= 60:
            rest = m.group(2).strip()
            if rest and rest[-1] not in '。？！；，':
                return 1
        if re.match(r'^(?:第[一二三四五六七八九十百千\d]+[章节部分]|[一二三四五六七八九十]+[、.．])', txt) and len(txt) <= 60:
            return 1
    return None


def has_numbering(p):
    """True if the paragraph carries Word auto-numbering (w:numPr)."""
    ppr = p._p.find(qn('w:pPr'))
    if ppr is None:
        return False
    return ppr.find(qn('w:numPr')) is not None


def is_section_break(p):
    """True if the paragraph ends a section (contains a sectPr in its pPr)."""
    ppr = p._p.find(qn('w:pPr'))
    if ppr is None:
        return False
    return ppr.find(qn('w:sectPr')) is not None


# --------------------------------------------------------------------------- #
# Run formatting
# --------------------------------------------------------------------------- #
def copy_run_format(src_run, dst_run):
    dst_run.bold = src_run.bold
    dst_run.italic = src_run.italic
    dst_run.underline = src_run.underline
    try:
        if src_run.font.superscript:
            dst_run.font.superscript = True
        if src_run.font.subscript:
            dst_run.font.subscript = True
    except Exception:
        pass


# --------------------------------------------------------------------------- #
# Tab stops (for dotted TOC leaders)
# --------------------------------------------------------------------------- #
def add_right_tab(paragraph, position_emu, leader='dot'):
    """Add a right-aligned tab stop with a dot leader to a paragraph."""
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn as _q

    pPr = paragraph._p.get_or_add_pPr()
    tabs = pPr.find(_q('w:tabs'))
    if tabs is None:
        tabs = OxmlElement('w:tabs')
        pPr.append(tabs)
    tab = OxmlElement('w:tab')
    tab.set(_q('w:val'), 'right')
    tab.set(_q('w:pos'), str(int(position_emu)))
    if leader in ('dot', 'dots'):
        tab.set(_q('w:leader'), 'dot')
    tabs.append(tab)
