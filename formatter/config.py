# -*- coding: utf-8 -*-
"""Template discovery and per-template configuration."""
import os
from dataclasses import dataclass, field


DEFAULTS = {
    'numbering': 'auto',            # 'auto' = strip source numbers, use template auto-numbering
    'image': {
        'max_width_ratio': 1.0,     # max width = usable text width * ratio
        'enlarge': False,           # never enlarge small images
        'center': True,             # center image-only paragraphs
    },
    'toc': {
        'levels': [1, 2],           # which heading levels go into the TOC
        'placeholder': '00',        # placeholder page number
        'tab_leader': True,         # dotted leader before page number
    },
    'table_style': None,            # force a table style name (None = copy source formatting)
    'body_start_style': 'Heading 1',
    'title': {
        'placeholder_regex': r'[X×Ｘx]{1,3}分?系统',  # text to replace on cover/title with source title
        'source': 'auto',            # 'auto' | 'filename' | None (disabled)
    },
    'match': {
        'keywords': [],
        'min_similarity': 0.35,
    },
}


@dataclass
class TemplateConfig:
    name: str
    file: str                       # absolute path to the .docx
    numbering: str = 'auto'
    image_max_width_ratio: float = 1.0
    image_enlarge: bool = False
    image_center: bool = True
    toc_levels: list = field(default_factory=lambda: [1, 2])
    toc_placeholder: str = '00'
    toc_tab_leader: bool = True
    table_style: str = None
    body_start_style: str = 'Heading 1'
    title_placeholder_regex: str = r'X{1,3}分?系统'
    title_source: str = 'auto'
    match_keywords: list = field(default_factory=list)
    min_similarity: float = 0.35


def discover_templates(templates_dir):
    """Return sorted list of absolute paths to .docx files in a directory."""
    out = []
    for name in sorted(os.listdir(templates_dir)):
        if name.lower().endswith('.docx') and not name.startswith('~$'):
            out.append(os.path.join(templates_dir, name))
    return out


def _merge(defaults, overrides):
    """Shallow merge for our two-level dict structure."""
    d = dict(defaults)
    for k, v in (overrides or {}).items():
        if isinstance(v, dict) and isinstance(d.get(k), dict):
            d[k] = _merge(d[k], v)
        else:
            d[k] = v
    return d


def load_configs(templates_dir, config_path=None):
    """Build TemplateConfig objects for every template in `templates_dir`.

    config_path (optional YAML) may contain a ``defaults`` block and a
    ``templates`` block keyed by template filename.
    """
    import yaml

    cfg_data = {}
    if config_path and os.path.exists(config_path):
        with open(config_path, 'r', encoding='utf-8') as f:
            cfg_data = yaml.safe_load(f) or {}

    global_defaults = _merge(DEFAULTS, cfg_data.get('defaults', {}))
    per_template = cfg_data.get('templates', {}) or {}

    configs = []
    for path in discover_templates(templates_dir):
        filename = os.path.basename(path)
        overrides = per_template.get(filename, {})
        merged = _merge(global_defaults, overrides)

        img = merged.get('image', {})
        toc = merged.get('toc', {})
        title = merged.get('title', {})
        match = merged.get('match', {})

        configs.append(TemplateConfig(
            name=merged.get('name') or os.path.splitext(filename)[0],
            file=path,
            numbering=merged.get('numbering', 'auto'),
            image_max_width_ratio=float(img.get('max_width_ratio', 1.0)),
            image_enlarge=bool(img.get('enlarge', False)),
            image_center=bool(img.get('center', True)),
            toc_levels=list(toc.get('levels', [1, 2])),
            toc_placeholder=str(toc.get('placeholder', '00')),
            toc_tab_leader=bool(toc.get('tab_leader', True)),
            table_style=merged.get('table_style'),
            body_start_style=merged.get('body_start_style', 'Heading 1'),
            title_placeholder_regex=title.get('placeholder_regex', r'[X×Ｘx]{1,3}分?系统'),
            title_source=title.get('source', 'auto'),
            match_keywords=list(match.get('keywords', [])),
            min_similarity=float(match.get('min_similarity', 0.35)),
        ))
    return configs
