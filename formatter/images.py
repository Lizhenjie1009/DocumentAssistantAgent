# -*- coding: utf-8 -*-
"""Image extraction from source runs and re-insertion into the new document."""
from io import BytesIO

from docx.oxml.ns import qn
from docx.shared import Emu

from .util import qn as _qn


def _blobs_in(run, source_doc):
    """Yield (blob, ext) for every image referenced by the run (inline or anchored)."""
    for blip in run._r.iter(qn('a:blip')):
        embed = blip.get(qn('r:embed'))
        if not embed:
            continue
        try:
            part = source_doc.part.related_parts[embed]
        except KeyError:
            continue
        blob = getattr(part, 'blob', None)
        if blob is None:
            continue
        ext = (part.partname.ext or '').lstrip('.') or 'png'
        yield blob, ext


def _display_size(run):
    """Return (cx_emu, cy_emu) from the first inline/anchor extent, or (None, None)."""
    for ext in run._r.iter(qn('wp:extent')):
        try:
            return int(ext.get('cx')), int(ext.get('cy'))
        except (TypeError, ValueError):
            pass
    return None, None


def _intrinsic_width_emu(blob):
    """Fallback: intrinsic pixel width converted to EMU at 96 DPI."""
    try:
        from PIL import Image
        with Image.open(BytesIO(blob)) as im:
            px = im.width
        return int(px * 914400 / 96)
    except Exception:
        return None


def copy_images_into(src_run, dst_paragraph, source_doc, cfg, usable_width_emu):
    """Re-insert all images of `src_run` into `dst_paragraph`, scaled per config.

    Returns the number of images inserted.
    """
    max_w = int(usable_width_emu * cfg.image_max_width_ratio)
    n = 0
    for blob, ext in _blobs_in(src_run, source_doc):
        cx, cy = _display_size(src_run)
        if cx and cx > 0:
            target_w = cx if (cx <= max_w and not cfg.image_enlarge) else max_w
        else:
            intrinsic = _intrinsic_width_emu(blob)
            if intrinsic and intrinsic <= max_w and not cfg.image_enlarge:
                target_w = intrinsic
            else:
                target_w = max_w
        new_run = dst_paragraph.add_run()
        try:
            new_run.add_picture(BytesIO(blob), width=Emu(target_w))
            n += 1
        except Exception:
            pass
    return n
