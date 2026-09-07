# -*- coding: utf-8 -*-
"""Run-level content copying (text + images), shared by body and table mapping."""
from .util import copy_run_format
from .images import copy_images_into


def copy_runs(src_p, dst_p, source_doc, cfg, usable_width_emu, copy_format=True):
    """Copy every run of `src_p` into `dst_p` (text with inline format, images scaled).

    Returns True if at least one image was inserted.
    """
    has_image = False
    for run in src_p.runs:
        n_img = copy_images_into(run, dst_p, source_doc, cfg, usable_width_emu)
        if n_img:
            has_image = True
        text = run.text
        if text:
            nr = dst_p.add_run(text)
            if copy_format:
                copy_run_format(run, nr)
    return has_image
