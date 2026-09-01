# -*- coding: utf-8 -*-
"""Helpers for editing LaTeX source .docx before MathType conversion.

Use when an agent rewrites manuscript paragraphs: clear stale OLE/OMath,
restore body/display typography, then run convert_docx_latex_to_equations.py.

Journal defaults below match typical Chinese IMRaD Word templates (ËÎÌå/TNR ÎåºÅ).
Override constants for other venues.
"""
from __future__ import annotations

import json
from pathlib import Path
from typing import Iterable

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.oxml.ns import qn
from docx.shared import Cm, Pt

SONG = "\u5b8b\u4f53"
TIMES = "Times New Roman"
SZ_BODY = 10.5
LINE_BODY = 15.6
FIRST_INDENT_CM = 0.74

EMBED_TAGS = (
    qn("w:object"),
    qn("m:oMath"),
    qn("m:oMathPara"),
    qn("w:drawing"),
    qn("w:pict"),
)


def set_run_fonts(run, *, east_asia: str = SONG, ascii_font: str = TIMES, size_pt: float = SZ_BODY, bold: bool = False) -> None:
    run.bold = bold
    run.font.size = Pt(size_pt)
    run.font.name = ascii_font
    rpr = run._element.get_or_add_rPr()
    rfonts = rpr.get_or_add_rFonts()
    rfonts.set(qn("w:ascii"), ascii_font)
    rfonts.set(qn("w:hAnsi"), ascii_font)
    rfonts.set(qn("w:eastAsia"), east_asia)
    rfonts.set(qn("w:cs"), ascii_font)


def set_body_para_format(para) -> None:
    pf = para.paragraph_format
    pf.line_spacing_rule = WD_LINE_SPACING.EXACTLY
    pf.line_spacing = Pt(LINE_BODY)
    pf.space_before = Pt(0)
    pf.space_after = Pt(0)
    pf.first_line_indent = Cm(FIRST_INDENT_CM)
    para.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY


def set_display_para_format(para) -> None:
    pf = para.paragraph_format
    pf.line_spacing_rule = WD_LINE_SPACING.EXACTLY
    pf.line_spacing = Pt(LINE_BODY)
    pf.space_before = Pt(0)
    pf.space_after = Pt(0)
    pf.first_line_indent = None
    para.alignment = WD_ALIGN_PARAGRAPH.CENTER


def clear_para_content(para) -> None:
    """Remove runs/objects but keep w:pPr (section typography)."""
    p = para._p
    for child in list(p):
        if child.tag.endswith("}pPr"):
            continue
        p.remove(child)


def para_has_embedded(para) -> bool:
    return any(True for tag in EMBED_TAGS for _ in para._p.iter(tag))


def set_para_latex_text(para, text: str, *, display: bool = False) -> None:
    """Replace paragraph with LaTeX plain text + fonts + paragraph format."""
    clear_para_content(para)
    run = para.add_run(text)
    set_run_fonts(run, bold=False)
    if display:
        set_display_para_format(para)
    else:
        set_body_para_format(para)


def verify_paragraphs_latex(doc: Document, indices: Iterable[int], *, require_dollar: bool = True) -> list[str]:
    issues: list[str] = []
    for i in indices:
        if i < 0 or i >= len(doc.paragraphs):
            issues.append(f"para {i}: out of range")
            continue
        para = doc.paragraphs[i]
        t = para.text or ""
        if para_has_embedded(para):
            issues.append(f"para {i}: stale OLE/OMath")
        if require_dollar and "$" not in t:
            issues.append(f"para {i}: missing LaTeX $ delimiters")
        run = para.runs[0] if para.runs else None
        if run is not None:
            rpr = run._element.rPr
            rfonts = rpr.rFonts if rpr is not None else None
            east = rfonts.get(qn("w:eastAsia")) if rfonts is not None else None
            if not east:
                issues.append(f"para {i}: missing eastAsia font on run")
    return issues


def verify_log_section(
    log_path: Path,
    *,
    word_para_lo: int,
    word_para_hi: int,
    min_inline: int = 1,
    min_display: int = 0,
) -> list[str]:
    """Word paragraph indices are 1-based (COM Paragraphs index)."""
    data = json.loads(log_path.read_text(encoding="utf-8"))
    s = data.get("summary", {})
    issues: list[str] = []
    if int(s.get("failed", 0)) > 0:
        issues.append(f"failed={s.get('failed')}")
    if int(s.get("inlineConverted", 0)) < min_inline and min_inline > 0:
        issues.append(f"inlineConverted={s.get('inlineConverted')} (expected >={min_inline})")
    if int(s.get("displayConverted", 0)) < min_display:
        issues.append(f"displayConverted={s.get('displayConverted')} (expected >={min_display})")
    skipped = []
    for e in data.get("events", []):
        if e.get("status") != "skipped-existing-object":
            continue
        pi = int(e.get("paragraphIndex", 0))
        if word_para_lo <= pi <= word_para_hi:
            skipped.append(pi)
    if skipped:
        issues.append(f"skipped-existing-object in section Word paras: {sorted(set(skipped))}")
    return issues
