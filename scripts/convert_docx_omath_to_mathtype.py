# -*- coding: utf-8 -*-
"""Convert Word / WPS OMML equations in a .docx to MathType OLE.

Pipeline (thesis TAB+(n) safe):

1. Copy source → target
2. OOXML: replace ``m:oMath`` / ``m:oMathPara`` with ``$...$`` / ``$$...$$``;
   strip Word-typed ``(n)`` / tabs on numbered display paragraphs → ``$$...$$``
3. Open Word; end-of-doc TeX Toggle + OLE paste-back for every dollar span
4. For display formulas, insert MathType right-side equation numbers
   (``MTCommand_InsertEqnNum``), not Word plain-text ``(n)``
5. Always remove unzip temp dirs

Requires: Windows, Word, MathType add-in, pywin32, lxml, pandoc.
"""
from __future__ import annotations

import argparse
import json
import re
import shutil
import sys
import tempfile
import zipfile
from datetime import datetime
from pathlib import Path

from lxml import etree

from omath_to_latex import (
    mathml_to_latex,
    omml_element_to_mathml,
    require_pandoc,
    resolve_omml2mml_xsl,
    sanitize_mathtype_latex,
)
from word_com_utils import (
    WORD_EQNUM_RE,
    convert_dollars_via_end_toggle,
    kill_office_processes,
    normalize_formula_scope,
    open_word_document,
)

from convert_docx_latex_to_equations import (
    assert_file_is_closed,
    resolve_default_output_path,
    resolve_word_executable,
)

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
M_NS = "http://schemas.openxmlformats.org/officeDocument/2006/math"
NSMAP = {"w": W_NS, "m": M_NS}


def _make_run_with_text(text: str) -> etree._Element:
    r = etree.Element("{%s}r" % W_NS)
    rpr = etree.SubElement(r, "{%s}rPr" % W_NS)
    rfonts = etree.SubElement(rpr, "{%s}rFonts" % W_NS)
    rfonts.set("{%s}ascii" % W_NS, "Times New Roman")
    rfonts.set("{%s}hAnsi" % W_NS, "Times New Roman")
    rfonts.set("{%s}eastAsia" % W_NS, "宋体")
    rfonts.set("{%s}cs" % W_NS, "Times New Roman")
    t = etree.SubElement(r, "{%s}t" % W_NS)
    t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    t.text = text
    return r


def _omath_node_to_latex(node: etree._Element, xsl: Path, pandoc: str) -> str:
    mathml = omml_element_to_mathml(node, xsl_path=xsl)
    latex = sanitize_mathtype_latex(mathml_to_latex(mathml, pandoc_exe=pandoc))
    # Keep dollar-span regex / Word Range.Text matching stable (no raw newlines).
    latex = re.sub(r"[\r\n\x0b]+", " ", latex)
    latex = re.sub(r"\s+", " ", latex).strip()
    return latex


def _wp_enclosing(node: etree._Element) -> etree._Element | None:
    wp = node
    while wp is not None and wp.tag != "{%s}p" % W_NS:
        wp = wp.getparent()
    return wp


def _wp_non_math_text(wp: etree._Element) -> str:
    """Plain text outside ``m:oMath`` / ``m:oMathPara`` (prose / Word eqnum / tabs runs)."""
    parts: list[str] = []
    for t in wp.iter("{%s}t" % W_NS):
        if any(
            etree.QName(a).localname in ("oMath", "oMathPara")
            for a in t.iterancestors()
        ):
            continue
        if t.text:
            parts.append(t.text)
    return "".join(parts)


def _wp_plain_text(wp: etree._Element) -> str:
    """Backward-compatible: non-math prose only (avoids counting OMML as body text)."""
    return _wp_non_math_text(wp)


def _wp_has_word_eqnum(wp: etree._Element) -> bool:
    """True if paragraph has Word-typed ``(n)`` / ``（n）`` (not MathType SEQ)."""
    return bool(WORD_EQNUM_RE.search(_wp_non_math_text(wp)))


def _wp_is_formula_only(wp: etree._Element) -> bool:
    """Paragraph is essentially only OMML (+ tabs/spaces/eqnum), i.e. a display line."""
    text = _wp_non_math_text(wp)
    stripped = WORD_EQNUM_RE.sub("", text)
    stripped = re.sub(r"[\s\u00a0\u3000]+", "", stripped)
    return stripped == ""


def _wp_is_display_host(wp: etree._Element | None) -> bool:
    """True only for standalone formula lines (not in-sentence oMathPara + prose)."""
    if wp is None:
        return False
    return _wp_is_formula_only(wp) or (
        _wp_has_word_eqnum(wp) and _wp_is_formula_only(wp)
    )


def _flatten_latex_for_dollar(latex: str) -> str:
    """Collapse newlines so Word paragraph scan / DOLLAR_RE can see the span."""
    return re.sub(r"[\r\n]+", " ", latex).strip()


def _replace_node_with_dollar_run(node: etree._Element, latex: str, *, display: bool) -> None:
    """Replace an ``m:oMath`` / ``m:oMathPara`` node with a dollar text run in place."""
    parent = node.getparent()
    if parent is None:
        return
    idx = list(parent).index(node)
    parent.remove(node)
    flat = _flatten_latex_for_dollar(latex)
    wrapped = f"$${flat}$$" if display else f"${flat}$"
    parent.insert(idx, _make_run_with_text(wrapped))


def _replace_wp_with_display_dollar(wp: etree._Element, latex: str) -> None:
    """Keep ``w:pPr``; clear body; insert ``$$latex$$`` (no Word ``(n)``).

    Alignment left: MathType right-number layout later uses center/right tab stops.
    Only call for true formula-only paragraphs — never for prose hosts.
    """
    for child in list(wp):
        if child.tag != "{%s}pPr" % W_NS:
            wp.remove(child)
    ppr = wp.find("{%s}pPr" % W_NS)
    if ppr is None:
        ppr = etree.Element("{%s}pPr" % W_NS)
        wp.insert(0, ppr)
    jc = ppr.find("{%s}jc" % W_NS)
    if jc is None:
        jc = etree.SubElement(ppr, "{%s}jc" % W_NS)
    jc.set("{%s}val" % W_NS, "left")
    # Drop inherited center/right tabs that assume full page width (dual-column safe).
    tabs = ppr.find("{%s}tabs" % W_NS)
    if tabs is not None:
        ppr.remove(tabs)
    flat = _flatten_latex_for_dollar(latex)
    wp.append(_make_run_with_text(f"$${flat}$$"))


def rewrite_omath_to_latex_dollars(
    docx_path: Path,
    *,
    xsl_path: Path | None = None,
    pandoc_exe: str | None = None,
    work_dir: Path | None = None,
    formula_scope: str = "all",
) -> dict:
    """Replace OMML in ``docx_path`` with ``$latex$`` / ``$$latex$$`` text runs (in place).

    Numbered display paragraphs (Word ``(n)`` / tabs + lone OMML) become ``$$...$$``
    with Word numbers stripped so MathType can insert right-side numbers later.

    ``formula_scope``: ``inline`` | ``display`` | ``all`` — only rewrite matching OMML.
    Unselected OMML is left untouched for a later independent pass.
    """
    scope = normalize_formula_scope(formula_scope)
    xsl = resolve_omml2mml_xsl(xsl_path)
    pandoc = pandoc_exe or require_pandoc()

    own_tmp = work_dir is None
    tmp_dir = Path(work_dir) if work_dir else Path(tempfile.mkdtemp(prefix="wlmt_omath_unzip_"))
    stats = {
        "omathPara": 0,
        "omathInline": 0,
        "omathDisplayNumbered": 0,
        "wordEqnumStripped": 0,
        "failed": 0,
        "skippedByScope": 0,
        "formulaScope": scope,
        "formulas": [],
    }
    try:
        shutil.rmtree(tmp_dir, ignore_errors=True)
        tmp_dir.mkdir(parents=True, exist_ok=True)

        with zipfile.ZipFile(docx_path, "r") as zin:
            zin.extractall(tmp_dir)

        doc_xml = tmp_dir / "word" / "document.xml"
        root = etree.parse(str(doc_xml)).getroot()

        for para in list(root.xpath("//m:oMathPara", namespaces=NSMAP)):
            oms = para.xpath(".//m:oMath", namespaces=NSMAP)
            if not oms:
                continue
            try:
                latex = _omath_node_to_latex(oms[0], xsl, pandoc)
                if not latex:
                    raise ValueError("empty latex")
                # oMathPara may sit inside w:p that still has prose (Word quirk:
                # end-of-sentence "display" math). Only formula-only hosts are
                # true single-line display; otherwise rewrite as inline $...$.
                wp = _wp_enclosing(para)
                is_display = _wp_is_display_host(wp)
                if is_display and scope == "inline":
                    stats["skippedByScope"] += 1
                    continue
                if (not is_display) and scope == "display":
                    stats["skippedByScope"] += 1
                    continue
                if is_display:
                    if _wp_has_word_eqnum(wp):
                        stats["wordEqnumStripped"] += 1
                    _replace_wp_with_display_dollar(wp, latex)
                    stats["omathPara"] += 1
                    stats["omathDisplayNumbered"] += 1
                    stats["formulas"].append(
                        {"display": True, "numbered": True, "latex": latex}
                    )
                else:
                    _replace_node_with_dollar_run(para, latex, display=False)
                    stats["omathInline"] += 1
                    stats["formulas"].append(
                        {"display": False, "numbered": False, "latex": latex}
                    )
            except Exception as exc:  # noqa: BLE001
                stats["failed"] += 1
                stats["formulas"].append({"display": True, "error": str(exc)})

        for om in list(root.xpath("//m:oMath", namespaces=NSMAP)):
            if om.getparent() is None:
                continue
            # Skip oMath still nested under oMathPara (handled above / left for scope).
            if any(
                etree.QName(a).localname == "oMathPara" for a in om.iterancestors()
            ):
                continue
            try:
                latex = _omath_node_to_latex(om, xsl, pandoc)
                if not latex:
                    raise ValueError("empty latex")
                wp = _wp_enclosing(om)
                # Thesis style: lone oMath + Word (n) / tabs → display $$ + strip number.
                # In-sentence oMath with prose must stay inline — never rewrite whole para.
                is_display = _wp_is_display_host(wp)
                if is_display and scope == "inline":
                    stats["skippedByScope"] += 1
                    continue
                if (not is_display) and scope == "display":
                    stats["skippedByScope"] += 1
                    continue
                if is_display:
                    if _wp_has_word_eqnum(wp):
                        stats["wordEqnumStripped"] += 1
                    _replace_wp_with_display_dollar(wp, latex)
                    stats["omathPara"] += 1
                    stats["omathDisplayNumbered"] += 1
                    stats["formulas"].append(
                        {"display": True, "numbered": True, "latex": latex}
                    )
                else:
                    _replace_node_with_dollar_run(om, latex, display=False)
                    stats["omathInline"] += 1
                    stats["formulas"].append(
                        {"display": False, "numbered": False, "latex": latex}
                    )
            except Exception as exc:  # noqa: BLE001
                stats["failed"] += 1
                stats["formulas"].append({"display": False, "error": str(exc)})

        doc_xml.write_bytes(
            etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)
        )

        repack = tmp_dir.parent / (tmp_dir.name + "_repack.docx")
        if repack.exists():
            repack.unlink()
        with zipfile.ZipFile(repack, "w", compression=zipfile.ZIP_DEFLATED) as zout:
            for f in tmp_dir.rglob("*"):
                if f.is_file():
                    zout.write(f, f.relative_to(tmp_dir).as_posix())
        shutil.move(str(repack), str(docx_path))
    finally:
        if own_tmp:
            shutil.rmtree(tmp_dir, ignore_errors=True)

    stats["omathFound"] = stats["omathPara"] + stats["omathInline"] + stats["failed"]
    stats["omathRewritten"] = stats["omathPara"] + stats["omathInline"]
    return stats


def convert_latex_dollars_via_end_toggle(
    word,
    document,
    *,
    mathtype_numbers: bool = True,
    formula_scope: str = "all",
    word_exe: Path | None = None,
) -> tuple[dict, list[dict]]:
    """OMath path: end-toggle paste-back + MathType right numbers for ``$$...$$``."""
    summary, events = convert_dollars_via_end_toggle(
        word,
        document,
        sanitize=sanitize_mathtype_latex,
        kind="omath",
        mathtype_numbers=mathtype_numbers,
        formula_scope=formula_scope,
        word_exe=word_exe,
    )
    # Map keys expected by older callers
    mapped = {
        "omathFound": 0,
        "dollarFound": summary["dollarFound"],
        "omathConverted": summary["converted"],
        "omathFailed": summary["failed"],
        "eqnNumbersInserted": summary.get("eqnNumbersInserted", 0),
        "eqnNumbersFailed": summary.get("eqnNumbersFailed", 0),
        "formulaScope": summary.get("formulaScope", formula_scope),
        "skippedMathTypeOle": 0,
        "rpcReopens": summary.get("rpcReopens", 0),
        "_word": summary.get("_word"),
        "_document": summary.get("_document"),
    }
    return mapped, events


def convert_omath_document(word, document, *, xsl_path=None, pandoc_exe=None):
    del xsl_path, pandoc_exe
    return convert_latex_dollars_via_end_toggle(word, document)


def convert_docx_omath(
    source: Path,
    target: Path,
    log_path: Path,
    *,
    word_exe: Path | None = None,
    open_wait_sec: float = 5.0,
    omml2mml_xsl: Path | None = None,
    pandoc_exe: str | None = None,
    mathtype_numbers: bool = True,
    formula_scope: str = "all",
) -> dict:
    scope = normalize_formula_scope(formula_scope)
    assert_file_is_closed(target, "Target document")
    if target.exists():
        target.unlink()
    shutil.copy2(source, target)

    resolve_omml2mml_xsl(omml2mml_xsl)
    pandoc = pandoc_exe or require_pandoc()

    rewrite_stats = rewrite_omath_to_latex_dollars(
        target,
        xsl_path=omml2mml_xsl,
        pandoc_exe=pandoc,
        formula_scope=scope,
    )

    kill_office_processes(force_all=True)
    word_path = resolve_word_executable(word_exe)
    summary = {
        "source": str(source.resolve()),
        "target": str(target.resolve()),
        "engine": "python-omath-xml-end-toggle-pasteback",
        "sourceMode": "omath",
        "formulaScope": scope,
        "mathtypeNumbers": bool(mathtype_numbers and scope in {"display", "all"}),
        "wordExe": str(word_path),
        "pandoc": pandoc,
        "startedAt": datetime.now().strftime("%Y-%m-%dT%H:%M:%S"),
        "omathFound": int(rewrite_stats.get("omathFound", 0)),
        "omathRewritten": int(rewrite_stats.get("omathRewritten", 0)),
        "omathRewriteFailed": int(rewrite_stats.get("failed", 0)),
        "omathDisplayNumbered": int(rewrite_stats.get("omathDisplayNumbered", 0)),
        "wordEqnumStripped": int(rewrite_stats.get("wordEqnumStripped", 0)),
        "skippedByScope": int(rewrite_stats.get("skippedByScope", 0)),
        "dollarFound": 0,
        "omathConverted": 0,
        "omathFailed": 0,
        "eqnNumbersInserted": 0,
        "eqnNumbersFailed": 0,
        "remainingOMath": 0,
        "inlineShapes": 0,
        "rpcReopens": 0,
        "openedFullName": "",
    }
    events: list[dict] = []

    word, document = open_word_document(target, word_path, open_wait_sec=open_wait_sec)
    try:
        try:
            summary["openedFullName"] = str(document.FullName)
        except Exception:  # noqa: BLE001
            summary["openedFullName"] = ""
        toggle_summary, toggle_events = convert_latex_dollars_via_end_toggle(
            word,
            document,
            mathtype_numbers=mathtype_numbers,
            formula_scope=scope,
            word_exe=word_path,
        )
        # RPC reopen may have replaced COM handles.
        if toggle_summary.get("_document") is not None:
            document = toggle_summary["_document"]
        if toggle_summary.get("_word") is not None:
            word = toggle_summary["_word"]
        summary["dollarFound"] = int(toggle_summary.get("dollarFound", 0))
        summary["omathConverted"] = int(toggle_summary.get("omathConverted", 0))
        summary["omathFailed"] = int(toggle_summary.get("omathFailed", 0)) + int(
            rewrite_stats.get("failed", 0)
        )
        summary["eqnNumbersInserted"] = int(toggle_summary.get("eqnNumbersInserted", 0))
        summary["eqnNumbersFailed"] = int(toggle_summary.get("eqnNumbersFailed", 0))
        summary["rpcReopens"] = int(toggle_summary.get("rpcReopens", 0))
        events.extend(toggle_events)
        try:
            summary["remainingOMath"] = int(document.OMaths.Count)
            summary["inlineShapes"] = int(document.InlineShapes.Count)
        except Exception:  # noqa: BLE001
            pass
        document.Save()
    finally:
        try:
            document.Close(False)
        except Exception:  # noqa: BLE001
            pass
        try:
            word.Quit()
        except Exception:  # noqa: BLE001
            pass
        kill_office_processes(force_all=True)

    summary["finishedAt"] = datetime.now().strftime("%Y-%m-%dT%H:%M:%S")
    output = {
        "summary": summary,
        "events": events,
        "rewriteFormulas": rewrite_stats.get("formulas", []),
    }
    log_path.parent.mkdir(parents=True, exist_ok=True)
    log_path.write_text(json.dumps(output, ensure_ascii=False, indent=2), encoding="utf-8")
    return output


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(
        description="Convert Word/WPS OMath (OMML) equations in a docx to MathType OLE."
    )
    parser.add_argument("--source", type=Path, required=True)
    parser.add_argument("--target", type=Path)
    parser.add_argument("--log", type=Path)
    parser.add_argument("--word-exe", type=Path)
    parser.add_argument("--open-wait-sec", type=float, default=5.0)
    parser.add_argument("--omml2mml-xsl", type=Path)
    parser.add_argument("--pandoc-exe", type=str)
    parser.add_argument(
        "--mathtype-numbers",
        dest="mathtype_numbers",
        action="store_true",
        default=True,
        help="Insert MathType right-side equation numbers for $$ display formulas (default).",
    )
    parser.add_argument(
        "--no-mathtype-numbers",
        dest="mathtype_numbers",
        action="store_false",
        help="Do not insert MathType equation numbers.",
    )
    parser.add_argument(
        "--formula-scope",
        choices=("all", "inline", "display"),
        default="all",
        help="all=段内+单行; inline=仅段内; display=仅单行(可右编号).",
    )
    args = parser.parse_args(argv)

    source = args.source.resolve()
    if not source.is_file():
        print(f"Source not found: {source}", file=sys.stderr)
        return 1
    target = args.target.resolve() if args.target else resolve_default_output_path(source, "_formula.docx")
    log_path = (
        args.log.resolve()
        if args.log
        else resolve_default_output_path(source, "_omath_formula_log.json")
    )
    if target == source:
        print("Target must differ from source.", file=sys.stderr)
        return 1

    result = convert_docx_omath(
        source,
        target,
        log_path,
        word_exe=args.word_exe,
        open_wait_sec=args.open_wait_sec,
        omml2mml_xsl=args.omml2mml_xsl,
        pandoc_exe=args.pandoc_exe,
        mathtype_numbers=args.mathtype_numbers,
        formula_scope=args.formula_scope,
    )
    s = result["summary"]
    print(
        f"scope={s.get('formulaScope', '?')} "
        f"omathFound={s['omathFound']} rewritten={s.get('omathRewritten', '?')} "
        f"converted={s['omathConverted']} failed={s['omathFailed']} "
        f"eqnNum={s.get('eqnNumbersInserted', 0)} "
        f"wordEqnumStripped={s.get('wordEqnumStripped', 0)} "
        f"remainingOMath={s.get('remainingOMath', '?')} shapes={s.get('inlineShapes', '?')}"
    )
    print(f"target={target}")
    print(f"log={log_path}")
    return 0 if s["omathFailed"] == 0 else 2


if __name__ == "__main__":
    raise SystemExit(main())
