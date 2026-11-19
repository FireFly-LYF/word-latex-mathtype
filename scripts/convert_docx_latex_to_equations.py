# -*- coding: utf-8 -*-
"""Convert LaTeX formula text in docx to MathType OLE via headless Word TeX toggle.

Default path for word-latex-mathtype on Windows:

1. Open the target docx with ``WINWORD.EXE /q`` (no visible Word window).
2. Attach with ``win32com.client.dynamic.Dispatch`` (avoids TYPE_E_CANTLOADLIBRARY).
3. Wrap LaTeX as ``$...$`` / ``$$...$$`` and run ``MTCommand_TeXToggle``.
4. Produce MathType OLE objects (ProgID ``Equation.DSMT4``, DSMT7 payload) without
   opening the MathType GUI (no ads/dialogs, safe for batch).
5. Fall back to Word OMath when toggle fails.

Do **not** use MathType InsertInline/InsertDisp GUI automation for batch jobs.
"""
from __future__ import annotations

import argparse
import json
import re
import shutil
import subprocess
import sys
import time
import winreg
from datetime import datetime
from pathlib import Path

import pythoncom
import win32com.client
import win32com.client.dynamic as win32dynamic


def _force_dynamic_dispatch(obj, ob_user_name=None, result_clsid=None):
    if obj is None:
        return None
    try:
        return win32dynamic.Dispatch(obj)
    except Exception:  # noqa: BLE001
        return obj


win32com.client._get_good_object_ = _force_dynamic_dispatch
win32com.client._get_good_single_object_ = _force_dynamic_dispatch
win32com.client.__WrapDispatch = (
    lambda dispatch, userName=None, resultCLSID=None, typeinfo=None, clsctx=pythoncom.CLSCTX_SERVER: win32dynamic.Dispatch(
        dispatch
    )
)

WD_ALIGN_PARAGRAPH_CENTER = 1

INLINE_PATTERNS = [
    r"t_\{\\mathrm\{on\}\}",
    r"t_\{\\mathrm\{off\}\}",
    r"\(t_\{\\mathrm\{on\}\},t_\{\\mathrm\{off\}\}\)",
    r"\(t_\{\\mathrm\{on\}\},\s*t_\{\\mathrm\{off\}\}\)",
    r"t_\{\\mathrm\{off\}\}\\ge\s*t_\{\\mathrm\{on\}\}",
    r"\\Delta\s*P_d",
    r"\\Delta\s*P_\{d,t\}",
    r"P_\{d,t\}",
    r"P_d",
    r"\\mathrm\{SINR\}(?:\^\{\\mathrm\{norm\}\})?(?:_t)?",
    r"SINR(?:\^\{\\mathrm\{norm\}\})?(?:_t)?",
    r"w_\{pd\}",
    r"w_\{s\}",
    r"g\(P_\{d,t\}\)",
    r"\[0,t_\{\\mathrm\{on\}\}\)",
    r"\[t_\{\\mathrm\{on\}\},t_\{\\mathrm\{off\}\}\)",
    r"\[t_\{\\mathrm\{off\}\},1\]",
    r"\[t_\{\\mathrm\{off\}\},1\)",
    r"ratio\^\\alpha",
    r"Q_[A-Za-z0-9{}]+\\times\s*Q_[A-Za-z0-9{}]+(?:\\times\s*Q_[A-Za-z0-9{}]+)*",
    r"[A-Za-z](?:_[A-Za-z0-9{}\\]+)+(?:\([^()\r\n]{0,80}\))?(?:\s*=\s*[^$\u4e00-\u9fff\u3000-\u303f]{1,80})",
    r"\\alpha\s*[<>=]\s*1",
    r"\\alpha",
    r"k_[0-9]+",
    r"S_\{[^}]+\}",
]

CODE_ID_RE = re.compile(
    r"^(build|load|read|apply|initialize|finalize|perform|grow|remove|calibrate|integrate|mie)_"
    r"[A-Za-z0-9_]*(?:\s*/\s*(build|load|read|apply|initialize|finalize|perform|grow|remove|"
    r"calibrate|integrate|mie)_[A-Za-z0-9_]*)*$"
)
ASCII_ID_RE = re.compile(
    r"^[A-Za-z][A-Za-z0-9_]*(?:\s*/\s*[A-Za-z][A-Za-z0-9_]*)*(?:\s+[\u4e00-\u9fff]+)?$"
)

WORD_EXE_CANDIDATES = (
    Path(r"C:\Program Files\Microsoft Office\root\Office16\WINWORD.EXE"),
    Path(r"C:\Program Files (x86)\Microsoft Office\root\Office16\WINWORD.EXE"),
)


def resolve_default_output_path(input_path: Path, suffix: str) -> Path:
    stem = input_path.stem
    if suffix.endswith(".json") or suffix.endswith(".docx"):
        return input_path.with_name(stem + suffix)
    return input_path.with_name(stem + suffix + input_path.suffix)


def resolve_word_executable(explicit: Path | None = None) -> Path:
    if explicit is not None:
        explicit = explicit.resolve()
        if not explicit.is_file():
            raise FileNotFoundError(f"Word executable not found: {explicit}")
        return explicit

    for candidate in WORD_EXE_CANDIDATES:
        if candidate.is_file():
            return candidate

    for hive in (winreg.HKEY_LOCAL_MACHINE, winreg.HKEY_CURRENT_USER):
        for subkey in (
            r"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\WINWORD.EXE",
            r"SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\App Paths\WINWORD.EXE",
        ):
            try:
                with winreg.OpenKey(hive, subkey) as key:
                    value, _ = winreg.QueryValueEx(key, "")
                    path = Path(str(value))
                    if path.is_file():
                        return path
            except OSError:
                continue

    raise FileNotFoundError("Could not locate WINWORD.EXE. Pass --word-exe explicitly.")


def normalize_formula_text(text: str) -> str:
    clean = re.sub(r"[\u0007\u000b]", " ", text)
    clean = re.sub(r"[\r\n]", " ", clean)
    clean = re.sub(r"\s+", " ", clean)
    return clean.strip()


def get_paragraph_body(paragraph) -> str:
    return re.sub(r"(\r|\a)+$", "", paragraph.Range.Text)


def has_delimited_math(text: str) -> bool:
    r"""True if author-marked $...$ / $$...$$ / \(...\) / \[...\] is present."""
    if not text:
        return False
    if "$$" in text or r"\[" in text or r"\(" in text:
        return True
    # Inline $...$ (not $$); allow simple identifiers like $i$, $N$.
    return bool(re.search(r"(?<!\$)\$(?!\$)(.+?)(?<!\$)\$(?!\$)", text))


def has_formula_signature(text: str) -> bool:
    if not text or not text.strip():
        return False
    # Author-delimited math always counts — including bare $i$ / $N$ in CJK prose.
    # Without this gate, the main loop silently skips those paragraphs (no log event).
    if has_delimited_math(text):
        return True
    return bool(
        re.search(r"\\[A-Za-z]+", text)
        or re.search(r"_[{A-Za-z0-9]", text)
        or re.search(r"\^[{A-Za-z0-9]", text)
        or re.search(r"[=<>]", text)
        or re.search(r"[{}]", text)
    )


def is_ascii_identifier_like(text: str) -> bool:
    return bool(ASCII_ID_RE.match(text.strip()))


def is_display_formula_paragraph(text: str) -> bool:
    normalized = normalize_formula_text(text)
    if not has_formula_signature(normalized):
        return False
    if is_ascii_identifier_like(normalized):
        return False
    # Explicit $$...$$ / \[...\] display delimiters always count as display.
    if re.match(r"^\$\$.*\$\$$", normalized) or re.match(r"^\\\[.*\\\]$", normalized):
        return True
    if ";" in normalized and not re.search(r"\\|[{}]|\^", normalized):
        return False
    if not re.search(r"\\|_|\^|[{}]", normalized):
        return False
    if not re.search(r"[\u4e00-\u9fff]", normalized):
        return True
    return bool(re.match(r"^(\$\$|\\\[|\\|\$?[A-Za-z].*[=<>])", normalized))


# Word/Unicode math that looks convertible but is NOT TeX Toggle input.
_UNICODE_MATH_MARKERS = re.compile(
    r"[\u03b1-\u03c9\u0391-\u03a9]"  # Greek letters (α, τ, …)
    r"|[\u2080-\u2089\u2090-\u209c\u2070-\u2079]"  # unicode super/subscripts
    r"|[\u2212\u00d7\u00b7\u200b]"  # minus / times / ZWSP from Word equation paste
)


def looks_like_unconverted_unicode_math(text: str) -> bool:
    """True for pasted Unicode equations like ``r(t)=αs sT(t−τs)+…`` without LaTeX."""
    normalized = normalize_formula_text(text)
    if not normalized or "$" in normalized or "\\" in normalized:
        return False
    if re.search(r"[\u4e00-\u9fff]", normalized):
        return False
    if "=" not in normalized:
        return False
    if _UNICODE_MATH_MARKERS.search(normalized):
        return True
    return bool(re.search(r"[A-Za-z]\([^)]{0,24}\)\s*=", normalized) and len(normalized) <= 160)


def should_skip_inline_match(text: str) -> bool:
    trimmed = text.strip()
    if not trimmed:
        return True
    if re.search(r"_|\\|[\^\{\}]", trimmed):
        return False
    if CODE_ID_RE.match(trimmed):
        return True
    return is_ascii_identifier_like(trimmed)


def get_inline_formula_matches(text: str) -> list[dict]:
    candidates: list[dict] = []

    # Prefer explicit $...$ / \(...\) wrappers (single-line).
    # Delimited math is author-marked: never skip as plain identifiers.
    for match in re.finditer(r"(?<!\$)\$(?!\$)(.+?)(?<!\$)\$(?!\$)", text):
        value = match.group(1).strip()
        if not value:
            continue
        candidates.append(
            {
                "start": match.start(),
                "length": match.end() - match.start(),
                "end": match.end(),
                "text": value,
                "delimited": True,
            }
        )

    for pattern in INLINE_PATTERNS:
        for match in re.finditer(pattern, text):
            value = match.group(0).strip()
            if should_skip_inline_match(value):
                continue
            candidates.append(
                {
                    "start": match.start(),
                    "length": match.end() - match.start(),
                    "end": match.end(),
                    "text": value,
                    "delimited": False,
                }
            )

    candidates.sort(key=lambda c: (c["start"], -c["length"]))
    selected: list[dict] = []
    for candidate in candidates:
        overlap = any(
            not (candidate["end"] <= existing["start"] or candidate["start"] >= existing["end"])
            for existing in selected
        )
        if not overlap:
            selected.append(candidate)
    selected.sort(key=lambda c: c["start"], reverse=True)
    return selected


def wrap_for_mathtype(formula: str, mode: str = "Inline") -> str:
    trimmed = formula.strip()
    if re.match(r"^\$.*\$$", trimmed):
        return trimmed
    if mode == "Display":
        return f"$${trimmed}$$"
    return f"${trimmed}$"


def find_inline_shape_near_start(document, start: int, max_start: int):
    for i in range(1, document.InlineShapes.Count + 1):
        shape = document.InlineShapes(i)
        if start <= shape.Range.Start <= max_start:
            return shape
    return None


def remove_adjacent_literal_dollars(inline_shape) -> None:
    document = inline_shape.Range.Document
    while True:
        removed = False
        if inline_shape.Range.Start > 0:
            leading = document.Range(inline_shape.Range.Start - 1, inline_shape.Range.Start)
            if leading.Text == "$":
                leading.Text = ""
                removed = True
        trailing = document.Range(inline_shape.Range.End, inline_shape.Range.End + 1)
        if trailing.Text == "$":
            trailing.Text = ""
            removed = True
        if not removed:
            break


def try_convert_range_with_mathtype(word, doc_range, formula: str, mode: str = "Inline") -> dict:
    """Headless MathType via Word TeX toggle (no MathType GUI)."""
    document = doc_range.Document
    wrapped = wrap_for_mathtype(formula, mode)
    start = doc_range.Start
    before = document.InlineShapes.Count

    doc_range.Text = wrapped
    working = document.Range(start, start + len(wrapped))
    working.Select()

    mt_error = ""
    for command in ("MTCommand_TeXToggle", "MTCommand_OnTexToggle"):
        try:
            word.Run(command)
            mt_error = ""
            break
        except Exception as exc:  # noqa: BLE001
            mt_error = str(exc)
    else:
        return {
            "success": False,
            "message": mt_error or "MathType toggle commands unavailable.",
            "range_start": start,
            "range_end": start + len(wrapped),
        }

    elapsed = 0
    while elapsed < 6000:
        time.sleep(0.3)
        elapsed += 300
        if document.InlineShapes.Count > before:
            time.sleep(0.8)
            shape = find_inline_shape_near_start(document, start, start + len(wrapped) + 2)
            if shape is None and document.InlineShapes.Count > 0:
                shape = document.InlineShapes(document.InlineShapes.Count)
            if shape is not None:
                # Display wraps with $$; inline matches often sit inside existing $...$.
                remove_adjacent_literal_dollars(shape)
            end = shape.Range.End if shape is not None else start + len(wrapped)
            return {
                "success": True,
                "message": "MathType TeX toggle succeeded.",
                "range_start": start,
                "range_end": end,
            }

    return {
        "success": False,
        "message": "MathType did not create an inline shape.",
        "range_start": start,
        "range_end": start + len(wrapped),
    }


def convert_latex_to_word_linear(formula: str) -> str:
    linear = normalize_formula_text(formula).strip("$")
    frac_pat = re.compile(r"\\frac\{([^{}]+)\}\{([^{}]+)\}")
    while frac_pat.search(linear):
        linear = frac_pat.sub(r"((\1)/(\2))", linear)
    macro_pat = re.compile(r"\\(?:mathrm|mathbf|mathcal|mathit|operatorname|text)\{([^{}]*)\}")
    while True:
        nxt = macro_pat.sub(r"\1", linear)
        if nxt == linear:
            break
        linear = nxt
    repl = [
        (r"\\begin\{cases\}", ""),
        (r"\\end\{cases\}", ""),
        (r"\\left", ""),
        (r"\\right", ""),
        (r"\\qquad", " "),
        (r"\\quad", " "),
        (r"\\,", " "),
        (r"\\!", ""),
        (r"\\cdot", "*"),
        (r"\\times", " x "),
        (r"\\geq?", ">="),
        (r"\\leq?", "<="),
        (r"\\neq", "!="),
        (r"\\Rightarrow", "=>"),
        (r"\\to", "->"),
        (r"\\Re", "Re"),
        (r"\\Im", "Im"),
    ]
    for pat, rep in repl:
        linear = re.sub(pat, rep, linear)
    linear = re.sub(r"\\sqrt\{([^{}]+)\}", r"sqrt(\1)", linear)
    linear = linear.replace("\\\\", "; ").replace("&", " ")
    linear = linear.replace("{", "(").replace("}", ")")
    linear = linear.replace("\\", "")
    linear = re.sub(r"\s+", " ", linear)
    return linear.strip()


def try_convert_range_with_word_equation(doc_range, formula: str) -> dict:
    document = doc_range.Document
    linear = convert_latex_to_word_linear(formula)
    if not linear:
        return {"success": False, "message": "Fallback linear text is empty."}
    before = document.OMaths.Count
    start = doc_range.Start
    doc_range.Text = linear
    working = document.Range(start, start + len(linear))
    try:
        document.OMaths.Add(working)
        document.OMaths.BuildUp()
    except Exception as exc:  # noqa: BLE001
        return {"success": False, "message": str(exc)}
    if document.OMaths.Count > before:
        omath = document.OMaths(document.OMaths.Count)
        return {"success": True, "message": "Word OMath fallback succeeded.", "range_end": omath.Range.End}
    return {"success": False, "message": "Word OMath fallback did not create a new equation."}


def open_word_document(doc_path: Path, word_exe: Path, *, open_wait_sec: float = 5.0):
    """Open docx through WINWORD.EXE /q, attach COM headlessly (tracks PID)."""
    from word_com_utils import open_word_document as _open

    return _open(doc_path, word_exe, open_wait_sec=open_wait_sec)


def kill_office_processes(*, force_all: bool = False) -> None:
    """Kill MathType; kill only managed WINWORD PIDs unless ``force_all``."""
    from word_com_utils import kill_office_processes as _kill

    _kill(force_all=force_all)


def assert_file_is_closed(path: Path, label: str) -> None:
    if not path.exists():
        return
    try:
        with path.open("r+b"):
            pass
    except OSError as exc:
        raise RuntimeError(f"{label} is locked. Close Word and retry: {path}") from exc


def convert_docx(
    source: Path,
    target: Path,
    log_path: Path,
    *,
    word_exe: Path | None = None,
    open_wait_sec: float = 5.0,
    prefer_word: bool = False,
    mathtype_mode: str = "toggle",
    source_mode: str = "latex",
    toggle_strategy: str = "auto",
    fallback: str = "none",
    omml2mml_xsl: Path | None = None,
    pandoc_exe: str | None = None,
    formula_scope: str = "all",
    mathtype_numbers: bool = True,
) -> dict:
    from word_com_utils import normalize_formula_scope

    assert_file_is_closed(target, "Target document")
    if target.exists():
        target.unlink()
    shutil.copy2(source, target)

    if prefer_word:
        mathtype_mode = "word"
    if mathtype_mode not in {"toggle", "word"}:
        raise ValueError(f"Unsupported mathtype_mode: {mathtype_mode!r}. Use 'toggle' or 'word'.")
    if source_mode not in {"auto", "latex", "omath"}:
        raise ValueError(f"Unsupported source_mode: {source_mode!r}. Use auto|latex|omath.")
    if toggle_strategy not in {"auto", "in-place", "end-paste"}:
        raise ValueError(f"Unsupported toggle_strategy: {toggle_strategy!r}.")
    if fallback not in {"none", "word"}:
        raise ValueError(f"Unsupported fallback: {fallback!r}. Use none|word.")
    scope = normalize_formula_scope(formula_scope)

    if source_mode == "omath" and mathtype_mode == "toggle":
        from convert_docx_omath_to_mathtype import convert_docx_omath

        return convert_docx_omath(
            source,
            target,
            log_path,
            word_exe=word_exe,
            open_wait_sec=open_wait_sec,
            omml2mml_xsl=omml2mml_xsl,
            pandoc_exe=pandoc_exe,
            mathtype_numbers=mathtype_numbers,
            formula_scope=scope,
        )

    omath_rewrite_stats = None
    if source_mode == "auto" and mathtype_mode == "toggle":
        from convert_docx_omath_to_mathtype import rewrite_omath_to_latex_dollars

        omath_rewrite_stats = rewrite_omath_to_latex_dollars(
            target,
            xsl_path=omml2mml_xsl,
            pandoc_exe=pandoc_exe,
            formula_scope=scope,
        )

    kill_office_processes(force_all=False)

    word_path = resolve_word_executable(word_exe)
    effective_strategy = toggle_strategy
    if source_mode == "auto" and toggle_strategy == "auto":
        effective_strategy = "end-paste"

    summary = {
        "source": str(source.resolve()),
        "target": str(target.resolve()),
        "engine": "python-win32com-headless-tex-toggle",
        "mathtypeMode": mathtype_mode,
        "sourceMode": source_mode,
        "formulaScope": scope,
        "mathtypeNumbers": bool(mathtype_numbers and scope in {"display", "all"}),
        "toggleStrategy": effective_strategy,
        "fallback": fallback,
        "preferWord": mathtype_mode == "word",
        "wordExe": str(word_path),
        "startedAt": datetime.now().strftime("%Y-%m-%dT%H:%M:%S"),
        "mathTypeAddIn": mathtype_mode == "toggle",
        "totalParagraphs": 0,
        "displayConverted": 0,
        "inlineConverted": 0,
        "mathTypeConverted": 0,
        "wordConverted": 0,
        "omathFound": int((omath_rewrite_stats or {}).get("omathFound", 0)),
        "omathRewritten": int((omath_rewrite_stats or {}).get("omathRewritten", 0)),
        "omathConverted": 0,
        "omathFailed": int((omath_rewrite_stats or {}).get("failed", 0)),
        "eqnNumbersInserted": 0,
        "eqnNumbersFailed": 0,
        "skipped": 0,
        "failed": 0,
        "unicodeMathSuspects": 0,
        "existingObjectParagraphs": 0,
    }
    events: list[dict] = []

    word, document = open_word_document(target, word_path, open_wait_sec=open_wait_sec)
    try:
        summary["totalParagraphs"] = document.Paragraphs.Count

        from word_com_utils import convert_dollars_via_end_toggle, try_end_toggle_one

        if mathtype_mode == "toggle" and effective_strategy == "end-paste":
            from omath_to_latex import sanitize_mathtype_latex

            dollar_summary, dollar_events = convert_dollars_via_end_toggle(
                word,
                document,
                sanitize=sanitize_mathtype_latex,
                kind="dollar",
                mathtype_numbers=mathtype_numbers,
                formula_scope=scope,
            )
            summary["mathTypeConverted"] += int(dollar_summary["converted"])
            summary["failed"] += int(dollar_summary["failed"])
            summary["eqnNumbersInserted"] += int(dollar_summary.get("eqnNumbersInserted", 0))
            summary["eqnNumbersFailed"] += int(dollar_summary.get("eqnNumbersFailed", 0))
            if source_mode == "auto":
                summary["omathConverted"] += int(dollar_summary["converted"])
                summary["omathFailed"] += int(dollar_summary["failed"])
            for ev in dollar_events:
                if ev.get("status") == "end-toggle-pasteback":
                    if ev.get("display"):
                        summary["displayConverted"] += 1
                    else:
                        summary["inlineConverted"] += 1
            events.extend(dollar_events)

        def convert_formula(rng, formula: str, mode: str) -> dict | None:
            if mathtype_mode == "word":
                return None
            display = mode == "Display"
            bare = formula.strip().strip("$")
            if effective_strategy == "end-paste":
                return try_end_toggle_one(
                    word, document, int(rng.Start), int(rng.End), bare, display=display
                )
            if effective_strategy == "in-place":
                return try_convert_range_with_mathtype(word, rng, formula, mode)
            result = try_convert_range_with_mathtype(word, rng, formula, mode)
            if result and result.get("success"):
                return result
            end_result = try_end_toggle_one(
                word, document, int(rng.Start), int(rng.End), bare, display=display
            )
            if end_result.get("success"):
                end_result["message"] = (
                    f"In-place failed ({(result or {}).get('message', '')}); end-paste ok."
                )
                return end_result
            return result or end_result

        def record_mathtype_success(kind: str, index: int, formula: str, result: dict, status: str) -> None:
            if kind == "display":
                summary["displayConverted"] += 1
            else:
                summary["inlineConverted"] += 1
            summary["mathTypeConverted"] += 1
            events.append(
                {
                    "kind": kind,
                    "paragraphIndex": index,
                    "engine": "MathType",
                    "status": status,
                    "formula": formula,
                    "message": result["message"],
                }
            )

        def apply_fallback_or_fail(kind: str, index: int, formula: str, rng, mt_msg: str) -> None:
            if fallback == "word":
                fb = try_convert_range_with_word_equation(rng, formula)
                if fb.get("success"):
                    if kind == "display":
                        try:
                            rng.Paragraphs(1).Alignment = WD_ALIGN_PARAGRAPH_CENTER
                        except Exception:  # noqa: BLE001
                            pass
                        summary["displayConverted"] += 1
                    else:
                        summary["inlineConverted"] += 1
                    summary["wordConverted"] += 1
                    events.append(
                        {
                            "kind": kind,
                            "paragraphIndex": index,
                            "engine": "Word",
                            "status": "word-omath",
                            "formula": formula,
                            "message": fb["message"],
                        }
                    )
                    return
                mt_msg = f"{fb.get('message', '')} | MathType={mt_msg}"
            summary["failed"] += 1
            events.append(
                {
                    "kind": kind,
                    "paragraphIndex": index,
                    "engine": "none",
                    "status": "failed",
                    "formula": formula,
                    "message": mt_msg,
                }
            )

        for index in range(document.Paragraphs.Count, 0, -1):
            paragraph = document.Paragraphs(index)
            body = get_paragraph_body(paragraph)
            if not body.strip():
                continue

            if paragraph.Range.OMaths.Count > 0 or paragraph.Range.InlineShapes.Count > 0:
                summary["skipped"] += 1
                summary["existingObjectParagraphs"] += 1
                events.append(
                    {
                        "kind": "paragraph",
                        "paragraphIndex": index,
                        "engine": "none",
                        "status": "skipped-existing-object",
                        "formula": normalize_formula_text(body),
                    }
                )
                continue

            if looks_like_unconverted_unicode_math(body) and not is_display_formula_paragraph(body):
                summary["skipped"] += 1
                summary["unicodeMathSuspects"] += 1
                events.append(
                    {
                        "kind": "display",
                        "paragraphIndex": index,
                        "engine": "none",
                        "status": "skipped-unicode-math",
                        "formula": normalize_formula_text(body),
                        "message": (
                            "Looks like pasted Unicode math, not LaTeX. "
                            "Rewrite as $$...$$ with \\alpha, \\tau, etc. before converting."
                        ),
                    }
                )
                continue

            if effective_strategy == "end-paste" and ("$" in body):
                continue

            # Delimited $...$ alone (e.g. CJK + $i$) must still be converted even if
            # legacy signature heuristics would miss them; keep both checks.
            inline_matches = get_inline_formula_matches(body)
            if not has_formula_signature(body) and not inline_matches:
                continue

            if is_display_formula_paragraph(body):
                if scope == "inline":
                    continue
                formula = normalize_formula_text(body)
                if is_ascii_identifier_like(formula):
                    summary["skipped"] += 1
                    events.append(
                        {
                            "kind": "display",
                            "paragraphIndex": index,
                            "engine": "none",
                            "status": "skipped-identifier",
                            "formula": formula,
                        }
                    )
                    continue

                rng = document.Range(paragraph.Range.Start, paragraph.Range.Start + len(body))
                result = convert_formula(rng, formula, "Display")
                if result and result.get("success"):
                    try:
                        paragraph.Alignment = WD_ALIGN_PARAGRAPH_CENTER
                    except Exception:  # noqa: BLE001
                        pass
                    record_mathtype_success("display", index, formula, result, "display-mathtype")
                    continue

                mt_msg = result["message"] if result else "MathType not attempted."
                end = (
                    result["range_end"]
                    if result and "range_end" in result
                    else paragraph.Range.Start + len(body)
                )
                rng = document.Range(paragraph.Range.Start, end)
                apply_fallback_or_fail("display", index, formula, rng, mt_msg)
                continue

            if scope == "display":
                continue

            matches = inline_matches
            if not matches:
                continue

            for match in matches:
                formula = match["text"].strip()
                if (not match.get("delimited")) and should_skip_inline_match(formula):
                    summary["skipped"] += 1
                    events.append(
                        {
                            "kind": "inline",
                            "paragraphIndex": index,
                            "engine": "none",
                            "status": "skipped-identifier",
                            "formula": formula,
                        }
                    )
                    continue
                if effective_strategy == "end-paste" and match.get("delimited"):
                    continue

                start = paragraph.Range.Start + match["start"]
                rng = document.Range(start, start + match["length"])
                result = convert_formula(rng, formula, "Inline")
                if result and result.get("success"):
                    record_mathtype_success("inline", index, formula, result, "inline-mathtype")
                    continue

                mt_msg = result["message"] if result else "MathType not attempted."
                end = (
                    result["range_end"]
                    if result and "range_end" in result
                    else start + match["length"]
                )
                rng = document.Range(start, end)
                apply_fallback_or_fail("inline", index, formula, rng, mt_msg)

        # Post-pass: MathType right numbers after in-place / auto fallback path.
        # end-paste already numbers inside convert_dollars_via_end_toggle; this covers
        # in-place TeX toggle which previously left eqnNumbersInserted=0.
        if (
            mathtype_mode == "toggle"
            and bool(mathtype_numbers)
            and scope in {"display", "all"}
            and int(summary.get("eqnNumbersInserted", 0)) == 0
            and int(summary.get("mathTypeConverted", 0)) > 0
        ):
            try:
                from word_com_utils import insert_mathtype_numbers_for_lone_displays

                num_summary = insert_mathtype_numbers_for_lone_displays(word, document)
                summary["eqnNumbersInserted"] = int(num_summary.get("inserted", 0))
                summary["eqnNumbersFailed"] = int(num_summary.get("failed", 0))
                events.append(
                    {
                        "kind": "paragraph",
                        "engine": "MathType",
                        "status": "inplace-lone-display-numbers",
                        "message": (
                            f"Post-pass InsertEqnNum: inserted={summary['eqnNumbersInserted']} "
                            f"failed={summary['eqnNumbersFailed']} skipped={num_summary.get('skipped', 0)}"
                        ),
                    }
                )
            except Exception as exc:  # noqa: BLE001
                summary["eqnNumbersFailed"] = int(summary.get("eqnNumbersFailed", 0)) + 1
                events.append(
                    {
                        "kind": "paragraph",
                        "engine": "none",
                        "status": "inplace-numbers-failed",
                        "message": str(exc),
                    }
                )

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
        kill_office_processes(force_all=False)

    summary["finishedAt"] = datetime.now().strftime("%Y-%m-%dT%H:%M:%S")
    output = {"summary": summary, "events": events}
    log_path.parent.mkdir(parents=True, exist_ok=True)
    log_path.write_text(json.dumps(output, ensure_ascii=False, indent=2), encoding="utf-8")
    return output


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(
        description="Convert LaTeX in docx to MathType OLE (headless TeX toggle) or Word OMath."
    )
    parser.add_argument("--source", type=Path, required=True, help="Source .docx with LaTeX text")
    parser.add_argument("--target", type=Path, help="Output .docx (default: <source>_formula.docx)")
    parser.add_argument("--log", type=Path, help="JSON log path (default: <source>_formula_log.json)")
    parser.add_argument("--word-exe", type=Path, help="Path to WINWORD.EXE")
    parser.add_argument("--open-wait-sec", type=float, default=5.0, help="Seconds to wait after CLI open")
    parser.add_argument(
        "--prefer-word",
        action="store_true",
        help="Use Word native OMath only (skip MathType).",
    )
    parser.add_argument(
        "--mathtype-mode",
        choices=("toggle", "word"),
        default="toggle",
        help="toggle = headless MathType TeX toggle (default); word = OMath only.",
    )
    parser.add_argument(
        "--source-mode",
        choices=("auto", "latex", "omath"),
        default="latex",
        help=(
            "latex = LaTeX text only (default); "
            "omath = Word/WPS OMML only; "
            "auto = rewrite OMath then convert all $...$ (end-paste) + bare heuristics."
        ),
    )
    parser.add_argument(
        "--toggle-strategy",
        choices=("auto", "in-place", "end-paste"),
        default="auto",
        help=(
            "auto = try in-place then end-paste (latex); "
            "end-paste = thesis-safe (forced for source-mode auto); "
            "in-place = legacy TeX toggle at formula location."
        ),
    )
    parser.add_argument(
        "--fallback",
        choices=("none", "word"),
        default="none",
        help="none = fail if MathType fails (default); word = fall back to Word OMath.",
    )
    parser.add_argument("--omml2mml-xsl", type=Path, help="Path to OMML2MML.XSL (omath/auto)")
    parser.add_argument("--pandoc-exe", type=str, help="Path to pandoc (omath/auto)")
    parser.add_argument(
        "--formula-scope",
        choices=("all", "inline", "display"),
        default="all",
        help="all=段内+单行; inline=仅段内 $...$; display=仅单行 $$...$$(可右编号).",
    )
    parser.add_argument(
        "--mathtype-numbers",
        dest="mathtype_numbers",
        action="store_true",
        default=True,
        help="For display formulas, insert MathType right numbers (default on).",
    )
    parser.add_argument(
        "--no-mathtype-numbers",
        dest="mathtype_numbers",
        action="store_false",
        help="Do not insert MathType equation numbers.",
    )
    args = parser.parse_args(argv)

    source = args.source.resolve()
    if not source.is_file():
        print(f"Source not found: {source}", file=sys.stderr)
        return 1

    target = args.target.resolve() if args.target else resolve_default_output_path(source, "_formula.docx")
    log_path = args.log.resolve() if args.log else resolve_default_output_path(source, "_formula_log.json")
    if target == source:
        print("Target must differ from source.", file=sys.stderr)
        return 1

    result = convert_docx(
        source,
        target,
        log_path,
        word_exe=args.word_exe,
        open_wait_sec=args.open_wait_sec,
        prefer_word=args.prefer_word,
        mathtype_mode=args.mathtype_mode,
        source_mode=args.source_mode,
        toggle_strategy=args.toggle_strategy,
        fallback=args.fallback,
        omml2mml_xsl=args.omml2mml_xsl,
        pandoc_exe=args.pandoc_exe,
        formula_scope=args.formula_scope,
        mathtype_numbers=args.mathtype_numbers,
    )
    s = result["summary"]
    print(
        f"sourceMode={s.get('sourceMode', '?')} scope={s.get('formulaScope', '?')} "
        f"strategy={s.get('toggleStrategy', '?')} "
        f"fallback={s.get('fallback', '?')} mode={s.get('mathtypeMode', '?')} "
        f"display={s['displayConverted']} inline={s['inlineConverted']} "
        f"omath={s.get('omathConverted', 0)}/{s.get('omathFound', 0)} "
        f"eqnNum={s.get('eqnNumbersInserted', 0)} "
        f"mathtype={s['mathTypeConverted']} word={s['wordConverted']} "
        f"skipped={s['skipped']} failed={s['failed']} "
        f"unicodeSuspects={s.get('unicodeMathSuspects', 0)} "
        f"existingOLE={s.get('existingObjectParagraphs', 0)}"
    )
    if s.get("unicodeMathSuspects", 0):
        print(
            "WARNING: source has Unicode-math paragraphs (not LaTeX). "
            "Rewrite as $$...$$ with \\\\commands before converting; see log status skipped-unicode-math.",
            file=sys.stderr,
        )
    if s.get("existingObjectParagraphs", 0) and s["displayConverted"] + s["inlineConverted"] == 0:
        print(
            "WARNING: many paragraphs already contain OLE/OMath and were skipped. "
            "Convert from a clean LaTeX source, not a previous _formula.docx merge.",
            file=sys.stderr,
        )
    print(f"target={target}")
    print(f"log={log_path}")
    return 0 if s["failed"] == 0 else 2


if __name__ == "__main__":
    raise SystemExit(main())
