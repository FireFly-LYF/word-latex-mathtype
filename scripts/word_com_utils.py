# -*- coding: utf-8 -*-
"""Shared Word/MathType COM helpers for word-latex-mathtype.

- Track WINWORD PID started by this skill and prefer killing only that process
- End-of-document TeX Toggle + OLE paste-back (thesis TAB+(n) safe)
- MathType right-side equation numbers (not Word-typed ``(1)``)
- Dollar-span collection / cleanup helpers
"""
from __future__ import annotations

import re
import subprocess
import time
import winreg
from pathlib import Path

import pythoncom
import win32com.client.dynamic as win32dynamic

WD_ALIGN_PARAGRAPH_LEFT = 0
WD_ALIGN_PARAGRAPH_CENTER = 1
WD_ALIGN_TAB_CENTER = 1
WD_ALIGN_TAB_RIGHT = 2
HELPER_MARKER = "⟦WLMT⟧"
# Allow newlines inside spans (pandoc matrix / multi-line OMML→LaTeX).
DOLLAR_RE = re.compile(
    r"\$\$([^$]+?)\$\$|(?<!\$)\$([^$\r]+?)\$(?!\$)",
    re.DOTALL,
)
WORD_EQNUM_RE = re.compile(r"[\(（]\s*\d+\s*[\)）]")

# Module-level PIDs started by open_word_document in this process.
_managed_word_pids: set[int] = set()


def _list_pids(image_name: str) -> set[int]:
    proc = subprocess.run(
        ["tasklist", "/FI", f"IMAGENAME eq {image_name}", "/FO", "CSV", "/NH"],
        capture_output=True,
        text=True,
        encoding="utf-8",
        errors="ignore",
        check=False,
    )
    pids: set[int] = set()
    for line in (proc.stdout or "").splitlines():
        # "WINWORD.EXE","1234",...
        parts = [p.strip().strip('"') for p in line.split(",")]
        if len(parts) >= 2 and parts[0].lower() == image_name.lower():
            try:
                pids.add(int(parts[1]))
            except ValueError:
                pass
    return pids


def _kill_pids(pids: set[int]) -> None:
    for pid in sorted(pids):
        subprocess.run(
            ["taskkill", "/PID", str(pid), "/F"],
            check=False,
            capture_output=True,
        )


def kill_mathtype() -> None:
    subprocess.run(["taskkill", "/IM", "MathType.exe", "/F"], check=False, capture_output=True)


def kill_office_processes(*, force_all: bool = False) -> None:
    """Kill MathType always; kill Word either managed PIDs only or all instances."""
    kill_mathtype()
    if force_all or not _managed_word_pids:
        subprocess.run(["taskkill", "/IM", "WINWORD.EXE", "/F"], check=False, capture_output=True)
        _managed_word_pids.clear()
    else:
        alive = _list_pids("WINWORD.EXE")
        _kill_pids(_managed_word_pids & alive)
        _managed_word_pids.clear()
    time.sleep(0.6)


def register_word_pid(pid: int | None) -> None:
    if pid:
        _managed_word_pids.add(int(pid))


def _document_matches_path(document, doc_path: Path) -> bool:
    try:
        full = Path(str(document.FullName)).resolve()
        want = Path(doc_path).resolve()
        return full.as_posix().casefold() == want.as_posix().casefold()
    except Exception:  # noqa: BLE001
        return False


def is_rpc_com_error(exc: BaseException) -> bool:
    """True when Word/MathType COM server died or Range.Text/Select is dead."""
    text = str(exc).lower()
    needles = (
        "rpc",
        "1726",
        "1722",
        "<unknown>.text",
        "<unknown>.select",
        "<unknown>.count",
        "server is unavailable",
        "远程过程调用",
        "rpc 服务器不可用",
        "被呼叫方拒绝",
    )
    if any(n in text for n in needles):
        return True
    try:
        import pywintypes  # type: ignore

        if isinstance(exc, pywintypes.com_error) and exc.args:
            hr = int(exc.args[0])
            return hr in (-2147023174, -2147023170, -2147417848, -2147023169, -2147418111)
    except Exception:  # noqa: BLE001
        pass
    return False


def open_word_document(doc_path: Path, word_exe: Path, *, open_wait_sec: float = 5.0):
    """Open docx via ``WINWORD.EXE path /q`` (loads MathType), then COM-attach.

    Proven paperA path from today's skill tests. Caller should
    ``kill_office_processes(force_all=True)`` first. Verifies ``FullName``.
    """
    doc_path = Path(doc_path).resolve()
    try:
        lock = doc_path.with_name("~$" + doc_path.name)
        if lock.exists():
            lock.unlink()
    except Exception:  # noqa: BLE001
        pass

    before = _list_pids("WINWORD.EXE")
    subprocess.Popen([str(word_exe), str(doc_path), "/q"], close_fds=True)
    time.sleep(max(float(open_wait_sec), 5.0))
    after = _list_pids("WINWORD.EXE")
    new_pids = after - before
    if new_pids:
        _managed_word_pids.update(new_pids)
    elif after:
        _managed_word_pids.update(after)

    pythoncom.CoInitialize()
    word = win32dynamic.Dispatch("Word.Application")
    try:
        word.Visible = False
    except Exception:  # noqa: BLE001
        pass
    try:
        word.DisplayAlerts = 0
    except Exception:  # noqa: BLE001
        pass

    document = None
    for _ in range(25):
        try:
            if int(word.Documents.Count) >= 1:
                active = word.ActiveDocument
                if _document_matches_path(active, doc_path):
                    document = active
                    break
                for i in range(1, int(word.Documents.Count) + 1):
                    doc = word.Documents(i)
                    if _document_matches_path(doc, doc_path):
                        document = doc
                        break
                if document is not None:
                    break
        except Exception:  # noqa: BLE001
            document = None
        time.sleep(0.4)

    if document is None:
        document = word.Documents.Open(
            str(doc_path),
            ConfirmConversions=False,
            ReadOnly=False,
            AddToRecentFiles=False,
        )
        time.sleep(1.0)

    if not _document_matches_path(document, doc_path):
        raise RuntimeError(
            f"Word opened wrong document: got {getattr(document, 'FullName', '?')!r}, "
            f"expected {str(doc_path)!r}"
        )
    try:
        document.Activate()
    except Exception:  # noqa: BLE001
        pass
    return word, document


def reopen_word_document(doc_path: Path, word_exe: Path, *, open_wait_sec: float = 5.0):
    """Kill Office and reopen ``doc_path`` after RPC / Word death."""
    kill_office_processes(force_all=True)
    time.sleep(1.5)
    return open_word_document(doc_path, word_exe, open_wait_sec=open_wait_sec)


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


def _delete_from_marker_to_end(document) -> None:
    """Remove helper content inserted after HELPER_MARKER near document end."""
    try:
        find = document.Content.Find
        find.ClearFormatting()
        ok = find.Execute(FindText=HELPER_MARKER, Forward=False, Wrap=0)  # wdFindStop=0
        if not ok:
            return
        start = int(find.Parent.Start)
        end = int(document.Content.End) - 1
        if end > start:
            document.Range(start, end).Delete()
    except Exception:  # noqa: BLE001
        # Fallback: delete last shape if very near end
        try:
            if int(document.InlineShapes.Count) > 0:
                last = document.InlineShapes(document.InlineShapes.Count)
                if int(last.Range.Start) > int(document.Content.End) - 200:
                    last.Delete()
        except Exception:  # noqa: BLE001
            pass


def ensure_mathtype_eqnnum_defaults(*, on_right: bool = True, format_str: str = "(#E1)") -> None:
    """Set MathType WordCommands registry: right-side numbers, no warning dialog."""
    key = winreg.CreateKey(winreg.HKEY_CURRENT_USER, r"Software\Design Science\DSMT7\WordCommands")
    try:
        winreg.SetValueEx(key, "EqnNumsOnRight", 0, winreg.REG_DWORD, 1 if on_right else 0)
        winreg.SetValueEx(key, "NoEqnNumWarningDlg", 0, winreg.REG_DWORD, 1)
        if format_str:
            try:
                winreg.SetValueEx(key, "DefaultEqnNumFormat", 0, winreg.REG_SZ, format_str)
            except OSError:
                pass
    finally:
        winreg.CloseKey(key)


ensure_mathtype_eqnnum_prefs = ensure_mathtype_eqnnum_defaults


def insert_mathtype_right_eqnnum(word, document, shape) -> dict:
    """Insert MathType equation number after ``shape`` (right side via registry)."""
    ensure_mathtype_eqnnum_defaults(on_right=True)
    try:
        document.Range(int(shape.Range.End), int(shape.Range.End)).Select()
    except Exception as exc:  # noqa: BLE001
        return {"success": False, "message": f"select after OLE failed: {exc}"}
    last = "InsertEqnNum unavailable"
    for cmd in ("MTCommand_InsertEqnNum", "MTCommand_OnInsertEquationNumber"):
        try:
            word.Run(cmd)
            time.sleep(0.35)
            return {"success": True, "message": f"{cmd} ok", "command": cmd}
        except Exception as exc:  # noqa: BLE001
            last = str(exc)
    return {"success": False, "message": last}


insert_mathtype_equation_number = insert_mathtype_right_eqnnum


def _page_usable_width_pts(page_setup) -> float:
    return float(page_setup.PageWidth) - float(page_setup.LeftMargin) - float(
        page_setup.RightMargin
    )


def _column_text_width_pts(document, para) -> tuple[float, int]:
    """Return (text width pts, column count) for tab stops.

    Multi-column sections must use **column** width, not full page usable width;
    otherwise center/right tabs land outside the column and wreck dual-column layout.
    """
    try:
        ps = para.Range.PageSetup
    except Exception:  # noqa: BLE001
        ps = document.PageSetup
    page_usable = _page_usable_width_pts(ps)
    n_cols = 1
    try:
        cols = ps.TextColumns
        n_cols = max(1, int(cols.Count))
    except Exception:  # noqa: BLE001
        return page_usable, 1
    if n_cols < 2:
        return page_usable, n_cols
    # Prefer explicit per-column width when Word exposes it.
    try:
        widths = []
        for i in range(1, n_cols + 1):
            try:
                widths.append(float(cols.Item(i).Width))
            except Exception:  # noqa: BLE001
                widths.append(0.0)
        positive = [w for w in widths if w > 1.0]
        if positive:
            # Equal-width columns share one width; unequal → use the narrowest
            # so the right tab never spills past a column edge.
            return min(positive), n_cols
    except Exception:  # noqa: BLE001
        pass
    try:
        spacing = float(cols.Spacing)
    except Exception:  # noqa: BLE001
        spacing = 0.0
    col_w = (page_usable - (n_cols - 1) * spacing) / float(n_cols)
    return max(col_w, 1.0), n_cols


def apply_mathtype_right_number_layout(document, shape) -> dict:
    """Force classic MathType right-number layout: ``TAB + OLE + TAB + (n)``.

    ``MTCommand_InsertEqnNum`` alone often leaves the number glued to the OLE.
    A proper right number needs left-aligned paragraph + center/right tab stops
    and real tab characters around the equation.

    Tab stops use **column** text width when ``TextColumns.Count >= 2``.
    """
    try:
        para = shape.Range.Paragraphs(1)
    except Exception as exc:  # noqa: BLE001
        return {"success": False, "message": f"no paragraph: {exc}"}

    # Only touch lone display paragraphs; never restyle body prose / inline hosts.
    if not _is_lone_display_equation(document, shape):
        return {"success": False, "message": "skip layout: not lone display paragraph"}

    try:
        para.Alignment = WD_ALIGN_PARAGRAPH_LEFT
    except Exception:  # noqa: BLE001
        pass

    try:
        usable, n_cols = _column_text_width_pts(document, para)
        # Built-in "equation" style often carries full-page tab stops — skip it
        # in multi-column sections so we do not blow past the column edge.
        if n_cols < 2:
            try:
                para.Style = "equation"
            except Exception:  # noqa: BLE001
                pass
        # TabStops.Position is measured from the column/left margin (points).
        center_pos = usable / 2.0
        right_pos = usable
        keep = False
        try:
            for i in range(1, int(para.TabStops.Count) + 1):
                ts = para.TabStops(i)
                if int(ts.Alignment) == WD_ALIGN_TAB_CENTER and abs(
                    float(ts.Position) - center_pos
                ) <= 8.0:
                    keep = True
                    break
        except Exception:  # noqa: BLE001
            keep = False
        if not keep:
            try:
                para.TabStops.ClearAll()
            except Exception:  # noqa: BLE001
                pass
            para.TabStops.Add(Position=center_pos, Alignment=WD_ALIGN_TAB_CENTER)
            para.TabStops.Add(Position=right_pos, Alignment=WD_ALIGN_TAB_RIGHT)
    except Exception as exc:  # noqa: BLE001
        return {"success": False, "message": f"tab stops failed: {exc}"}

    # Refresh shape: first DSMT4 in this paragraph.
    try:
        shape = None
        p_start = int(para.Range.Start)
        p_end = int(para.Range.End)
        for i in range(1, int(document.InlineShapes.Count) + 1):
            sh = document.InlineShapes(i)
            if p_start <= int(sh.Range.Start) < p_end and _is_dsmt4(sh):
                shape = sh
                break
        if shape is None:
            return {"success": False, "message": "OLE not found in paragraph"}
    except Exception as exc:  # noqa: BLE001
        return {"success": False, "message": f"refresh shape failed: {exc}"}

    try:
        # Tab after OLE (pushes number to right tab stop).
        sh_end = int(shape.Range.End)
        after = document.Range(sh_end, min(sh_end + 2, int(para.Range.End)))
        if not (after.Text or "").startswith("\t"):
            document.Range(sh_end, sh_end).InsertAfter("\t")

        # Tab before OLE (centers equation on center tab stop).
        shape = para.Range.InlineShapes(1)
        sh_start = int(shape.Range.Start)
        before = document.Range(int(para.Range.Start), sh_start)
        if "\t" not in (before.Text or ""):
            document.Range(sh_start, sh_start).InsertBefore("\t")
    except Exception as exc:  # noqa: BLE001
        return {"success": False, "message": f"insert tabs failed: {exc}"}

    return {"success": True, "message": "TAB+OLE+TAB right layout applied"}


def update_mathtype_eqn_numbers(word, document) -> dict:
    """Refresh MathType/Word fields after inserting equation numbers."""
    last = "UpdateEqns unavailable"
    for cmd in ("MTCommand_UpdateEqns", "MTCommand_OnUpdateEquationNumbers"):
        try:
            word.Run(cmd)
            time.sleep(0.4)
            try:
                document.Fields.Update()
            except Exception:  # noqa: BLE001
                pass
            return {"success": True, "message": f"{cmd} ok", "command": cmd}
        except Exception as exc:  # noqa: BLE001
            last = str(exc)
    try:
        document.Fields.Update()
    except Exception:  # noqa: BLE001
        pass
    return {"success": False, "message": last}


update_mathtype_equation_numbers = update_mathtype_eqn_numbers


def tex_toggle_at_end(word, document, latex: str, *, display: bool, wait_loops: int = 30):
    """Create MathType OLE at a clean end paragraph; return InlineShape or None.

    Raises on RPC/COM death so the caller can reopen Word; returns ``None`` for
    soft TeX-toggle failures.
    """
    latex = re.sub(r"[\r\n\x0b]+", " ", latex or "")
    latex = re.sub(r"\s+", " ", latex).strip()
    wrapped = f"$${latex}$$" if display else f"${latex}$"
    end = document.Content.End - 1
    document.Range(end, end).InsertAfter("\r" + HELPER_MARKER)
    start = document.Content.End - 1
    document.Range(start, start).InsertAfter(wrapped)
    work = document.Range(start, start + len(wrapped))
    try:
        work.Font.Name = "Times New Roman"
        work.Font.Size = 10.5
    except Exception as exc:  # noqa: BLE001
        if is_rpc_com_error(exc):
            raise
    try:
        document.Activate()
    except Exception as exc:  # noqa: BLE001
        if is_rpc_com_error(exc):
            raise
    selected = False
    for _ in range(3):
        try:
            work.Select()
            selected = True
            break
        except Exception as exc:  # noqa: BLE001
            if is_rpc_com_error(exc):
                raise
            time.sleep(0.35)
            try:
                work = document.Range(start, start + len(wrapped))
            except Exception as exc2:  # noqa: BLE001
                if is_rpc_com_error(exc2):
                    raise
    if not selected:
        _delete_from_marker_to_end(document)
        return None
    before = int(document.InlineShapes.Count)
    try:
        word.Run("MTCommand_TeXToggle")
    except Exception as exc:  # noqa: BLE001
        if is_rpc_com_error(exc):
            raise
        try:
            _delete_from_marker_to_end(document)
        except Exception:  # noqa: BLE001
            pass
        return None

    shape = None
    for _ in range(wait_loops):
        time.sleep(0.25)
        try:
            count = int(document.InlineShapes.Count)
        except Exception as exc:  # noqa: BLE001
            if is_rpc_com_error(exc):
                raise
            break
        if count > before:
            time.sleep(0.35)
            shape = document.InlineShapes(document.InlineShapes.Count)
            try:
                remove_adjacent_literal_dollars(shape)
            except Exception as exc:  # noqa: BLE001
                if is_rpc_com_error(exc):
                    raise
            break
    if shape is None:
        try:
            _delete_from_marker_to_end(document)
        except Exception as exc:  # noqa: BLE001
            if is_rpc_com_error(exc):
                raise
    return shape



def _scrub_residual_dollar_text(paragraph) -> None:
    """Remove leftover ``$...$`` / ``$$...$$`` text after OLE paste-back.

    Lone-display numbering requires empty prose beside the OLE; leftover LaTeX
    also shows as duplicated formulas in visual QA.
    """
    try:
        document = paragraph.Range.Document
        p0 = int(paragraph.Range.Start)
        raw = paragraph.Range.Text or ""
    except Exception:  # noqa: BLE001
        return
    for match in reversed(list(DOLLAR_RE.finditer(raw))):
        try:
            frag = document.Range(p0 + match.start(), p0 + match.end())
            if int(frag.InlineShapes.Count) > 0:
                continue
            try:
                if int(frag.Fields.Count) > 0:
                    continue
            except Exception:  # noqa: BLE001
                pass
            frag.Text = ""
        except Exception:  # noqa: BLE001
            continue


def paste_ole_over_range(document, abs_start: int, abs_end: int, shape, *, display: bool):
    """Copy ``shape`` OLE over ``[abs_start, abs_end)``, restore light typography, cleanup helper.

    Returns the pasted InlineShape when found, else None.

    Inline paste preserves host paragraph alignment. Display paste does **not**
    force center/equation style onto body prose — lone-display right-number
    layout is applied later via ``apply_mathtype_right_number_layout``.
    """
    shape.Range.Copy()
    target = document.Range(abs_start, abs_end)
    # Capture nearby font size / alignment before paste replaces the span.
    try:
        font_size = float(target.Font.Size)
    except Exception:  # noqa: BLE001
        font_size = 10.5
    host_align = None
    try:
        host_align = int(target.Paragraphs(1).Alignment)
    except Exception:  # noqa: BLE001
        host_align = None
    # CRITICAL: Range.Paste() on some Word builds inserts OLE without deleting
    # the selected LaTeX text, leaving $$...$$ + OLE side-by-side. Clear first.
    try:
        target.Text = ""
    except Exception:  # noqa: BLE001
        try:
            target.Delete()
        except Exception:  # noqa: BLE001
            pass
    insert_at = document.Range(abs_start, abs_start)
    insert_at.Paste()
    pasted = None
    # After paste, the inserted shape sits at abs_start.
    try:
        for i in range(1, int(document.InlineShapes.Count) + 1):
            sh = document.InlineShapes(i)
            if abs_start - 2 <= int(sh.Range.Start) <= abs_start + 8:
                pasted = sh
                break
        if pasted is not None:
            try:
                pasted.Range.Font.Size = font_size
            except Exception:  # noqa: BLE001
                pass
            try:
                para = pasted.Range.Paragraphs(1)
                if display and _is_lone_display_equation(document, pasted):
                    # Leave left-aligned for subsequent TAB+OLE+TAB right-number
                    # layout; do not center (breaks dual-column + right tabs).
                    para.Alignment = WD_ALIGN_PARAGRAPH_LEFT
                elif host_align is not None:
                    # Inline / in-prose: restore original paragraph alignment.
                    para.Alignment = host_align
            except Exception:  # noqa: BLE001
                pass
    except Exception:  # noqa: BLE001
        pass
    _delete_from_marker_to_end(document)
    if pasted is not None and display:
        try:
            _scrub_residual_dollar_text(pasted.Range.Paragraphs(1))
        except Exception:  # noqa: BLE001
            pass
    return pasted


def _find_shape_near(document, abs_start: int):
    try:
        for i in range(1, int(document.InlineShapes.Count) + 1):
            sh = document.InlineShapes(i)
            if abs_start - 2 <= int(sh.Range.Start) <= abs_start + 8:
                return sh
    except Exception:  # noqa: BLE001
        return None
    return None


def normalize_formula_scope(scope: str | None) -> str:
    """Normalize scope to ``inline`` | ``display`` | ``all``."""
    value = (scope or "all").strip().lower()
    aliases = {
        "all": "all",
        "both": "all",
        "inline": "inline",
        "in-line": "inline",
        "段内": "inline",
        "display": "display",
        "single": "display",
        "single-line": "display",
        "单行": "display",
        "numbered": "display",
    }
    if value not in aliases:
        raise ValueError(
            f"Unsupported formula_scope: {scope!r}. Use inline|display|all."
        )
    return aliases[value]


def collect_dollar_formulas(document, *, formula_scope: str = "all") -> list[dict]:
    """Scan ``document.Content.Text`` once for ``$...$`` / ``$$...$$``.

    Avoids per-paragraph COM walks that commonly RPC-crash on dual-column papers.
    """
    scope = normalize_formula_scope(formula_scope)
    items: list[dict] = []
    try:
        content = document.Content
        base = int(content.Start)
        full = content.Text or ""
    except Exception as exc:  # noqa: BLE001
        if is_rpc_com_error(exc):
            raise
        return items
    for m in DOLLAR_RE.finditer(full):
        if m.group(1) is not None:
            latex = m.group(1).strip()
            display = True
        else:
            latex = (m.group(2) or "").strip()
            display = False
        if not latex:
            continue
        window = full[max(0, m.start() - 24) : m.end() + 8]
        if HELPER_MARKER in window:
            continue
        if scope == "inline" and display:
            continue
        if scope == "display" and not display:
            continue
        latex = re.sub(r"[\r\n\x0b]+", " ", latex)
        latex = re.sub(r"\s+", " ", latex).strip()
        para_idx = full.count("\r", 0, m.start()) + 1
        items.append(
            {
                "para": para_idx,
                "abs_start": base + m.start(),
                "abs_end": base + m.end(),
                "latex": latex,
                "display": display,
                "add_number": display,
            }
        )
    items.sort(key=lambda x: x["abs_start"], reverse=True)
    return items


def convert_dollars_via_end_toggle(
    word,
    document,
    *,
    sanitize=None,
    kind: str = "formula",
    mathtype_numbers: bool = False,
    formula_scope: str = "all",
    word_exe: Path | None = None,
    max_rpc_reopens: int = 4,
    proactive_reopen_every: int = 0,
) -> tuple[dict, list[dict]]:
    """Convert ``$...$`` / ``$$...$$`` via end-toggle + paste-back.

    ``formula_scope``:
      - ``inline``: only ``$...$`` (段内)
      - ``display``: only ``$$...$$`` (单行/独立成段，可右编号)
      - ``all``: both

    When ``mathtype_numbers`` is True, each converted display formula gets a
    MathType right-side equation number via ``MTCommand_InsertEqnNum``.

    On Word RPC / ``<unknown>.Text`` death, save + ``reopen_word_document``
    (up to ``max_rpc_reopens``). Also reopens every ``proactive_reopen_every``
    successful conversions to keep dual-column batches alive.
    """
    scope = normalize_formula_scope(formula_scope)
    summary = {
        "dollarFound": 0,
        "converted": 0,
        "failed": 0,
        "eqnNumbersInserted": 0,
        "eqnNumbersFailed": 0,
        "formulaScope": scope,
        "rpcReopens": 0,
    }
    events: list[dict] = []
    want_numbers = bool(mathtype_numbers and scope in {"display", "all"})
    if want_numbers:
        ensure_mathtype_eqnnum_defaults(on_right=True)

    doc_path = None
    try:
        doc_path = Path(str(document.FullName)).resolve()
    except Exception:  # noqa: BLE001
        doc_path = None

    def _reopen(reason: str, display_flag: bool, para_i: int, latex_s: str) -> None:
        nonlocal word, document, items, idx, rpc_reopens
        rpc_reopens += 1
        summary["rpcReopens"] = rpc_reopens
        events.append(
            {
                "kind": kind,
                "engine": "none",
                "status": "rpc-reopen",
                "display": display_flag,
                "paragraphIndex": para_i,
                "formula": latex_s,
                "message": reason,
            }
        )
        try:
            document.Save()
        except Exception:  # noqa: BLE001
            pass
        try:
            document.Close(False)
        except Exception:  # noqa: BLE001
            pass
        try:
            word.Quit()
        except Exception:  # noqa: BLE001
            pass
        word, document = reopen_word_document(doc_path, Path(word_exe))
        if want_numbers:
            ensure_mathtype_eqnnum_defaults(on_right=True)
        items = collect_dollar_formulas(document, formula_scope=scope)
        summary["dollarFound"] = max(
            summary["dollarFound"], len(items) + summary["converted"]
        )
        idx = 0
        summary["_word"] = word
        summary["_document"] = document

    items: list[dict] = []
    try:
        items = collect_dollar_formulas(document, formula_scope=scope)
    except Exception as exc:  # noqa: BLE001
        if (
            is_rpc_com_error(exc)
            and doc_path is not None
            and word_exe is not None
        ):
            word, document = reopen_word_document(doc_path, Path(word_exe))
            summary["rpcReopens"] = 1
            summary["_word"] = word
            summary["_document"] = document
            items = collect_dollar_formulas(document, formula_scope=scope)
        else:
            raise
    summary["dollarFound"] = len(items)
    idx = 0
    rpc_reopens = int(summary.get("rpcReopens", 0))
    since_reopen = 0
    while idx < len(items):
        item = items[idx]
        latex = item["latex"]
        # Collapse newlines / soft breaks so MathType TeX toggle stays stable.
        latex = re.sub(r"[\r\n\x0b]+", " ", latex)
        latex = re.sub(r"\s+", " ", latex).strip()
        if sanitize is not None:
            latex = sanitize(latex)
        display = bool(item["display"])
        try:
            shape = tex_toggle_at_end(word, document, latex, display=display)
            if shape is None:
                summary["failed"] += 1
                events.append(
                    {
                        "kind": kind,
                        "engine": "none",
                        "status": "failed-tex-toggle",
                        "display": display,
                        "paragraphIndex": item["para"],
                        "formula": latex,
                        "message": "TeX toggle at document end did not create OLE.",
                    }
                )
                idx += 1
                continue
            paste_ole_over_range(
                document,
                item["abs_start"],
                item["abs_end"],
                shape,
                display=display,
            )
            summary["converted"] += 1
            since_reopen += 1
            events.append(
                {
                    "kind": kind,
                    "engine": "MathType",
                    "status": "end-toggle-pasteback",
                    "display": display,
                    "paragraphIndex": item["para"],
                    "formula": latex,
                    "message": "End-toggle + OLE paste-back succeeded.",
                    "eqnNumber": False,
                    "addNumber": bool(want_numbers and item.get("add_number") and display),
                }
            )
            # Checkpoint long runs so RPC death does not lose all OLE pastes.
            if summary["converted"] % 5 == 0:
                try:
                    document.Save()
                except Exception:  # noqa: BLE001
                    pass
            idx += 1
            # Proactive Word recycle on large dual-column / long inline batches.
            if (
                proactive_reopen_every > 0
                and since_reopen >= proactive_reopen_every
                and doc_path is not None
                and word_exe is not None
                and rpc_reopens < max_rpc_reopens
                and idx < len(items)
            ):
                _reopen(
                    f"Proactive reopen after {since_reopen} conversions "
                    f"(#{rpc_reopens + 1})",
                    display,
                    item["para"],
                    latex,
                )
                since_reopen = 0
                continue
        except Exception as exc:  # noqa: BLE001
            if (
                is_rpc_com_error(exc)
                and doc_path is not None
                and word_exe is not None
                and rpc_reopens < max_rpc_reopens
            ):
                _reopen(
                    f"Word COM died; reopen #{rpc_reopens + 1}: {exc}",
                    display,
                    item["para"],
                    latex,
                )
                since_reopen = 0
                continue
            summary["failed"] += 1
            events.append(
                {
                    "kind": kind,
                    "engine": "none",
                    "status": "failed-paste" if "paste" in str(exc).lower() else "failed-com",
                    "display": display,
                    "paragraphIndex": item["para"],
                    "formula": latex,
                    "message": str(exc),
                }
            )
            try:
                _delete_from_marker_to_end(document)
            except Exception:  # noqa: BLE001
                pass
            idx += 1

    if want_numbers:
        try:
            num_summary = insert_mathtype_numbers_for_lone_displays(word, document)
            summary["eqnNumbersInserted"] = int(num_summary.get("inserted", 0))
            summary["eqnNumbersFailed"] = int(num_summary.get("failed", 0))
            if summary["eqnNumbersInserted"] > 0:
                for ev in events:
                    if ev.get("addNumber") and ev.get("status") == "end-toggle-pasteback":
                        ev["eqnNumber"] = True
                        ev["message"] += " MathType right number (lone-display pass)."
        except Exception as exc:  # noqa: BLE001
            if is_rpc_com_error(exc):
                summary["eqnNumbersFailed"] = summary.get("eqnNumbersFailed", 0) + 1
                events.append(
                    {
                        "kind": kind,
                        "engine": "none",
                        "status": "rpc-numbers-failed",
                        "message": str(exc),
                    }
                )
            else:
                raise

    # Return live COM handles so caller can Save the reopened document.
    summary["_word"] = word
    summary["_document"] = document
    return summary, events


def _is_dsmt4(shape) -> bool:
    try:
        prog = str(shape.OLEFormat.ProgID)
        return "Equation.DSMT4" in prog or "DSMT" in prog
    except Exception:  # noqa: BLE001
        return False


def _paragraph_has_mteqn_field(paragraph) -> bool:
    try:
        for i in range(1, int(paragraph.Range.Fields.Count) + 1):
            code = str(paragraph.Range.Fields(i).Code.Text)
            if "MTEqn" in code or "MTPlaceRef" in code:
                return True
    except Exception:  # noqa: BLE001
        pass
    return False


def _is_lone_display_equation(document, shape) -> bool:
    """True when shape is the only OLE in its paragraph and little/no prose remains.

    After ``InsertEqnNum``, Word keeps field markers (``\x13``..``\x15``) and
    possibly MathType ``MTEqn`` fields beside the OLE. Those must NOT disqualify
    the paragraph — otherwise ``apply_mathtype_right_number_layout`` is skipped
    and the visual result is wrong (number glued left; formula not centered).
    """
    try:
        para = shape.Range.Paragraphs(1)
        p_start = int(para.Range.Start)
        p_end = int(para.Range.End)
        n_ole = 0
        for i in range(1, int(document.InlineShapes.Count) + 1):
            sh = document.InlineShapes(i)
            s = int(sh.Range.Start)
            if p_start <= s < p_end:
                n_ole += 1
        if n_ole != 1:
            return False
        # Strip OLE placeholder, tabs, whitespace, Word field markers, typed (n).
        body = re.sub(
            r"[\r\x07\x0b\x01\x08\t\s\u00a0\u3000\x13\x14\x15]+",
            "",
            para.Range.Text or "",
        )
        body = WORD_EQNUM_RE.sub("", body)
        if body == "":
            return True
        # MTEqn / MTPlaceRef fields count as numbering chrome, not prose.
        if _paragraph_has_mteqn_field(para) and not re.search(
            r"[\u4e00-\u9fffA-Za-z]", body
        ):
            return True
        return False
    except Exception:  # noqa: BLE001
        return False
        body = re.sub(r"[\r\x07\x0b\x01\x08\t\s\u00a0\u3000]+", "", para.Range.Text or "")
        body = WORD_EQNUM_RE.sub("", body)
        return body == ""
    except Exception:  # noqa: BLE001
        return False


def insert_mathtype_numbers_for_lone_displays(word, document) -> dict:
    """Insert MathType right numbers for lone display OLE paragraphs (document order)."""
    ensure_mathtype_eqnnum_defaults(on_right=True)
    result = {"inserted": 0, "failed": 0, "skipped": 0}
    shapes: list = []
    try:
        for i in range(1, int(document.InlineShapes.Count) + 1):
            sh = document.InlineShapes(i)
            if not _is_dsmt4(sh):
                continue
            if not _is_lone_display_equation(document, sh):
                continue
            if _paragraph_has_mteqn_field(sh.Range.Paragraphs(1)):
                result["skipped"] += 1
                continue
            shapes.append(sh)
    except Exception:  # noqa: BLE001
        return result

    shapes.sort(key=lambda sh: int(sh.Range.Start))
    for sh in shapes:
        num_res = insert_mathtype_right_eqnnum(word, document, sh)
        if not num_res.get("success"):
            result["failed"] += 1
            continue
        layout_res = apply_mathtype_right_number_layout(document, sh)
        if layout_res.get("success"):
            result["inserted"] += 1
        else:
            # Number field exists but layout incomplete — still count as inserted,
            # but record layout failure message on the result.
            result["inserted"] += 1
            result["layoutFailed"] = int(result.get("layoutFailed", 0)) + 1
            result["lastLayoutError"] = layout_res.get("message")
    if result["inserted"] > 0:
        update_mathtype_eqn_numbers(word, document)
    return result


def try_end_toggle_one(
    word,
    document,
    abs_start: int,
    abs_end: int,
    latex: str,
    *,
    display: bool,
    add_number: bool = False,
) -> dict:
    """Convert a single known span via end-toggle paste-back."""
    if add_number:
        ensure_mathtype_eqnnum_defaults(on_right=True)
    shape = tex_toggle_at_end(word, document, latex, display=display)
    if shape is None:
        return {"success": False, "message": "End-toggle did not create OLE."}
    try:
        pasted = paste_ole_over_range(document, abs_start, abs_end, shape, display=display)
        result = {"success": True, "message": "End-toggle + paste-back succeeded.", "eqnNumber": False}
        if add_number and display:
            target_shape = pasted or _find_shape_near(document, abs_start)
            if target_shape is not None:
                num_res = insert_mathtype_right_eqnnum(word, document, target_shape)
                result["eqnNumber"] = bool(num_res.get("success"))
                result["message"] += " " + str(num_res.get("message", ""))
                if result["eqnNumber"]:
                    update_mathtype_eqn_numbers(word, document)
            else:
                result["message"] += " MathType eqn number skipped (OLE not found)."
        return result
    except Exception as exc:  # noqa: BLE001
        try:
            shape.Delete()
        except Exception:  # noqa: BLE001
            pass
        _delete_from_marker_to_end(document)
        return {"success": False, "message": str(exc)}
