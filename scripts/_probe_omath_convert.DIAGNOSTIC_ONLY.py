# -*- coding: utf-8 -*-
from __future__ import annotations

import ctypes
import shutil
import subprocess
import tempfile
import time
import winreg
from ctypes import wintypes
from pathlib import Path

import pythoncom
import win32com.client.dynamic as D

OUT = Path(tempfile.mkdtemp(prefix="wlmt_probe_"))
LOG = OUT / "probe_log.txt"
SRC = OUT / "omath.docx"
WORD = Path(r"C:\Program Files\Microsoft Office\root\Office16\WINWORD.EXE")
user32 = ctypes.windll.user32
EnumWindowsProc = ctypes.WINFUNCTYPE(ctypes.c_bool, wintypes.HWND, wintypes.LPARAM)
BM_CLICK = 0x00F5
lines: list[str] = []


def log(msg: str) -> None:
    lines.append(msg)
    print(msg)


def kill() -> None:
    for n in ("MathType.exe", "WINWORD.EXE"):
        subprocess.run(["taskkill", "/IM", n, "/F"], check=False, capture_output=True)
    time.sleep(1.2)


def title(h: int) -> str:
    n = user32.GetWindowTextLengthW(h)
    b = ctypes.create_unicode_buffer(n + 1)
    user32.GetWindowTextW(h, b, n + 1)
    return b.value


def windows():
    out = []

    def cb(h, _):
        if user32.IsWindowVisible(h):
            t = title(h)
            if t:
                out.append((h, t))
        return True

    user32.EnumWindows(EnumWindowsProc(cb), 0)
    return out


def buttons(hwnd: int):
    child = None
    while True:
        child = user32.FindWindowExW(hwnd, child, None, None)
        if not child:
            break
        yield child, title(child)
        sub = None
        while True:
            sub = user32.FindWindowExW(child, sub, None, None)
            if not sub:
                break
            yield sub, title(sub)


def main() -> int:
    kill()
    key = winreg.CreateKey(winreg.HKEY_CURRENT_USER, r"Software\Design Science\DSMT7\WordCommands")
    winreg.SetValueEx(key, "ConvertFrom", 0, winreg.REG_DWORD, 8)
    winreg.SetValueEx(key, "ConvertTo", 0, winreg.REG_DWORD, 0)
    winreg.SetValueEx(key, "ConvertTranslator", 0, winreg.REG_DWORD, 0)
    winreg.SetValueEx(key, "ConvertMisc", 0, winreg.REG_DWORD, 0)
    winreg.CloseKey(key)

    pythoncom.CoInitialize()
    w = D.Dispatch("Word.Application")
    w.Visible = False
    # Keep alerts for Convert dialog path
    w.DisplayAlerts = -1
    doc = w.Documents.Add()
    doc.Content.Text = "E=mc^2"
    doc.OMaths.Add(doc.Range(0, 6))
    doc.OMaths.BuildUp()
    if SRC.exists():
        SRC.unlink()
    doc.SaveAs2(str(SRC))
    log(f"created omath={doc.OMaths.Count}")
    doc.Close(False)
    w.Quit()
    kill()

    subprocess.Popen([str(WORD), str(SRC)], close_fds=True)
    time.sleep(9)
    pythoncom.CoInitialize()
    w = D.Dispatch("Word.Application")
    for i in range(1, w.AddIns.Count + 1):
        a = w.AddIns(i)
        log(f"addin {a.Name} installed={a.Installed}")
        if "MathType" in a.Name and not a.Installed:
            a.Installed = True
            log(f"enabled {a.Name}")
    try:
        w.Visible = True
    except Exception as e:
        log(f"visible err {e}")

    doc = w.ActiveDocument
    log(f"before omath={doc.OMaths.Count} shapes={doc.InlineShapes.Count}")

    # Warm-load MathType implementation via TeXToggle once.
    try:
        w.Run("MTCommand_TeXToggle")
        log("TeXToggle ok")
    except Exception as e:
        log(f"TeXToggle fail {e}")
    time.sleep(2)

    opened = False
    for attempt in range(10):
        try:
            w.Run("MTCommand_OnConvertEquations")
            log(f"Convert opened attempt={attempt}")
            opened = True
            break
        except Exception as e:
            log(f"Convert fail attempt={attempt} {e}")
            time.sleep(1)

    if opened:
        deadline = time.time() + 25
        while time.time() < deadline:
            for h, t in windows():
                if any(k in t for k in ("Convert", "转换")):
                    log(f"dialog {t}")
                    for bh, bt in buttons(h):
                        if bt and any(k in bt for k in ("Convert", "转换", "确定", "OK")):
                            user32.SendMessageW(bh, BM_CLICK, 0, 0)
                            log(f"clicked {bt}")
                            deadline = time.time() + 8
            time.sleep(0.4)
        time.sleep(3)
        for h, t in windows():
            for bh, bt in buttons(h):
                if bt in ("OK", "确定", "Close", "关闭"):
                    user32.SendMessageW(bh, BM_CLICK, 0, 0)

    log(f"after omath={doc.OMaths.Count} shapes={doc.InlineShapes.Count}")
    for i in range(1, doc.InlineShapes.Count + 1):
        sh = doc.InlineShapes(i)
        try:
            log(f"shape{i} {sh.OLEFormat.ClassType}")
        except Exception as e:
            log(f"shape{i} {e}")

    try:
        doc.Close(False)
        w.Quit()
    except Exception:
        pass
    kill()
    LOG.write_text("\n".join(lines), encoding="utf-8")
    log(f"log={LOG}")
    # Always remove probe temp dir (log already printed above via lines).
    shutil.rmtree(OUT, ignore_errors=True)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
