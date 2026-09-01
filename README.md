# word-latex-mathtype

[![License: MIT](https://img.shields.io/badge/License-MIT-blue.svg)](LICENSE)
[![Platform: Windows](https://img.shields.io/badge/Platform-Windows-0078D6)](https://www.microsoft.com/windows)
[![Python 3](https://img.shields.io/badge/Python-3.10%2B-3776AB)](https://www.python.org/)

**Headless batch conversion** of LaTeX formula text in Word `.docx` files to **MathType OLE** objects on Windows — without opening the MathType GUI.

在 Windows 上把 Word 文档里的 LaTeX 公式文本**批量、后台**转换为 **MathType 可编辑对象**，不弹出 MathType 窗口（无广告干扰），适合论文/报告自动化排版。

---

## Features

- Silent Word automation: `WINWORD.EXE docx /q` + Python `win32com`
- MathType via **TeX Toggle** (`MTCommand_TeXToggle`) — no GUI, no manual close
- Produces `Equation.DSMT4` OLE with **DSMT7** payload (compatible with manual MathType equations)
- Display + inline LaTeX detection: prefers explicit `$...$` / `$$...$$`, then conservative bare heuristics
- Strips leftover `$` delimiters after TeX toggle (inline and display)
- Automatic fallback to **Word native OMath** when MathType toggle fails
- JSON conversion log for auditing
- Optional **Cursor Agent Skill** (`SKILL.md`) for AI-assisted workflows

## Requirements

| Component | Required |
|-----------|----------|
| Windows 10/11 | Yes |
| Microsoft Word | Yes |
| MathType Word add-in | Yes (for `--mathtype-mode toggle`) |
| Python 3.10+ | Yes |
| pywin32 | Yes (`pip install pywin32`) |

> **Do not** use `InsertInlineEqn` / `InsertDispEqn` for batch jobs — they open interactive MathType windows (ads, dialogs) and break unattended automation.

## Quick start

```powershell
git clone https://github.com/FireFly-LYF/word-latex-mathtype.git
cd word-latex-mathtype

# Install Python dependencies & verify Word
powershell -ExecutionPolicy Bypass -File .\scripts\setup_environment.ps1

# Convert (input.docx must contain LaTeX-like text, e.g. P_d, t_{\mathrm{on}})
python .\scripts\convert_docx_latex_to_equations.py `
  --source "C:\path\to\input.docx" `
  --target "C:\path\to\input_formula.docx" `
  --log "C:\path\to\input_formula_log.json"
```

PowerShell wrapper (same behavior):

```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\convert_docx_latex_to_equations.ps1 `
  -SourcePath "C:\path\to\input.docx"
```

Word OMath only (no MathType):

```powershell
python .\scripts\convert_docx_latex_to_equations.py --source input.docx --mathtype-mode word
```

## How it works

```
input.docx (LaTeX text in paragraphs)
        │
        ▼
  WINWORD.EXE /q          ← Word runs headless
        │
        ▼
  Python win32com dynamic   ← avoids PowerShell TYPE_E_CANTLOADLIBRARY
        │
        ▼
  Wrap: $inline$ / $$display$$
        │
        ▼
  MTCommand_TeXToggle       ← MathType OLE, no GUI
        │
        ├─ success → Equation.DSMT4 OLE
        └─ failure → Word OMath fallback
        │
        ▼
  output_formula.docx + output_formula_log.json
```

## CLI reference

| Argument | Default | Description |
|----------|---------|-------------|
| `--source` | (required) | Source `.docx` with LaTeX text |
| `--target` | `<source>_formula.docx` | Output path |
| `--log` | `<source>_formula_log.json` | JSON log |
| `--mathtype-mode` | `toggle` | `toggle` = MathType OLE; `word` = OMath only |
| `--word-exe` | auto-detect | Path to `WINWORD.EXE` |
| `--open-wait-sec` | `5` | Seconds to wait after CLI open |

## Log summary fields

| Field / status | Meaning |
|----------------|---------|
| `engine: python-win32com-headless-tex-toggle` | Normal headless path |
| `display-mathtype` / `inline-mathtype` | MathType conversion OK |
| `word-omath` | Fallback to Word native equation |
| `skipped-existing-object` | Paragraph already has OLE/OMath |
| `skipped-unicode-math` | Pasted Unicode equation (α/τ/…), not LaTeX — rewrite as `$$...$$` |
| `failed` | Both MathType and OMath failed |
| `unicodeMathSuspects` / `existingObjectParagraphs` | Summary counters for the above skips |

## Project layout

```text
word-latex-mathtype/
├── LICENSE
├── README.md                 ← you are here
├── SKILL.md                  ← Cursor Agent Skill instructions
├── agents/openai.yaml
└── scripts/
    ├── setup_environment.ps1
    ├── convert_docx_latex_to_equations.py      # default converter
    ├── convert_docx_latex_to_equations.ps1     # PS wrapper
    └── convert_docx_latex_to_equations_legacy.ps1  # legacy PS COM (CJK edge cases)
```

## Use as a Cursor Skill

Copy or clone this repo into your Cursor skills folder, e.g.:

```text
%USERPROFILE%\.cursor\skills\word-latex-mathtype\
```

Cursor will load `SKILL.md` when you ask to convert docx formulas to MathType.

## Source conventions

Prefer author-delimited math so detection is unambiguous:

| Kind | Write as |
|------|----------|
| Display | Own paragraph: `$$...$$` |
| Inline | `$...$` (including simple vars like `$t$`, `$N$`) |

Matching order inside the converter:

1. Display paragraphs (`$$...$$` / `\[...\]` / whole-paragraph LaTeX)
2. Explicit `$...$` spans — **always convert** (never treated as plain identifiers)
3. Bare heuristic `INLINE_PATTERNS` (legacy undelimited text only)

Always convert from a **clean LaTeX source** into a new `--target`. Paragraphs that already contain OLE/OMath are skipped as a whole (`skipped-existing-object`), so re-running on a half-converted `_formula.docx` will not finish remaining `$...$` text.

## FAQ

**PowerShell `New-Object Word.Application` fails with `TYPE_E_CANTLOADLIBRARY`**

Use the default **Python** path. Do not use `-UseLegacy` unless you need legacy CJK `\text{}` splitting.

**Target document is locked**

Close the output `.docx` in Word before re-running.

**Leftover `$$` shells or fragments like `$[,1)$`**

Older builds matched heuristic patterns *inside* `$...$` and left the dollars behind; inline mode also skipped stripping adjacent `$`. Current script matches full `$...$` spans first and strips adjacent `$` after every successful toggle. Re-convert from the clean source.

**Simple `$t$` / `$N$` left as literal dollars**

Author-delimited `$...$` must not be skipped as ASCII identifiers. Current script marks those matches `delimited=True` and always converts them.

**`display=0` but you see an equation that looks like `r(t)=α…`**

That is Unicode/Word-equation paste, not LaTeX. TeX Toggle will not convert it. Rewrite as `$$r(t)=\alpha_s\,s_T(t-\tau_s)+\\ldots$$`. The converter logs `skipped-unicode-math` and prints `unicodeSuspects` in the summary.

**Most paragraphs skipped / only a few new formulas convert**

The source was probably merged with a previous `_formula.docx` (OLE already present). Keep the editable master as LaTeX-only; write MathType output only to `*_formula.docx`.

**Formulas not detected**

Prefer wrapping with `$...$` / `$$...$$`. Undelimited text still needs LaTeX markers (`_`, `\mathrm`, `{`, etc.). Plain `Pd` without subscripts is intentionally skipped by the bare heuristic path.

**Complex Chinese `\text{...}` display formulas**

Try `-UseLegacy` on the `.ps1` wrapper, or edit manually.

## Legacy PowerShell path

For display formulas with mixed `\text{中文}` that need text/equation splitting:

```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\convert_docx_latex_to_equations.ps1 `
  -SourcePath input.docx -UseLegacy
```

## License

[MIT](LICENSE) — free to use, modify, and distribute.

## Contributing

Issues and pull requests welcome at [github.com/FireFly-LYF/word-latex-mathtype](https://github.com/FireFly-LYF/word-latex-mathtype).

---

## 中文简要说明

本工具通过 Word 后台 **TeX Toggle** 将 docx 中的 LaTeX 文本转为 MathType OLE，**不打开 MathType 界面**，适合批量处理论文公式。安装依赖后运行 `convert_docx_latex_to_equations.py` 即可。开源协议：MIT。
