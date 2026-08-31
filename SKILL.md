---
name: word-latex-mathtype
description: "Convert LaTeX text in Word .docx files into MathType OLE objects on Windows via headless Word TeX toggle (no MathType GUI). Use for docx formula conversion, batch LaTeX-to-MathType, or Word OMath fallback. 在 Windows 上通过 Word 后台 TeX Toggle 将 .docx 中的 LaTeX 转为 MathType OLE（不弹出 MathType 窗口），支持批量转换与 Word 原生公式回退。"
---

# Word LaTeX MathType

## 中文说明

### 概述

把 Word `.docx` 里的 LaTeX 公式文本批量转为 **可双击编辑的 MathType OLE 对象**（`Equation.DSMT4`，DSMT7 载荷）。

**默认：Python + win32com + Word 后台 TeX Toggle**

1. `WINWORD.EXE docx /q` 静默打开（Word 不可见）
2. `win32com.client.dynamic.Dispatch` 附着 COM
3. LaTeX 包裹为 `$...$` / `$$...$$`，调用 `MTCommand_TeXToggle`
4. **不打开 MathType GUI**（无广告、适合批量）
5. 失败时回退 Word OMath

**禁止**批量使用 `InsertInlineEqn` GUI——会弹 MathType 窗口且无法可靠自动关闭。

### 何时使用

- Word / docx / LaTeX / MathType 公式批量转换
- 需要后台、无人值守
- 需要 JSON 日志

### 前置条件

- Windows + Word + MathType 加载项
- Python 3 + pywin32
- 转换前关闭目标 docx

### 命令

```powershell
powershell -ExecutionPolicy Bypass -File scripts/setup_environment.ps1
```

```powershell
python scripts/convert_docx_latex_to_equations.py `
  --source "C:\path\to\input.docx" `
  --target "C:\path\to\input_formula.docx" `
  --log "C:\path\to\input_formula_log.json"
```

Word OMath 回退：`--mathtype-mode word`

### 脚本

- `convert_docx_latex_to_equations.py` — **默认主路径**
- `convert_docx_latex_to_equations.ps1` — 包装器
- `convert_docx_latex_to_equations_legacy.ps1` — 旧 PS1（CJK `\text{}` 边缘场景）

### 限制

- 仅 Windows；需 MathType 加载项
- 勿用 Insert GUI 批量转换
- 复杂中文混排可用 legacy 或手工

## English Guide

Headless `.docx` LaTeX → MathType OLE via `MTCommand_TeXToggle`. No MathType GUI.

```powershell
python scripts/convert_docx_latex_to_equations.py --source input.docx
```

Fallback: `--mathtype-mode word`. Legacy CJK: `-UseLegacy` on `.ps1`.
