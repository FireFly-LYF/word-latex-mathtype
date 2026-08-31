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

### 源稿约定（重要）

转换器按以下优先级识别公式：

1. **整段 display**：`$$...$$` / `\[...\]` 或整段可识别 LaTeX
2. **显式行内定界**：`$...$`（作者标记的数学，**一律转换**，含 `$t$`、`$N$` 等简单变量）
3. **无定界启发式**：`INLINE_PATTERNS`（切点、下标等）；仅用于未包 `$` 的遗留文本

推荐：行内公式一律写成 `$...$`，独立公式用 `$$...$$` 单独成段。不要只写裸 `t_{\mathrm{on}}` 再依赖启发式。

### 已知失败模式与修复（2026-08）

| 现象 | 根因 | 修复 |
|------|------|------|
| 动作空间出现空的 `$$`、残片 `$[,1)$` | 旧版只匹配定界符**内部**的启发式模式，转换后留下两侧 `$`；且行内未清相邻 `$` | 优先整段匹配 `$...$`；成功后对 display/inline 都调用 `remove_adjacent_literal_dollars` |
| `$t$`、`$B$`、`$N$` 仍留在正文 | `should_skip_inline_match` 把简单标识符当成非公式跳过 | 对 `delimited=True`（来自 `$...$`）的匹配不再 skip |
| `u_m=0$。` 被吞进公式 | 赋值启发式模式把 `$`/标点吃进匹配 | 收紧赋值模式字符类，排除 `$` 与 CJK 标点 |
| 重跑后半段仍是 LaTeX | 段内已有 OLE 时整段 `skipped-existing-object` | **始终从干净 LaTeX 源稿**生成新 target，勿对半成品 `_formula.docx` 再转 |

### Agent 执行清单

1. 确认源/目标 docx 均未打开（无 `~$` 锁）
2. `--source` = 含 `$...$` / `$$...$$` 的干净稿；`--target` 必须不同于 source
3. 用本 skill 目录下的 `scripts/convert_docx_latex_to_equations.py`（勿用过时副本）
4. 检查日志：`failed=0`；抽查 1.1 方法节 `dollar=False` 且 display 段有 OLE
5. 若曾用旧脚本得到破损 `_formula.docx`，删除后从源稿重转

### 限制

- 仅 Windows；需 MathType 加载项
- 勿用 Insert GUI 批量转换
- 复杂中文混排可用 legacy 或手工

## English Guide

Headless `.docx` LaTeX → MathType OLE via `MTCommand_TeXToggle`. No MathType GUI.

```powershell
python scripts/convert_docx_latex_to_equations.py --source input.docx
```

**Matching order:** display `$$...$$` → explicit inline `$...$` (always convert) → bare `INLINE_PATTERNS`. Always convert from a clean LaTeX source; paragraphs that already contain OLE are skipped entirely.

Fallback: `--mathtype-mode word`. Legacy CJK: `-UseLegacy` on `.ps1`.
