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

**禁止**把 Word 公式面板 / Unicode 数学（如 `r(t)=αs sT(t−τs)+…`）当作源。TeX Toggle 只吃 LaTeX 文本，例如：

```text
$$r(t)=\alpha_s\,s_T(t-\tau_s)+\alpha_j\,j(t)+n(t)$$
```

### 已知失败模式与修复

| 现象 | 根因 | 修复 |
|------|------|------|
| 动作空间空的 `$$`、残片 `$[,1)$` | 旧版只匹配 `$` 内部启发式；行内未清相邻 `$` | 优先整段 `$...$`；成功后剥相邻 `$` |
| `$t$` / `$N$` 残留 | 简单标识符被 skip | `delimited=True` 不再 skip |
| `u_m=0$。` 被吞 | 赋值启发式过宽 | 收紧字符类，排除 `$`/CJK 标点 |
| display=0，段内是 `α`/`τ` Unicode | 粘贴的 Unicode 公式无 `\`/`$`，不被识别 | 先改写成 `$$...$$` LaTeX；日志 `skipped-unicode-math` |
| 只转了新段、旧段大量 skipped | 源稿已合并过 OLE（把 `_formula` 写回主稿） | 主稿保持 LaTeX；OLE 只留在 `_formula.docx` |
| 改稿脚本失败仍跑了转换 | Agent 未先校验源稿就 convert | **先 verify 源段含 `$$`/`$`，再 convert** |
| 重跑半成品无效 | 有 OLE 的段整段跳过 | 始终从干净 LaTeX 源生成新 target |

### Agent 执行清单（门禁）

1. 确认源/目标 docx 均未打开（无 `~$` 锁）
2. **改完源稿后先抽查**：目标小节段落文本必须含 `$$...$$` 或 `$...$`，不能是 Unicode `α`/`τ`，也不能已是 OLE 空壳
3. 改稿脚本若失败（SyntaxError / 编码错误）：**停止**，修好并重跑改稿，禁止接着 convert
4. `--source` = 干净 LaTeX 主稿；`--target` ≠ source（通常 `*_formula.docx`）
5. 只用本 skill 目录下 `scripts/convert_docx_latex_to_equations.py`
6. 看日志：`failed=0`；关注 `unicodeSuspects` / `skipped-unicode-math` / `existingOLE`
7. 抽查输出：目标节 `dollar=False` 且 display 段 `ole≥1`

### Windows / 工具注意

- 写含 `$` 的 Python/LaTeX 时，勿用 PowerShell 双引号或 `python -c "...$r(t)..."`（`$` 会被展开）。优先：`.py` 文件 + `python path.py`，或 `@'...'@ | Set-Content`
- 含中文的 `.py` 若 Write 工具编码损坏，改用 `\uXXXX` 转义或 PowerShell `@' '@` 写入 UTF-8

### 限制

- 仅 Windows；需 MathType 加载项
- 勿用 Insert GUI 批量转换
- 复杂中文混排可用 legacy 或手工

## English Guide

Headless `.docx` LaTeX → MathType OLE via `MTCommand_TeXToggle`. No MathType GUI.

```powershell
python scripts/convert_docx_latex_to_equations.py --source input.docx
```

**Matching order:** display `$$...$$` → explicit inline `$...$` (always convert) → bare `INLINE_PATTERNS`.

**Gate:** After editing the source, verify paragraphs contain real LaTeX (`$$` / `$` / `\alpha`), then convert. Unicode-pasted math is logged as `skipped-unicode-math`. Never convert a document whose body already stores OLE for the sections you still need to change—keep LaTeX in the source and OLE only in `_formula.docx`.

Fallback: `--mathtype-mode word`. Legacy CJK: `-UseLegacy` on `.ps1`.
