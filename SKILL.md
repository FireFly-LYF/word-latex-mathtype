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
| **`failed=0` 但新段公式仍是 `$...$` 文本** | 段内**残留 OLE**（只改了 `run.text` 未删 `w:object`）；或 `display=0 inline=0` 全篇已是 OLE 壳 | 改稿用 **`clear_para_content` + 新 run + 字体/段落格式**（见下）；日志查 `skipped-existing-object` 是否落在你改的小节 |
| **改完 1.3 节字体变 Calibri / 无缩进** | 用 `OxmlElement('w:r')` 或只写 `runs[0].text` 未恢复 **w:rPr / w:pPr** | 改稿后必须 **`set_run_fonts` + `set_body_para_format`**；display 段用 `set_display_para_format` |
| 动作空间整段跳过、末尾残 `0,1` | 该段已有 InlineShapes，转换器整段 skip | 改稿前 **`para_has_embedded`** 必须为 False |
| 把 `_formula.docx` 存回主稿 | 主稿混入 OLE，后续无法再 TeX Toggle | **主稿永远 LaTeX**；OLE 只写 `*_formula.docx` |

### 改稿 + 转换标准流程（Agent 必遵）

**顺序不可颠倒：改 LaTeX 源 → 校验字体/OLE/定界符 → convert → 校验日志小节范围。**

1. **只改 `--source` 主稿**（`*_formula.docx` 只读预览/投稿，不作编辑源）
2. **重写段落**时用 `scripts/docx_latex_helpers.py`（或等价逻辑）：
   - `clear_para_content(para)` — 删 run/object，**保留 `w:pPr`**
   - `para.add_run(latex_text)` + `set_run_fonts(run)` — 恢复中西文字体（默认：宋体 + Times New Roman + 五号）
   - `set_body_para_format(para)` 或 `set_display_para_format(para)` — 恢复行距/对齐/首行缩进
   - **禁止**仅 `runs[0].text = ...` 或裸 `OxmlElement('w:t')` 无 rFonts
3. **转换前门禁**（源稿）：
   - 目标小节 `para_has_embedded(para)` 全为 False
   - 含公式段有 `$...$` / `$$...$$`（Unicode `α`/`τ` 不算）
   - 抽查 run 的 `w:eastAsia` 与相邻正文段一致
4. **转换**：`--source` ≠ `--target`；target 每次由脚本 `copy2` 覆盖
5. **转换后门禁**（日志 + 输出）：
   - `summary.failed == 0` **不够**
   - 改稿小节对应 **Word 1-based 段号**内无 `skipped-existing-object`
   - `inlineConverted` / `displayConverted` 与改稿量匹配（全 0 表示源稿已是 OLE 壳或未识别 LaTeX）
   - 公式版该节 `ole≥1`、正文 `$` 已消失

```python
# 最小示例（项目内可 from docx_latex_helpers import ...）
from docx_latex_helpers import (
    set_para_latex_text,
    para_has_embedded,
    verify_paragraphs_latex,
    verify_log_section,
)

set_para_latex_text(para, r"动作 … $(t_{\mathrm{on}},t_{\mathrm{off}})$ …", display=False)
set_para_latex_text(eq_para, r"$$r_t = w_{\mathrm{pd}}\,\Delta P_{d,t} + \cdots$$", display=True)
assert not para_has_embedded(para)
issues = verify_log_section(log_path, word_para_lo=40, word_para_hi=46, min_inline=5)
```

期刊模板字体/行距若不同，改 `docx_latex_helpers.py` 顶部常量，或从相邻未改段落复制 `w:pPr` 再写 run。

### Agent 执行清单（门禁）

1. 确认源/目标 docx 均未打开（无 `~$` 锁）
2. **改完源稿后先抽查**：目标小节段落文本必须含 `$$...$$` 或 `$...$`，不能是 Unicode `α`/`τ`，也不能已是 OLE 空壳
3. 改稿脚本若失败（SyntaxError / 编码错误）：**停止**，修好并重跑改稿，禁止接着 convert
4. `--source` = 干净 LaTeX 主稿；`--target` ≠ source（通常 `*_formula.docx`）
5. 只用本 skill 目录下 `scripts/convert_docx_latex_to_equations.py`
6. 看日志：`failed=0`；**且**改稿小节无 `skipped-existing-object`；**且** `inlineConverted`/`displayConverted` 非异常为 0
7. 抽查输出：目标节 `dollar=False` 且 display 段 `ole≥1`；**中文 run 仍为模板字体**（非 Calibri 默认）

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

**Gate:** After editing the source, verify paragraphs contain real LaTeX (`$$` / `$` / `\alpha`), **no stale OLE**, and **restored run/paragraph fonts** (use `docx_latex_helpers.py`). Then convert. Log must show `failed=0` **and** no `skipped-existing-object` in the edited section's Word paragraph index range—not `failed=0` alone when `inlineConverted=0`.

Fallback: `--mathtype-mode word`. Legacy CJK: `-UseLegacy` on `.ps1`.
