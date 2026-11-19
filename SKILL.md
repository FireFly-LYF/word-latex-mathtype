---
name: word-latex-mathtype
description: "Convert LaTeX, Word/WPS OMath, or mixed LaTeX+OMath in .docx to MathType OLE on Windows (headless TeX toggle). Supports independent or joint conversion of inline vs display (numbered) formulas via --source-mode and --formula-scope. 将纯LaTeX、纯Word公式或混排转为MathType；段内/单行可分可合。"
---

# Word LaTeX MathType

## 中文说明

### 概述

把 Word `.docx` 里的公式批量转为 **可双击编辑的 MathType OLE 对象**（`Equation.DSMT4`，DSMT7 载荷）。

**源格式（`--source-mode`）**

| 模式 | 含义 |
|------|------|
| `latex` | 纯 LaTeX（`$...$` / `$$...$$`） |
| `omath` | 纯 Word/WPS 公式（OMML） |
| `auto` | LaTeX + Word 公式混排：先改写 OMath 再统一转 |

**公式范围（`--formula-scope`）** — 可分可合

| 范围 | 含义 | MathType 右编号 |
|------|------|----------------|
| `all`（默认） | 段内 + 单行一起转 | 单行默认开 |
| `inline` | 仅段内 `$...$` | 否 |
| `display` | 仅单行 `$$...$$` | 默认开 |

**默认：Python + win32com + Word 后台 TeX Toggle**

1. `WINWORD.EXE docx /q` 静默打开（Word 不可见）
2. 文末 `MTCommand_TeXToggle` + OLE 回贴（不打开 MathType GUI）
3. 单行：`InsertEqnNum` + `TAB+OLE+TAB+(n)` 右编号版式

**禁止**批量使用 `InsertInlineEqn` GUI。

### 何时使用

- Word / docx / LaTeX / MathType 公式批量转换
- **Word / WPS 公式编辑器对象 → MathType**
- 需要后台、无人值守
- 需要 JSON 日志



### 前置条件

- Windows + Word + MathType 加载项
- Python 3 + pywin32
- OMath 路径额外需要：`lxml`、`pandoc`、Office `OMML2MML.XSL`
- 转换前关闭目标 docx



### 命令

```powershell
powershell -ExecutionPolicy Bypass -File scripts/setup_environment.ps1
```

统一主入口：

```powershell
# 纯 LaTeX，全部
python scripts/convert_docx_latex_to_equations.py --source in.docx --source-mode latex --formula-scope all

# 纯 Word 公式，仅单行（右编号）
python scripts/convert_docx_latex_to_equations.py --source in.docx --source-mode omath --formula-scope display

# 混排 LaTeX+Word
python scripts/convert_docx_latex_to_equations.py --source in.docx --source-mode auto --formula-scope all

# 仅段内
python scripts/convert_docx_latex_to_equations.py --source in.docx --source-mode auto --formula-scope inline
```

OMath 专用脚本（等价 `--source-mode omath`）：

```powershell
python scripts/convert_docx_omath_to_mathtype.py --source in.docx --formula-scope display
```

开关：

- `--source-mode latex|omath|auto`
- `--formula-scope all|inline|display`
- `--toggle-strategy auto|in-place|end-paste`（`auto` 源模式强制 `end-paste`；**带右编号的 display 优先** `--toggle-strategy end-paste`。`in-place` 结束时也会补跑 `insert_mathtype_numbers_for_lone_displays`，但仍建议 end-paste，因其与 `TAB+OLE+TAB+(n)` 布局同路径）
- `--fallback none|word`（默认 `none`）
- `--mathtype-numbers` / `--no-mathtype-numbers`（仅对单行 display 生效）
- 右编号（视觉硬标准）：段左对齐 + 中制表位(`usable/2`, Center) + 右制表位(`usable`, Right)；段内结构为 `TAB + 公式OLE + TAB + (n)`，**公式视觉居中、编号贴右页/栏边**。`InsertEqnNum`  alone 不够，必须再跑 `apply_mathtype_right_number_layout`。**双栏**时 `usable` = 栏宽（非整页宽）

验收：

```powershell
python scripts/verify_formula_docx.py --docx output.docx --min-dsmt4 1 --max-omath 0 --max-dollar-pairs 0 --min-mteqn 1
```

Word OMath 回退（不做 MathType）：`--mathtype-mode word`

### 脚本

- `convert_docx_latex_to_equations.py` — **默认主路径**（`--source-mode latex|omath|auto`）
- `convert_docx_omath_to_mathtype.py` — Word/WPS OMath → MathType
- `omath_to_latex.py` — OMML→MathML→LaTeX 辅助
- `word_com_utils.py` — Word PID 管理、文末 Toggle+OLE 回贴
- `verify_formula_docx.py` — 转换后验收门禁
- `convert_docx_latex_to_equations.ps1` — 包装器
- `convert_docx_latex_to_equations_legacy.ps1` — 旧 PS1（CJK `\text{}` 边缘场景）



### 源稿约定（重要）

转换器按以下优先级识别公式：

1. **整段 display**：`$$...$$` / `\[...\]` 或整段可识别 LaTeX
2. **显式行内定界**：`$...$`（作者标记的数学，**一律转换**，含 `$t$`、`$N$`、`$i$` 等简单变量；即使段内没有任何 `_`/`\`/`=`，`has_formula_signature` / `has_delimited_math` 也必须命中）
3. **无定界启发式**：`INLINE_PATTERNS`（切点、下标等）；仅用于未包 `$` 的遗留文本

推荐：行内公式一律写成 `$...$`，独立公式用 `$$...$$` 单独成段。不要只写裸 `t_{\mathrm{on}}` 再依赖启发式。

**禁止**把 Word 公式面板 / Unicode 数学（如 `r(t)=αs sT(t−τs)+…`）当作源。TeX Toggle 只吃 LaTeX 文本，例如：

```text
$$r(t)=\alpha_s\,s_T(t-\tau_s)+\alpha_j\,j(t)+n(t)$$
```



### 已知失败模式与修复


| 现象                                    | 根因                                                                           | 修复                                                                                         |
| ------------------------------------- | ---------------------------------------------------------------------------- | ------------------------------------------------------------------------------------------ |
| 动作空间空的 `$$`、残片 `$[,1)$`               | 旧版只匹配 `$` 内部启发式；行内未清相邻 `$`                                                   | 优先整段 `$...$`；成功后剥相邻 `$`                                                                    |
| `$t$` / `$N$` 残留（段内已有 `_`/`\` 等签名） | 旧版 `should_skip_inline_match` 把简单标识符当普通词                                  | `delimited=True` 不再 skip                                                                   |
| **整段仅** `$i$`/`$m$` **或中文+$i$ 未转、无 log** | `has_formula_signature` 不认裸 `$...$`，主循环静默 `continue`（`skipped=0`）          | `has_delimited_math`：显式 `$...$`/`$$` 一律算签名；有定界匹配时不因签名失败跳过                         |
| `u_m=0$。` 被吞                          | 赋值启发式过宽                                                                      | 收紧字符类，排除 `$`/CJK 标点                                                                        |
| display=0，段内是 `α`/`τ` Unicode         | 粘贴的 Unicode 公式无 `\`/`$`，不被识别                                                 | 先改写成 `$$...$$` LaTeX；日志 `skipped-unicode-math`                                             |
| 只转了新段、旧段大量 skipped                    | 源稿已合并过 OLE（把 `_formula` 写回主稿）                                                | 主稿保持 LaTeX；OLE 只留在 `_formula.docx`                                                         |
| 改稿脚本失败仍跑了转换                           | Agent 未先校验源稿就 convert                                                        | **先 verify 源段含** `$$`**/**`$`**，再 convert**                                                |
| 重跑半成品无效                               | 有 OLE 的段整段跳过                                                                 | 始终从干净 LaTeX 源生成新 target                                                                    |
| `failed=0` **但新段公式仍是** `$...$` **文本** | 段内**残留 OLE**（只改了 `run.text` 未删 `w:object`）；或 `display=0 inline=0` 全篇已是 OLE 壳 | 改稿用 `clear_para_content` **+ 新 run + 字体/段落格式**（见下）；日志查 `skipped-existing-object` 是否落在你改的小节 |
| **改完 1.3 节字体变 Calibri / 无缩进**         | 用 `OxmlElement('w:r')` 或只写 `runs[0].text` 未恢复 **w:rPr / w:pPr**              | 改稿后必须 `set_run_fonts` **+** `set_body_para_format`；display 段用 `set_display_para_format`    |
| 动作空间整段跳过、末尾残 `0,1`                    | 该段已有 InlineShapes，转换器整段 skip                                                 | 改稿前 `para_has_embedded` 必须为 False                                                          |
| **display 行距过紧 / 公式被裁切** | display 段误用正文 `EXACTLY 15.6pt` | **只改公式行**：`set_display_para_format`→`MULTIPLE 1.5`；`set_body_para_format` 保持正文原固定行距 |
| **公式旁残留红字** `\lvert`/`\rvert`/`\dfrac` | MathType TeX Toggle 不完全支持这些宏 | 源稿改写为 `\left|...\right|`、`\frac`；转换后肉眼查 PNG |
| **`mathtypeNumbers=true` 但 eqnNumbers=0** | 旧版 in-place 不跑编号；或 end-paste 后 lone-display 因字段符误判失败 | 优先 `--toggle-strategy end-paste`；convert 结束若仍为 0 再跑编号 pass；字段符须计入 lone-display |
| **end-paste 后 OLE 旁残留** `$$...$$` | `Range.Paste()` 未删除选区原文 | 先清空 Range 再粘贴；display 段 `_scrub_residual_dollar_text` |
| **in-place 转换后无右编号** | 编号只在 end-paste 的 lone-display pass | convert 结束若 `eqnNumbersInserted==0` 再跑 `insert_mathtype_numbers_for_lone_displays` |
| **编号贴在公式旁、公式不居中** | `InsertEqnNum` 后段内出现字段符，`_is_lone_display_equation` 误判并跳过 TAB 布局 | 字段符不算正文；强制 `TAB+OLE+TAB+(n)` + center/right 制表位；段左对齐 |
| **转换后文末几十上百空段** | end-toggle helper 段清理不净 | 转换后删除文末连续空段；抽查 `Paragraphs.Count` 与源稿接近 |
| 把 `_formula.docx` 存回主稿                | 主稿混入 OLE，后续无法再 TeX Toggle                                                    | **主稿永远 LaTeX**；OLE 只写 `*_formula.docx`                                                     |




### 改稿 + 转换标准流程（Agent 必遵）

**顺序不可颠倒：改 LaTeX 源 → 校验字体/OLE/定界符 → convert → 校验日志小节范围。**

1. **只改** `--source` **主稿**（`*_formula.docx` 只读预览/投稿，不作编辑源）
2. **重写段落**时用 `scripts/docx_latex_helpers.py`（或等价逻辑）：
  - `clear_para_content(para)` — 删 run/object，**保留** `w:pPr`
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

期刊模板字体/行距若不同，改 `docx_latex_helpers.py` 顶部常量，或从相邻未改段落复制 `w:pPr` 再写 run。**行距原则**：正文/`set_body_para_format` 不强制 1.5；仅独立公式行/`set_display_para_format` 用 `MULTIPLE 1.5` 避免 OLE 裁切。

### Agent 执行清单（门禁）

1. 确认源/目标 docx 均未打开（无 `~$` 锁）
2. **改完源稿后先抽查**：目标小节段落文本必须含 `$$...$$` 或 `$...$`，不能是 Unicode `α`/`τ`，也不能已是 OLE 空壳
3. 改稿脚本若失败（SyntaxError / 编码错误）：**停止**，修好并重跑改稿，禁止接着 convert
4. `--source` = 干净 LaTeX 主稿；`--target` ≠ source（通常 `*_formula.docx`）
5. 只用本 skill 目录下 `scripts/convert_docx_latex_to_equations.py`
6. 看日志：`failed=0`；**且**改稿小节无 `skipped-existing-object`；**且** `inlineConverted`/`displayConverted` 非异常为 0
7. 抽查输出：目标节 `dollar=False` 且 display 段 `ole≥1`；**中文 run 仍为模板字体**（非 Calibri 默认）
8. 若 `failed=0` 但仍见裸 `$i$`/`$m$`/`$N$`：属签名门旧 bug；确认脚本含 `has_delimited_math`，从干净 LaTeX 源重转
9. **右编号视觉验收（硬标准）**：display 段文本形如 `TAB + OLE + TAB + (n)`；制表位 Center@`usable/2` 与 Right@`usable`；**公式居中、编号贴右**。仅有 `(n)` 贴在公式旁 = 布局失败，需重跑 `apply_mathtype_right_number_layout`
10. 导 PDF/PNG 肉眼查：无残留 `$$`/`\lvert`、无文末空段洪泛、行距不为 EXACTLY 裁切





### 案例复盘（互扰模型.docx，2026-11）

一次完整 latex→MathType 过程中连续踩到的坑（均已修脚本/文档）：

1. **裸 `$i$` 静默跳过**：`has_formula_signature` 不认无 `_`/`\` 的定界公式 → 加 `has_delimited_math`
2. **行距 EXACTLY 15.6pt 裁切 OLE** → **仅公式行** `MULTIPLE 1.5`，正文保持原固定行距；display 源段左对齐留给制表位
3. **`\lvert`/`\dfrac` 转换残字** → 源稿只用 `\left|...\right|`、`\frac`
4. **end-paste 后 OLE 旁残留 `$$...$$`** → 粘贴前清空 Range + scrub
5. **编号不贴右 / 公式不居中**：`InsertEqnNum` 后的字段符让 lone-display 误判，跳过 `TAB+OLE+TAB+(n)` → 字段符计入 lone-display；布局后肉眼确认**公式居中、编号贴右**
6. **文末上百空段** → 转换后清理 trailing empty paragraphs

推荐命令（带右编号 display）：

```powershell
python scripts/convert_docx_latex_to_equations.py `
  --source clean_latex.docx `
  --target out_formula.docx `
  --toggle-strategy end-paste `
  --mathtype-numbers `
  --formula-scope all
```

### Windows / 工具注意

- 写含 `$` 的 Python/LaTeX 时，勿用 PowerShell 双引号或 `python -c "...$r(t)..."`（`$` 会被展开）。优先：`.py` 文件 + `python path.py`，或 `@'...'@ | Set-Content`
- 含中文的 `.py` 若 Write 工具编码损坏，改用 `\uXXXX` 转义或 PowerShell `@' '@` 写入 UTF-8



### Word / WPS 公式 → MathType（OMath 路径）

1. 源稿含 Word 公式编辑器对象（OMML）；WPS 请先另存为 `.docx` 并用 Word 打开确认是 OMath
2. 运行 `convert_docx_omath_to_mathtype.py` 或主脚本 `--source-mode omath`
3. **正确流水线（已修复）**：
  - OOXML：`m:oMath` → `$latex$` / `$$latex$$`；带 Word 手打 `(n)` 的独立公式段 → `$$...$$` **并剥离** `(n)`**/TAB**
  - 在**文末干净段落**执行 `MTCommand_TeXToggle` 得到 `Equation.DSMT4`
  - Copy OLE → Paste 覆盖原稿 `$...$` → 删除文末 helper
  - 对独立 display（`$$`）：`InsertEqnNum` + `TAB+OLE+TAB+(n)` **右编号版式**（制表位相对栏宽/左页边距；双栏用栏宽）
  - 段内有旁白的 `oMathPara` 按 **行内** `$...$` 处理，避免整段居中撑破双栏
4. 日志字段：`omathFound` / `omathRewritten` / `omathConverted` / `eqnNumbersInserted` / `omathFailed`
5. **临时文件**：OOXML unzip 目录用完即删；勿残留 `wlmt_omath_unzip_`*



### 已知失败模式（OMath）与修复


| 现象                                      | 根因                                             | 修复                                                      |
| --------------------------------------- | ---------------------------------------------- | ------------------------------------------------------- |
| `omathConverted=0`，`failed-tex-toggle`  | 在原 OMath 位置删公式再插 `$...$` 后原地 TeX Toggle        | **禁止原地 Toggle**；改文末 Toggle + OLE 回贴                     |
| 中间 LaTeX 稿 Toggle 失败、大量 `word-omath` 回退 | 学位论文常见 `TAB + 公式 + TAB + (1)` 行，原地 Toggle 静默无效 | 同样走文末 Toggle + paste-back；OMath 路径不回退 Word OMath        |
| 编号贴在公式旁、无右对齐空隙                          | 只插了 `InsertEqnNum`，缺 `TAB+OLE+TAB`             | `apply_mathtype_right_number_layout`                    |
| 有 TAB 但公式不居中、偏右                         | 制表位写成 `LeftMargin + usable/2`（相对页边又加了页边）       | 制表位用 `usable/2` 与 `usable`（相对**左页边距/栏左缘**）              |
| **双栏** display 撑破栏、正文被居中                | 制表位用整页 `usable`；`oMathPara`+旁白被当成单行并 `jc=center` | 栏宽制表位；仅 formula-only 段为 display；禁止对正文段强制居中/equation 样式 |
| 仍是 Word 手打 `(1)(2)…`                    | OOXML 未剥编号 / 未走 MathType 编号                    | 重写时剥 `(n)`；默认 `--mathtype-numbers`                      |
| 插入的 `$` 变成 `\uf024` 等 PUA               | 残留数学区域 / 错误字体                                  | OOXML 写入显式 TNR run；Toggle 前 `Font.Name=Times New Roman` |
| pandoc 产出 `\left. x \right.\hat{}`      | OMML 重音还原怪异                                    | `sanitize_mathtype_latex` 清洗                            |
| `--formula-scope inline` 且 `omathConverted=0` | 源稿 OMath 全是独立成段/带 `(n)` 的单行，无段内公式      | 属预期；行内交付物保持源稿，日志 `skippedByScope`≈全部              |
| 批量转 B 行内 `dollarFound=0` 但 XML 有 `$`     | Word 单实例复用，`ActiveDocument` 绑到上一篇            | `Documents.Open` + 校验 `FullName`；每任务前强制杀 Word；RPC 死则 reopen |




### 限制

- 仅 Windows；需 MathType 加载项
- 勿用 Insert GUI 批量转换
- OMath 路径需要 pandoc + lxml；复杂公式 LaTeX 还原可能有细节损失
- 复杂中文混排可用 legacy 或手工
- MathType 官方 Convert Equations 对话框自动化不稳定；OMath 走 OMML→LaTeX→**文末** TeX Toggle



## English Guide

Headless `.docx` LaTeX **or Word OMath** → MathType OLE via `MTCommand_TeXToggle`. No MathType GUI.

```powershell
python scripts/convert_docx_latex_to_equations.py --source input.docx
python scripts/convert_docx_omath_to_mathtype.py --source input.docx
python scripts/convert_docx_latex_to_equations.py --source input.docx --source-mode auto
```

**OMath path (fixed):** OOXML rewrite OMML to dollar LaTeX, then TeX Toggle on a clean end-of-document paragraph, Copy OLE, Paste over the original span. Do not TeX-toggle in-place on thesis TAB+eq+TAB+(n) rows or residual math zones.

**Matching order (latex mode):** display `$$...$$` → explicit inline `$...$` (always convert) → bare `INLINE_PATTERNS`.

**Gate:** After editing the source, verify paragraphs contain real LaTeX (`$$` / `$` / `\alpha`), **no stale OLE**, and **restored run/paragraph fonts** (use `docx_latex_helpers.py`). Then convert. Log must show `failed=0` **and** no `skipped-existing-object` in the edited section's Word paragraph index range—not `failed=0` alone when `inlineConverted=0`.

Fallback: `--mathtype-mode word`. Legacy CJK: `-UseLegacy` on `.ps1`.