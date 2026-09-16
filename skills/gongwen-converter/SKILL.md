---
name: gongwen-converter
description: 公文格式转换工作流——将 Markdown 一键转换为符合 GB/T 9704-2012《党政机关公文格式》的 Word 文档。当用户要求"把 md 转成公文格式 docx"、"生成公文 word"、"用公文模板转换文稿"、"把某份 docx 转成标准公文格式"、"写公文材料"、"按公文格式修正/美化表格"时使用本技能。内置 convert.py（Pandoc 转换+清间距+修引号）、create_reference.py（生成模板）、strip_spacing.py（后处理）、fix_gongwen_format.py（推荐主入口：转换+标题层级字体+表格美化+落款右对齐）、refill_images.py（回填源文档内嵌图片）、check_gongwen.py（格式校验）。搭配 doc-reader 技能读取已有材料、references/公文MD提示词.txt 约束 AI 起草格式。
---

# Gongwen Converter — 公文格式转换

## 概述

将 Markdown 文稿转换为符合 GB/T 9704-2012 的公文 .docx：二号方正小标宋标题、三号仿宋_GB2312 正文（固定行距 28 磅、首行缩进 2 字符）、黑体/楷体层级标题、A4 标准页边距。全流程自动处理段间距、中文引号方向、**表格框线与列宽、落款右对齐**。

**推荐入口是 `fix_gongwen_format.py`**——它包含 convert.py 的全部动作，并额外补齐三块最容易漏掉的成品化处理（标题层级字体与加粗、表格美化、落款对齐）。`convert.py` 保留作为轻量转换入口。

## 使用流程

### 1. 准备环境

- Python venv：`C:\Users\admin\.workbuddy\binaries\python\envs\default\Scripts\python.exe`（已装 python-docx、lxml、pywin32）
- Pandoc：`C:\Program Files\Pandoc\pandoc.exe`

### 2. 生成模板（仅首次或模板丢失时）

```bash
<python> <skill_dir>/scripts/create_reference.py
```

生成 `reference.docx` 到当前目录（公文样式模板）。

### 3. 转换文稿（推荐）

```bash
<python> <skill_dir>/scripts/fix_gongwen_format.py <输入.md> [输出.docx]
```

一步完成：Pandoc 转换 → 清段间距/修引号 → **标题层级字体 + 二级标题加粗 + 三级标题引导语加粗** → **表格美化** → **落款右对齐**。

对**已有 docx** 做同样的精修（不做 Pandoc 转换，但仍清段间距、修引号）：

```bash
<python> <skill_dir>/scripts/fix_gongwen_format.py --docx <输入.docx> [输出.docx]
```

省略输出路径时，md 模式输出同名 .docx；`--docx` 模式原地精修。目标文件被 Word 占用时自动改输出 `<名>_new.docx`。

**表格美化做了什么**（pandoc 默认表格是无框线的 Normal Table，直接用会很"素"）：

| 项 | 处理 |
|---|---|
| 框线 | 显式 `tblBorders` 全框线，外框 1pt / 内线 0.5pt，颜色黑 |
| 宽度 | `tblW` 固定为版心宽 8844 twip，`tblLayout=fixed`，表格居中 |
| 列宽 | 7 列预算表用实测调优值；其他表按各列最长内容加权自动分配 |
| 字体 | 表头黑体五号、表内仿宋_GB2312 五号，行距固定 16 磅 |
| 对齐 | 表头与短内容列居中、长文本列左对齐；垂直居中；表内取消首行缩进 |
| 跨页 | 表头行 `tblHeader` 跨页自动重复；单行 `cantSplit` 不拆行；表格前段落 `keepNext` 保证表格整体不拆到两页 |
| 间距 | 表格后正文加 12 磅段前间距，避免贴排 |

### 3b. 轻量转换（可选）

只要基础转换、不做表格与落款处理时：

```bash
<python> <skill_dir>/scripts/convert.py <输入.md> [输出.docx]
```

- 只给 .md 时默认输出同名 .docx
- convert.py 自动查找 pandoc（同目录 > 常见安装路径 > PATH），找不到时报错提示
- 转换后自动 `os.startfile` 打开生成的 docx

### 4. 单独后处理（可选）

若已有 docx 需要清理，可单独运行：

```bash
<python> <skill_dir>/scripts/strip_spacing.py <输入.docx>
```

### 5. 回填原文档中的图片（源文档含插图/思维导图时必做）

**关键**：转换走的是 Markdown → docx，Markdown 里没有图片，所以**源 docx 中的内嵌图片
（思维导图、插图、表格截图等）会被整体丢弃**。源文档含图片时必须补这一步：

```bash
<python> <skill_dir>/scripts/refill_images.py <源.docx> <目标.docx> [-w 15.6]
```

- 以「图片所在段落之前最近的非空文本段落」为锚点，在目标 docx 中定位同一段落并把图片插到其后
- `-w` 为图片宽度（cm），默认 15.6（A4 公文版心宽 = 21 − 2.8 − 2.6）
- 找不到锚点会告警跳过，不会破坏目标文档；可重复执行（先删同位置旧图）

### 6. 校验输出格式

```bash
<python> <skill_dir>/scripts/check_gongwen.py <目标.docx> [-o 报告.txt]
```

输出页面设置、逐段样式清单、各样式字体/字号/固定行距/首行缩进及主题字体残留检查，
结尾给出「全部通过 ✓ / 存在问题 ✗」。交付前建议跑一次。

## Markdown 编写规范（给 AI 起草时用）

起草阶段把 `references/公文MD提示词.txt` 全文作为约束发给 AI，关键规则：

- YAML front matter 写 `title`（不重复写标题、不加书名号）
- **段落之间必须用空行分隔**（Markdown 语法要求；转换工具会确保最终无多余段间距）
- **`#` → 一级标题（一、）；`##` → 二级标题（（一））——必须真的写成 Markdown 标题标记，不能只把「一、活动主题」当普通段落写！** 否则 pandoc 会把全文套成同一个正文样式，黑体/楷体层级全丢（这是"转换出来不像公文"的头号原因）
- 三级标题「1.XXXX。」写在正文段落里即可，脚本会自动把首个句号（含）之前的引导语加粗
- 全文不用粗体/斜体/列表等格式标记
- **表格用标准 Markdown 管道表**（表头行 + 分隔行 + 数据行），第一行即表头，转出来自动加框线、自动分配列宽；单元格留空写空即可，不要写 `-`、`/`、`无`
- 落款与成文日期各占一段放在全文末尾，脚本会自动右对齐（右空四字），无需在 Word 里手工调

## 文档迭代规则（强制，防止覆盖用户手动修改）

**修改已有的公文文档前，必须先读取用户当前的最新文件，在原文基础上迭代；绝不可基于自己的缓存/旧版本直接覆盖 docx。**

1. **先读再改**：每次修改 docx 前，先用 doc-reader 技能读取当前 docx 内容（`read_doc.py <文件.docx> -o 内容.txt`），与已同步的 md 对比。
   **注意：不要假设本地 md 仍是最新、也不要假设 docx 还是自己上次生成的那版**——用户很可能已在 WPS/Word 里改过并保存。稳妥做法是逐行 diff 当前 docx 与 md（`difflib.SequenceMatcher` 即可），把所有差异逐条判断是"用户改动"还是"软件改写"，用户改动一律合并回 md。识别 WPS 改写的特征见「踩坑记录」第 22 条。
2. **合并用户的手动修改**：用户在 Word 中可能直接改过 docx（格式、措辞、内容），这些修改**必须保留**。先把用户手动修改的部分合并进 md，再做本次修改并重新转换。
3. **转换输出前确认 docx 与 md 已同步**：确保 md 是转换的准确来源，不丢失用户改动。
4. **文件被 Word 占用**：若检测到 `~$` 锁文件（Word 正在打开该文档），输出到新文件名，并提示用户关闭 Word 后再覆盖，不要强行覆盖。
5. **最小改动原则**：用户强调"不要动我没说的部分"——只修改用户指定的部分，其余内容原样保留。
6. **读取工具**：读取/比对用 doc-reader 技能的 `read_doc.py`（.docx 走 python-docx，.doc 走 antiword/Word COM）。

## 关键注意事项（踩坑记录）

1. **Pandoc 必须用 `-f markdown-smart`**：禁用智能引号，否则中文引号全被转成右引号 `”`（convert.py 已内置）。
2. **引号修复逻辑**：strip_spacing.py 按段落顺序交替配对引号（奇数→左 `“`，偶数→右 `”`），旧文档也可用它修复。
3. **段间距两层清理**：document.xml（段落级）+ styles.xml（docDefaults 兜底默认值），漏掉后者 Word 仍会显示段后间距。
4. **首行缩进用 `w:firstLineChars`**（字符数）而非 `w:firstLine`（绝对磅值），字号变化时缩进仍为 2 字符。
5. **模板无页码**：用户已要求去掉页脚页码，create_reference.py 不含页码生成逻辑。
6. **字体依赖（本机实测）**：`仿宋_GB2312`、`黑体`、`方正小标宋简体` 已装；**`楷体_GB2312` 未装**，Word 会自动替换为「楷体」（PDF 里嵌入名显示 `KaiTi`），视觉可接受，不必为此改字体名。检查方法：看 `C:\Windows\Fonts` 下有无 `仿宋_gb2312.ttf`、`方正小标宋简.ttf`、`simhei.ttf`、`simkai.ttf`（注意 `simkai.ttf` 是"楷体"而非"楷体_GB2312"）。
7. **标题前后各空一行（回车空行）**：公文标题（Title）段落前后各插入一个真正的空段落（回车产生的空行），Heading 不加；不是段后间距。strip_spacing.py 的 `insert_blank_around_title` 实现，且幂等（已有空行则跳过）。
8. **必须定义 FirstParagraph / BodyText 样式**：Pandoc 转换时正文段落实际引用这两个样式（非 Normal），模板若不定义则 Word 回退显示异常（曾导致党建材料格式错乱）。create_reference.py 已将其定义为与 Normal 一致。
9. **必须清除 rFonts 的 theme 引用属性**：Pandoc 模板会给各样式带 `asciiTheme/eastAsiaTheme/hAnsiTheme/cstheme`，其优先级高于显式字体名，会把黑体/楷体/小标宋覆盖成主题字体（宋体）。convert.py / strip_spacing.py 均已自动清除（document.xml + styles.xml 两层，`strip_font_theme_attrs`），无需手动处理。
10. **排版紧凑三开关（分页粘连处置，压缩篇幅关键）**：python-docx 内建模板的 Heading1-9 自带 `w:keepLines`（段中不分页）和 `w:keepNext`（与下段同页），Word 默认还开孤行控制（`w:widowControl`），会把段落尾部最后几行整体挤到下一页、页尾留白严重。已在**多层**处置：create_reference.py 生成模板时删除样式的 keep 属性（`remove_keep_flags`，源头）、convert.py 内嵌副本生成模板后同样清除、后处理兜底用**置 0 法**——`disable_keep_together`（convert.py 与 strip_spacing.py 均有）按 CT_PPr schema 顺序把 `keepNext=0`/`keepLines=0`/`widowControl=0` 显式写入 document.xml 与 styles.xml 两层的每个段落 pPr（`PPR_CHILD_ORDER` 保证插入顺序合法，避免 Word 校验报错）。关闭后每页可多排 2~4 行，要把长文压到指定页数（如两页半）时此项必须开启。注意：Word 打开文档统计页数（ComputeStatistics）与打印分页一致，验证篇幅以 Word/PDF 实测为准，别只按字数估算。
11. **标题内不得夹带正文（正文不跟标题格式）**：若 md 把二级/三级标题和正文写在同一自然段（如 `## （一）强化组织领导。各级党组要切实履行主体责任……`），Pandoc 会整段渲染为 Heading 样式，导致本应正文的内容也跟着标题格式（楷体/黑体）。修复：`split_heading_with_inline_body`（strip_spacing.py）/ `_split_heading_inline_body`（convert.py 内嵌）会把标题到第一个「。」为止作为标题、句号后的内容拆成一个 FirstParagraph 正文段落。幂等：标题无句号或句号后无正文时不拆分，正常短标题不受影响。
12. **偶发 UnicodeDecodeError**：convert.py 偶尔报 `Pandoc 转换失败: None` + `UnicodeDecodeError: 'utf-8' codec can't decode...`（多为 pandoc stderr 输出非 UTF-8 字节所致）。convert.py 已对 subprocess 加 `encoding="utf-8", errors="replace"` 容错，如仍复现，重跑一次即可成功，无需改文件。
13. **源文档内嵌图片会被丢掉**：Markdown 转换链路不携带图片，源 docx 里的插图/思维导图/截图会全部丢失（且 `read_doc.py` 读出来是空段，容易误判为「原文就是空的」）。**判定方法**：查源 docx 包内是否有 `word/media/*`。有则转换完必须用 `refill_images.py` 回填。
14. **图片段落的行距必须改成 auto**：公文样式的「固定行距 28 磅」会把图片段落压成 28 磅高，**图片被裁掉只剩一条**。`refill_images.py` 已对新插入的图片段落显式设 `line_spacing = 1.0`（即 `lineRule="auto"; line=240`）、居中、无缩进；若手工插图务必照做。
15. **Windows 控制台输出中文乱码**：用 PowerShell 调 Python 脚本时，stdout 中文常变乱码（脚本本身写文件的内容是正常的 UTF-8）。判断结果请看脚本生成的报告文件，或先设 `$env:PYTHONIOENCODING="utf-8"`，不要因为控制台乱码就以为生成失败。
16. **md 标题没写 `#`/`##` 是"格式全丢"的头号原因**：只把「一、活动主题」写成普通段落时，pandoc 会把它套成 `BodyText`（body text），黑体/楷体层级字体、缩进层级全部失效——文档看起来就"不像公文"。原始 docx 一旦重新生成就难以察觉，务必在起草阶段就用真标题标记。
17. **pandoc 表格默认无框线**：pandoc 用 `tblStyle="Normal Table"`（`Table Grid` 之外的样式都不带框线），且列宽按内容"自然"分配，常出现「序号」两字竖排、备注列末行只剩一个字。**必须在后处理里显式写 `tblBorders`**，不要指望改样式表解决。
18. **表格列宽存在"临界值断崖"**：五号字（10.5pt）下，列可容字数 = (列宽 − 左右内边距) / 字号宽。差 0.2pt 就会把 7 字/行打成 6 字/行，直接把某一列挤成 3 行、末行只剩一个字。**别靠估算**——用 PDF 实测反推：`page.get_drawings()` 收集竖线 x 坐标，相邻差分 ×20 即得实际列宽（twip），再对照目标值调整。本技能 7 列预算表的调优值已写入脚本常量。
19. **表格防跨页三件套**：`tblHeader`（表头跨页重复）+ 行 `cantSplit`（单行不拆）+ 表格前段落 `keepNext`（表格整体不与前文拆开）。只用前两个时，表格仍可能从页面底部开始、只留 1 行到下一页。
20. **成品验证必须落到 PDF**：只看 docx XML 会漏掉分页、行距挤压、字体回退等问题。用 Word COM `Documents.Open(...).SaveAs2(path, FileFormat=17)` 转 PDF，再 `pymupdf` `get_pixmap(dpi≈105)` 出图逐页目视；要量化就 `get_text('blocks')` 取文本块 y 坐标（算每页留白 = 页面高 − 末行 y − 下边距）、`get_drawings()` 取表格线坐标。`page.get_fonts()` 可确认字体是否真被嵌入。
21. **注意留白的正常值**：A4 页面高 842pt，下边距 3.5cm ≈ 99pt，所以每页"下方空白"有 ~100pt 是**正常的边距**，不是排版浪费；判断有没有浪费要减去下边距再看。曾因此误判"页面没排满"而白折腾一轮。
22. **文件被 WPS 改过的识别特征**：文档 XML 里出现 `xmlns:wpsCustomData="http://www.wps.cn/officeDocument/2013/wpsCustomData"`、样式 ID 变成纯数字（`1`/`19`/`31`）、标题被 `w:br` 或空格拆成多段、`<w:t xml:space="preserve"> </w:t>` 空 run —— 说明用户在 WPS 里打开并重新保存过。**这时必须以 docx 当前内容为准**，先与本地 md 做逐行 diff，把用户改动合并回 md 再操作，绝不能拿旧 md 直接覆盖。

## 与 doc-reader 技能配合

起草前如需参考已有材料：

```bash
<python> ~/.workbuddy-ai/skills/doc-reader/scripts/read_doc.py <旧材料.doc(x)> -o 内容.txt
```

读取后用 `references/公文MD提示词.txt` 约束起草，再经 convert.py 转换。
