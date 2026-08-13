---
name: gongwen-converter
description: 公文格式转换工作流——将 Markdown 一键转换为符合 GB/T 9704-2012《党政机关公文格式》的 Word 文档。当用户要求"把 md 转成公文格式 docx"、"生成公文 word"、"用公文模板转换文稿"、"写公文材料"时使用本技能。内置 convert.py（Pandoc 转换+清间距+修引号）、create_reference.py（生成模板）、strip_spacing.py（后处理）。搭配 doc-reader 技能读取已有材料、references/公文MD提示词.txt 约束 AI 起草格式。
---

# Gongwen Converter — 公文格式转换

## 概述

将 Markdown 文稿转换为符合 GB/T 9704-2012 的公文 .docx：二号方正小标宋标题、三号仿宋_GB2312 正文（固定行距 28 磅、首行缩进 2 字符）、黑体/楷体层级标题、A4 标准页边距。全流程自动处理段间距与中文引号方向。

## 使用流程

### 1. 准备环境

- Python venv：`C:\Users\admin\.workbuddy\binaries\python\envs\default\Scripts\python.exe`（已装 python-docx、lxml、pywin32）
- Pandoc：`C:\Program Files\Pandoc\pandoc.exe`

### 2. 生成模板（仅首次或模板丢失时）

```bash
<python> <skill_dir>/scripts/create_reference.py
```

生成 `reference.docx` 到当前目录（公文样式模板）。

### 3. 转换文稿

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

## Markdown 编写规范（给 AI 起草时用）

起草阶段把 `references/公文MD提示词.txt` 全文作为约束发给 AI，关键规则：

- YAML front matter 写 `title`（不重复写标题、不加书名号）
- **段落之间必须用空行分隔**（Markdown 语法要求；转换工具会确保最终无多余段间距）
- `#` → 一级标题（一、）；`##` → 二级标题（（一））
- 全文不用粗体/斜体/列表等格式标记
- 落款/日期转换后在 Word 中手动右对齐

## 文档迭代规则（强制，防止覆盖用户手动修改）

**修改已有的公文文档前，必须先读取用户当前的最新文件，在原文基础上迭代；绝不可基于自己的缓存/旧版本直接覆盖 docx。**

1. **先读再改**：每次修改 docx 前，先用 doc-reader 技能读取当前 docx 内容（`read_doc.py <文件.docx> -o 内容.txt`），与已同步的 md 对比。
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
6. **字体依赖**：方正小标宋简体、仿宋_GB2312、楷体_GB2312 需系统已安装，否则 Word 会回退字体。
7. **标题前后各空一行（回车空行）**：公文标题（Title）段落前后各插入一个真正的空段落（回车产生的空行），Heading 不加；不是段后间距。strip_spacing.py 的 `insert_blank_around_title` 实现，且幂等（已有空行则跳过）。
8. **必须定义 FirstParagraph / BodyText 样式**：Pandoc 转换时正文段落实际引用这两个样式（非 Normal），模板若不定义则 Word 回退显示异常（曾导致党建材料格式错乱）。create_reference.py 已将其定义为与 Normal 一致。
9. **必须清除 rFonts 的 theme 引用属性**：Pandoc 模板会给各样式带 `asciiTheme/eastAsiaTheme/hAnsiTheme/cstheme`，其优先级高于显式字体名，会把黑体/楷体/小标宋覆盖成主题字体（宋体）。strip_spacing.py 的 `strip_font_theme_attrs` 已自动清除（document.xml + styles.xml 两层），无需手动处理。

## 与 doc-reader 技能配合

起草前如需参考已有材料：

```bash
<python> ~/.workbuddy-ai/skills/doc-reader/scripts/read_doc.py <旧材料.doc(x)> -o 内容.txt
```

读取后用 `references/公文MD提示词.txt` 约束起草，再经 convert.py 转换。
