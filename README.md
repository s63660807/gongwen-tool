# 公文格式转换工具

将 Markdown 一键转换为符合 **GB/T 9704-2012《党政机关公文格式》** 的 Word 文档。

## 功能特性

- **Markdown → 标准公文 docx**：Pandoc 转换 + 自定义样式模板
- **自动处理格式细节**：
  - 标题：二号方正小标宋简体，居中
  - 正文：三号仿宋_GB2312，固定行距28磅，首行缩进2字符
  - 一级标题：三号黑体；二级标题：三号楷体_GB2312
  - A4 页面，标准公文页边距（上3.7/下3.5/左2.8/右2.6 cm）
  - 自动清除段前段后间距（处理 document.xml + styles.xml 两层）
  - 自动修复中文引号方向（左引号“右引号”成对）
- **读取已有材料**：doc / docx 内容提取，供起草新文稿参考
- **可打包为独立 exe**：无需安装 Python（PyInstaller）

## 目录结构

```
├── scripts/
│   ├── create_reference.py   # 生成公文格式模板 reference.docx
│   ├── convert.py            # 一键转换：md → 标准公文 docx（主入口）
│   ├── strip_spacing.py      # 后处理：清除段间距 + 修复中文引号
│   └── read_doc.py           # 读取 doc/docx 已有材料内容
├── 公文MD提示词.txt           # AI 起草提示词（规范 md 输出格式）
└── 使用说明.txt               # 详细使用说明
```

## 快速开始

### 环境依赖

- Python 3.8+
- [Pandoc](https://pandoc.org/installing.html)（转换引擎）
- pip 安装：`pip install python-docx lxml`（读取/后处理）
- 读取 .doc 旧格式：需安装 Microsoft Word（Windows），并 `pip install pywin32`

### 1. 生成公文模板（首次）

```bash
python scripts/create_reference.py
# 生成 reference.docx（GB/T 9704-2012 样式模板）
```

### 2. 转换文稿

```bash
python scripts/convert.py 文稿.md
# 生成同名 .docx，自动完成：Pandoc 转换 → 清段间距 → 修引号
```

### 3. 读取已有材料

```bash
python scripts/read_doc.py 旧材料.docx          # 打印内容
python scripts/read_doc.py 旧材料.doc -o 内容.txt  # 保存为 txt
```

### 4. AI 起草（可选）

把 `公文MD提示词.txt` 的内容粘贴给 AI，让它按规范格式输出 Markdown，
再拖入本工具即可转换。

## Markdown 编写规范

```markdown
---
title: 关于XXX工作的通知
---

正文段落直接写，段落之间用空行分隔。

# 一、总体要求

## （一）基本原则

```

- `title` (YAML) → 公文标题（二号方正小标宋，居中）
- `#` → 一级标题（一、二、三、）
- `##` → 二级标题（（一）（二）（三））
- 正文段落 → 三号仿宋，首行缩进2字符
- 落款/日期转换后在 Word 中手动右对齐

## 打包为独立 exe（可选）

```bash
pip install pyinstaller
pyinstaller --onefile --name 公文转换工具 \
  --add-data "reference.docx;." \
  --hidden-import docx --hidden-import lxml \
  scripts/convert.py
```

将生成的 exe 与 `pandoc.exe`、`reference.docx` 放在同一目录即可分发，
目标机器无需安装 Python。

## 说明

- 主送机关、署名、成文日期等特殊段落需在转换后的 Word 中手动调整对齐方式
- 模板使用「仿宋_GB2312」「方正小标宋简体」「楷体_GB2312」字体，需系统已安装

## 许可

MIT License
