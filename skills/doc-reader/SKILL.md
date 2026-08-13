---
name: doc-reader
description: 读取 .doc/.docx 已有材料内容（段落、表格、标题层级），供起草公文等文稿时参考。当用户要求"读取/提取/查看某个 doc 或 docx 文件内容"、"写材料前参考已有材料"、"看看这份材料里写了什么"时使用本技能。支持旧版二进制 .doc（antiword + Word COM 回退）和新版 .docx（python-docx）。
---

# Doc Reader — doc/docx 内容读取

## 概述

从已有 Word 文档（.doc / .docx）中提取正文内容，输出为可读文本，供起草新文稿时参考。脚本封装了三种读取方式，自动选择可用方案，无需每次重新写代码。

## 使用方法

用项目 venv 的 Python 运行脚本（脚本随技能打包在 `scripts/read_doc.py`）：

```bash
# 打印到屏幕
<python> <skill_dir>/scripts/read_doc.py "文件路径.doc"

# 保存到 txt 文件
<python> <skill_dir>/scripts/read_doc.py "文件路径.docx" -o 输出.txt
```

其中 `<python>` 为可用 Python 解释器；本机默认：`C:\Users\admin\.workbuddy\binaries\python\envs\default\Scripts\python.exe`

## 读取策略（自动选择）

| 文件类型 | 优先方案 | 回退方案 |
|---------|---------|---------|
| `.docx` | python-docx（段落+表格，标题转 `#` 层级） | — |
| `.doc`（旧版二进制） | antiword（需复制到临时 ASCII 路径，因不支持中文路径） | Word COM（win32com，需本机安装 Microsoft Word + pywin32） |

## 关键注意事项（踩坑记录）

1. **antiword 不支持中文路径**：文件必须先复制到纯 ASCII 临时路径再调用，脚本已内置处理。
2. **antiword 0.37 解析新版 .doc 常失败**（报 "is not a Word Document"）：这是版本限制，脚本会自动回退到 Word COM。
3. **Word COM 是主力方案**：本机已装 Microsoft Office，pywin32 已安装于 venv。若脚本报 `ImportError: win32com`，先执行 `pip install pywin32`。
4. **输出格式**：docx 的标题段落会转为 `#`/`##` 层级，正文为纯文本；表格行以 ` | ` 连接。

## 复用场景

- 起草公文/总结/报告前，提取历史材料的框架和表述
- 比较新旧版本文档内容
- 从长文档中定位特定章节（输出到 txt 后可用 grep 检索）
