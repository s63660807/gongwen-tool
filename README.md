# 公文格式转换工具

将 Markdown 一键转换为符合 **GB/T 9704-2012《党政机关公文格式》** 的 Word 文档。

## 功能特性

- **Markdown → 标准公文 docx**：Pandoc 转换 + 自定义样式模板
- **自动处理格式细节**：
  - 标题：二号方正小标宋简体，居中，前后各空一行
  - 正文：三号仿宋_GB2312，固定行距 28 磅，首行缩进 2 字符
  - 一级标题（一、）：三号黑体
  - 二级标题（（一））：三号楷体_GB2312 加粗
  - 三级标题（1.）：三号仿宋，引导语加粗
  - A4 页面，标准公文页边距（上 3.7 / 下 3.5 / 左 2.8 / 右 2.6 cm）
  - 清除段前段后间距（document.xml + styles.xml 两层）
  - 修复中文引号方向（左引号“右引号”成对）
- **表格美化**（Pandoc 默认输出的表格是没有框线的，直接用会很"素"）：
  - 显式全框线，外框 1pt / 内线 0.5pt
  - 表宽固定为版心宽 15.6cm，表格居中，`fixed` 布局
  - 列宽：标准 7 列预算表用实测调优值；其他表按各列最长内容加权自动分配
  - 表头黑体五号、表内仿宋_GB2312 五号、行距固定 16 磅
  - 垂直居中，表内取消首行缩进；表头跨页自动重复，表格整体不拆到两页
- **落款自动右对齐**：发文机关署名与成文日期右空四字
- **读取已有材料**：doc / docx 内容提取，供起草新文稿参考
- **格式校验**：一键体检页面设置、逐段样式、样式字体字号行距、表格框线列宽
- **可打包为独立 exe**：无需安装 Python（PyInstaller）

## 目录结构

```
├── skills/                            # WorkBuddy 技能包（推荐的用法）
│   ├── doc-reader/                    # 读取 doc/docx 已有材料
│   └── gongwen-converter/
│       ├── SKILL.md                   # 技能说明（完整踩坑记录）
│       ├── references/公文MD提示词.txt  # 约束 AI 起草的 md 格式
│       └── scripts/
│           ├── fix_gongwen_format.py  # ★ 推荐主入口：转换+层级字体+表格美化+落款对齐
│           ├── convert.py             # 轻量转换：md → 公文 docx
│           ├── create_reference.py    # 生成公文格式模板 reference.docx
│           ├── strip_spacing.py       # 后处理：清段间距 + 修中文引号
│           ├── refill_images.py       # 回填源文档内嵌图片
│           ├── check_gongwen.py       # 格式校验（含表格校验）
│           └── reference.docx         # 公文样式模板
├── scripts/                           # 独立运行的精简脚本副本
│   ├── fix_gongwen_format.py
│   ├── convert.py
│   ├── create_reference.py
│   ├── strip_spacing.py
│   └── read_doc.py
├── 公文MD提示词.txt                    # AI 起草提示词（规范 md 输出格式）
└── 使用说明.txt                        # 详细使用说明
```

## 快速开始

### 环境依赖

- Python 3.8+
- [Pandoc](https://pandoc.org/installing.html)（转换引擎）
- pip 安装：`pip install python-docx lxml`（读取/后处理）
- 读取 .doc 旧格式：需安装 Microsoft Word（Windows），并 `pip install pywin32`
- 可选（成品验证）：`pip install pywin32 pymupdf`，可把 docx 转 PDF 逐页核版

### 1. 生成公文模板（首次）

```bash
python skills/gongwen-converter/scripts/create_reference.py
# 生成 reference.docx（GB/T 9704-2012 样式模板）
```

### 2. 转换文稿（推荐）

```bash
python skills/gongwen-converter/scripts/fix_gongwen_format.py 文稿.md
```

一步完成：Pandoc 转换 → 清段间距/修引号 → 标题层级字体 → 表格美化 → 落款右对齐。

对**已有 docx** 做同样精修（不做 Pandoc 转换）：

```bash
python skills/gongwen-converter/scripts/fix_gongwen_format.py --docx 已转换.docx
```

### 3. 轻量转换（可选）

只要基础转换、不做表格与落款处理时：

```bash
python skills/gongwen-converter/scripts/convert.py 文稿.md
# 生成同名 .docx，自动打开
```

### 4. 校验格式

```bash
python skills/gongwen-converter/scripts/check_gongwen.py 文稿.docx -o 校验报告.txt
```

输出页面设置、逐段样式、样式字体/字号/固定行距/首行缩进、主题字体残留检查，
以及表格框线/表宽/居中/表头重复/字体字号校验；结尾给出「全部通过 ✓ / 存在问题 ✗」。
退出码 0=通过，2=存在问题。交付前建议跑一次。

### 5. 回填源文档中的图片（源文档含插图时必做）

转换走的是 Markdown → docx，Markdown 里没有图片，源 docx 的内嵌图片会被整体丢弃：

```bash
python skills/gongwen-converter/scripts/refill_images.py 源.docx 目标.docx [-w 15.6]
```

### 6. 读取已有材料

```bash
python skills/gongwen-converter/scripts/read_doc.py 旧材料.docx
python skills/doc-reader/scripts/read_doc.py 旧材料.doc -o 内容.txt
```

### 7. AI 起草（可选）

把 `公文MD提示词.txt` 的内容粘贴给 AI，让它按规范格式输出 Markdown，再交给本工具转换。

## Markdown 编写规范

```markdown
---
title: 关于XXX工作的通知
---

正文段落直接写，段落之间用空行分隔。

# 一、总体要求

## （一）基本原则

| 序号 | 项目 | 说明 |
| --- | --- | --- |
| 1 | XX | XX |

汕尾市水务局

2026年9月15日
```

- `title` (YAML) → 公文标题（二号方正小标宋，居中，前后各空一行）
- **`#` → 一级标题（一、二、三、）；`##` → 二级标题（（一）（二）（三））**
  —— **必须真的写成 Markdown 标题标记**，只把「一、总体要求」当普通段落写会让全文套用同一样式、层级字体全丢
- 三级标题「1.XXXX。」写在正文段落里即可，脚本自动把首个句号（含）之前的引导语加粗
- 正文段落 → 三号仿宋，首行缩进 2 字符
- 表格 → 标准 Markdown 管道表，第一行即表头，自动加框线、自动分配列宽；单元格留空写空即可
- 落款与成文日期各占一段放在末尾，脚本自动右对齐

## 打包为独立 exe（可选）

```bash
pip install pyinstaller
pyinstaller --onefile --name 公文转换工具 \
  --add-data "reference.docx;." \
  --hidden-import docx --hidden-import lxml \
  scripts/fix_gongwen_format.py
```

将生成的 exe 与 `pandoc.exe`、`reference.docx` 放在同一目录即可分发，
目标机器无需安装 Python。

## 文档迭代规则（重要）

修改**已有**的公文文档（而非新建）时，务必遵守以下约束，避免覆盖用户在 Word 中的手动修改：

1. **先读再改**：修改 docx 前，先用 `read_doc.py` 读取当前 docx 的最新内容，与 md 对比。
   **不要假设本地 md 仍是最新、也不要假设 docx 还是自己上次生成的那版**——用户很可能已在
   WPS/Word 里改过并保存。稳妥做法是逐行 diff 当前 docx 与 md，把差异逐条判断是"用户改动"
   还是"软件改写"，用户改动一律合并回 md。
2. **合并用户手动修改**：用户在 Word 里可能直接改过 docx（格式、措辞、内容），这些修改必须保留。先把用户改动合并进 md，再做本次修改并重新转换。
3. **转换前确认同步**：确保 md 与 docx 内容一致，md 是准确来源。
4. **文件被占用**：若 Word 正打开该文档（存在 `~$` 锁文件），输出到新文件名，提示用户关闭 Word 后再覆盖。
5. **最小改动**：只修改用户指定的部分，其余内容原样保留（用户强调"不要动我没说的部分"）。
6. **识别 WPS 改写**：文档 XML 里出现 `xmlns:wpsCustomData`、样式 ID 变成纯数字、标题被拆成多段、
   出现 `<w:t xml:space="preserve"> </w:t>` 空 run —— 说明用户在 WPS 里重新保存过。

## 说明

- 模板使用「仿宋_GB2312」「黑体」「方正小标宋简体」「楷体_GB2312」字体，需系统已安装；
  未装 `楷体_GB2312` 时 Word 会自动替换为「楷体」，视觉可接受
- 成品核版建议转 PDF 逐页看：只看 docx XML 会漏掉分页、行距挤压、字体回退等问题

## 许可

MIT License
