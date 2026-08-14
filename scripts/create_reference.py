#!/usr/bin/env python3
"""
生成符合《党政机关公文格式》(GB/T 9704-2012) 的 Pandoc 参考文档。
该文档作为 pandoc --reference-doc 的模板，用于将 Markdown 转换为规范公文。

规范来源: GB/T 9704-2012 第7条"公文格式各要素编排规则"

样式 ── Pandoc 映射:
  Title       ← 文档标题 (YAML title / % title)
  Heading 1   ← # (一级标题: 一、二、三、)
  Heading 2   ← ## (二级标题: (一)(二)(三))
  Heading 3   ← ### (三级标题: 1. 2. 3.)
  Heading 4   ← #### (四级标题: (1)(2)(3))
  Normal      ← 正文段落

Pandoc 限制:
  主送机关 (顶格)、署名日期 (右对齐)、附件标注 等特殊段落无法自动映射，
  需转换后在 Word 中手动调整，或使用 Lua 过滤器。
"""

from docx import Document
from docx.shared import Pt, Cm, RGBColor, Emu
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_TAB_ALIGNMENT, WD_TAB_LEADER
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import os

OUTPUT_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "reference.docx")

# 如果文件被占用则保存到备份名称
def safe_save(doc, path):
    try:
        doc.save(path)
    except PermissionError:
        base, ext = os.path.splitext(path)
        alt = base + "_new" + ext
        doc.save(alt)
        return alt
    return path


# ═══════════════════════════════════════════════════════════════════
# 工具函数
# ═══════════════════════════════════════════════════════════════════

def set_style_font(style, latin_font, ea_font, size_pt, bold=False, italic=False):
    """设置样式的中文字体 (w:eastAsia) 和西文字体 (w:ascii/w:hAnsi) 及字号"""
    style.font.size = Pt(size_pt)
    style.font.bold = bold
    style.font.italic = italic
    style.font.name = latin_font
    style.font.color.rgb = RGBColor(0, 0, 0)  # 全文黑色

    rPr = style.element.get_or_add_rPr()
    rFonts = rPr.find(qn("w:rFonts"))
    if rFonts is None:
        rFonts = OxmlElement("w:rFonts")
        rPr.insert(0, rFonts)
    # 先清除主题引用，避免覆盖显式设置
    for attr in ["w:asciiTheme", "w:eastAsiaTheme", "w:hAnsiTheme", "w:cstheme"]:
        try:
            del rFonts.attrib[qn(attr)]
        except KeyError:
            pass
    rFonts.set(qn("w:eastAsia"), ea_font)
    rFonts.set(qn("w:ascii"), latin_font)
    rFonts.set(qn("w:hAnsi"), latin_font)
    rFonts.set(qn("w:cs"), latin_font)


def set_line_spacing_exact(style, pt_value):
    """设置固定行距 (lineRule=exact)，单位磅"""
    pf = style.paragraph_format
    pf.line_spacing = Pt(pt_value)
    # 显式写入 XML 确保 exact 规则
    pPr = style.element.get_or_add_pPr()
    spacing = pPr.find(qn("w:spacing"))
    if spacing is None:
        spacing = OxmlElement("w:spacing")
        pPr.append(spacing)
    spacing.set(qn("w:lineRule"), "exact")
    spacing.set(qn("w:line"), str(int(pt_value * 20)))   # twips = pt × 20


def set_paragraph_spacing(style, before=0, after=0):
    """段前 / 段后间距 (磅)，并清除 Word 自动间距"""
    pf = style.paragraph_format
    pf.space_before = Pt(before)
    pf.space_after = Pt(after)

    # XML 层面写入段前段后间距（twips = pt × 20）
    pPr = style.element.get_or_add_pPr()
    spacing = pPr.find(qn("w:spacing"))
    if spacing is None:
        spacing = OxmlElement("w:spacing")
        pPr.append(spacing)
    spacing.set(qn("w:before"), str(int(before * 20)))
    spacing.set(qn("w:after"), str(int(after * 20)))
    spacing.set(qn("w:beforeAutospacing"), "0")
    spacing.set(qn("w:afterAutospacing"), "0")


def set_alignment(style, alignment):
    style.paragraph_format.alignment = alignment


def set_first_line_indent(style, chars):
    """首行缩进 (字符数)，通过 XML w:firstLineChars 设置，符合 GB/T 9704 要求"""
    # 清除 python-docx 的绝对缩进
    style.paragraph_format.first_line_indent = None
    # XML 层面设置字符缩进
    pPr = style.element.get_or_add_pPr()
    ind = pPr.find(qn("w:ind"))
    if ind is None:
        ind = OxmlElement("w:ind")
        pPr.append(ind)
    ind.set(qn("w:firstLineChars"), str(int(chars * 100)))  # 100 = 1 字符


def remove_keep_flags(style):
    """移除样式的「段中不分页」(w:keepLines) 与「与下段同页」(w:keepNext)。

    这两种分页控制来自 python-docx 内建模板自带在 Heading 系列样式上，
    会导致 Word 中段落「段中不分页」「与下段同页」被勾选，与公文排版习惯不符
    （公文段落间不应做这种分页粘连）。在生成模板后统一清除。
    """
    pPr = style.element.find(qn("w:pPr"))
    if pPr is None:
        return
    for tag in ("w:keepLines", "w:keepNext", "w:keep_with_next"):
        elem = pPr.find(qn(tag))
        if elem is not None:
            pPr.remove(elem)


# ═══════════════════════════════════════════════════════════════════
# 创建文档
# ═══════════════════════════════════════════════════════════════════
doc = Document()

# ── 页面设置 ──
# A4: 210×297mm，标准公文页边距
for section in doc.sections:
    section.page_width  = Cm(21.0)
    section.page_height = Cm(29.7)
    section.top_margin    = Cm(3.7)   # 37mm±1mm
    section.bottom_margin = Cm(3.5)   # 35mm±1mm
    section.left_margin   = Cm(2.8)   # 28mm±1mm
    section.right_margin  = Cm(2.6)   # 26mm±1mm

# ── 删除文档默认空段落 ──
for p in doc.paragraphs:
    p._element.getparent().remove(p._element)

# ═══════════════════════════════════════════════════════════════════
# 样式定义
# ═══════════════════════════════════════════════════════════════════

# ── Normal (正文) ──
# 3号仿宋_GB2312，固定行距28磅，首行缩进2字符
# 注意: 请确认系统已安装「仿宋_GB2312」字体
normal = doc.styles["Normal"]
set_style_font(normal, "Times New Roman", "仿宋_GB2312", 16)
set_line_spacing_exact(normal, 28)
set_paragraph_spacing(normal, 0, 0)
set_alignment(normal, WD_ALIGN_PARAGRAPH.JUSTIFY)
set_first_line_indent(normal, 2)

# ── FirstParagraph / BodyText (Pandoc 正文段落实际使用的样式) ──
# Pandoc 转换时正文段落引用 FirstParagraph / BodyText 样式，
# 必须显式定义与 Normal 一致，否则 Word 回退显示异常
from docx.enum.style import WD_STYLE_TYPE
for _style_name in ("First Paragraph", "Body Text"):
    _style = doc.styles.get_by_id(None) if False else None
    try:
        _style = doc.styles[_style_name]
    except KeyError:
        _style = doc.styles.add_style(_style_name, WD_STYLE_TYPE.PARAGRAPH)
        _style.base_style = doc.styles["Normal"]
    set_style_font(_style, "Times New Roman", "仿宋_GB2312", 16)
    set_line_spacing_exact(_style, 28)
    set_paragraph_spacing(_style, 0, 0)
    set_alignment(_style, WD_ALIGN_PARAGRAPH.JUSTIFY)
    set_first_line_indent(_style, 2)

# ── Title (公文标题) ──
# 2号方正小标宋简体，居中，无缩进（标题后空行由后处理插入空段落实现）
title_style = doc.styles["Title"]
set_style_font(title_style, "Times New Roman", "方正小标宋简体", 22, bold=False)
set_line_spacing_exact(title_style, 28)
set_paragraph_spacing(title_style, 0, 0)
set_alignment(title_style, WD_ALIGN_PARAGRAPH.CENTER)
set_first_line_indent(title_style, 0)

# 移除 Title 样式默认的底部横线（Word 模板自带蓝色实线）
pPr = title_style.element.get_or_add_pPr()
pBdr = pPr.find(qn("w:pBdr"))
if pBdr is not None:
    pPr.remove(pBdr)

# ── Heading 1 (一级标题) ──
# 三号黑体，首行缩进2字符，序号 "一、"
h1 = doc.styles["Heading 1"]
set_style_font(h1, "Times New Roman", "黑体", 16, bold=False)
set_line_spacing_exact(h1, 28)
set_paragraph_spacing(h1, 0, 0)
set_alignment(h1, WD_ALIGN_PARAGRAPH.JUSTIFY)
set_first_line_indent(h1, 2)          # ✅ 修正：标题也缩进

# ── Heading 2 (二级标题) ──
# 三号楷体_GB2312，首行缩进2字符，序号 "(二)"
h2 = doc.styles["Heading 2"]
set_style_font(h2, "Times New Roman", "楷体_GB2312", 16, bold=False)
set_line_spacing_exact(h2, 28)
set_paragraph_spacing(h2, 0, 0)
set_alignment(h2, WD_ALIGN_PARAGRAPH.JUSTIFY)
set_first_line_indent(h2, 2)          # ✅ 修正

# ── Heading 3 (三级标题) ──
# 三号仿宋_GB2312，首行缩进2字符，序号 "3."
h3 = doc.styles["Heading 3"]
set_style_font(h3, "Times New Roman", "仿宋_GB2312", 16, bold=False)
set_line_spacing_exact(h3, 28)
set_paragraph_spacing(h3, 0, 0)
set_alignment(h3, WD_ALIGN_PARAGRAPH.JUSTIFY)
set_first_line_indent(h3, 2)          # ✅ 修正

# ── Heading 4 (四级标题) ──
# 三号仿宋_GB2312，首行缩进2字符，序号 "(4)"
h4 = doc.styles["Heading 4"]
set_style_font(h4, "Times New Roman", "仿宋_GB2312", 16, bold=False)
set_line_spacing_exact(h4, 28)
set_paragraph_spacing(h4, 0, 0)
set_alignment(h4, WD_ALIGN_PARAGRAPH.JUSTIFY)
set_first_line_indent(h4, 2)          # ✅ 修正

# ═══════════════════════════════════════════════════════════════════
# 尝试创建特殊格式的自定义样式 (Pandoc 不会自动映射，需手动或 Lua)
# ═══════════════════════════════════════════════════════════════════

def add_custom_style(doc, name, base_style, latin_font, ea_font, size_pt,
                     indent_chars=2, alignment=WD_ALIGN_PARAGRAPH.JUSTIFY, bold=False):
    """向文档添加自定义段落样式"""
    from docx.enum.style import WD_STYLE_TYPE
    style = doc.styles.add_style(name, WD_STYLE_TYPE.PARAGRAPH)
    style.base_style = doc.styles[base_style]
    set_style_font(style, latin_font, ea_font, size_pt, bold=bold)
    set_line_spacing_exact(style, 28)
    set_paragraph_spacing(style, 0, 0)
    set_alignment(style, alignment)
    set_first_line_indent(style, indent_chars)
    return style

# 主送机关 (顶格，3号仿宋_GB2312) — 需转换后手动应用
add_custom_style(doc, "Addressee", "Normal",
                 "Times New Roman", "仿宋_GB2312", 16,
                 indent_chars=0, alignment=WD_ALIGN_PARAGRAPH.LEFT)

# 发文机关署名 (右对齐，3号仿宋_GB2312)
add_custom_style(doc, "Signature", "Normal",
                 "Times New Roman", "仿宋_GB2312", 16,
                 indent_chars=0, alignment=WD_ALIGN_PARAGRAPH.RIGHT)

# 成文日期 (右空四字，相当于右对齐加右缩进，3号仿宋_GB2312)
date_style = add_custom_style(doc, "Date", "Normal",
                              "Times New Roman", "仿宋_GB2312", 16,
                              indent_chars=0, alignment=WD_ALIGN_PARAGRAPH.RIGHT)
# 右缩进约4字符 = 64pt
date_style.paragraph_format.right_indent = Pt(64)

# ═══════════════════════════════════════════════════════════════════
# 添加示例内容 (仅用于预览样式；Pandoc 转换时只用样式定义)
# ═══════════════════════════════════════════════════════════════════

def add_body(text):
    """添加正文段落"""
    doc.add_paragraph(text, style="Normal")

# ── 标题 ──
doc.add_paragraph("关于进一步加强政务信息化建设工作的通知", style="Title")

# ── 主送机关 (参考 GB/T 9704 顶格) ──
p = doc.add_paragraph(
    "各省、自治区、直辖市人民政府，国务院各部委、各直属机构：",
    style="Addressee"
)

# ── 正文 ──
add_body(
    "为深入贯彻落实党中央、国务院关于数字中国建设的决策部署，加快推进政务信息化建设，"
    "提升政府治理能力和公共服务水平，现就有关事项通知如下。"
)

# ── 一级标题 ──
doc.add_paragraph("一、总体要求", style="Heading 1")
add_body(
    "以习近平新时代中国特色社会主义思想为指导，全面贯彻党的二十大精神，"
    "坚持以人民为中心的发展思想，以数字化转型整体驱动生产方式、生活方式和治理方式变革。"
)

# ── 二级标题 ──
doc.add_paragraph("（一）基本原则", style="Heading 2")
add_body(
    "坚持统筹规划、分步实施。按照全国一盘棋的思路，加强顶层设计，统一标准规范，"
    "有计划、有步骤地推进政务信息化建设。"
)

doc.add_paragraph("（二）发展目标", style="Heading 2")
add_body(
    "到2026年底，基本建成覆盖全国、互联互通、集约高效的政务信息化体系，"
    "政务数据共享开放水平显著提升，数字政府建设取得明显成效。"
)

# ── 一级标题 2 ──
doc.add_paragraph("二、重点任务", style="Heading 1")
add_body("围绕以下三个方面的重点任务，统筹推进各项工作落实。")

# ── 三级标题 ──
doc.add_paragraph("3. 构建一体化政务数据平台", style="Heading 3")
add_body(
    "依托国家电子政务外网，建设统一的数据共享交换平台，实现跨层级、跨地域、跨系统、"
    "跨部门、跨业务的数据共享和业务协同。具体包括以下方面："
)

# ── 四级标题 ──
doc.add_paragraph("（1）完善数据资源目录", style="Heading 4")
add_body(
    "按照统一标准编制政务数据资源目录，明确数据来源、更新频率、共享范围和开放属性，"
    "形成覆盖全面、动态更新的数据资源体系。"
)

doc.add_paragraph("（2）推进数据共享开放", style="Heading 4")
add_body(
    "建立健全数据共享协调机制，推动公共数据向社会开放，优先开放交通、医疗、教育、"
    "环保等民生密切相关领域的数据资源。"
)

# ── 三级标题 2 ──
doc.add_paragraph("4. 提升政务服务效能", style="Heading 3")
add_body(
    "深化'互联网+政务服务'，推动更多政务服务事项网上办、掌上办、一次办，"
    "持续优化营商环境。"
)

doc.add_paragraph("（1）拓展线上服务范围", style="Heading 4")
add_body(
    "到2026年6月底前，实现县级以上政务服务事项网上可办率不低于95%，"
    "全程网办率不低于80%。"
)

doc.add_paragraph("（2）优化服务流程", style="Heading 4")
add_body(
    "推行'一件事一次办'改革，围绕企业和群众高频办事场景，整合优化业务流程，"
    "减少申报材料，压缩办理时限。"
)

# ── 一级标题 3 ──
doc.add_paragraph("三、保障措施", style="Heading 1")

doc.add_paragraph("（一）加强组织领导", style="Heading 2")
add_body(
    "各地区、各部门要高度重视政务信息化建设工作，主要负责同志要亲自抓，"
    "明确分管领导和责任部门，建立健全工作机制。"
)

doc.add_paragraph("（二）强化资金保障", style="Heading 2")
add_body(
    "各级财政部门要将政务信息化建设经费纳入年度预算，保障项目建设和运维需要。"
    "鼓励社会资本参与政务信息化建设运营。"
)

doc.add_paragraph("（三）加强考核评估", style="Heading 2")
add_body(
    "建立政务信息化建设考核评估机制，定期通报工作进展，"
    "对工作推进不力的地区和部门进行约谈督促。"
)

# ── 结尾正文 ──
add_body(
    "各地区、各部门要根据本通知要求，结合实际制定具体实施方案，"
    "于2026年X月X日前报送工作推进情况。"
)

# ── 署名 / 日期 (特殊格式) ──
doc.add_paragraph("国务院办公厅", style="Signature")
doc.add_paragraph("2026年X月X日", style="Date")

# ═══════════════════════════════════════════════════════════════════
# 保存
# ═══════════════════════════════════════════════════════════════════

# ── 清除全部样式的「段中不分页」/「与下段同页」──
# python-docx 内建模板的 Heading1-9 自带 w:keepLines(段中不分页) 和
# w:keepNext(与下段同页)，公文段落不应做这种分页粘连，统一清除。
_removed = {"w:keepLines": 0, "w:keepNext": 0, "w:keep_with_next": 0}
for _style in doc.styles:
    if _style.type != 1:  # 只处理段落样式
        continue
    _pPr = _style.element.find(qn("w:pPr"))
    if _pPr is None:
        continue
    for _tag in ("w:keepLines", "w:keepNext", "w:keep_with_next"):
        _el = _pPr.find(qn(_tag))
        if _el is not None:
            _pPr.remove(_el)
            _removed[_tag] += 1
if any(_removed.values()):
    print(f"[INFO] 已清除样式的分页粘连属性: {_removed}")

saved_path = safe_save(doc, OUTPUT_PATH)
print(f"[OK] 参考文档: {saved_path}")
print(f"     大小: {os.path.getsize(saved_path)} bytes")
if saved_path != OUTPUT_PATH:
    print(f"     ⚠ 原文件 {OUTPUT_PATH} 被占用，已保存为备用文件。")
    print(f"     关闭原文件后可将备用文件重命名为 {os.path.basename(OUTPUT_PATH)}")

print("""
样式一览 (GB/T 9704-2012):
  ┌──────────────┬──────┬──────────────────┬──────────┬──────────┬──────────┐
  │ 样式         │ 字号 │ 中文字体         │ 西文字体 │ 行距     │ 缩进     │
  ├──────────────┼──────┼──────────────────┼──────────┼──────────┼──────────┤
  │ Title        │ 二号 │ 方正小标宋简体   │ TNR      │ 固定28磅 │ 无       │
  │ Normal       │ 三号 │ 仿宋_GB2312      │ TNR      │ 固定28磅 │ 2字符    │
  │ Heading 1    │ 三号 │ 黑体             │ TNR      │ 固定28磅 │ 2字符    │
  │ Heading 2    │ 三号 │ 楷体_GB2312      │ TNR      │ 固定28磅 │ 2字符    │
  │ Heading 3    │ 三号 │ 仿宋_GB2312      │ TNR      │ 固定28磅 │ 2字符    │
  │ Heading 4    │ 三号 │ 仿宋_GB2312      │ TNR      │ 固定28磅 │ 2字符    │
  │ Addressee *  │ 三号 │ 仿宋_GB2312      │ TNR      │ 固定28磅 │ 无(顶格) │
  │ Signature *  │ 三号 │ 仿宋_GB2312      │ TNR      │ 固定28磅 │ 无(右对齐)│
  │ Date *       │ 三号 │ 仿宋_GB2312      │ TNR      │ 固定28磅 │ 无(右对齐)│
  └──────────────┴──────┴──────────────────┴──────────┴──────────┴──────────┘
  * 标记的为自定义样式，Pandoc 不会自动映射，需转换后手动应用。

页面设置:
  A4 (210×297mm)  |  上 3.7 / 下 3.5 / 左 2.8 / 右 2.6 cm
""")
