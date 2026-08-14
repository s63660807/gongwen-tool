#!/usr/bin/env python3
"""
公文格式转换工具 — Markdown → 标准公文 .docx (GB/T 9704-2012)

用法:
    convert.py <输入.md> [输出.docx]

    如果只提供 .md 文件，默认输出同名 .docx（同目录）。

依赖:
    pandoc.exe  — 放在本程序同目录下，或已加入系统 PATH
    reference.docx  — 首次运行自动生成（符合 GB/T 9704-2012）
"""

import sys
import os
import subprocess
import shutil


def exe_dir():
    """获取程序所在目录（PyInstaller 打包后仍正确）"""
    if getattr(sys, 'frozen', False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


def find_pandoc():
    """查找 pandoc 可执行文件：同目录 > 常见安装路径 > PATH"""
    local = os.path.join(exe_dir(), "pandoc.exe")
    if os.path.exists(local):
        return local

    # 常见安装路径
    common_paths = [
        os.path.join(os.environ.get("ProgramFiles", "C:\\Program Files"), "Pandoc", "pandoc.exe"),
        os.path.join(os.environ.get("LOCALAPPDATA", ""), "Pandoc", "pandoc.exe"),
        r"C:\Program Files\Pandoc\pandoc.exe",
        r"C:\Program Files (x86)\Pandoc\pandoc.exe",
    ]
    for p in common_paths:
        if os.path.exists(p):
            return p

    found = shutil.which("pandoc")
    if found:
        return found
    return None


def generate_reference(output_path):
    """生成 reference.docx 模板（内嵌 create_reference 逻辑，无需外部依赖）"""
    try:
        from create_reference import OUTPUT_PATH, safe_save
    except ImportError:
        pass

    from docx import Document
    from docx.shared import Pt, Cm, RGBColor
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.oxml.ns import qn
    from docx.oxml import OxmlElement

    def safe_save_local(doc, path):
        try:
            doc.save(path)
        except PermissionError:
            base, ext = os.path.splitext(path)
            alt = base + "_new" + ext
            doc.save(alt)
            return alt
        return path

    def set_style_font(style, latin_font, ea_font, size_pt, bold=False, italic=False):
        style.font.size = Pt(size_pt)
        style.font.bold = bold
        style.font.italic = italic
        style.font.name = latin_font
        style.font.color.rgb = RGBColor(0, 0, 0)
        rPr = style.element.get_or_add_rPr()
        rFonts = rPr.find(qn("w:rFonts"))
        if rFonts is None:
            rFonts = OxmlElement("w:rFonts")
            rPr.insert(0, rFonts)
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
        pf = style.paragraph_format
        pf.line_spacing = Pt(pt_value)
        pPr = style.element.get_or_add_pPr()
        spacing = pPr.find(qn("w:spacing"))
        if spacing is None:
            spacing = OxmlElement("w:spacing")
            pPr.append(spacing)
        spacing.set(qn("w:lineRule"), "exact")
        spacing.set(qn("w:line"), str(int(pt_value * 20)))

    def set_paragraph_spacing(style, before=0, after=0):
        pf = style.paragraph_format
        pf.space_before = Pt(before)
        pf.space_after = Pt(after)
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
        style.paragraph_format.first_line_indent = None
        pPr = style.element.get_or_add_pPr()
        ind = pPr.find(qn("w:ind"))
        if ind is None:
            ind = OxmlElement("w:ind")
            pPr.append(ind)
        ind.set(qn("w:firstLineChars"), str(int(chars * 100)))

    doc = Document()
    for section in doc.sections:
        section.page_width = Cm(21.0)
        section.page_height = Cm(29.7)
        section.top_margin = Cm(3.7)
        section.bottom_margin = Cm(3.5)
        section.left_margin = Cm(2.8)
        section.right_margin = Cm(2.6)

    for p in doc.paragraphs:
        p._element.getparent().remove(p._element)

    # Normal
    normal = doc.styles["Normal"]
    set_style_font(normal, "Times New Roman", "仿宋_GB2312", 16)
    set_line_spacing_exact(normal, 28)
    set_paragraph_spacing(normal, 0, 0)
    set_alignment(normal, WD_ALIGN_PARAGRAPH.JUSTIFY)
    set_first_line_indent(normal, 2)

    # FirstParagraph / BodyText（Pandoc 正文实际引用的样式，需与 Normal 一致）
    from docx.enum.style import WD_STYLE_TYPE
    for _style_name in ("First Paragraph", "Body Text"):
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

    # Title（标题后空行由后处理插入空段落实现）
    title_style = doc.styles["Title"]
    set_style_font(title_style, "Times New Roman", "方正小标宋简体", 22, bold=False)
    set_line_spacing_exact(title_style, 28)
    set_paragraph_spacing(title_style, 0, 0)
    set_alignment(title_style, WD_ALIGN_PARAGRAPH.CENTER)
    set_first_line_indent(title_style, 0)
    pPr = title_style.element.get_or_add_pPr()
    pBdr = pPr.find(qn("w:pBdr"))
    if pBdr is not None:
        pPr.remove(pBdr)

    # Heading 1
    h1 = doc.styles["Heading 1"]
    set_style_font(h1, "Times New Roman", "黑体", 16, bold=False)
    set_line_spacing_exact(h1, 28)
    set_paragraph_spacing(h1, 0, 0)
    set_alignment(h1, WD_ALIGN_PARAGRAPH.JUSTIFY)
    set_first_line_indent(h1, 2)

    # Heading 2
    h2 = doc.styles["Heading 2"]
    set_style_font(h2, "Times New Roman", "楷体_GB2312", 16, bold=False)
    set_line_spacing_exact(h2, 28)
    set_paragraph_spacing(h2, 0, 0)
    set_alignment(h2, WD_ALIGN_PARAGRAPH.JUSTIFY)
    set_first_line_indent(h2, 2)

    # Heading 3
    h3 = doc.styles["Heading 3"]
    set_style_font(h3, "Times New Roman", "仿宋_GB2312", 16, bold=False)
    set_line_spacing_exact(h3, 28)
    set_paragraph_spacing(h3, 0, 0)
    set_alignment(h3, WD_ALIGN_PARAGRAPH.JUSTIFY)
    set_first_line_indent(h3, 2)

    # Heading 4
    h4 = doc.styles["Heading 4"]
    set_style_font(h4, "Times New Roman", "仿宋_GB2312", 16, bold=False)
    set_line_spacing_exact(h4, 28)
    set_paragraph_spacing(h4, 0, 0)
    set_alignment(h4, WD_ALIGN_PARAGRAPH.JUSTIFY)
    set_first_line_indent(h4, 2)

    # 清除全部样式的「段中不分页」(w:keepLines) 与「与下段同页」(w:keepNext)
    # python-docx 内建模板的 Heading1-9 自带这两个分页属性，公文段落不应做分页粘连。
    for _s in doc.styles:
        if getattr(_s, "type", 1) != 1:  # 只处理段落样式
            continue
        _pp = _s.element.find(qn("w:pPr"))
        if _pp is None:
            continue
        for _tag in ("w:keepLines", "w:keepNext", "w:keep_with_next"):
            _el = _pp.find(qn(_tag))
            if _el is not None:
                _pp.remove(_el)

    return safe_save_local(doc, output_path)


def strip_spacing_docx(docx_path):
    """清除 docx 中所有段前/段后间距，Title 后插入空行，修复中文引号方向，拆分标题内夹带正文"""
    import zipfile
    import tempfile
    from lxml import etree
    import re as _re

    NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
    DOUBLE_QUOTES = {'"': True, '\u201c': True, '\u201d': True}
    SINGLE_QUOTES = {"'": True, '\u2018': True, '\u2019': True}
    tmp = tempfile.mkdtemp()

    # ── 标题内夹带正文拆分（修复正文被套用标题格式） ──
    _CN_NUM = '一二三四五六七八九'
    _RE_NUM = _re.compile(r'^\d+[\.、]')
    _RE_H1 = _re.compile(rf'^[{_CN_NUM}十]+、')
    _RE_H2 = _re.compile(rf'^（[{_CN_NUM}十]+）')

    def _strip_prefix(full, style_val):
        """去掉标题序号前缀（「一、」「（一）」「1. 」），返回标题正文；非标题开头原样返回。"""
        if style_val in ('Heading1', 'Heading2', 'Heading3', 'Compact'):
            for m in (_RE_NUM.match(full), _RE_H1.match(full), _RE_H2.match(full)):
                if m:
                    return full[len(m.group(0)):]
        return full

    def _set_text(para, new_text):
        """把段落文本替换为单一 new_text，保留首 run 格式。"""
        runs = list(para.iter(f'{{{NS}}}r'))
        if runs:
            r0 = runs[0]
            for tn in list(r0.findall(f'{{{NS}}}t')):
                r0.remove(tn)
            t = etree.SubElement(r0, f'{{{NS}}}t')
            t.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
            t.text = new_text
            for r in runs[1:]:
                para.remove(r)
        else:
            r = etree.Element(f'{{{NS}}}r')
            t = etree.SubElement(r, f'{{{NS}}}t')
            t.text = new_text
            pPr = para.find(f'{{{NS}}}pPr')
            if pPr is not None:
                pPr.addnext(r)
            else:
                para.append(r)

    def _split_heading_inline_body(tree):
        """Heading1/2/3（及 Compact）段落若在序号后第一个「。」后仍有正文，则拆分标题与正文。"""
        body = tree.find(f'{{{NS}}}body')
        if body is None:
            return
        for para in list(body.iter(f'{{{NS}}}p')):
            pPr = para.find(f'{{{NS}}}pPr')
            if pPr is None:
                continue
            pStyle = pPr.find(f'{{{NS}}}pStyle')
            if pStyle is None:
                continue
            style_val = pStyle.get(f'{{{NS}}}val')
            if style_val not in ('Heading1', 'Heading2', 'Heading3', 'Compact'):
                continue
            texts = [t for t in para.iter(f'{{{NS}}}t') if t.text]
            full = ''.join(t.text for t in texts)
            title_body = _strip_prefix(full, style_val)
            if title_body is full and style_val == 'Compact':
                continue  # Compact 非数字列表则跳过
            dot_idx = title_body.find('。')
            if dot_idx < 0:
                continue
            after = title_body[dot_idx + 1:].strip()
            if not after:
                continue
            serial_len = len(full) - len(title_body)
            title_txt = full[: serial_len + dot_idx + 1]
            body_txt = full[serial_len + dot_idx + 1:]
            new_p = etree.Element(f'{{{NS}}}p')
            new_pPr = etree.SubElement(new_p, f'{{{NS}}}pPr')
            new_st = etree.SubElement(new_pPr, f'{{{NS}}}pStyle')
            new_st.set(f'{{{NS}}}val', 'FirstParagraph')
            new_r = etree.SubElement(new_p, f'{{{NS}}}r')
            new_t = etree.SubElement(new_r, f'{{{NS}}}t')
            new_t.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
            new_t.text = body_txt
            _set_text(para, title_txt)
            para.addnext(new_p)

    with zipfile.ZipFile(docx_path, 'r') as z:
        z.extractall(tmp)

    for xml_name in ['word/document.xml', 'word/styles.xml']:
        xml_path = os.path.join(tmp, xml_name)
        if not os.path.exists(xml_path):
            continue
        tree = etree.parse(xml_path)
        for spacing in tree.iter(f'{{{NS}}}spacing'):
            for attr in ['before', 'after', 'beforeAutospacing', 'afterAutospacing']:
                spacing.attrib[f'{{{NS}}}{attr}'] = '0'
        # 移除 rFonts 的 theme 引用属性（避免主题字体覆盖公文要求字体）
        for rfonts in tree.iter(f'{{{NS}}}rFonts'):
            for attr in ['asciiTheme', 'eastAsiaTheme', 'hAnsiTheme', 'cstheme']:
                if rfonts.get(f'{{{NS}}}{attr}') is not None:
                    del rfonts.attrib[f'{{{NS}}}{attr}']
        tree.write(xml_path, xml_declaration=True, encoding='UTF-8')

    # 修复引号方向 + Title 前后插入空行 + 拆分标题内夹带正文（仅 document.xml 正文）
    doc_xml = os.path.join(tmp, 'word/document.xml')
    if os.path.exists(doc_xml):
        tree = etree.parse(doc_xml)
        # 拆分标题内夹带正文（修复正文被套用标题格式）
        _split_heading_inline_body(tree)
        # Title 前后各插入一个空段落（回车空行），已存在则跳过
        body = tree.find(f'{{{NS}}}body')
        if body is not None:
            paras = list(body.iter(f'{{{NS}}}p'))
            for i, para in enumerate(paras):
                pPr = para.find(f'{{{NS}}}pPr')
                if pPr is None:
                    continue
                pStyle = pPr.find(f'{{{NS}}}pStyle')
                if pStyle is None or pStyle.get(f'{{{NS}}}val') != 'Title':
                    continue

                def is_blank(p):
                    return p is not None and not p.findall(f'{{{NS}}}r')

                prev_para = paras[i - 1] if i > 0 else None
                if prev_para is None or not is_blank(prev_para):
                    blank = etree.Element(f'{{{NS}}}p')
                    para.addprevious(blank)

                next_para = paras[i + 1] if i + 1 < len(paras) else None
                if next_para is None or not is_blank(next_para):
                    blank = etree.Element(f'{{{NS}}}p')
                    para.addnext(blank)
                break
        for para in tree.iter(f'{{{NS}}}p'):
            text_nodes = [t for t in para.iter(f'{{{NS}}}t') if t.text]
            if not text_nodes:
                continue
            open_double = False
            open_single = False
            for node in text_nodes:
                chars = list(node.text)
                for i, ch in enumerate(chars):
                    if ch in DOUBLE_QUOTES:
                        chars[i] = '\u201c' if not open_double else '\u201d'
                        open_double = not open_double
                    elif ch in SINGLE_QUOTES:
                        chars[i] = '\u2018' if not open_single else '\u2019'
                        open_single = not open_single
                node.text = ''.join(chars)
        tree.write(doc_xml, xml_declaration=True, encoding='UTF-8')

    # 替换原文件
    with zipfile.ZipFile(docx_path, 'w', zipfile.ZIP_DEFLATED) as zout:
        for root_dir, dirs, files in os.walk(tmp):
            for f in files:
                full = os.path.join(root_dir, f)
                arcname = os.path.relpath(full, tmp)
                zout.write(full, arcname)

    shutil.rmtree(tmp)


def main():
    if len(sys.argv) < 2:
        print("公文格式转换工具 — Markdown → 标准公文 .docx (GB/T 9704-2012)")
        print()
        print("用法:")
        print("    convert.exe <输入.md> [输出.docx]")
        print()
        print("    - 输入:  Markdown 文件，YAML front matter 中写 title")
        print("    - 输出:  标准公文格式 .docx（默认与输入同名、同目录）")
        print()
        print("依赖: pandoc.exe 需放在本程序同目录下，或已加入系统 PATH")
        sys.exit(0)

    md_path = sys.argv[1]
    if len(sys.argv) >= 3:
        docx_path = sys.argv[2]
    else:
        base, _ = os.path.splitext(md_path)
        docx_path = base + ".docx"

    if not os.path.exists(md_path):
        print(f"[错误] 文件不存在: {md_path}")
        sys.exit(1)

    # 1. 查找 pandoc
    pandoc = find_pandoc()
    if not pandoc:
        print("[错误] 找不到 pandoc.exe！")
        print("       请将 pandoc.exe 放在本程序同目录下，或将其加入系统 PATH。")
        print("       下载地址: https://pandoc.org/installing.html")
        sys.exit(1)

    # 2. 确保 reference.docx 存在
    ref_path = os.path.join(exe_dir(), "reference.docx")
    if not os.path.exists(ref_path):
        print("[信息] 未找到 reference.docx，正在自动生成...")
        ref_path = generate_reference(ref_path)
        print(f"[信息] 已生成: {ref_path}")

    # 3. Pandoc 转换（-f markdown-smart 禁用智能引号，避免中文引号被误转为右引号）
    print(f"[转换] {md_path} → {docx_path}")
    result = subprocess.run(
        [pandoc, "-f", "markdown-smart", md_path, "-o", docx_path, "--reference-doc=" + ref_path],
        capture_output=True, text=True
    )
    if result.returncode != 0:
        print(f"[错误] Pandoc 转换失败: {result.stderr}")
        sys.exit(1)

    # 4. 清除段间距
    print("[清理] 清除段前/段后间距...")
    strip_spacing_docx(docx_path)

    print(f"[完成] {docx_path}")
    os.startfile(docx_path)  # 自动打开生成的文档


if __name__ == "__main__":
    main()
