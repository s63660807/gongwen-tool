#!/usr/bin/env python3
"""
后处理 Pandoc 生成的 docx：
1. 清除段前/段后间距（保留标题样式的段前段后空行；清除 document.xml 段落层 + styles.xml 的 docDefaults 和其他样式层）
2. 修复中文引号：将 ASCII 直引号/全右弯引号按成对规则替换为规范的左引号“和右引号”
"""
import sys
import zipfile
import os
import tempfile
import shutil
from lxml import etree

NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'

# 引号字符映射
DOUBLE_QUOTES = {'"': True, '\u201c': True, '\u201d': True}   # " “ ”
SINGLE_QUOTES = {"'": True, '\u2018': True, '\u2019': True}   # ' ‘ ’


def strip_spacing_in_tree(tree, xml_path=''):
    """清除 XML 树中所有 w:spacing 的段前/段后间距（保留 line/lineRule）

    - document.xml：段落级 spacing
    - styles.xml：docDefaults + 所有样式层
    """
    SPACING_ATTRS = ['before', 'after', 'beforeAutospacing', 'afterAutospacing']
    for spacing in tree.iter(f'{{{NS}}}spacing'):
        for attr in SPACING_ATTRS:
            spacing.attrib[f'{{{NS}}}{attr}'] = '0'


def strip_font_theme_attrs(tree):
    """移除所有 rFonts 的 theme 引用属性（asciiTheme/eastAsiaTheme/hAnsiTheme/cstheme）。

    theme 属性优先级高于显式字体名，会导致 Word 用主题字体（如宋体）覆盖
    公文要求的黑体/楷体/小标宋。Pandoc 生成的 docx 常带这些属性，必须清除。
    """
    THEME_ATTRS = ['asciiTheme', 'eastAsiaTheme', 'hAnsiTheme', 'cstheme']
    count = 0
    for rfonts in tree.iter(f'{{{NS}}}rFonts'):
        for attr in THEME_ATTRS:
            if rfonts.get(f'{{{NS}}}{attr}') is not None:
                del rfonts.attrib[f'{{{NS}}}{attr}']
                count += 1
    return count


def strip_keep_flags_in_tree(tree, xml_path=''):
    """移除 XML 树中所有「段中不分页」(w:keepLines) 与「与下段同页」(w:keepNext)。

    兜底清除：即使 reference.docx 模板或 Pandoc 转换遗漏了这两个属性，
    这里也能保证最终 docx 段落不带「段中不分页」「与下段同页」分页粘连。
    """
    count = 0
    for pPr in tree.iter(f'{{{NS}}}pPr'):
        for tag in ('keepLines', 'keepNext', 'keep_with_next'):
            el = pPr.find(f'{{{NS}}}{tag}')
            if el is not None:
                pPr.remove(el)
                count += 1
    return count


def insert_blank_around_title(tree):
    """在公文标题（Title 样式）段落前后各插入一个真正的空段落（回车空行）。

    Title 前、后各空一行；若已存在空段落则跳过对应插入（幂等）。
    """
    body = tree.find(f'{{{NS}}}body')
    if body is None:
        return
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

        # 前一段（Title 之前）插入空段落
        prev_para = paras[i - 1] if i > 0 else None
        if prev_para is None or not is_blank(prev_para):
            blank = etree.Element(f'{{{NS}}}p')
            para.addprevious(blank)

        # 后一段（Title 之后）插入空段落
        next_para = paras[i + 1] if i + 1 < len(paras) else None
        if next_para is None or not is_blank(next_para):
            blank = etree.Element(f'{{{NS}}}p')
            para.addnext(blank)
        return


def fix_quotes_in_tree(tree):
    """修复正文中所有引号方向：按段落顺序交替，奇数个为左引号，偶数个为右引号"""
    for para in tree.iter(f'{{{NS}}}p'):
        # 收集该段落内所有文本节点（按文档顺序）
        text_nodes = []
        for t in para.iter(f'{{{NS}}}t'):
            if t.text:
                text_nodes.append(t)
        if not text_nodes:
            continue

        # 逐字符处理，维护引号开关状态
        open_double = False
        open_single = False
        for node in text_nodes:
            chars = list(node.text)
            for i, ch in enumerate(chars):
                if ch in DOUBLE_QUOTES:
                    if not open_double:
                        chars[i] = '\u201c'
                        open_double = True
                    else:
                        chars[i] = '\u201d'
                        open_double = False
                elif ch in SINGLE_QUOTES:
                    if not open_single:
                        chars[i] = '\u2018'
                        open_single = True
                    else:
                        chars[i] = '\u2019'
                        open_single = False
            node.text = ''.join(chars)


def strip_spacing(input_path, output_path=None):
    if output_path is None:
        base, ext = os.path.splitext(input_path)
        output_path = base + "_clean" + ext

    tmp = tempfile.mkdtemp()
    with zipfile.ZipFile(input_path, 'r') as z:
        z.extractall(tmp)

    nsmap = {'w': NS}

    # 1. 处理 document.xml — 段落级 spacing + 标题前后空行 + 引号修复 + 字体theme清理 + 分页粘连清理
    doc_path = os.path.join(tmp, 'word/document.xml')
    if os.path.exists(doc_path):
        tree = etree.parse(doc_path)
        strip_spacing_in_tree(tree, xml_path='word/document.xml')
        insert_blank_around_title(tree)
        fix_quotes_in_tree(tree)
        strip_font_theme_attrs(tree)
        strip_keep_flags_in_tree(tree, xml_path='word/document.xml')
        tree.write(doc_path, xml_declaration=True, encoding='UTF-8')

    # 2. 处理 styles.xml — docDefaults + 所有样式 + 字体theme清理 + 分页粘连清理
    styles_path = os.path.join(tmp, 'word/styles.xml')
    if os.path.exists(styles_path):
        tree = etree.parse(styles_path)
        strip_spacing_in_tree(tree, xml_path='word/styles.xml')
        strip_font_theme_attrs(tree)
        strip_keep_flags_in_tree(tree, xml_path='word/styles.xml')
        tree.write(styles_path, xml_declaration=True, encoding='UTF-8')

    # 重新打包
    with zipfile.ZipFile(output_path, 'w', zipfile.ZIP_DEFLATED) as zout:
        for root_dir, dirs, files in os.walk(tmp):
            for f in files:
                full = os.path.join(root_dir, f)
                arcname = os.path.relpath(full, tmp)
                zout.write(full, arcname)

    shutil.rmtree(tmp)
    print(f"[OK] 段间距已清零、标题前后空行已插入、引号与字体已修复、分页粘连(段中不分页/与下段同页)已清除: {output_path}")
    return output_path


if __name__ == '__main__':
    if len(sys.argv) < 2:
        print("用法: python strip_spacing.py <input.docx> [output.docx]")
        sys.exit(1)

    inp = sys.argv[1]
    out = sys.argv[2] if len(sys.argv) > 2 else None
    strip_spacing(inp, out)
