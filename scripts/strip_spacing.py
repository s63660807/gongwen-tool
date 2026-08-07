#!/usr/bin/env python3
"""
后处理 Pandoc 生成的 docx：
1. 彻底清除所有段前/段后间距（document.xml 段落层 + styles.xml 的 docDefaults 和所有样式层）
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


def strip_spacing_in_tree(tree):
    """清除 XML 树中所有 w:spacing 元素的段前/段后间距（保留 line/lineRule）"""
    SPACING_ATTRS = ['before', 'after', 'beforeAutospacing', 'afterAutospacing']
    for spacing in tree.iter(f'{{{NS}}}spacing'):
        for attr in SPACING_ATTRS:
            spacing.attrib[f'{{{NS}}}{attr}'] = '0'


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

    # 1. 处理 document.xml — 段落级 spacing + 引号修复
    doc_path = os.path.join(tmp, 'word/document.xml')
    if os.path.exists(doc_path):
        tree = etree.parse(doc_path)
        strip_spacing_in_tree(tree)
        fix_quotes_in_tree(tree)
        tree.write(doc_path, xml_declaration=True, encoding='UTF-8')

    # 2. 处理 styles.xml — docDefaults + 所有样式
    styles_path = os.path.join(tmp, 'word/styles.xml')
    if os.path.exists(styles_path):
        tree = etree.parse(styles_path)
        strip_spacing_in_tree(tree)
        tree.write(styles_path, xml_declaration=True, encoding='UTF-8')

    # 重新打包
    with zipfile.ZipFile(output_path, 'w', zipfile.ZIP_DEFLATED) as zout:
        for root_dir, dirs, files in os.walk(tmp):
            for f in files:
                full = os.path.join(root_dir, f)
                arcname = os.path.relpath(full, tmp)
                zout.write(full, arcname)

    shutil.rmtree(tmp)
    print(f"[OK] 段间距已清零、引号已修复: {output_path}")
    return output_path


if __name__ == '__main__':
    if len(sys.argv) < 2:
        print("用法: python strip_spacing.py <input.docx> [output.docx]")
        sys.exit(1)

    inp = sys.argv[1]
    out = sys.argv[2] if len(sys.argv) > 2 else None
    strip_spacing(inp, out)
