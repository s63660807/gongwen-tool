#!/usr/bin/env python3
"""
后处理 Pandoc 生成的 docx：
1. 清除段前/段后间距（保留标题样式的段前段后空行；清除 document.xml 段落层 + styles.xml 的 docDefaults 和其他样式层）
2. 修复中文引号：将 ASCII 直引号/全右弯引号按成对规则替换为规范的左引号“和右引号”
3. 关闭所有段落的孤行控制(widowControl)、段中不分页(keepLines)、与下段同页(keepNext)，避免段尾跳页浪费版面
4. 拆分「标题与正文写在同一自然段」的段落（标题内夹带正文），避免正文被套用标题格式
5. 移除 rFonts 的 theme 引用属性（asciiTheme/eastAsiaTheme/hAnsiTheme/cstheme），避免主题字体覆盖公文要求字体
"""
import sys
import zipfile
import os
import tempfile
import shutil
import re as _re
from lxml import etree

NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'

# 引号字符映射
DOUBLE_QUOTES = {'"': True, '\u201c': True, '\u201d': True}   # " “ ”
SINGLE_QUOTES = {"'": True, '\u2018': True, '\u2019': True}   # ' ‘ ’

# CT_PPr 子元素 schema 顺序（按序插入 pPr 属性，避免 Word 校验报错）
PPR_CHILD_ORDER = [
    'pStyle', 'keepNext', 'keepLines', 'pageBreakBefore', 'framePr',
    'widowControl', 'numPr', 'suppressLineNumbers', 'pBdr', 'shd', 'tabs',
    'suppressAutoHyphens', 'kinsoku', 'wordWrap', 'overflowPunct', 'topLinePunct',
    'autoSpaceDE', 'autoSpaceDN', 'bidi', 'adjustRightInd', 'snapToGrid',
    'spacing', 'ind', 'contextualSpacing', 'mirrorIndents', 'suppressOverlap',
    'jc', 'textDirection', 'textAlignment', 'textboxTightWrap', 'outlineLvl',
    'divId', 'cnfStyle', 'rPr', 'sectPr', 'pPrChange',
]


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


# 中文数字，用于识别「一、二、…」「（一）（二）…」「（十）（廿）…」等公文标题序号
_CN_NUM = '一二三四五六七八九'

# 预编译：数字点/顿号 → 如 "1." 或 "1、"
HEADING_RE_NUM = _re.compile(r'^\d+[\.、]')
HEADING_RE_H1 = _re.compile(rf'^[{_CN_NUM}十]+、')
HEADING_RE_H2 = _re.compile(rf'^（[{_CN_NUM}十]+）')


def _strip_serial_prefix(full, style_val):
    """去掉标题段落开头的公文序号前缀（「一、」「（一）」「1. 」等），返回标题正文字段。

    若开头不是对应层级的序号，则原样返回（调用方据此视为普通段落）。
    """
    if style_val in ('Heading1', 'Heading2', 'Heading3', 'Compact'):
        # Heading3 / Compact 常用数字点「1. 」或数字顿号「1、」
        m = HEADING_RE_NUM.match(full)
        if m:
            return full[len(m.group(0)):]
        # Heading1：一、二、……
        m = HEADING_RE_H1.match(full)
        if m:
            return full[len(m.group(0)):]
        # Heading2 / 括号序列：（一）（二）……（十）（十一）……
        m = HEADING_RE_H2.match(full)
        if m:
            return full[len(m.group(0)):]
    return full


def split_heading_with_inline_body(tree):
    """修复「标题与正文写在同一自然段」导致的正文被套用标题格式问题。

    场景：md 中把二级/三级标题和正文写在同一自然段（如
    `## （一）强化组织领导。各级党组要切实履行主体责任……`），Pandoc
    整段渲染为 Heading 样式，导致正文也跟着标题格式（楷体/黑体）。

    规则：对 Heading1/2/3（及 Compact 列表标题）段落，去掉序号前缀后，
    若找到第一个「。」且其后仍有非空正文，则把标题部分（到句号为止）
    保留为标题，句号后的内容拆成一个新的正文段落（FirstParagraph 样式）。

    幂等：标题本身无句号或句号后无正文时不拆分；重复运行不影响已拆分结果。
    """
    body = tree.find(f'{{{NS}}}body')
    if body is None:
        return 0
    split_count = 0
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
        # 收集段落全部文本
        texts = [t for t in para.iter(f'{{{NS}}}t') if t.text]
        full = ''.join(t.text for t in texts)
        # 去掉序号前缀，得到标题正文字段
        title_body = _strip_serial_prefix(full, style_val)
        if title_body is full:
            # 无序号前缀，但 Heading1/2/3 一般都有；视为普通标题跳过
            # 例外：Compact（列表）严格限定数字开头，无则跳过
            if style_val == 'Compact':
                continue
        # 找第一个句号
        dot_idx = title_body.find('。')
        if dot_idx < 0:
            continue  # 标题本身无句号，正常短标题，跳过
        after = title_body[dot_idx + 1:].strip()
        if not after:
            continue  # 句号后无正文，正常标题，跳过
        # ===== 需要拆分 =====
        # 标题部分 = 序号 + 标题正文到句号；正文部分 = 句号后内容
        serial_len = len(full) - len(title_body)
        title_txt = full[: serial_len + dot_idx + 1]       # 含句号
        body_txt = full[serial_len + dot_idx + 1:]          # 句号后内容
        # 新建正文段落（FirstParagraph 样式）
        new_p = etree.Element(f'{{{NS}}}p')
        new_pPr = etree.Element(f'{{{NS}}}pPr')
        new_style = etree.SubElement(new_pPr, f'{{{NS}}}pStyle')
        new_style.set(f'{{{NS}}}val', 'FirstParagraph')
        new_p.append(new_pPr)
        new_r = etree.SubElement(new_p, f'{{{NS}}}r')
        new_t = etree.SubElement(new_r, f'{{{NS}}}t')
        new_t.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
        new_t.text = body_txt
        # 原段落文本改为标题部分
        _set_para_text(para, title_txt)
        # 新正文段插入原段之后
        para.addnext(new_p)
        split_count += 1
    return split_count


def _set_para_text(para, new_text):
    """把段落内的文本替换为单一文本 new_text，保留第一个 run 的格式。"""
    runs = list(para.iter(f'{{{NS}}}r'))
    if runs:
        r0 = runs[0]
        # 移除 r0 内所有 t，重设第一个
        for tn in list(r0.findall(f'{{{NS}}}t')):
            r0.remove(tn)
        t = etree.SubElement(r0, f'{{{NS}}}t')
        t.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
        t.text = new_text
        # 移除其他 run
        for r in runs[1:]:
            para.remove(r)
    else:
        # 无 run，新建一个并插到 pPr 后
        r = etree.Element(f'{{{NS}}}r')
        t = etree.SubElement(r, f'{{{NS}}}t')
        t.text = new_text
        pPr = para.find(f'{{{NS}}}pPr')
        if pPr is not None:
            pPr.addnext(r)
        else:
            para.append(r)


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


def _insert_ppr_flag(pPr, tag, ns):
    """在 w:pPr 中按 schema 顺序设置/创建元素并置 w:val='0'（幂等）"""
    el = pPr.find(f'{{{ns}}}{tag}')
    if el is None:
        el = etree.Element(f'{{{ns}}}{tag}')
        order = PPR_CHILD_ORDER.index(tag)
        pos = 0
        for child in pPr:
            cname = etree.QName(child).localname
            if cname in PPR_CHILD_ORDER and PPR_CHILD_ORDER.index(cname) > order:
                break
            pos += 1
        pPr.insert(pos, el)
    el.set(f'{{{ns}}}val', '0')


def disable_keep_together(tree, ns=NS):
    """关闭所有段落的孤行控制(widowControl)、段中不分页(keepLines)、与下段同页(keepNext)。

    Word 默认的孤行控制/段中不分页/与下段同页会把段落尾部整段挤到下一页，
    造成页面大量空白；公文压缩篇幅时必须全部关闭，让排版紧凑。
    """
    for pPr in tree.iter(f'{{{ns}}}pPr'):
        for tag in ('keepNext', 'keepLines', 'widowControl'):
            _insert_ppr_flag(pPr, tag, ns)


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

    # 1. 处理 document.xml — 段落级 spacing + 标题内夹带正文拆分 + 标题前后空行 + 引号修复 + 字体theme清理 + 关闭孤行/段中/同页控制
    doc_path = os.path.join(tmp, 'word/document.xml')
    if os.path.exists(doc_path):
        tree = etree.parse(doc_path)
        strip_spacing_in_tree(tree, xml_path='word/document.xml')
        split_heading_with_inline_body(tree)
        insert_blank_around_title(tree)
        fix_quotes_in_tree(tree)
        strip_font_theme_attrs(tree)
        disable_keep_together(tree)
        tree.write(doc_path, xml_declaration=True, encoding='UTF-8')

    # 2. 处理 styles.xml — docDefaults + 所有样式 + 字体theme清理 + 关闭孤行/段中/同页控制
    styles_path = os.path.join(tmp, 'word/styles.xml')
    if os.path.exists(styles_path):
        tree = etree.parse(styles_path)
        strip_spacing_in_tree(tree, xml_path='word/styles.xml')
        strip_font_theme_attrs(tree)
        disable_keep_together(tree)
        tree.write(styles_path, xml_declaration=True, encoding='UTF-8')

    # 重新打包
    with zipfile.ZipFile(output_path, 'w', zipfile.ZIP_DEFLATED) as zout:
        for root_dir, dirs, files in os.walk(tmp):
            for f in files:
                full = os.path.join(root_dir, f)
                arcname = os.path.relpath(full, tmp)
                zout.write(full, arcname)

    shutil.rmtree(tmp)
    print(f"[OK] 段间距已清零、标题前后空行已插入、引号与字体已修复、孤行/段中/同页控制已关闭、标题内夹带正文已拆分: {output_path}")
    return output_path


if __name__ == '__main__':
    if len(sys.argv) < 2:
        print("用法: python strip_spacing.py <input.docx> [output.docx]")
        sys.exit(1)

    inp = sys.argv[1]
    out = sys.argv[2] if len(sys.argv) > 2 else None
    strip_spacing(inp, out)
