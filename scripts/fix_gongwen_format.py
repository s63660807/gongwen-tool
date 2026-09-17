#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
公文格式精修（GB/T 9704-2012 落地版）

在 convert.py（Pandoc 转换 + 清段间距 + 修引号）之后，补齐标准公文转换
流程里最容易漏掉的三块：

  1. 标题层级字体/加粗 —— 二级标题（一）楷体加粗、三级标题「1.XXXX。」引导语加粗
  2. 表格美化 —— 全框线 + 固定列宽 + 居中 + 表头重复 + 垂直居中 + 表内字号字体
  3. 落款对齐 —— 发文机关署名与成文日期右对齐（右空四字）

用法（三选一）：

    # A. 一步到位：md -> 精修后的标准公文 docx
    fix_gongwen_format.py <输入.md> [输出.docx]

    # B. 对已有 docx 做精修（不含 Pandoc 转换，仍会清段间距/修引号）
    fix_gongwen_format.py --docx <输入.docx> [输出.docx]

    # C. 落款被挤到单独一页时，加 --fit-signoff 做紧凑处置（见下）
    fix_gongwen_format.py <输入.md> [输出.docx] --fit-signoff[=1|2]

依赖：同目录下有 reference.docx、convert.py、strip_spacing.py。
"""
import os
import re
import sys
import copy
import shutil
import subprocess

from docx import Document
from docx.shared import Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

# ── 公文字体（GB/T 9704-2012）────────────────────────────────
EA_TITLE = '方正小标宋简体'   # 文件标题
EA_H1 = '黑体'               # 一级标题 一、
EA_H2 = '楷体_GB2312'        # 二级标题 （一）
EA_BODY = '仿宋_GB2312'      # 正文 / 三级标题
LATIN = 'Times New Roman'

SZ_TITLE = 22     # 二号
SZ_BODY = 16      # 三号
SZ_TABLE = 10.5   # 五号

LINE_BODY = 28    # 正文行距固定值 28 磅
LINE_TABLE = 16   # 表内行距固定值 16 磅
LINE_TABLE_TIGHT = 15   # 紧凑模式表内行距（落款单独占页时用）
LINE_BODY_TIGHT = 27    # 紧凑模式正文行距（GB/T 9704 允许为容纳落款调整行距）

# ── 表格几何 ──────────────────────────────────────────────────
# A4 版心宽 = 21 - 2.8(左) - 2.6(右) = 15.6cm = 8844 twip
TABLE_W = 8844
CELL_MAR = 60          # 单元格左右内边距（twip），小一点可多容纳一字
MIN_COL_W = 500

# 标准 7 列预算表（序号/预算项目/规格说明/数量/单价/金额/备注）实测调优值：
# 按五号字逐列核算「列可容字数 = (列宽 - 内边距) / 字号宽」，保证每格内容 1~2 行排完，
# 避免出现「序号」竖排、备注列末行只剩一个字这类破相。
BUDGET_COL_W = [620, 1050, 1900, 1150, 1240, 1240, 1644]
BUDGET_COL_ALIGN = [0, 0, 1, 0, 0, 0, 1]   # 0=居中 1=左对齐


def _el(tag, **attrs):
    e = OxmlElement(tag)
    for k, v in attrs.items():
        e.set(qn('w:' + k), str(v))
    return e


# ── 字体/段落工具 ────────────────────────────────────────────
def set_run(run, ea, size_pt, bold=False, latin=LATIN):
    """设置 run 的中英文字体、字号、加粗，并清掉会覆盖字体的 theme 引用"""
    run.font.size = Pt(size_pt)
    run.font.bold = bold
    run.font.color.rgb = RGBColor(0, 0, 0)
    run.font.name = latin
    rPr = run._element.get_or_add_rPr()
    rFonts = rPr.find(qn('w:rFonts'))
    if rFonts is None:
        rFonts = OxmlElement('w:rFonts')
        rPr.insert(0, rFonts)
    for a in ('asciiTheme', 'eastAsiaTheme', 'hAnsiTheme', 'cstheme'):
        if rFonts.get(qn('w:' + a)) is not None:
            del rFonts.attrib[qn('w:' + a)]
    rFonts.set(qn('w:eastAsia'), ea)
    rFonts.set(qn('w:ascii'), latin)
    rFonts.set(qn('w:hAnsi'), latin)
    rFonts.set(qn('w:cs'), latin)


def set_para_spacing(para, line_pt=None, before=0, after=0):
    """段前/段后置 0，行距设为固定值 line_pt 磅"""
    pf = para.paragraph_format
    pf.space_before = Pt(before)
    pf.space_after = Pt(after)
    pPr = para._element.get_or_add_pPr()
    sp = pPr.find(qn('w:spacing'))
    if sp is None:
        sp = OxmlElement('w:spacing')
        pPr.append(sp)
    sp.set(qn('w:before'), str(int(before * 20)))
    sp.set(qn('w:after'), str(int(after * 20)))
    sp.set(qn('w:beforeAutospacing'), '0')
    sp.set(qn('w:afterAutospacing'), '0')
    if line_pt:
        sp.set(qn('w:line'), str(int(line_pt * 20)))
        sp.set(qn('w:lineRule'), 'exact')


def set_indent_chars(para, first_line=0):
    """按「字符」设置首行缩进（字号变化时仍为固定字数）"""
    pPr = para._element.get_or_add_pPr()
    ind = pPr.find(qn('w:ind'))
    if ind is None:
        ind = OxmlElement('w:ind')
        pPr.append(ind)
    ind.set(qn('w:firstLineChars'), str(int(first_line * 100)))
    if first_line == 0:
        ind.set(qn('w:firstLine'), '0')


# ── 表格：列宽推算 ──────────────────────────────────────────
def _disp_width(text):
    """按显示宽度计字数：全角/宽字符计 1，半角计 0.5"""
    return sum(0.5 if ord(c) < 0x1100 else 1.0 for c in text)


def auto_col_widths(tbl):
    """列宽策略：7 列预算表用实测调优值；其他表按各列最长内容加权分配。

    返回 (widths, aligns)。aligns 中 0=居中、1=左对齐：
    内容宽度 ≤ 6 个全角字的列居中（表头短、数字类），超出的列左对齐。
    """
    ncol = len(tbl.columns)

    if ncol == len(BUDGET_COL_W):
        return BUDGET_COL_W[:], BUDGET_COL_ALIGN[:]

    weights = []
    for ci in range(ncol):
        longest = 0.0
        for row in tbl.rows:
            if ci < len(row.cells):
                longest = max(longest, _disp_width(row.cells[ci].text.strip()))
        weights.append(max(longest, 2.0))

    raw = [w / sum(weights) * TABLE_W for w in weights]
    # 保底 + 上限（单列不超过整表 40%，避免一列吃掉半张表）
    widths = [max(MIN_COL_W, min(w, TABLE_W * 0.40)) for w in raw]
    # 归一化回总宽
    scale = TABLE_W / sum(widths)
    widths = [int(round(w * scale)) for w in widths]
    widths[-1] += TABLE_W - sum(widths)   # 余数补到最后列

    aligns = [0 if _disp_width(tbl.rows[0].cells[ci].text.strip()) <= 6 else 1
              for ci in range(ncol)]
    return widths, aligns


def beautify_table(tbl):
    """表格美化：全框线、居中、固定列宽、表头黑体、表内仿宋五号、垂直居中、表头跨页重复"""
    tbl_el = tbl._element
    tblPr = tbl_el.find(qn('w:tblPr'))
    if tblPr is None:
        tblPr = OxmlElement('w:tblPr')
        tbl_el.insert(0, tblPr)

    # 丢弃 pandoc 的表格样式（Normal Table 无框线），改用显式框线
    for tag in ('w:tblStyle', 'w:tblBorders', 'w:tblW', 'w:tblInd',
                'w:tblLayout', 'w:jc', 'w:tblCellMar'):
        for e in tblPr.findall(qn(tag)):
            tblPr.remove(e)

    borders = OxmlElement('w:tblBorders')
    for edge, sz in (('top', 8), ('left', 8), ('bottom', 8), ('right', 8),
                     ('insideH', 4), ('insideV', 4)):
        borders.append(_el('w:' + edge, val='single', sz=sz, space='0', color='000000'))
    tblPr.append(borders)

    widths, aligns = auto_col_widths(tbl)

    tblPr.append(_el('w:tblW', w=sum(widths), type='dxa'))
    tblPr.append(_el('w:tblInd', w=0, type='dxa'))
    tblPr.append(_el('w:tblLayout', type='fixed'))
    tblPr.append(_el('w:jc', val='center'))

    mar = OxmlElement('w:tblCellMar')
    for side in ('top', 'left', 'bottom', 'right'):
        mar.append(_el('w:' + side, w=CELL_MAR, type='dxa'))
    tblPr.append(mar)

    grid = tbl_el.find(qn('w:tblGrid'))
    if grid is not None:
        for gc, wv in zip(grid.findall(qn('w:gridCol')), widths):
            gc.set(qn('w:w'), str(wv))

    for ri, row in enumerate(tbl.rows):
        is_head = (ri == 0)
        trPr = row._element.get_or_add_trPr()
        for tag in ('w:trHeight', 'w:tblHeader', 'w:cantSplit'):
            for e in trPr.findall(qn(tag)):
                trPr.remove(e)
        trPr.append(_el('w:trHeight', val=454, hRule='atLeast'))   # ≥0.8cm
        if is_head:
            trPr.append(OxmlElement('w:tblHeader'))   # 跨页时自动重复表头
        trPr.append(OxmlElement('w:cantSplit'))       # 单行不跨页

        for ci, cell in enumerate(row.cells):
            tcPr = cell._element.get_or_add_tcPr()
            for tag in ('w:tcW', 'w:vAlign'):
                for e in tcPr.findall(qn(tag)):
                    tcPr.remove(e)
            tcPr.append(_el('w:tcW', w=widths[ci] if ci < len(widths) else MIN_COL_W,
                            type='dxa'))
            tcPr.append(_el('w:vAlign', val='center'))

            align = aligns[ci] if ci < len(aligns) else 0
            for para in cell.paragraphs:
                para.alignment = (WD_ALIGN_PARAGRAPH.CENTER if align == 0
                                  else WD_ALIGN_PARAGRAPH.LEFT)
                set_indent_chars(para, 0)          # 表内不缩进
                set_para_spacing(para, line_pt=LINE_TABLE)
                for r in para.runs:
                    set_run(r, EA_H1 if is_head else EA_BODY, SZ_TABLE, bold=False)


def keep_table_with_prev(doc):
    """给表格前一段落加 keepNext，让表格整体不与前文拆到两页"""
    for tbl in doc.element.body.iter(qn('w:tbl')):
        prev = tbl.getprevious()
        if prev is None or prev.tag != qn('w:p'):
            continue
        pPr = prev.find(qn('w:pPr'))
        if pPr is None:
            pPr = OxmlElement('w:pPr')
            prev.insert(0, pPr)
        for e in pPr.findall(qn('w:keepNext')):
            pPr.remove(e)
        pPr.append(OxmlElement('w:keepNext'))


def space_after_table(doc):
    """表格后正文加 12 磅段前间距，避免表格与下文贴排"""
    for tbl in doc.element.body.iter(qn('w:tbl')):
        nxt = tbl.getnext()
        if nxt is None or nxt.tag != qn('w:p'):
            continue
        pPr = nxt.find(qn('w:pPr'))
        if pPr is None:
            pPr = OxmlElement('w:pPr')
            nxt.insert(0, pPr)
        spacing = pPr.find(qn('w:spacing'))
        if spacing is None:
            spacing = OxmlElement('w:spacing')
            pPr.append(spacing)
        spacing.set(qn('w:before'), '240')


# ── 标题层级 ────────────────────────────────────────────────
def apply_title_fonts(doc):
    """文件标题：方正小标宋简体二号居中（显式设置，不依赖模板样式）"""
    for p in doc.paragraphs:
        if p.style.name not in ('Title',):
            continue
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        set_indent_chars(p, 0)
        for r in p.runs:
            set_run(r, EA_TITLE, SZ_TITLE, bold=False, latin=EA_TITLE)


def apply_heading_fonts(doc):
    """一级标题黑体三号；二级标题楷体_GB2312 三号加粗；均首行缩进 2 字符"""
    for p in doc.paragraphs:
        name = p.style.name
        if name in ('Heading 1', 'Heading1'):
            set_para_spacing(p, line_pt=LINE_BODY)
            set_indent_chars(p, 2)
            for r in p.runs:
                set_run(r, EA_H1, SZ_BODY, bold=False, latin=EA_H1)
        elif name in ('Heading 2', 'Heading2'):
            set_para_spacing(p, line_pt=LINE_BODY)
            set_indent_chars(p, 2)
            for r in p.runs:
                set_run(r, EA_H2, SZ_BODY, bold=True, latin=EA_H2)


def bold_serial_lead(doc):
    """三级标题「1.引导语。」加粗，其后的正文不加粗

    公文三级标题为仿宋三号加粗。md 里写成「1.铸魂强基，……。正文……」这种
    引导语与正文同段的形式时，把首个句号（含）之前的部分加粗。
    """
    pat = re.compile(r'^\d+\.')
    for p in doc.paragraphs:
        if p.style.name.startswith('Heading'):
            continue
        txt = p.text
        if not pat.match(txt):
            continue
        idx = txt.find('。')
        if idx < 0 or idx + 1 >= len(txt):
            continue          # 整段就是标题，无需拆分
        lead, rest = txt[:idx + 1], txt[idx + 1:]
        runs = p.runs
        if not runs:
            continue
        first = runs[0]
        for extra in runs[1:]:
            extra._element.getparent().remove(extra._element)
        first.text = lead
        set_run(first, EA_BODY, SZ_BODY, bold=True)

        new_r = copy.deepcopy(first._element)
        first._element.addnext(new_r)
        from docx.text.run import Run
        nr = Run(new_r, p)
        nr.text = rest
        set_run(nr, EA_BODY, SZ_BODY, bold=False)


def normalize_body(doc):
    """正文段落统一 28 磅固定行距（字体/缩进由样式提供）"""
    for p in doc.paragraphs:
        if p.text.strip():
            set_para_spacing(p, line_pt=LINE_BODY)


def right_align_signoff(doc):
    """落款右对齐：发文机关署名与成文日期均右空四字（GB/T 9704-2012）

    取全文最后两个非空段落，按「署名、日期」处理。
    """
    paras = [p for p in doc.paragraphs if p.text.strip()]
    if len(paras) < 2:
        return
    date_p, org_p = paras[-1], paras[-2]

    for p in (org_p, date_p):
        p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        set_indent_chars(p, 0)
        pPr = p._element.get_or_add_pPr()
        ind = pPr.find(qn('w:ind'))
        if ind is None:
            ind = OxmlElement('w:ind')
            pPr.append(ind)
        ind.set(qn('w:rightChars'), '400')                 # 右空 4 字
        ind.set(qn('w:right'), str(4 * SZ_BODY * 20))      # 对应绝对磅值
        for r in p.runs:
            set_run(r, EA_BODY, SZ_BODY, bold=False)


# ── 落款同页处置 ────────────────────────────────────────────
def tighten_for_signoff(doc, level=1):
    """紧凑处置：把发文机关署名与成文日期拉回正文最后一页。

    GB/T 9704-2012 允许「当公文排版后所剩空白处不能容下印章或成文日期时，
    可以采取调整行距、字间距的方法解决，务使印章与正文同处一页」。
    正文排得越满越容易出现「落款单独占一页」——末尾只剩几行空白时尤其明显。

    level 1：取消表格后段前间距 + 收紧表格单元格上下内边距与表内行距（约省 30~40pt）
    level 2：在 level 1 之上再把正文行距由 28 磅压到 27 磅（每页约省 22pt）

    用法上先试 level 1，转 PDF 看落款是否已回正文页；不够再 level 2。
    """
    # 表格：收紧上下内边距 + 表内行距 + 取消表后段前间距
    for tbl in doc.tables:
        el = tbl._element
        tblPr = el.find(qn('w:tblPr'))
        mar = tblPr.find(qn('w:tblCellMar')) if tblPr is not None else None
        if mar is not None:
            for side in ('top', 'bottom'):
                e = mar.find(qn('w:' + side))
                if e is not None:
                    e.set(qn('w:w'), '20')
        nxt = el.getnext()
        if nxt is not None and nxt.tag == qn('w:p'):
            pPr = nxt.find(qn('w:pPr'))
            sp = pPr.find(qn('w:spacing')) if pPr is not None else None
            if sp is not None:
                sp.set(qn('w:before'), '0')
        for row in tbl.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    pPr = p._element.find(qn('w:pPr'))
                    sp = pPr.find(qn('w:spacing')) if pPr is not None else None
                    if sp is not None:
                        sp.set(qn('w:line'), str(int(LINE_TABLE_TIGHT * 20)))
                        sp.set(qn('w:lineRule'), 'exact')

    if level < 2:
        return

    # 正文行距 28 磅 → 27 磅（只动原本就是 28 磅固定值的段落）
    old = str(int(LINE_BODY * 20))
    new = str(int(LINE_BODY_TIGHT * 20))
    for p in doc.paragraphs:
        pPr = p._element.find(qn('w:pPr'))
        sp = pPr.find(qn('w:spacing')) if pPr is not None else None
        if sp is not None and sp.get(qn('w:line')) == old \
                and sp.get(qn('w:lineRule')) == 'exact':
            sp.set(qn('w:line'), new)


# ── 主流程 ──────────────────────────────────────────────────
def fix_docx(docx_path, fit_signoff=0):
    """对已有 docx 做公文化精修（原地保存）

    fit_signoff: 0=不处理；1/2=落款单独占页时的紧凑处置级别（见 tighten_for_signoff）
    """
    doc = Document(docx_path)
    apply_title_fonts(doc)
    apply_heading_fonts(doc)
    bold_serial_lead(doc)
    normalize_body(doc)
    if doc.tables:
        for t in doc.tables:
            beautify_table(t)
        keep_table_with_prev(doc)
        space_after_table(doc)
    right_align_signoff(doc)
    if fit_signoff:
        tighten_for_signoff(doc, fit_signoff)
    doc.save(docx_path)
    return docx_path


def convert(md_path, out_path, fit_signoff=0):
    """md -> Pandoc 转换 -> 清段间距/修引号 -> 公文化精修 -> out_path"""
    here = os.path.dirname(os.path.abspath(__file__))
    if here not in sys.path:
        sys.path.insert(0, here)
    from convert import find_pandoc, strip_spacing_docx

    pandoc = find_pandoc()
    if not pandoc:
        print('[错误] 找不到 pandoc.exe，请安装或加入 PATH')
        sys.exit(1)
    ref = os.path.join(here, 'reference.docx')
    if not os.path.exists(ref):
        print('[错误] 缺少模板:', ref, '（先运行 create_reference.py）')
        sys.exit(1)

    tmp = out_path + '.tmp.docx'
    r = subprocess.run([pandoc, '-f', 'markdown-smart', md_path, '-o', tmp,
                        '--reference-doc=' + ref],
                       capture_output=True, text=True,
                       encoding='utf-8', errors='replace')
    if r.returncode != 0:
        print('[错误] pandoc 转换失败:', r.stderr)
        sys.exit(1)

    strip_spacing_docx(tmp)
    fix_docx(tmp, fit_signoff)

    if os.path.exists(out_path):
        try:
            os.remove(out_path)
        except PermissionError:
            alt = os.path.splitext(out_path)[0] + '_new.docx'
            print('[提示] 目标文件被占用，改输出:', alt)
            out_path = alt
    shutil.move(tmp, out_path)
    print('[完成] 已生成:', out_path)
    return out_path


def main():
    argv = sys.argv[1:]
    if not argv or argv[0] in ('-h', '--help'):
        print(__doc__)
        sys.exit(0)

    # 解析 --fit-signoff[=N]（默认 1）
    fit = 0
    for a in list(argv):
        if a == '--fit-signoff':
            fit = 1
            argv.remove(a)
        elif a.startswith('--fit-signoff='):
            fit = int(a.split('=', 1)[1])
            argv.remove(a)

    if argv and argv[0] == '--docx':
        if len(argv) < 2:
            print('用法: fix_gongwen_format.py --docx <输入.docx> [输出.docx] [--fit-signoff[=1|2]]')
            sys.exit(1)
        src = argv[1]
        dst = argv[2] if len(argv) > 2 else src
        if src != dst:
            shutil.copy2(src, dst)
        print('[完成] 已精修:', fix_docx(dst, fit))
        return

    if not argv:
        print(__doc__)
        sys.exit(0)
    md_path = argv[0]
    out_path = argv[1] if len(argv) > 1 else os.path.splitext(md_path)[0] + '.docx'
    if not os.path.exists(md_path):
        print('[错误] 文件不存在:', md_path)
        sys.exit(1)
    convert(md_path, out_path, fit)


if __name__ == '__main__':
    main()
