# -*- coding: utf-8 -*-
"""生成符合「全网格三线表变体」版式的回归结果表（.docx）。

版式来源：从一份定稿的 Word 文档内部 XML 反解得到，所有数值均为实测值，
见 docs/format-spec.md。

用法：
    python make_reg_table.py                    # 用内置示例数据生成
    python make_reg_table.py --out 表2.docx      # 指定输出文件名
    from make_reg_table import build, EXAMPLE   # 当库用
    build(EXAMPLE, "out.docx")

只改 SPEC 里的数据即可换表，版式（边框/字号/合并/间距）全部由样式常量控制。
"""
import argparse

from docx import Document
from docx.enum.table import WD_ROW_HEIGHT_RULE, WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, Twips

# ── 版式常量（实测值，勿随意改；改动前请对照 docs/format-spec.md）───────────────
CN_FONT = "宋体"                 # 中文字体
EN_FONT = "Times New Roman"     # 西文字体
BODY_PT = 9                     # 表体字号（小五）
TITLE_PT = 10.5                 # 表题字号（五号）
CELL_W = 1436                   # 单列宽 twips = 2.53 cm
TABLE_W = 8616                  # 表宽 twips ≈ 15.19 cm（A4 版心 15.2 cm）
BORDER_THICK = "12"             # 上下外框 1.5 pt（sz 为 1/8 磅）
BORDER_THIN = "4"               # 内部横竖线 0.5 pt
CELL_MAR_LR = "108"             # 单元格左右内边距 twips = 0.19 cm
HDR_H1, HDR_H2 = 408, 351       # 两行表头的行高 twips（0.72 / 0.62 cm）

# OOXML schema 规定的子元素顺序，插入顺序错了 Word 会报“内容有问题”
TCPR_ORDER = ["w:cnfStyle", "w:tcW", "w:gridSpan", "w:hMerge", "w:vMerge", "w:tcBorders",
              "w:shd", "w:noWrap", "w:tcMar", "w:textDirection", "w:tcFitText", "w:vAlign",
              "w:hideMark", "w:headers", "w:cellIns", "w:cellDel", "w:cellMerge", "w:tcPrChange"]
TBLPR_ORDER = ["w:tblStyle", "w:tblpPr", "w:tblOverlap", "w:bidiVisual",
               "w:tblStyleRowBandSize", "w:tblStyleColBandSize", "w:tblW", "w:jc",
               "w:tblCellSpacing", "w:tblInd", "w:tblBorders", "w:shd", "w:tblLayout",
               "w:tblCellMar", "w:tblLook", "w:tblCaption", "w:tblDescription"]
TRPR_ORDER = ["w:cnfStyle", "w:divId", "w:gridBefore", "w:gridAfter", "w:wBefore",
              "w:wAfter", "w:cantSplit", "w:trHeight", "w:tblHeader", "w:tblCellSpacing",
              "w:jc", "w:hidden"]

# ── 数据：换表只改这里 ─────────────────────────────────────────────────────────
EXAMPLE = {
    "title": "表1  基准回归结果",          # 表号与标题之间两个半角空格
    "groups": [("ESG", 4), ("lnTS", 1)],  # 被解释变量分组：名称 + 占几列
    # 行定义：
    #   {"label": 标签, "cells": [(系数, 星号, 标准误), ...]}  → 系数+标准误上下两行
    #   {"label": 标签, "cells": ["是", "否", ...]}            → 单行文字
    #   "align": "left" 仅变量名行使用；"italic": True 让变量名斜体
    "rows": [
        {"label": "CBEC×Post", "italic": True, "align": "left",
         "cells": [("0.1312", "***", "(0.0337)"), ("0.1316", "***", "(0.0397)"),
                   ("0.1868", "***", "(0.0245)"), ("0.1052", "***", "(0.0370)"),
                   ("0.0081", "***", "(0.0024)")]},
        {"label": "控制变量", "cells": ["是"] * 5},
        {"label": "年份固定效应", "cells": ["否", "是", "否", "是", "是"]},
        {"label": "企业固定效应", "cells": ["否", "否", "是", "是", "是"]},
        {"label": "样本量", "cells": ["10\u2009877"] * 4 + ["10\u2009841"]},
        {"label": "R2", "cells": ["0.2154", "0.2252", "0.6835", "0.6892", "0.6225"]},
    ],
    # “注：”与星号由程序自动加，这里只写星号之后的说明
    "note": "分别表示在 10%、5%、1% 的水平上显著，括号内为聚类稳健标准误，下表同。",
}


def _insert_ordered(parent, child, order):
    """按 schema 顺序插入子元素，并去掉同名旧元素。"""
    tag = "w:" + child.tag.split("}")[-1]
    idx = order.index(tag)
    for existing in list(parent):
        etag = "w:" + existing.tag.split("}")[-1]
        if etag == tag:
            parent.remove(existing)
            continue
        if etag in order and order.index(etag) > idx:
            existing.addprevious(child)
            return child
    parent.append(child)
    return child


def _run_font(run, size=BODY_PT, bold=False, italic=False, sup=False):
    run.font.size = Pt(size)
    run.font.name = EN_FONT
    rPr = run._element.get_or_add_rPr()
    rf = rPr.find(qn("w:rFonts"))
    if rf is None:
        rf = OxmlElement("w:rFonts")
        rPr.insert(0, rf)
    rf.set(qn("w:ascii"), EN_FONT)
    rf.set(qn("w:hAnsi"), EN_FONT)
    rf.set(qn("w:eastAsia"), CN_FONT)
    run.bold = bold
    run.italic = italic
    if sup:
        run.font.superscript = True


def _cell_para(cell, first=False, align=WD_ALIGN_PARAGRAPH.CENTER,
               before=Pt(2), after=Pt(2)):
    p = cell.paragraphs[0] if first else cell.add_paragraph()
    p.alignment = align
    pf = p.paragraph_format
    pf.space_before = before
    pf.space_after = after
    pf.line_spacing = 1.0
    return p


def _vcenter(cell):
    tcPr = cell._tc.get_or_add_tcPr()
    va = OxmlElement("w:vAlign")
    va.set(qn("w:val"), "center")
    _insert_ordered(tcPr, va, TCPR_ORDER)


def _setup_pr(tblPr):
    """表级属性：宽度 auto、固定列宽、边框、单元格内边距。"""
    tw = OxmlElement("w:tblW")
    tw.set(qn("w:w"), "0")
    tw.set(qn("w:type"), "auto")
    _insert_ordered(tblPr, tw, TBLPR_ORDER)

    layout = OxmlElement("w:tblLayout")
    layout.set(qn("w:type"), "fixed")
    _insert_ordered(tblPr, layout, TBLPR_ORDER)

    # 上下 1.5 pt 封边、左右无外框、内部横竖线 0.5 pt —— 这套版式的核心特征
    borders = OxmlElement("w:tblBorders")
    for tag, val, sz in (("top", "single", BORDER_THICK),
                         ("left", "none", "0"),
                         ("bottom", "single", BORDER_THICK),
                         ("right", "none", "0"),
                         ("insideH", "single", BORDER_THIN),
                         ("insideV", "single", BORDER_THIN)):
        el = OxmlElement("w:" + tag)
        el.set(qn("w:val"), val)
        el.set(qn("w:sz"), sz)
        el.set(qn("w:space"), "0")
        el.set(qn("w:color"), "000000" if val == "single" else "auto")
        borders.append(el)
    _insert_ordered(tblPr, borders, TBLPR_ORDER)

    mar = OxmlElement("w:tblCellMar")
    for tag, w in (("top", "0"), ("left", CELL_MAR_LR),
                   ("bottom", "0"), ("right", CELL_MAR_LR)):
        el = OxmlElement("w:" + tag)
        el.set(qn("w:w"), w)
        el.set(qn("w:type"), "dxa")
        mar.append(el)
    _insert_ordered(tblPr, mar, TBLPR_ORDER)


def build(spec, out_path="regression_table.docx"):
    """按 spec 生成 docx。spec 结构见上面 EXAMPLE。"""
    groups = spec["groups"]
    n_data = sum(n for _, n in groups)
    ncol = n_data + 1
    rows = spec["rows"]

    doc = Document()
    sec = doc.sections[0]
    sec.page_width, sec.page_height = Cm(21), Cm(29.7)
    sec.top_margin = sec.bottom_margin = Cm(2.54)
    sec.left_margin = sec.right_margin = Cm(2.9)

    normal = doc.styles["Normal"]
    normal.font.name = EN_FONT
    normal.font.size = Pt(BODY_PT)
    normal.element.rPr.rFonts.set(qn("w:eastAsia"), CN_FONT)

    # 表题
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(6)
    p.paragraph_format.space_after = Pt(3)
    _run_font(p.add_run(spec["title"]), TITLE_PT, bold=True)

    table = doc.add_table(rows=2 + len(rows), cols=ncol)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False
    _setup_pr(table._tbl.tblPr)

    for col in table.columns:
        col.width = Twips(CELL_W)
    table.rows[0].height = Twips(HDR_H1)
    table.rows[1].height = Twips(HDR_H2)
    for row in table.rows[:2]:
        row.height_rule = WD_ROW_HEIGHT_RULE.AT_LEAST
    for row in table.rows:
        trPr = row._tr.get_or_add_trPr()
        jc = OxmlElement("w:jc")
        jc.set(qn("w:val"), "center")
        _insert_ordered(trPr, jc, TRPR_ORDER)
        for c in row.cells:
            c.width = Twips(CELL_W)
            _vcenter(c)

    # 表头第一行：变量（纵向合并两行）+ 被解释变量分组（横向合并）
    merged = table.cell(0, 0).merge(table.cell(1, 0))
    _run_font(_cell_para(merged, first=True).add_run("变量"))
    col = 1
    for name, span in groups:
        cell = (table.cell(0, col).merge(table.cell(0, col + span - 1))
                if span > 1 else table.cell(0, col))
        _run_font(_cell_para(cell, first=True).add_run(name), italic=True)
        col += span

    # 表头第二行：(1)…(n)
    for i in range(n_data):
        _run_font(_cell_para(table.cell(1, i + 1), first=True).add_run("(%d)" % (i + 1)),
                  italic=True)

    # 数据行
    for ri, row_spec in enumerate(rows, start=2):
        label = row_spec["label"]
        lc = table.cell(ri, 0)
        align = (WD_ALIGN_PARAGRAPH.LEFT if row_spec.get("align") == "left"
                 else WD_ALIGN_PARAGRAPH.CENTER)
        if label == "R2":                       # R 斜体 + 2 上标
            para = _cell_para(lc, first=True, align=align)
            _run_font(para.add_run("R"), italic=True)
            _run_font(para.add_run("2"), sup=True)
        elif row_spec.get("italic"):            # 变量名斜体，乘号 × 正体
            para = _cell_para(lc, first=True, align=align)
            for chunk in label.split("×"):
                if para.runs:
                    _run_font(para.add_run("×"))
                _run_font(para.add_run(chunk), italic=True)
        else:
            _run_font(_cell_para(lc, first=True, align=align).add_run(label))

        for ci, val in enumerate(row_spec["cells"], start=1):
            cell = table.cell(ri, ci)
            if isinstance(val, tuple):          # 系数 + 星号（上标）/ 标准误（次行）
                coef, star, se = val
                p1 = _cell_para(cell, first=True, before=Pt(1), after=Pt(0))
                _run_font(p1.add_run(coef))
                if star:
                    _run_font(p1.add_run(star), sup=True)
                _run_font(_cell_para(cell, before=Pt(0), after=Pt(1)).add_run(se))
            else:
                _run_font(_cell_para(cell, first=True).add_run(val))

    # 表注
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p.paragraph_format.space_before = Pt(4)
    p.paragraph_format.line_spacing = 1.0
    _run_font(p.add_run("注："))
    for star in ("*", "**", "***"):
        _run_font(p.add_run(star), sup=True)
        if star != "***":
            _run_font(p.add_run("、"))
    _run_font(p.add_run(spec["note"]))

    doc.save(out_path)
    return out_path


def main():
    ap = argparse.ArgumentParser(description="生成学术回归结果表（.docx）")
    ap.add_argument("--out", default="regression_table.docx", help="输出文件名")
    args = ap.parse_args()
    print("saved:", build(EXAMPLE, args.out))


if __name__ == "__main__":
    main()
