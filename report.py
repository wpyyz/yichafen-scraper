# -*- coding: utf-8 -*-
"""
把查询结果整理成一张「能直接看」的成绩表。

默认（report 关）只输出一张平表，和以前一样。
打开 report 后，额外生成：

    各科平均分   每个数字列的 均分/最高/最低/极差/标准差
    我的成绩     把你自己的六科单独拎出来，带班内排名和与均分差
    覆盖情况     这次查了谁、几行、缺什么
    未命中姓名   查不到的，附原因

这些东西都从结果本身算出来，不需要另一份模板。
"""

import re
import statistics

from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter

HF = PatternFill("solid", fgColor="1F4E79")
HFONT = Font(bold=True, color="FFFFFF", size=11)
THIN = Side(style="thin", color="BFBFBF")
BD = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

# 这些列看着像数字，但不是分数，算平均分要跳过
NOT_SCORE = re.compile(r"考号|学号|编号|准考证|排名|位次|名次|班级|班别|年级|"
                       r"电话|手机|身份证|时间|日期|序号|性别")
RANK_WORDS = ("排名", "名次", "位次")


def _num(v):
    """能转成数字就转，转不了返回 None。空串、'—'、'无' 都当没有。"""
    if v is None:
        return None
    s = str(v).strip()
    if not s or s in ("—", "-", "无", "None", "nan"):
        return None
    try:
        f = float(s)
    except ValueError:
        return None
    return int(f) if f == int(f) else f


def _score_cols(rows, cols):
    """挑出真正的分数列：整列基本都能转成数字，且名字不像学号/排名。"""
    out = []
    for c in cols:
        if NOT_SCORE.search(str(c)):
            continue
        vals = [_num(r.get(c)) for r in rows]
        vals = [v for v in vals if v is not None]
        if not vals:
            continue
        if len(vals) >= max(3, len(rows) * 2 // 3):
            out.append(c)
    return out


def _pick(cols, *names):
    for c in cols:
        if str(c).strip() in names:
            return c
    return None


def _total_col(cols):
    for c in cols:
        s = str(c)
        if "总分" in s or "合计" in s or "总成绩" in s:
            return c
    return None


def _rank_cols(cols):
    return [c for c in cols if any(w in str(c) for w in RANK_WORDS)]


def _stats(vals):
    vals = [x for x in vals if x is not None]
    if not vals:
        return None
    return {
        "n": len(vals),
        "avg": round(statistics.mean(vals), 2),
        "max": max(vals),
        "min": min(vals),
        "rng": max(vals) - min(vals),
        "sd": round(statistics.pstdev(vals), 2) if len(vals) > 1 else 0,
    }


def _style_header(ws, ncol):
    for i in range(1, ncol + 1):
        c = ws.cell(row=1, column=i)
        c.fill, c.font = HF, HFONT
        c.alignment = Alignment(horizontal="center", vertical="center")
        c.border = BD


def _fit(ws, ncol, rows_hint=None):
    widths = {}
    for row in ws.iter_rows(values_only=True):
        for i, v in enumerate(row, 1):
            if i > ncol:
                break
            w = len(str(v)) if v is not None else 0
            widths[i] = max(widths.get(i, 6), min(w * 2, 40))
    for i in range(1, ncol + 1):
        ws.column_dimensions[get_column_letter(i)].width = max(8, widths.get(i, 10))


def build(rows, out_path, meta=None, failed=None, me=None, main_title="成绩数据"):
    """
    rows     : [{列名: 值}, ...]
    failed   : [(姓名, 原因), ...]
    me       : 想单独拎出来的名字，None/空 就不生成「我的成绩」
    """
    meta = meta or {}
    failed = failed or []

    cols = []
    for r in rows:
        for k in r:
            if k not in cols:
                cols.append(k)

    score_cols = _score_cols(rows, cols)
    total = _total_col(cols)
    name_c = _pick(cols, "姓名", "名字", "学生姓名", "儿童姓名", "考生姓名")

    # 「查询条件」是我们为了对账加的内部列，跟姓名重复，不进表
    if name_c and "查询条件" in cols:
        cols = [c for c in cols if c != "查询条件"]

    # 源数据自带班排名（易查分就有）→ 直接沿用，别再叠一列
    given_rank = None
    for c in cols:
        s = str(c)
        if "班排名" in s or "班内排名" in s:
            given_rank = c
            break

    # 班级里没给名次时的退路：先看总分，再看校级名次
    school_rank = None
    for c in cols:
        s = str(c)
        if c != given_rank and any(k in s for k in ("校排名", "年级排名", "全校排名", "校名次")):
            school_rank = c
            break

    # ---------- 排序 ----------
    # 有班排名 → 照人家原本的名次排；
    # 没有 → 按总分从高到低（顺便算个班排名）；
    # 连总分都没有 → 按校级名次从小到大。
    ranked = {}
    if rows:
        if given_rank:
            rows = sorted(rows, key=lambda r: (_num(r.get(given_rank)) is None,
                                               _num(r.get(given_rank)) or 0))
        elif total:
            ordered = sorted(rows, key=lambda r: -(_num(r.get(total)) or -1))
            prev, rank = None, 0
            for i, r in enumerate(ordered, 1):
                v = _num(r.get(total)) or 0
                if v != prev:
                    rank, prev = i, v
                ranked[id(r)] = rank
            rows = ordered
        elif school_rank:
            rows = sorted(rows, key=lambda r: (_num(r.get(school_rank)) is None,
                                               _num(r.get(school_rank)) or 0))

    wb = Workbook()

    # ---------- 1. 主表 ----------
    ws = wb.active
    ws.title = main_title[:31]
    head = (["班排名"] if ranked else []) + cols
    ws.append(head)
    _style_header(ws, len(head))
    for r in rows:
        line = ([ranked[id(r)]] if ranked else []) + [r.get(c, "") for c in cols]
        ws.append(line)
    for row in ws.iter_rows(min_row=2, max_row=ws.max_row, max_col=len(head)):
        for cell in row:
            cell.border = BD
            cell.alignment = Alignment(horizontal="center", vertical="center")
    ws.freeze_panes = "A2"
    if rows:
        ws.auto_filter.ref = "A1:%s%d" % (get_column_letter(len(head)), ws.max_row)
    _fit(ws, len(head))

    # ---------- 2. 各科平均分 ----------
    stat_cols = [c for c in score_cols if c != total] + ([total] if total in score_cols else [])
    if stat_cols:
        ws2 = wb.create_sheet("各科平均分")
        ws2.append(["科目", "人数", "平均分", "最高分", "最低分", "极差", "标准差"])
        _style_header(ws2, 7)
        for c in stat_cols:
            st = _stats([_num(r.get(c)) for r in rows])
            if not st:
                continue
            ws2.append([c, st["n"], st["avg"], st["max"], st["min"], st["rng"], st["sd"]])
        for row in ws2.iter_rows(min_row=2, max_row=ws2.max_row, max_col=7):
            for cell in row:
                cell.border = BD
                cell.alignment = Alignment(horizontal="center", vertical="center")
        _fit(ws2, 7)

    # ---------- 3. 我的成绩 ----------
    mine = None
    if me and name_c and rows:
        for r in rows:
            if str(r.get(name_c, "")).strip() == str(me).strip():
                mine = r
                break
    if me:
        ws3 = wb.create_sheet("我的成绩")
        ws3["A1"] = str(me)
        ws3["A1"].font = Font(bold=True, size=13)
        if mine is None:
            ws3["A2"] = "这次结果里没找到这个名字"
        else:
            bits, used = [], set()
            if ranked:
                bits.append("班排名 %s" % ranked[id(mine)])
                used.add("班排名")
            elif given_rank:
                bits.append("班排名 %s" % mine.get(given_rank))
                used.add(given_rank)
            if total:
                bits.append("总分 %s" % mine.get(total))
                used.add(total)
            for rc in _rank_cols(cols):
                if rc in used:
                    continue
                bits.append("%s %s" % (rc, mine.get(rc)))
            ws3["A2"] = " / ".join(bits)

            ws3["A4"], ws3["B4"], ws3["C4"] = "科目", "得分", "班内排名"
            ws3["D4"], ws3["E4"] = "班级平均", "与均分差"
            for i in range(1, 6):
                c = ws3.cell(row=4, column=i)
                c.fill, c.font = HF, HFONT
                c.alignment = Alignment(horizontal="center", vertical="center")

            r_i = 5
            for c in stat_cols:
                if c == total:          # 总分已经在上面那行写过，这里只列单科
                    continue
                sc = _num(mine.get(c))
                allv = [x for x in (_num(r.get(c)) for r in rows) if x is not None]
                if sc is None or not allv:
                    continue
                srt = sorted(allv, reverse=True)
                av = round(statistics.mean(allv), 2)
                ws3.cell(r_i, 1).value = c
                ws3.cell(r_i, 2).value = sc
                ws3.cell(r_i, 3).value = srt.index(sc) + 1
                ws3.cell(r_i, 4).value = av
                ws3.cell(r_i, 5).value = round(sc - av, 2)
                r_i += 1
            for row in ws3.iter_rows(min_row=5, max_row=r_i - 1, max_col=5):
                for cell in row:
                    cell.border = BD
                    cell.alignment = Alignment(horizontal="center", vertical="center")
        _fit(ws3, 5)

    # ---------- 4. 覆盖情况 ----------
    ws4 = wb.create_sheet("覆盖情况")
    ws4.append(["项目", "值"])
    _style_header(ws4, 2)
    for k, v in meta.items():
        ws4.append([k, str(v)])
    ws4.append(["人数", len(rows)])
    ws4.append(["列数", len(cols)])
    for row in ws4.iter_rows(min_row=2, max_row=ws4.max_row, max_col=2):
        for cell in row:
            cell.border = BD
            cell.alignment = Alignment(horizontal="center", vertical="center")
    ws4.column_dimensions["A"].width = 18
    ws4.column_dimensions["B"].width = 46

    # ---------- 5. 未命中姓名 ----------
    if failed:
        ws5 = wb.create_sheet("未命中姓名")
        ws5.append(["查询过的姓名", "结果"])
        _style_header(ws5, 2)
        for n, e in failed:
            ws5.append([n, str(e)])
        ws5.column_dimensions["A"].width = 20
        ws5.column_dimensions["B"].width = 46

    wb.save(out_path)
    return out_path
