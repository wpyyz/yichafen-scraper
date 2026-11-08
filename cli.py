#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
yichafen-cli —— 批量查询易查分成绩并导出 Excel

最少只需要给名字：

    python cli.py --names names.txt

查询码 / openid 的选取顺序（成绩一个月出一次，每次都要能自动跟上最新那份）：

    1. --uq-code 显式指定
    2. 微信缓存里最后打开的那一个（多个取最新）
    3. 都没有 → 用配置文件里上次用过的
    4. 再没有 → 报错，提示先在微信电脑端打开一次

运行前会把「本次查的是哪一份」打在屏幕上，确认无误再开始。
"""

import argparse
import json
import os
import random
import sys
import time

from ycf import YiChaFen, CooldownError, QueryError, record_rows
import discover
import queries

DEFAULT_CACHE = "ycf-cache.json"
CONFIG = os.path.join(os.path.dirname(os.path.abspath(__file__)), "ycf-config.json")
OUT_DIR = "成绩输出"


def _read_text(path):
    """按常见编码依次尝试读取，兼容记事本的 ANSI(GBK) 和 UTF-8(带/不带 BOM)。"""
    for enc in ("utf-8-sig", "gbk", "utf-16"):
        try:
            with open(path, encoding=enc) as f:
                return f.read()
        except (UnicodeDecodeError, UnicodeError):
            continue
    with open(path, encoding="utf-8", errors="replace") as f:
        return f.read()


def load_names(arg):
    if len(arg) == 1 and os.path.isfile(arg[0]):
        out = []
        # 记事本存 ANSI 是 GBK，存 UTF-8 可能带 BOM，两种都得能读
        for line in _read_text(arg[0]).splitlines():
            line = line.strip().lstrip("\ufeff")
            if line and not line.startswith("#"):
                out.append(line)
        return out
    return [a.strip().lstrip("\ufeff") for a in arg if a.strip()]


# 缓存按「查询码」分区存放：{"by_uq": {uqCode: {姓名: 结果}}}
# 必须分区！否则查过 A 班之后，B 班里同名的人会被当成已有结果直接跳过。
def load_cache(path):
    if os.path.exists(path):
        try:
            with open(path, encoding="utf-8-sig") as f:
                data = json.load(f)
        except Exception:
            return {"by_uq": {}}
        if not isinstance(data, dict) or "by_uq" not in data:
            # 老格式（按姓名平铺、无法判断属于哪次查询）→ 直接丢弃重抓
            return {"by_uq": {}}
        return data
    return {"by_uq": {}}


def cache_items(all_cache, uq_code):
    return (all_cache.get("by_uq") or {}).get(uq_code) or {}


def cache_put(all_cache, uq_code, items):
    all_cache.setdefault("by_uq", {})[uq_code] = items


def save_cache(path, all_cache):
    tmp = path + ".tmp"
    with open(tmp, "w", encoding="utf-8") as f:
        json.dump(all_cache, f, ensure_ascii=False, indent=1)
    os.replace(tmp, path)


def load_config():
    if os.path.exists(CONFIG):
        try:
            with open(CONFIG, encoding="utf-8-sig") as f:
                return json.load(f)
        except Exception:
            pass
    return {}


def save_config(openid, uq_code, tu_code=None, uq_ts=None):
    try:
        with open(CONFIG, "w", encoding="utf-8") as f:  # 不带 BOM
            json.dump({"openid": openid, "uq_code": uq_code,
                       "tu_code": tu_code, "uq_ts": uq_ts},
                      f, ensure_ascii=False, indent=1)
    except Exception:
        pass


def build_xlsx(rows, out_path, meta=None, failed=None):
    try:
        from openpyxl import Workbook
        from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
        from openpyxl.utils import get_column_letter
    except ImportError:
        print("!! 需要 openpyxl 才能导出 Excel：pip install openpyxl", file=sys.stderr)
        return None

    cols = []
    for r in rows:
        for k in r:
            if k not in cols:
                cols.append(k)

    wb = Workbook()
    ws = wb.active
    ws.title = "成绩数据"
    hf = PatternFill("solid", fgColor="1F4E79")
    hfont = Font(bold=True, color="FFFFFF", size=11)
    thin = Side(style="thin", color="BFBFBF")
    bd = Border(left=thin, right=thin, top=thin, bottom=thin)

    if cols:
        ws.append(cols)
        for i in range(1, len(cols) + 1):
            c = ws.cell(row=1, column=i)
            c.fill, c.font = hf, hfont
            c.alignment = Alignment(horizontal="center", vertical="center")
            c.border = bd
        for r in rows:
            ws.append([r.get(c, "") for c in cols])
        for i, c in enumerate(cols, 1):
            w = max([len(str(c))] + [len(str(r.get(c, ""))) for r in rows])
            ws.column_dimensions[get_column_letter(i)].width = min(max(8, w * 2), 30)
        for row in ws.iter_rows(min_row=2, max_row=ws.max_row, max_col=len(cols)):
            for cell in row:
                cell.border = bd
                cell.alignment = Alignment(horizontal="center", vertical="center")
        ws.freeze_panes = "A2"
        ws.auto_filter.ref = "A1:%s%d" % (get_column_letter(len(cols)), ws.max_row)

    ws2 = wb.create_sheet("说明")
    ws2.append(["项目", "内容"])
    for k, v in (meta or {}).items():
        ws2.append([k, str(v)])
    if failed:
        ws2.append([])
        ws2.append(["未命中", "原因"])
        for n, e in failed:
            ws2.append([n, str(e)])
    for i in (1, 2):
        c = ws2.cell(row=1, column=i)
        c.fill, c.font = hf, hfont
    ws2.column_dimensions["A"].width = 18
    ws2.column_dimensions["B"].width = 60

    wb.save(out_path)
    return out_path


def auto_out(base="成绩.xlsx", dirname=OUT_DIR):
    """结果收进「成绩输出」子目录并自动编号，不散在根目录也不覆盖历史。"""
    d = os.path.join(os.path.dirname(os.path.abspath(__file__)), dirname)
    os.makedirs(d, exist_ok=True)
    stem, ext = os.path.splitext(base)
    i = 1
    while os.path.exists(os.path.join(d, "%s%d%s" % (stem, i, ext))):
        i += 1
    return os.path.join(d, "%s%d%s" % (stem, i, ext))


def resolve_target(args):
    """
    返回 (openid, uqCode, 来源说明)。
    缓存和配置之间比时间戳，谁新用谁。
    """
    saved = load_config()

    cache_uq, cache_oid = [], None
    if not args.uq_code or not (args.openid or saved.get("openid")):
        print("[*] 扫描微信本地缓存…")
        cache_uq = discover.recent_uq(args.wechat_root or None)
        cache_oid = discover.recent_openid(args.wechat_root or None)

    openid = args.openid or saved.get("openid") or cache_oid
    if not openid:
        print("")
        print("[!] 找不到 openidCode。")
        print("    微信这部分缓存会被清理，重启微信后常常就没了。")
        print("    办法：先在【微信电脑端】打开一次那个查询小程序、随便查一个人，")
        print("          然后立刻回来重新运行 —— 新写的缓存就能抓到。")
        print("          或者用 --openid 手动指定。")
        raise SystemExit("缺少 openidCode")

    candidates = []          # (时间戳, uqCode, 来源说明)
    if args.uq_code:
        # 命令行是绝对优先，不参与时间比较
        return openid, args.uq_code, "命令行 --uq-code", time.time(), []
    for uq, mt, _p in cache_uq:
        candidates.append((mt, uq, "微信缓存，最后打开于 %s"
                           % time.strftime("%Y-%m-%d %H:%M", time.localtime(mt))))
    if saved.get("uq_code"):
        candidates.append((saved.get("uq_ts") or 0, saved["uq_code"], "上次用过的"))

    if not candidates:
        raise SystemExit("找不到查询码。请用 --uq-code 指定，或先在微信里打开一次该查询。")

    candidates.sort(key=lambda x: -x[0])
    ts, uq, source = candidates[0]

    others = [(c, u) for c, u, _ in candidates[1:] if u != uq][:5]

    return openid, uq, source, ts, others


def main():
    ap = argparse.ArgumentParser(
        description="批量查询易查分成绩并导出 Excel",
        epilog="最少用法：python cli.py --names names.txt")
    ap.add_argument("--names", nargs="+", required=True, help="姓名列表文件，或直接给名字")
    ap.add_argument("--uq-code", help="查询码；不给就自动取缓存里最新的")
    ap.add_argument("--openid", help="微信 openidCode；不给会自动发现")
    ap.add_argument("--tu-code", help="老师账号码（仅记录用）")
    ap.add_argument("--field", default="s_xingming", help="查询字段名，默认 s_xingming")
    ap.add_argument("--out", default=None, help="输出 xlsx；不给就自动编号")
    ap.add_argument("--cache", default=DEFAULT_CACHE, help="断点续抓缓存")
    ap.add_argument("--delay", type=float, default=12.0, help="每人间隔秒（默认 12）")
    ap.add_argument("--cooldown", type=float, default=1900, help="触发限流后等待秒")
    ap.add_argument("--proxy", help="可选代理")
    ap.add_argument("--wechat-root", action="append", default=[],
                    help="微信数据目录，可多次指定")
    ap.add_argument("--pick", type=int, help="在有多个查询时直接选第几个（配合上面的列表）")
    ap.add_argument("--yes", action="store_true", help="跳过确认，直接开跑")
    args = ap.parse_args()

    names = load_names(args.names)
    if not names:
        raise SystemExit("没有可查询的姓名")

    openid, uq_code, source, ts, others = resolve_target(args)

    # 有多个查询时：列出全部（含当前），按编号就能换
    if others:
        all_rows = queries.rows_for(openid, [(ts, uq_code)] + list(others), proxy=args.proxy)
        all_rows.sort(key=lambda r: -r["sort"])
        width = max([len(r["name"]) for r in all_rows] + [10])
        print("")
        print("[i] 这个账号下的成绩（编号直接选）：")
        for i, r in enumerate(all_rows, 1):
            mark = "  ← 当前" if r["uq"] == uq_code else ""
            print("    %d) %-*s   发布 %-10s 最近打开 %s%s"
                  % (i, width, r["name"], r["pub"], r["opened"] or "—", mark))
        print("")

        chosen = None
        if args.pick:
            if 1 <= args.pick <= len(all_rows):
                chosen = all_rows[args.pick - 1]
            else:
                print("    --pick 超出范围，继续用当前这份。")
        elif not args.yes and sys.stdin is not None and sys.stdin.isatty():
            try:
                raw = input("    输入编号换一份，直接回车查当前这份: ").strip()
            except (EOFError, KeyboardInterrupt):
                raw = ""
            if raw.isdigit() and 1 <= int(raw) <= len(all_rows):
                chosen = all_rows[int(raw) - 1]
        else:
            print("    想换的话双击「选择查询.bat」，或用 --pick 编号。")

        if chosen and chosen["uq"] != uq_code:
            uq_code = chosen["uq"]
            source = "手动选择"
            ts = time.time()
            print("    已换成：%s" % chosen["name"])
        print("")

    # 选定之后才建客户端、取标题（因为可能刚换过一份）
    all_cache = load_cache(args.cache)
    cache = cache_items(all_cache, uq_code)
    client = YiChaFen(uq_code, openid, proxy=args.proxy)

    info = client.refresh_kkey().get("queryInfo") or {}
    title = info.get("name") or uq_code
    update = info.get("update_time") or "未知"

    bar = "=" * 48
    print("")
    print(bar)
    print("  本次查询：%s" % title)
    print("  查询码  ：%s" % uq_code)
    print("  发布时间：%s" % update)
    print("  凭据来源：%s" % source)
    print("  待查人数：%d 人 × %.0f 秒" % (len(names), args.delay))
    print(bar)
    print("")

    # 交互式运行时给一次反悔机会；被 bat 调用（非 TTY）就直接跑
    if not args.yes and sys.stdin is not None and sys.stdin.isatty():
        try:
            ans = input("确认查询这一份吗？直接回车继续，输入 n 取消: ").strip().lower()
        except (EOFError, KeyboardInterrupt):
            ans = ""
        if ans in ("n", "no"):
            print("已取消。想换一份，用 --uq-code 指定其他查询码。")
            return 2

    save_config(openid, uq_code, args.tu_code, ts)

    failed = []
    for i, name in enumerate(names, 1):
        if cache.get(name, {}).get("ok"):
            print("  [%d/%d] %s（缓存）" % (i, len(names), name))
            continue

        data = None
        while True:
            try:
                data = client.query(**{args.field: name})
                break
            except CooldownError:
                print("  !! 触发限流，等待 %.0f 分钟后继续" % (args.cooldown / 60))
                time.sleep(args.cooldown)
                try:
                    client.refresh_kkey()
                except Exception:
                    pass
            except QueryError as e:
                failed.append((name, e))
                print("  [%d/%d] %s -> %s" % (i, len(names), name, e))
                break
            except Exception as e:
                failed.append((name, e))
                print("  [%d/%d] %s -> 异常 %s" % (i, len(names), name, e))
                break

        if data is not None:
            cache[name] = {"ok": True, "data": data}
            cache_put(all_cache, uq_code, cache)
            save_cache(args.cache, all_cache)
            print("  [%d/%d] %s -> %d 行" % (i, len(names), name, len(record_rows(data))))

        time.sleep(args.delay + random.random() * 2)

    rows = []
    for name in names:
        entry = cache.get(name, {})
        if entry.get("ok"):
            for r in record_rows(entry["data"]):
                r.setdefault("查询条件", name)
                rows.append(r)

    meta = {
        "查询名称": title,
        "查询码": uq_code,
        "发布时间": update,
        "凭据来源": source,
        "抓取时间": time.strftime("%Y-%m-%d %H:%M:%S"),
        "数据行数": len(rows),
        "未命中数": len(failed),
    }
    out_path = args.out or auto_out()
    out = build_xlsx(rows, out_path, meta, failed)
    print("\n[+] %d 行 -> %s" % (len(rows), out or "(未导出)"))
    if rows:
        print("    提示：跑「名单收集.bat」可以把查过的姓名整理进 班级名单.txt，下次直接复用。")
    return 0 if rows else 1


if __name__ == "__main__":
    sys.exit(main())
