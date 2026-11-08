#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
yichafen-cli —— 批量查询易查分成绩并导出 Excel

最少只需要给名字：

    python cli.py --names names.txt

uqCode / openidCode 会自动从微信本地缓存里找，找不到再让你手填。

    # 手动指定也行
    python cli.py --names names.txt --uq-code YOUR_UQ_CODE --openid <你的openid>
"""

import argparse
import json
import os
import random
import sys
import time

from ycf import YiChaFen, CooldownError, QueryError, record_rows
import discover

DEFAULT_CACHE = "ycf-cache.json"
CONFIG = os.path.join(os.path.dirname(os.path.abspath(__file__)), "ycf-config.json")


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


def load_cache(path):
    if os.path.exists(path):
        try:
            with open(path, encoding="utf-8-sig") as f:
                return json.load(f)
        except Exception:
            return {}
    return {}


def save_cache(path, cache):
    tmp = path + ".tmp"
    with open(tmp, "w", encoding="utf-8") as f:
        json.dump(cache, f, ensure_ascii=False, indent=1)
    os.replace(tmp, path)


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


def load_config():
    """上一次成功用过的凭据，存在脚本同目录，方便二次运行。"""
    if os.path.exists(CONFIG):
        try:
            with open(CONFIG, encoding="utf-8-sig") as f:
                return json.load(f)
        except Exception:
            pass
    return {}


def save_config(openid, uq_code, tu_code):
    try:
        with open(CONFIG, "w", encoding="utf-8") as f:  # 写不带 BOM 的 UTF-8
            json.dump({"openid": openid, "uq_code": uq_code, "tu_code": tu_code},
                      f, ensure_ascii=False, indent=1)
    except Exception:
        pass


def resolve_target(args):
    """确定 openid / uqCode / tuCode。优先级：命令行 > 上次配置 > 微信缓存。"""
    saved = load_config()
    openid = args.openid or saved.get("openid")
    tu = args.tu_code or saved.get("tu_code")
    uq = args.uq_code or saved.get("uq_code")

    if not openid or not uq:
        print("[*] 正在从微信本地缓存里找查询凭据…")
        cands = discover.candidates(args.wechat_root or None)
        if cands:
            print(discover.describe(cands[:3]))
            openid = openid or cands[0][0]
            tu = tu or cands[0][1]
        else:
            print("[!] 没找到。微信的这部分缓存会被定期清理（重启微信后往往就没了）。")
            print("    最稳的办法：先在微信里打开一次那个查询小程序、随便查一个人，")
            print("    然后立刻回到这里重新运行 —— 新写的缓存就能被抓到。")
            print("    或者用 --openid / --uq-code 手动指定。")

    if not openid:
        raise SystemExit("缺少 openidCode。")

    if uq:
        save_config(openid, uq, tu)
        return openid, uq, tu

    if not tu:
        raise SystemExit("缺少 tuCode，无法列出查询项目。请用 --uq-code 直接指定查询码。")

    client = YiChaFen("", openid, proxy=args.proxy)
    print("[*] 用 tuCode=%s 拉取可用查询项目…" % tu)
    subjects = client.list_subjects(tu)
    if not subjects:
        raise SystemExit("这个账号下没有可用的查询项目。")

    print("\n可用查询项目：")
    for i, s in enumerate(subjects, 1):
        print("  %d) %-24s uqCode=%s  更新=%s"
              % (i, s.get("name"), s.get("uq_code"), s.get("update_time")))
    print()

    if args.pick:
        idx = args.pick
        if not (1 <= idx <= len(subjects)):
            raise SystemExit("--pick 超出范围")
        uq = subjects[idx - 1]["uq_code"]
        save_config(openid, uq, tu)
        return openid, uq, tu

    if len(subjects) == 1:
        uq = subjects[0]["uq_code"]
        save_config(openid, uq, tu)
        return openid, uq, tu

    raise SystemExit("有多个查询项目，请用 --pick N 选一个，或 --uq-code 直接指定。")


OUT_DIR = "成绩输出"


def auto_out(base="成绩.xlsx", dirname=OUT_DIR):
    """把结果收进「成绩输出」子目录，并自动编号，既不散在根目录也不覆盖历史。"""
    d = os.path.join(os.path.dirname(os.path.abspath(__file__)), dirname)
    os.makedirs(d, exist_ok=True)
    stem, ext = os.path.splitext(base)
    i = 1
    while os.path.exists(os.path.join(d, "%s%d%s" % (stem, i, ext))):
        i += 1
    return os.path.join(d, "%s%d%s" % (stem, i, ext))


def main():
    ap = argparse.ArgumentParser(
        description="批量查询易查分成绩并导出 Excel",
        epilog="最少用法：python cli.py --names names.txt")
    ap.add_argument("--names", nargs="+", required=True, help="姓名列表文件，或直接给名字")
    ap.add_argument("--uq-code", help="查询码；不给会自动发现")
    ap.add_argument("--openid", help="微信 openidCode；不给会自动发现")
    ap.add_argument("--tu-code", help="老师账号码；配合自动发现用")
    ap.add_argument("--pick", type=int, help="自动发现到多个查询项目时，选第几个")
    ap.add_argument("--field", default="s_xingming", help="查询字段名，默认 s_xingming")
    ap.add_argument("--out", default=None, help="输出 xlsx 路径；不给就自动编号，不覆盖历史结果")
    ap.add_argument("--cache", default=DEFAULT_CACHE, help="断点续抓缓存")
    ap.add_argument("--delay", type=float, default=12.0, help="每人间隔秒（默认 12）")
    ap.add_argument("--cooldown", type=float, default=1900, help="触发限流后等待秒")
    ap.add_argument("--proxy", help="可选代理，如 socks5h://127.0.0.1:1080")
    ap.add_argument("--wechat-root", action="append", default=[],
                    help="微信数据目录，可多次指定")
    args = ap.parse_args()

    names = load_names(args.names)
    if not names:
        raise SystemExit("没有可查询的姓名")

    openid, uq_code, tu = resolve_target(args)

    cache = load_cache(args.cache)
    client = YiChaFen(uq_code, openid, proxy=args.proxy)

    info = client.refresh_kkey().get("queryInfo") or {}
    title = info.get("name") or uq_code
    print("\n[+] 目标：%s（更新 %s）" % (title, info.get("update_time")))
    print("[+] 待查 %d 人，间隔 %.1fs" % (len(names), args.delay))

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
            save_cache(args.cache, cache)
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
        "查询码": uq_code,
        "标题": title,
        "更新时间": info.get("update_time"),
        "抓取时间": time.strftime("%Y-%m-%d %H:%M:%S"),
        "数据行数": len(rows),
        "未命中数": len(failed),
    }
    out_path = args.out or auto_out()
    out = build_xlsx(rows, out_path, meta, failed)
    print("\n[+] %d 行 -> %s" % (len(rows), out or "(未导出)"))
    return 0 if rows else 1


if __name__ == "__main__":
    sys.exit(main())
