# -*- coding: utf-8 -*-
"""
列出 / 切换「要查哪一份成绩」。

易查分的每一份成绩有自己的 uqCode，光看码根本不知道对应哪次考试。
这个模块的做法是：拿 openid 挨个去问一下服务端，把中文名和发布时间取回来，
拼成一张人类看得懂的表，再让你选。

用法：
    python queries.py                # 列出所有已知查询
    python queries.py --pick 2       # 直接选第 2 个，写进配置
"""

import argparse
import json
import os
import sys
import time

import discover
from ycf import YiChaFen, CooldownError, QueryError

CONFIG = os.path.join(os.path.dirname(os.path.abspath(__file__)), "ycf-config.json")


def load_config():
    if os.path.exists(CONFIG):
        try:
            with open(CONFIG, encoding="utf-8-sig") as f:
                return json.load(f)
        except Exception:
            pass
    return {}


def save_config(openid, uq_code, uq_ts=None, tu_code=None):
    with open(CONFIG, "w", encoding="utf-8") as f:
        json.dump({"openid": openid, "uq_code": uq_code,
                   "tu_code": tu_code, "uq_ts": uq_ts},
                  f, ensure_ascii=False, indent=1)


def fetch_title(openid, uq_code, proxy=None, timeout=20):
    """问服务端这份查询叫什么、什么时候发布的。失败就返回 (None, None)。"""
    try:
        c = YiChaFen(uq_code, openid, proxy=proxy, timeout=timeout)
        info = c.refresh_kkey().get("queryInfo") or {}
        return info.get("name"), info.get("update_time")
    except (CooldownError, QueryError, Exception):
        return None, None


def parse_time(s):
    """服务端有两种日期格式：2026.11.07 和 2026-11-07 11:36:25，都要认。"""
    if not s:
        return 0.0
    s = str(s).strip().replace(".", "-")
    for fmt in ("%Y-%m-%d %H:%M:%S", "%Y-%m-%d %H:%M", "%Y-%m-%d"):
        try:
            return time.mktime(time.strptime(s[:len(time.strftime(fmt, time.localtime(0)))], fmt))
        except Exception:
            continue
    return 0.0


def from_subject_list(openid, tu, proxy=None):
    """
    用老师账号码问服务端要「这个账号下发布过的全部查询」。
    这是最完整的来源，名称也是服务端给的，不会猜错。
    """
    try:
        c = YiChaFen("", openid, proxy=proxy)
        subs = c.list_subjects(tu) or []
    except Exception:
        return []
    out = []
    for s in subs:
        uq = s.get("uq_code")
        if uq:
            out.append((uq, s.get("name"), s.get("update_time")))
    return out


def collect(openid, roots=None, extra=None, proxy=None, quiet=False):
    """
    返回 [(uq, 名称, 发布时间, 时间戳)]，按时间倒序。

    优先用 subjectList（完整、带名字），再合并微信缓存里出现过的
    （缓存能提供「最后打开时间」，用来判断哪份最新）。
    """
    cache_uq = {uq: mt for uq, mt, _p in discover.recent_uq(roots)}

    tu = discover.recent_tu(roots)
    if not tu:
        cfg = load_config()
        tu = cfg.get("tu_code")

    known = {}
    if tu:
        if not quiet:
            print("  正在向服务端拉取该账号下的查询清单…")
        for uq, name, upd in from_subject_list(openid, tu, proxy=proxy):
            known[uq] = (name, upd)

    # 合并：subjectList 里没有、但缓存/配置里出现过的，单独去问名字
    for uq in list(cache_uq.keys()) + [u for u in (extra or []) if u]:
        if uq in known:
            continue
        if not quiet:
            print("  查询 %s …" % uq, end="", flush=True)
        name, upd = fetch_title(openid, uq, proxy=proxy)
        if not quiet:
            print(" %s" % (name or "（取不到，可能已失效）"))
        known[uq] = (name, upd)
        time.sleep(1.2)

    out = []
    for uq, (name, upd) in known.items():
        opened_mt = cache_uq.get(uq, 0.0)
        out.append({
            "uq": uq,
            "name": name or "（名称未知）",
            "pub": upd or "未知",
            "opened": time.strftime("%m-%d %H:%M", time.localtime(opened_mt)) if opened_mt else "",
            "sort": opened_mt or parse_time(upd),
        })

    out.sort(key=lambda x: -x["sort"])
    return out


def render(rows, current=None, show_cmd=True):
    if not rows:
        return "  （没有找到任何可用的查询）"

    lines = []
    width = max([len(r["name"]) for r in rows] + [10])
    for i, r in enumerate(rows, 1):
        mark = "  ← 当前在用" if r["uq"] == current else ""
        opened = r["opened"] or "—"
        lines.append("  %d) %-*s  %s   发布 %-10s 最近打开 %s%s"
                     % (i, width, r["name"], r["uq"], r["pub"], opened, mark))
    if show_cmd and rows:
        lines.append("")
        lines.append("  换用第 N 个：--uq-code %s" % rows[0]["uq"])
        lines.append("  例：python cli.py --names 班级名单.txt --uq-code %s" % rows[0]["uq"])
    return "\n".join(lines)


def describe_others(openid, others, current, proxy=None):
    """给 cli.py 用：把备选查询连中文名一起列出来。"""
    if not others:
        return ""
    rows = []
    for ts, uq in others:
        name, upd = fetch_title(openid, uq, proxy=proxy)
        rows.append({"uq": uq, "name": name or "（名称未知）",
                     "pub": upd or "未知",
                     "opened": time.strftime("%m-%d %H:%M", time.localtime(ts)) if ts else "",
                     "sort": ts})
        time.sleep(1.0)

    lines = ["", "[i] 还能查这些（按最近打开排序）："]
    width = max([len(r["name"]) for r in rows] + [10])
    for i, r in enumerate(rows, 1):
        lines.append("    %d) %-*s  %s   发布 %-10s 最近打开 %s"
                     % (i, width, r["name"], r["uq"], r["pub"], r["opened"] or "—"))
    lines.append("")
    lines.append("    想换某一份，把 --uq-code 后面换成它的查询码，例如：")
    lines.append("      python cli.py --names 班级名单.txt --uq-code %s" % rows[0]["uq"])
    lines.append("    双击 bat 的用户：运行「选择查询.bat」，按编号选就行。")
    lines.append("")
    return "\n".join(lines)


def main():
    ap = argparse.ArgumentParser(description="列出 / 切换要查哪一份成绩")
    ap.add_argument("--openid", help="微信 openidCode；不给会自动发现")
    ap.add_argument("--pick", type=int, help="直接选第几个并写进配置")
    ap.add_argument("--wechat-root", action="append", default=[])
    ap.add_argument("--proxy")
    args = ap.parse_args()

    cfg = load_config()
    openid = args.openid or cfg.get("openid")
    if not openid:
        print("没找到 openidCode，先扫描微信缓存…")
        openid = discover.recent_openid(args.wechat_root or None)
    if not openid:
        print("")
        print("找不到微信身份。请先在【微信电脑端】打开一次那个查询小程序，再运行本工具。")
        return 1

    print("")
    print("正在读取各份查询的名称…")
    rows = collect(openid, roots=args.wechat_root or None,
                   extra=[cfg.get("uq_code"), args.uq_code if hasattr(args, "uq_code") else None],
                   proxy=args.proxy)
    if not rows:
        print("")
        print("没找到任何查询。请先在微信电脑端打开一次要查的那个小程序。")
        return 1

    print("")
    print("=" * 60)
    print("  已知的查询")
    print("=" * 60)
    print(render(rows, current=cfg.get("uq_code"), show_cmd=False))
    print("=" * 60)

    pick = args.pick
    if pick is None:
        try:
            raw = input("\n要查第几个？直接回车 = 保持现在的 %s: "
                        % (cfg.get("uq_code") or rows[0][0]))
        except (EOFError, KeyboardInterrupt):
            raw = ""
        raw = (raw or "").strip()
        if not raw:
            if cfg.get("uq_code"):
                print("保持不变。")
                return 0
            pick = 1
        else:
            if not raw.isdigit():
                print("输入的不是数字，退出。")
                return 1
            pick = int(raw)

    if not (1 <= pick <= len(rows)):
        print("编号超出范围。")
        return 1

    chosen = rows[pick - 1]
    uq, name, upd = chosen["uq"], chosen["name"], chosen["pub"]
    save_config(openid, uq, uq_ts=time.time(), tu_code=cfg.get("tu_code"))
    print("")
    print("已设为：%s" % name)
    print("查询码：%s" % uq)
    print("发布会：%s" % (upd or "未知"))
    print("")
    print("下次双击「一键查分.bat」就查这一份。")
    return 0


if __name__ == "__main__":
    sys.exit(main())
