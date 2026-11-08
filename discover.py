# -*- coding: utf-8 -*-
"""
从微信本地缓存里自动找出易查分的 uqCode / openidCode。

原理：小程序每次发请求，都会把完整 URL 记进本地统计文件
（`%APPDATA%\\Tencent\\xwechat\\net\\kvcomm\\*.statistic` 之类），
里面是 URL 编码后的明文，不是加密数据库。

注意：这两个值不一定出现在同一条 URL 里，得分别收集再配对。
例如：
    .../qz?uqCode=YOUR_UQ_CODE&openidCode=<你的openid>      ← 两个都有
    .../getUserInfoByTuCode?tuCode=YOUR_TU_CODE&openidCode=<你的openid>   ← 只有 openid
"""

import argparse
import os
import re
import time
import urllib.parse

URL_RE = re.compile(
    rb"(?:https?%3A%2F%2F|https?://)(?:www\.)?yichafen\.com[^\x00-\x20\"'<>|]{0,400}"
)

MAX_FILE_BYTES = 80 * 1024 * 1024
SKIP_DIRS = {"cache", "FileStorage", "Video", "Image", "attach"}


def default_roots():
    roots = []
    for env in ("APPDATA", "LOCALAPPDATA"):
        base = os.environ.get(env)
        if base:
            roots.append(os.path.join(base, "Tencent"))
    userprofile = os.environ.get("USERPROFILE")
    if userprofile:
        roots.append(os.path.join(userprofile, "Documents", "WeChat Files"))
    for drive in ("C:", "D:", "E:", "F:", "G:"):
        p = os.path.join(drive + os.sep, "wxltjl")
        if os.path.isdir(p):
            roots.append(p)
    seen, out = set(), []
    for r in roots:
        if r and r not in seen and os.path.isdir(r):
            seen.add(r)
            out.append(r)
    return out


def _params(text):
    try:
        q = urllib.parse.parse_qs(urllib.parse.urlparse(text).query)
    except Exception:
        return {}
    return {k: v[0] for k, v in q.items() if v}


def scan(roots=None, max_files=50000, quiet=True):
    """
    返回 {
      "uq":  {uqCode: (mtime, path)},
      "oid": {openidCode: (mtime, path)},
      "pairs": {(uqCode, openidCode): (mtime, path)},
      "tu":  {tuCode: (mtime, path)},
    }
    """
    roots = roots or default_roots()
    acc = {"uq": {}, "oid": {}, "pairs": {}, "tu": {}}
    seen = 0

    for root in roots:
        for dirpath, dirnames, filenames in os.walk(root):
            dirnames[:] = [d for d in dirnames if d not in SKIP_DIRS]
            for fn in filenames:
                seen += 1
                if seen > max_files:
                    break
                path = os.path.join(dirpath, fn)
                try:
                    if os.path.getsize(path) > MAX_FILE_BYTES:
                        continue
                    with open(path, "rb") as f:
                        blob = f.read()
                except Exception:
                    continue
                if b"yichafen" not in blob:
                    continue

                mtime = os.path.getmtime(path)
                for m in URL_RE.finditer(blob):
                    text = urllib.parse.unquote(m.group(0).decode("utf-8", "ignore"))
                    ps = _params(text)
                    uq = ps.get("uqCode")
                    oid = ps.get("openidCode")
                    tu = ps.get("tuCode")

                    if uq and (uq not in acc["uq"] or acc["uq"][uq][0] < mtime):
                        acc["uq"][uq] = (mtime, path)
                    if oid and (oid not in acc["oid"] or acc["oid"][oid][0] < mtime):
                        acc["oid"][oid] = (mtime, path)
                    if tu and (tu not in acc["tu"] or acc["tu"][tu][0] < mtime):
                        acc["tu"][tu] = (mtime, path)
                    if uq and oid:
                        key = (uq, oid)
                        if key not in acc["pairs"] or acc["pairs"][key][0] < mtime:
                            acc["pairs"][key] = (mtime, path)
    return acc


def candidates(roots=None):
    """
    返回 [(openidCode, tuCode, mtime, 来源说明)]，按时间倒序。
    uqCode 不保证还在缓存里（会被轮转），所以这里不强求；
    拿到 openid + tuCode 之后可以用 subjectList 接口列出所有查询项目。
    """
    acc = scan(roots=roots)
    if not acc["oid"]:
        return []

    newest_tu = None
    if acc["tu"]:
        newest_tu, _ = max(acc["tu"].items(), key=lambda kv: kv[1][0])

    out = []
    for oid, (mt, path) in acc["oid"].items():
        out.append((oid, newest_tu, mt, path))
    out.sort(key=lambda x: -x[2])
    return out


def describe(rows):
    if not rows:
        return "没在本地缓存里找到易查分记录。用 --wechat-root 指定微信数据目录试试。"
    lines = []
    for i, (oid, tu, mt, src) in enumerate(rows, 1):
        lines.append("%2d) openid=%s…  tuCode=%s   %s"
                     % (i, oid[:16], tu or "-", time.strftime("%Y-%m-%d %H:%M", time.localtime(mt))))
        lines.append("      %s" % src)
    return "\n".join(lines)


if __name__ == "__main__":
    ap = argparse.ArgumentParser(description="从微信本地缓存里找出易查分 uqCode / openidCode")
    ap.add_argument("--wechat-root", action="append", default=[],
                    help="微信数据目录，可多次指定；不给就自动猜")
    args = ap.parse_args()
    print(describe(candidates(args.wechat_root or None)))

def recent_uq(roots=None):
    """
    返回缓存里出现过的 uqCode，按最后出现时间倒序：
        [(uqCode, mtime, 来源路径)]
    用来判断「用户最近打开的是哪一次查询」。
    """
    acc = scan(roots=roots)
    out = [(uq, mt, path) for uq, (mt, path) in acc["uq"].items()]
    out.sort(key=lambda x: -x[1])
    return out


def recent_openid(roots=None):
    """最新的 openidCode，没有就 None。"""
    acc = scan(roots=roots)
    if not acc["oid"]:
        return None
    oid, (mt, path) = max(acc["oid"].items(), key=lambda kv: kv[1][0])
    return oid
