# -*- coding: utf-8 -*-
"""
ycf —— 易查分（yichafen.com）小程序接口客户端

只用标准库，不依赖 requests。

接口链路（逆向自小程序的 utils/request.js）：

    1) GET  /qz            -> 取 kkey（签名令牌，约 2 小时有效）
    2) POST /verifyParams  -> 用查询条件换 resultToken
    3) GET  /subjectResultV4 -> 用 resultToken 取成绩

三个坑：
    * 必须带 Referer: https://servicewechat.com/<appid>/<ver>/page-frame.html，
      否则接口直接 404「请求来源不合法」。
    * openidCode 和 clientProtocolVersion 必须放在 **body** 里（放 query 上会报
      「授权参数错误」）。
    * 高频请求会触发「当前您的网络环境不稳定，请半小时后再来查询」的冷却，
      窗口按最后一次请求计算。默认限速 12 秒/人比较安全。
"""

import json
import random
import ssl
import time
import urllib.error
import urllib.parse
import urllib.request

API_BASE = "https://www.yichafen.com/Miniprogram/Api"
APPID = "wx9300e397f1f28ec7"
CLIENT_PROTOCOL_VERSION = "1001"

DEFAULT_UA = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
    "(KHTML, like Gecko) Chrome/131.0 Safari/537.36 "
    "MicroMessenger/7.0.20.1781(0x6700143B) WindowsWechat(0x63090a13)"
)

COOLDOWN_MARK = "网络环境不稳定"


class CooldownError(RuntimeError):
    """触发了服务端的反爬冷却。"""


class QueryError(RuntimeError):
    """业务层错误（姓名不对、时间不对等）。"""


class YiChaFen:
    def __init__(self, uq_code, openid_code, appid=APPID,
                 client_protocol_version=CLIENT_PROTOCOL_VERSION,
                 user_agent=DEFAULT_UA, timeout=30, proxy=None):
        self.uq_code = uq_code
        self.openid_code = openid_code
        self.appid = appid
        self.cpv = client_protocol_version
        self.timeout = timeout
        self.headers = {
            "User-Agent": user_agent,
            "Referer": "https://servicewechat.com/%s/68/page-frame.html" % appid,
            "Origin": "https://servicewechat.com",
            "Accept": "application/json, text/plain, */*",
        }
        self._kkey = None
        self._ssl = ssl.create_default_context()

        if proxy:
            # proxy 形如 socks5h://127.0.0.1:1080 或 http://127.0.0.1:8080
            handler = urllib.request.ProxyHandler({"http": proxy, "https": proxy})
        else:
            # 关键：忽略系统环境里的 http(s)_proxy，避免被本地代理链路污染
            handler = urllib.request.ProxyHandler({})
        self._opener = urllib.request.build_opener(handler)

    # ---------------- 底层 ----------------

    def _request(self, method, path, params=None, data=None, tries=3):
        url = API_BASE + path
        if params:
            url += "?" + urllib.parse.urlencode(params)

        body = None
        headers = dict(self.headers)
        if data is not None:
            body = urllib.parse.urlencode(data).encode("utf-8")
            headers["Content-Type"] = "application/x-www-form-urlencoded"

        last = None
        for i in range(tries):
            try:
                req = urllib.request.Request(url, data=body, headers=headers, method=method)
                with self._opener.open(req, timeout=self.timeout) as resp:
                    raw = resp.read().decode("utf-8", "replace")
                if COOLDOWN_MARK in raw:
                    raise CooldownError("触发限流：网络环境不稳定")
                return json.loads(raw)
            except CooldownError:
                raise
            except urllib.error.HTTPError as e:
                text = e.read().decode("utf-8", "replace")
                if COOLDOWN_MARK in text:
                    raise CooldownError("触发限流：网络环境不稳定")
                last = "HTTP %s %s" % (e.code, text[:120])
            except Exception as e:  # 网络抖动
                last = str(e)
            time.sleep(1.5 + i * 2 + random.random())
        raise QueryError("请求失败：%s" % last)

    # ---------------- 业务 ----------------

    def refresh_kkey(self):
        j = self._request("GET", "/qz", params={
            "uqCode": self.uq_code,
            "doorToken": "",
            "openidCode": self.openid_code,
            "clientProtocolVersion": self.cpv,
        })
        data = j.get("data") or {}
        if not data.get("kkey"):
            raise QueryError("取 kkey 失败：%s" % json.dumps(j, ensure_ascii=False)[:200])
        self._kkey = data["kkey"]
        return data

    @property
    def kkey(self):
        if not self._kkey:
            self.refresh_kkey()
        return self._kkey

    def list_subjects(self, tu_code):
        """列出该老师账号下所有已发布的查询项目（含 uq_code）。"""
        j = self._request("GET", "/subjectList", params={
            "tuCode": tu_code,
            "doorToken": "",
            "openidCode": self.openid_code,
            "clientProtocolVersion": self.cpv,
        })
        return (j.get("data") or {}).get("queryList") or []
    def get_token(self, **conditions):
        """conditions 形如 s_xingming="张三"，键名来自 /qz 的 queryFields。"""
        payload = dict(conditions)
        payload["openidCode"] = self.openid_code
        payload["clientProtocolVersion"] = self.cpv

        j = self._request("POST", "/verifyParams",
                          params={"kkey": self.kkey, "uqCode": self.uq_code},
                          data=payload)
        data = j.get("data") or {}
        if j.get("status") == 1 and data.get("token"):
            return data["token"]
        raise QueryError(j.get("info") or "errCode=%s" % j.get("errCode"))

    def get_result(self, token):
        j = self._request("GET", "/subjectResultV4", params={
            "uqCode": self.uq_code,
            "token": token,
            "openidCode": self.openid_code,
            "clientProtocolVersion": self.cpv,
        })
        if j.get("status") != 1:
            raise QueryError(j.get("info") or "取结果失败")
        return j.get("data") or {}

    def query(self, **conditions):
        """一步到位：条件 -> 结果 data。"""
        return self.get_result(self.get_token(**conditions))


# ---------------- 结果解析 ----------------

def record_rows(result_data, key_name="s_xingming"):
    """把 recordList 拍平成 [dict]，列名用中文表头。"""
    rows = []
    for item in result_data.get("recordList") or []:
        row = {}
        for col in item.get("columnList") or []:
            name = col.get("name") or col.get("pinyin")
            row[str(name)] = col.get("value")
        if row:
            rows.append(row)
    return rows
