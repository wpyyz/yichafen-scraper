# yichafen-scraper

批量查询 **易查分（yichafen.com）** 成绩并导出 Excel。

易查分的查询入口在微信小程序里，网页端通常被老师关掉。这个工具把小程序那三个接口还原出来，方便一次性把整班成绩整理成表格。

> 只用于查询你有权查看的成绩。

---

## 快速开始（Windows）

**不用会敲命令：**

1. 装 Python —— https://www.python.org/downloads/ ，安装时**勾上最下面的 `Add Python to PATH`**
2. 把姓名写进一个 txt，一行一个（可以照抄 `姓名示例.txt`）
3. **把那个 txt 拖到 `一键查分.bat` 上**，松手就行

想调查询间隔（防限流），双击 `设置间隔.bat` 输入秒数，存在 `interval.txt`。
默认 12 秒/人，一个班 70 人约 14 分钟。

首次运行会自动装 `openpyxl`，跑完在同目录生成 `成绩.xlsx`，并问你要不要直接打开。

## 命令行用法

```bash
pip install openpyxl

# 只要给名字，其余全自动
python cli.py --names names.txt
```

第一次跑会做两件事：

1. 从微信本地缓存里找 `openidCode` 和 `tuCode`
2. 用它们调 `subjectList`，列出这个账号下所有查询项目，让你选

```
可用查询项目：
  1) 某某班期中成绩     uqCode=YOUR_UQ_CODE  更新=2026.11.07
  2) 英语周测                uqCode=YOUR_UQ_CODE  更新=2026.10.23
  3) 某某班月考成绩         uqCode=YOUR_UQ_CODE  更新=2026.10.01
```

有多个时用 `--pick 1` 指定。也可以完全手动：

```bash
python cli.py --names 张三 李四 --uq-code YOUR_UQ_CODE --openid <你的openid>
```

产物：`成绩输出\成绩1.xlsx`（主表 + 说明页，含未命中清单）。每次运行自动编号，不会覆盖历史结果。

### 参数

| 参数 | 说明 | 默认 |
|---|---|---|
| `--names` | 姓名文件（一行一个）或直接给名字 | 必填 |
| `--pick` | 自动发现到多个查询项目时选第几个 | 无 |
| `--uq-code` | 手动指定查询码 | 自动发现 |
| `--openid` | 手动指定 openidCode | 自动发现 |
| `--tu-code` | 手动指定老师账号码 | 自动发现 |
| `--field` | 查询字段名 | `s_xingming` |
| `--out` | 输出 xlsx | `成绩.xlsx` |
| `--cache` | 断点续抓缓存 | `ycf-cache.json` |
| `--delay` | 每人间隔秒 | `12` |
| `--wechat-root` | 微信数据目录（自动猜不到时指定） | 自动 |

## 接口

三个接口，都是从小程序 `utils/request.js` 里还原的：

| 步骤 | 方法 | 路径 | 作用 |
|---|---|---|---|
| 1 | GET | `/Miniprogram/Api/subjectList` | 列出查询项目（拿 uqCode） |
| 2 | GET | `/Miniprogram/Api/qz` | 拿 `kkey`（签名令牌） |
| 3 | POST | `/Miniprogram/Api/verifyParams` | 查询条件 → `resultToken` |
| 4 | GET | `/Miniprogram/Api/subjectResultV4` | `resultToken` → 成绩 |

### 两个必须注意的细节

**1. 要带小程序 Referer，否则 404**

```
Referer: https://servicewechat.com/<appid>/<version>/page-frame.html
```

不带这个头，接口返回 404「请求来源不合法」。

**2. `openidCode` 和 `clientProtocolVersion` 要放在 body 里**

```http
POST /verifyParams?kkey=...&uqCode=...

s_xingming=张三&openidCode=...&clientProtocolVersion=1001
```

放 query 上会返回「授权参数错误」，和参数真的写错时提示一样，容易误判。

### 关于限流

实测到的现象：

- 短时间内查得多了，接口返回 **HTTP 502**，页面文字是「当前您的网络环境不稳定，请半小时后再来查询」
- 同一时刻，`/qz`、`/subjectList` 这些接口**仍然正常**，只有查询接口被挡
- 等约 30 分钟后恢复

**没有验证**的部分：到底按什么维度限（IP？openid？UA？还是组合），没做对照实验，不做结论。实测过的无效尝试：换 User-Agent、换出口 IP，都没能绕过。

触发之后如果继续每隔几分钟重试，会一直解不开——推测是按「最后一次请求」计时，但这一点同样没严格验证，只是现象相符。

所以默认 `--delay 12`（12 秒/人）。70 个人按这个节奏跑完没触发限流。

## 关于自动发现凭据

小程序每次请求都会把完整 URL 记进本地统计文件，是**明文**（不是加密数据库）：

```
%APPDATA%\Tencent\xwechat\net\kvcomm\*.statistic
```

里面形如 `https%3A%2F%2Fwww.yichafen.com%2FMiniprogram%2FApi%2Fqz%3FuqCode%3D...`，
URL 解码就能拿到 `uqCode`、`openidCode`、`tuCode`。

注意这三个值**不一定在同一条 URL 里**：`uqCode` 容易随缓存轮转消失，`openidCode` 和 `tuCode` 通常还在——所以先拿这两个，再用 `subjectList` 反查 `uqCode`。

### ★ 这段缓存会被清掉

微信重启或定期清理时，`kvcomm` 里的文件会轮换，**之前抓到的 yichafen 记录可能就没了**。

所以：

- **第一次运行前，先在微信里打开一次那个查询小程序**（随便查一个人），让缓存新鲜
- 成功一次之后，凭据会存进同目录的 `ycf-config.json`，以后就不用再找了
- 实在找不到，用 `--openid` / `--uq-code` 手动指定，也会记进配置文件

## 两个容易踩的小坑

**姓名文件带 BOM**

Windows 记事本存 UTF-8 会加 BOM，第一个名字会变成 `\ufeff张三`，查询必然失败。程序里已经用 `utf-8-sig` 处理掉了，但你手动给 `--names` 传名字时也要注意别贴进不可见字符。

**git 会把 .bat 存成 LF**

本仓库用 `.gitattributes` 强制 `*.bat text eol=crlf`，克隆下来才是 CRLF。如果你自己 fork 后删了这个文件，bat 在中文 cmd 里会报一堆「'----------' 不是内部或外部命令」。

## 作为库用

```python
from ycf import YiChaFen, record_rows

c = YiChaFen(uq_code="YOUR_UQ_CODE", openid_code="<你的openid>")
data = c.query(s_xingming="张三")
print(record_rows(data))
```

## 前提：必须有可用的微信身份

这一点容易被误解，先说清楚。

`openidCode` 是**服务端校验**的，实测：

| openidCode | `/qz` 返回 |
|---|---|
| 真实（本机微信的） | 正常下发 kkey |
| 伪造（随便编一串） | `授权参数错误` |
| 空 | `授权参数错误` |

所以：

- **没有微信登录环境的机器上，这工具查不了任何东西** —— 本地找不到 openid 会直接退出，一个请求都不发
- 手动塞一个编造的 openid 也没用，第一步就被服务端拒了
- 它替代的是「手动一个个点」，**不是**「没有身份也能查」

两个东西的分工：

```
openidCode  →  你是谁（校验，伪造不了）
uqCode      →  查哪一份（老师发布时生成，每班不同）
```

有微信 + 知道 uqCode，就能查任意可访问的查询；只有 uqCode 没微信，什么也做不了。

## 限制

- 查询条件由老师配置，通常只有姓名，且**精确匹配**——不支持模糊、通配。
- 拿不到正确姓名就查不出来（比如同学在群里用的是网名）。这属于数据源本身的限制。
- 令牌会过期，长期跑请重新查而不是复用缓存。
- 微信缓存的解析依赖 `*.statistic` 这类文件的明文格式，微信改版后可能失效。

## License

MIT
