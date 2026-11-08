# yichafen-scraper

批量查询 **易查分（yichafen.com）** 成绩并导出 Excel。

易查分的查询入口在微信小程序里，网页端通常被老师关掉。这个工具把小程序那三个接口还原出来，方便一次性把整班成绩整理成表格。

> 只用于查询你有权查看的成绩。

---

## 这个工具到底在做什么

易查分的正常设计是「一人一档」：家长用自己的微信进来，**只能查到自己孩子那一行**。老师发布时也不会把名单给任何家长。

本工具做的是另一件事：**拿一个身份，把整份名单逐个人名挨着查一遍**，从而把全班所有人的成绩全部拉下来。

关键在于，它不需要破解登录、不需要老师账号 —— 服务端只校验「这个姓名在不在名单里」，**不校验「这个 openid 是不是这个学生的家长」**。

实测佐证：同一个 openid 连续查了 72 个不同姓名，全部返回成功。
`one_w_one_p`（一人一档）那个开关只拦「同一条件重复查」，不拦「换个人名再查」。

所以这是**绕过业务限制**，不是效率优化。说清楚比较好。


`openidCode` 是**服务端校验**的，实测：

| openidCode | `/qz` 返回 |
|---|---|
| 真实（本机微信的） | 正常下发 kkey |
| 伪造（随便编一串） | `授权参数错误` |
| 空 | `授权参数错误` |

所以：

- **没有微信登录环境的机器上，这工具查不了任何东西** —— 本地找不到 openid 会直接退出，一个请求都不发
- 手动塞一个编造的 openid 也没用，第一步就被服务端拒了
- 它替代的是「只能查自己孩子」这条限制，**不是**「没有身份也能查」

两个东西的分工：

```
openidCode  →  你是谁（校验，伪造不了）
uqCode      →  查哪一份（老师发布时生成，每班不同）
```

有微信 + 知道 uqCode，就能查任意可访问的查询；只有 uqCode 没微信，什么也做不了。

## 每次查的是哪一份？

成绩一个月出一次，每次都是新链接，所以查询码不能固定死。

选取规则（缓存与配置比时间戳，谁新用谁）：

```
1. --uq-code 显式指定        ← 绝对优先
2. 微信缓存里最后打开的那一个   ← 多个取最新
3. 都没有 → 用上次用过的查询码
4. 再没有 → 报错，提示先去微信电脑端打开一次
```

运行前会打出横幅，确认无误再开始：

```
================================================
  本次查询：某某班期中成绩
  查询码  ：xxxxxxx
  发布时间：2026-12-01 10:00:00
  凭据来源：微信缓存，最后打开于 2026-12-01 10:20
  待查人数：72 人 × 12 秒
================================================
```

交互式运行时还会问一句「确认查询这一份吗」；被 bat 调用时自动跳过（传 `--yes`）。

**老师发了新链接怎么办**：先在**微信电脑端**点开一次那个新查询（随便查一个人），再运行脚本，就会自动跟上。不用删配置、不用改文件。

> 依赖的是 Windows/Mac 版微信的本地缓存，手机微信抓不到。

### 换一份查

光看查询码根本不知道对应哪次考试，所以别手抄。

双击 **`选择查询.bat`**：它用 `subjectList` 把该账号下**所有查询连中文名、发布日期一起列出来**，
标出当前在用的那份，输入编号就能切换。

```
============================================================
  已知的查询
============================================================
  1) 某某班第一次练习     aaaa1111   发布 2026.10.01  最近打开 11-08 20:40
  2) 英语周测            bbbb2222   发布 2026.10.23  最近打开 11-08 19:40
  3) 某某班第二次练习成绩  cccc3333   发布 2026.11.07  最近打开 —   ← 当前在用
============================================================

  要查第几个？直接回车 = 保持现在的 cccc3333:
```

命令行等价写法（`--uq-code` 后面**直接跟查询码，不加引号**）：

```bash
# 查「英语周测」这一份，名单用 班级名单.txt，间隔 20 秒
python cli.py --names 班级名单.txt --uq-code bbbb2222 --delay 20
```

| 参数 | 含义 |
|---|---|
| `--names` | 名单文件，一行一个姓名 |
| `--uq-code` | 要查哪一份，从 `选择查询.bat` 里抄 |
| `--delay` | 每人间隔秒，可不写（默认 12） |

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

## 完整的接口清单（反编译小程序得到）

包在 `%APPDATA%\Tencent\xwechat\radium\users\<id>\applet\packages\wx9300e397f1f28ec7\<ver>\__APP__.wxapkg`，
微信 4.x 是 `V1MMWX` 格式（AES-256-CBC + 按 wxid 派生 XOR），解开后 `app-service.js` 里能拿到全部请求调用点。

**客户端一共只有这 30 个接口：**

```
/qz                  /verifyParams        /subjectResultV4     /subjectList
/queryResultPageV2   /refreshQueryTicket  /typeSuccessV2       /scoreAnalysis
/favScoreAnalysis    /favQueryListV2      /favScoreQueryResult /doDeleteFavScore
/doLockResult        /doSaveEditedFields  /doSubmitSignData     /doVerifyDoorPwd
/doVerifyGroupDoorPwd /getDoorConfig      /getChangeLongPwdInfo /changeLongPwd
/checkSmsCode        /checkTeacherCode    /getMiniprogramInfo  /getUserInfoByTuCode
/gp                  /dfV3                /submitDataFormV1    /submitComment
/statisBtnsClickCnt  /vipSpecifyAccnout
```

## 为什么「拿不到名单」

经常有人问：能不能不提供姓名，直接把这批人的名字扒下来？

**不能。** 下面每一条都实测过，全部堵死：

| 类别 | 思路 | 结果 |
|---|---|---|
| 条件注入 | 空值 / 通配符 `%` `_` `*` | `[ 姓名 ] 不能为空` 或「姓名不正确」 |
| 条件注入 | 单字、前缀（王 / 王鹏） | 精确匹配，一律「姓名不正确」 |
| 条件注入 | `unconditional=1` 无条件查询 | 服务端忽略该参数 |
| 条件注入 | 同名参数重复提交（`s_xingming=%&s_xingming=张三`） | 取最后一个值，无绕过 |
| 条件注入 | JSON 对象 `{"$ne":null}` | 当字面量比对，「姓名不正确」 |
| 类型混淆 | `s_xingming[]=张三` 数组注入 | `查询参数格式错误` |
| 参数篡改 | 附加 `isTeacher` / `admin` / `debug` / `one_w_one_p` / `mpType` | 全部被忽略，不存在的名字照样拒绝 |
| 身份伪造 | 伪造 `openidCode` | `/qz` 直接 `授权参数错误` |
| 身份伪造 | 换 `clientProtocolVersion`（1/2/1000/9999） | 无差别 |
| 身份伪造 | 换 Referer 里的 appid | `请求来源不合法` |
| 越权 | `uqCode → tuCode` 反推 | **可行**（`certificationUrl` 泄露），但 tuCode 只能列查询名称 |
| 越权 | 教师编号 `/checkTeacherCode` | 需正确编号，且只放行进入查询页，不含数据 |
| 越权 | 开门口令 `/doVerifyDoorPwd` `/getDoorConfig` | 参数错误 / 404 |
| 分页 | `/queryResultPageV2` `pageNum=2,3` | 第 1 页 1 条，之后 0 条，`totalPage=null` |
| 导出 | `resultExcel` / `getFileByKey` | POST `resultExcel` 拿 `fileKey`，绑 `resultToken`，单人文件 |
| 联想 | 补全接口（12 个候选路径） | 全部 404 |
| 站点 | 老师站 `xxxx.yichafen.com` 各路径 | 302 到 `jumpminip.html`；`/Index/certification` 是公开证明页，无数据 |
原因在客户端代码里写得很清楚：**每个取数据的接口都要求 `token`，而 `token` 只能由「正确且非空的查询条件」换取**。
`/queryResultPageV2`、`/scoreAnalysis` 也都是 `{token, ...}` 形态，token 绑死了条件，翻页只会翻同一个人的多行。

学生端小程序里没有下发名单的接口——那份名单在教师端。

**所以名单只能从别处来**：班级群成员、成绩单、老师发的表格。
这也正是本工具要求你给一份姓名文件的原因。
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


## 限制

- 查询条件由老师配置，通常只有姓名，且**精确匹配**——不支持模糊、通配。
- 拿不到正确姓名就查不出来（比如同学在群里用的是网名）。这属于数据源本身的限制。
- 令牌会过期，长期跑请重新查而不是复用缓存。
- 微信缓存的解析依赖 `*.statistic` 这类文件的明文格式，微信改版后可能失效。

## License

MIT
