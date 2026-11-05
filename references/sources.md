# 三信源清单与检索要点（2026-09-24 全量核查后修订）

> 检索目标：私募基金管理人、私募基金产品、私募从业人员的处罚/监管案例。
> **硬性要求：每次执行必须「全量遍历」，不允许抽样、不允许只查重点局、不允许只取前几页。**
> 覆盖度自检口径见文末「五、覆盖度自检」。

## 一、中基协（中国证券投资基金业协会）

- 官网：https://www.amac.org.cn
- 栏目（两个，都必须全量遍历）：
  - 受处分机构：`https://www.amac.org.cn/zlgl/jlcf/scfjg/`
  - 受处分人员：`https://www.amac.org.cn/zlgl/jlcf/scfry/`
- **列表分页（可直接 urllib 抓取，无需浏览器）**：
  - 列表首页 HTML 里有 `createPageHTML(<总页数>, 0, "index","html")`，据此得总页数 N。
  - 第 1 页 `index.html`，第 i 页 `index_{i}.html`（i=1..N-1）。
  - 条目正则：`href="([^"]+?P0\d+\.pdf)"[^>]*>(?:\s*<[^>]+>\s*)*([^<]{4,160})`，
    只保留 URL 含 `/zlgl/jlcf/` 的（页面里混有协会新闻等无关 PDF 链接）。
- 措施类型：纪律处分、公开谴责、暂停备案、注销/撤销登记、取消会员资格、纪律处分复核决定。
- **决定书 PDF 是扫描件，无文字层**（`page.get_text()` 返回空）。必须走 OCR：
  ```
  pip install rapidocr-onnxruntime
  from rapidocr_onnxruntime import RapidOCR
  ocr = RapidOCR(); res,_ = ocr(page.get_pixmap(dpi=160).tobytes('png'))
  ```
  取前 3 页即可判断违规类型、文号、当事人与落款日期。
- 文号正则：`中基协处分[〔\[（(]\s*(\d{4})\s*[〕\]）)]\s*(\d+)\s*号`（OCR 常把〔〕识别成（）或[]，正则要宽松）。

## 二、各地证监局（中国证监会派出机构）——**必须覆盖全部 36 家**

- 证监会官网：https://www.csrc.gov.cn → 「派出机构」进入各局子站。
- **列表接口（可直接 urllib 抓取，无需浏览器）**：
  ```
  https://www.csrc.gov.cn/searchList/{channelId}?_isAgg=true&_isJson=true
      &_pageSize=20&_template=index&_rangeTimeGte=&_channelName=&page={n}
  ```
  - `_pageSize` 实际被服务端限制为 **20**，按 `data.total` 翻页即可。
  - 返回体字段：`data.results[]` 每项含 `title / url / content / publishedTimeStr / manuscriptId`，
    以及 `domainMetaList[].resultList[]` 里的元数据（`fwrq` 发文日期、`wh` 文号）。
  - 列表按日期倒序，遇到第一条早于窗口起点的记录即可停止翻页。
  - 正文页 `https://www.csrc.gov.cn/<局拼音>/c<N>/c<N>/<id>/content.shtml` 为静态页，
    用 Chrome `--headless=new --print-to-pdf` 打印成 PDF 留存。
- **栏目 ID 每局独立，绝不共享**（曾误判为「全局共享」导致只抓到上海局，务必按表取值）：

| 局 | 拼音 | 行政监管措施 channelId | 行政处罚决定 channelId |
|---|---|---|---|
| 上海 | `shanghai` | `9ebc4198232e496e8bebf1b1bb1778ef` | `c8318fc200764e38b30116c2d5f72b4b` |
| 云南 | `yunnan` | `ebb6ce034bb84594a0f0c06fccecde9b` | `99481e931fc94d13b8592449acb386a4` |
| 内蒙古 | `neimenggu` | `4c96602a020c4398adfc7123a48f2ee1` | `e4f2dbcdd85c4bac9a9e9a6f2d093d19` |
| 北京 | `beijing` | `ea82afe8502d46aebea7346f0729dd85` | `53be9cf2273744cda5c5780e0dded972` |
| 厦门 | `xiamen` | `e3feac23205440e5b10ed81efa9dd94b` | `0c1b68108fca4de0b240caaa32f901ed` |
| 吉林 | `jilin` | `366f3072cbe143bebd533ef37f04eff5` | `9a3520fd5ea644c2ad6e05acd0a05401` |
| 四川 | `sichuan` | `c8ec4969173148c19b6d52afc6500b49` | `6cd6f0ccbd2f49c6af32878d34c54ae6` |
| 大连 | `dalian` | `fdf5c88757e04212a6dd66f396b21168` | `fd0d6ccfc90640ec8af49b99d4af14f2` |
| 天津 | `tianjin` | `f374ea9e12af484691ef888847fb6615` | `8dce98784326429e9c97515b634723fe` |
| 宁夏 | `ningxia` | `90ada2ec65a44e3797cbb858e62ba748` | `45da89683b9d4a4c9657ebc12da2a73c` |
| 宁波 | `ningbo` | `fd693136e4cd4936a2ddb2255e5c1257` | `8e2d5a7bf989489fb9caa99a39bfaa70` |
| 安徽 | `anhui` | `5efe954d7f8049128e04a5141fc23da6` | `83fb5344bde849ea8973f3715a680a15` |
| 山东 | `shandong` | `f8979e6c0b544e1992fea51d3c8739e3` | `3a1acab4996548faa7532b957ebf4f8e` |
| 山西 | `shanxi` | `7de444553c114e0f9dd0b5de7733cf5f` | `085ed9d1e1dc437cb9945e961fb2f5a0` |
| 广东 | `guangdong` | `29a1328dc824498b89eae8e63c38837f` | `02a93424320e46dea2631da827f96174` |
| 广西 | `guangxi` | `9b57abde8c1541239b97224e26ba1da3` | `f12807ebb0f948afaf7070627fadad7e` |
| 新疆 | `xinjiang` | `95e320eada3b4b2a8a0b319f991d392c` | `3d0077b1d74e4965ade6bdf997d3fce3` |
| 江苏 | `jiangsu` | `9d6dfc15aa6144c8aea00f9878bdb1ce` | `b335834574a34e43874ef3dba68d5be5` |
| 江西 | `jiangxi` | `9520f36bb312424bbeffb08c4542d83c` | `97691c60bc9a4b96af8b79df53b13b74` |
| 河北 | `hebei` | `4091867fdc664e169267aefc763c6d1d` | `e838879760e84c668062433e2cdbc389` |
| 河南 | `henan` | `92dfd6c412c2470bb93ed753c3404bc4` | `9724ed53f22a4333b91c65c349edaf48` |
| 浙江 | `zhejiang` | `83e130eb74f74d458c81027715955391` | `441d2ae10ef240cbbcfca758a73356f2` |
| 海南 | `hainan` | `c42cd21f50b84d16a8614fdf1cc3e478` | `6d27bb42929c46ae8487a14f41fab43b` |
| 深圳 | `shenzhen` | `58959eb1bd68458088cac63f46a5fa40` | `a3d847e5be4d40a5baaf387be4a56e9b` |
| 湖北 | `hubei` | `5a4e3e893a5f45d3a3f450cb8f96c635` | `f0a7388893fa4c15a2d6482510e2fb31` |
| 湖南 | `hunan` | `2f087f7010fb43788a9015f21448447a` | `9182876b7fe841d3a2990998f5b756f3` |
| 甘肃 | `gansu` | `ab8c66c4e8b24a169ad3b8f93d6d7fab` | `8e13b6a296324a60bc247a4a8a1df6a6` |
| 福建 | `fujian` | `e0d6aa56b3564735930b1cc078646d49` | `3df69151384a46cf8c8d11584cff5e94` |
| 西藏 | `xizang` | `5860856427aa48a98445caad530c723d` | `10f9fd2824c444ffafeaed5192682d55` |
| 贵州 | `guizhou` | `7f610f00f3b6435f9e7b730f125db5e6` | `8bcfadbce6b74f178c37e6eafa9438b0` |
| 辽宁 | `liaoning` | `58df751f822a42a0acb1abd7471c71f8` | `94cc260b27d2430b8c39c21e4627cf2e` |
| 重庆 | `chongqing` | `55dbc14f9bea476bb09743d5f1c8c842` | `febe5cf9074b4ce6a52fd3d34d7a5cba` |
| 陕西 | `shaanxi` | `f6c680055f9e43e7addf9493680a6112` | `f0dad1ecc157416ea412e4a0608e04bd` |
| 青岛 | `qingdao` | `453f88a0ab9642f88908e41ceb78003d` | `e3e00f2b72414f1c9ea951eeb010a6cf` |
| 青海 | `qinghai` | `989b26dfc95a4188b64b3f9d4895ddd9` | `439a663f89484376be17d4dcae953254` |
| 黑龙江 | `heilongjiang` | `e41f4e2fc9b34f5bab55489864313654` | `370916bda2524acdb187191fbba13f44` |

- **栏目 ID 自发现方法**（若上表失效或需新增派出机构）：
  1. 取 `https://www.csrc.gov.cn/<局拼音>/xhtml/js/render.js`，正则 `requestChildChannel\('(c\d+)'`
     与 `getChannelList\?channelCode=(c\d+)` 得到根栏目码；
  2. `https://www.csrc.gov.cn/getChannelList?channelCode=<码>` 返回的 `results` 是**单个对象**
     （不是数组），其 `children[]` 含 `channelName / channelCode / channelId`，递归 2–3 层；
  3. 匹配 `channelName` 含「行政监管措施」「行政处罚决定」的条目取 `channelId`。
- **私募筛选口径**：决定书正文（接口 `content` 字段直接返回全文）含「私募」即收录；
  再剔除标题主体为证券公司/期货公司/会计师事务所/上市公司的条目（其正文引述「私募资管」属误命中）。
- 措施类型：出具警示函、责令改正、监管谈话、责令增加内部合规检查次数、行政处罚、市场禁入等。
- 查重键：文号（`wh` 字段）或 机构名+决定日期。

## 三、交易所

### 上交所 https://www.sse.com.cn

> **2026-09-29 补全**：此前只跑 2 个栏目，且因只认 `.pdf` 漏掉大量 `.doc` 附件。见 SKILL.md 0.4.4。

- **6 个监管栏目全部要跑**（服务端渲染静态 HTML，直接解析）：

  | 栏目 | 路径 | 条数/页数 | 私募相关 |
  |---|---|---|---|
  | 会员监管-监管措施 | `/regulation/members/measures` | 44 / 3 | **主要来源**（IPO 网下询价违规） |
  | 会员监管-纪律处分 | `/regulation/members/disposition` | 11 / 1 | 全为券商/期货，最新 2019 年 |
  | 交易监管-纪律处分 | `/regulation/trading/disposition` | 56 / 4 | **限制交易纪律处分，私募产品在此** |
  | 发行上市审核-监管措施 | `/regulation/listing/measures` | 125 / 9 | 发行人/保代/会计师，无私募 |
  | 自律管理复核 | `/regulation/review` | 42 / 3 | 无私募 |
  | 自律监管与市场服务事项 | `/regulation/ma` | 8 / 1 | 无私募 |

- **分页**：页面内 `createPageHTML('paging',<页数>,<当前>,'s_list','shtml',<总数>)`，
  第 n 页 = `<栏目>/s_list_<n>.shtml`（**不是** `index_n.shtml`，后者 404）。
- **表格有两种结构，解析器必须都支持**：
  - `<tr><td>类型</td><td><a>标题</a></td><td>对象</td><td>日期</td>`（members/measures、listing/measures、review）
  - `<dl><dd><span>日期</span><a>标题</a></dd>`（trading/disposition）
- **附件扩展名不只是 `.pdf`**：限制交易纪律处分大量是 `.doc`。
  正则写 `href="(/regulation/[^"]+?\.(?:pdf|docx?|xlsx?))\s*"`（注意 `href` 值尾部可能有空格）。
  只认 `.pdf` 时交易监管-纪律处分只抓到 4/56 条。
- PDF 直链下载需带 `Referer: https://www.sse.com.cn/`。
- 部分决定书为**扫描件无文字层**（如农银汇理、工银安盛、长信基金那几份），按「一」的 OCR 方法处理；
  但这三家是公募/保险资管，按 SKILL.md 0.3 **不收录**。

### 深交所 https://www.szse.cn

> **2026-09-29 重大修正（此前覆盖不全 + 假阴性，务必照做，详见 SKILL.md 0.4）**
> 1. **附件真实域名是 `reportdocs.static.szse.cn`**，`www.szse.cn/UpFiles/...` 一律 **404**
>    （曾静默产出 322 个 0 字节文件 → 扫描"命中 0" → 漏收 2 条私募案例）。
>    点击列表「详细内容」后新页面 URL 即真实地址。换域名后**无需 Referer**。
> 2. **`1800_jgxxgk` 列表不返回被处罚主体名**（只有上市公司代码/简称/日期/措施/涉及对象），
>    **必须下载 PDF 抽正文**再用正则 `关于对(.{2,60}?)(?:予以|采取|给予|的)?(?:监管函|监管措施|…)`
>    抽主体名做私募判定。仅扫全文关键词会有大量假阳性（正文引用法规/他方主体）。
> 3. 栏目不止 `1800_jgxxgk` 一个，见下表。

**栏目与 CATALOGID**（取自栏目页 DOM 的 `data-s-catalog-id` 属性）：

> **2026-09-29 第二次修正：TABKEY 是数据集开关。** 同一 CATALOGID 下 tab1/tab2 是两套数据。
> `1800_jgxxgk_cf`（纪律处分）用 `TABKEY=tab1` 返回 `pagecount=0`、`data=[]`，
> 曾被误判为"无数据"；**换成 `TABKEY=tab2` 立刻返回 145 页 2894 条**。
> 每个 CATALOGID 必须显式声明 TABKEY 列表并逐个遍历。

| 栏目 | 路径 | CATALOGID | **TABKEY** | 字段（各不相同） | 窗口内私募实测 |
|---|---|---|---|---|---|
| 监管措施（上市公司） | `/disclosure/supervision/measure/measure/index.html` | `1800_jgxxgk` | tab1 | `fhrq`/`hjbt`/`gkxx_sjdx`/`gkxx_jgcs` | **有**（本轮 3 条） |
| 会员及其他交易参与人监管 | `/disclosure/supervision/members/index.html` | `members_supervision` | tab1 | `fhrq`/`hjbt` | 0（全为保代/券商） |
| 中介机构监管 | `/disclosure/supervision/supervise/index.html` | `ZJJGJG` | tab1 | `punishdate`/`title` | 0（全为会所/评估/券商） |
| **纪律处分（上市公司）** | 监管措施页 | `1800_jgxxgk_cf` | **tab2** | `xx_fwrq`/`xx_bt`/`ck`/`xx_cflb`（标题 `xx_bt` **直接含主体名**） | 窗口内 126 条，全为上市公司+董监高，**私募 0** |
| **债券监管措施** | 监管措施页 | `ZQ_JGCS` | tab1 | `fhrq`/`hjbt`/`dx`/`hh`/`lx`（对象 `dx`、文号 `hh`） | 窗口内 13 条，全为发行人（国资运营集团/AMC），**私募 0** |
| 债券纪律处分 | — | `ZQ_JGCL` | tab1/tab2 | — | 两种 TABKEY 均返回 0，视为无数据 |
| 问询函件（不纳入） | — | `main_wxhj` | tab1 | `gsdm`/`gsjc`/`fhrq`/`hjlb` | 370 页，属问询非处罚措施 |
| 限制交易与纪律处分 | `/disclosure/supervision/transaction/restrict/index.html` | —（**静态渲染在 div，非 table**） | — | — | 最新条目 2024-05-22，窗口内 0 条 |

> **⚠ 频控（2026-09-29）**：`www.szse.cn/api/...` 与 `reportdocs.static.szse.cn` **都有单 IP 频控**。
> 触发后表现为：附件全部 403（连刚才成功的 URL 也 403），继而列表接口返回 `pagecount=1 / data=[]`
> 或直接 `RemoteDisconnected`——**与"该栏目确实没有数据"无法从现象上区分**。
> 对策：并发 ≤3、请求间隔 0.8s、403 指数退避重试 4 次、**复用 `--cache-dir` 缓存**；
> 出现大面积失败时**立即停止**，用 1~2 个曾经成功的 URL 做探针，探针也失败即确认为限流，等待后重跑。
> 详见 SKILL.md 0.4.7。

- **列表接口（已验证可用，无需浏览器渲染）**：
  ```
  https://www.szse.cn/api/report/ShowReport/data?SHOWTYPE=JSON&CATALOGID=<上表id>
      &TABKEY=tab1&PAGENO={n}&random=0.5
  ```
  需带 `Referer`（对应栏目页 URL）。**返回顶层是 list**，`j[0]['metadata']['pagecount']`
  给总页数，`j[0]['data']` 是记录数组（不要按 dict 取 `j.get(...)`，会报
  `'list' object has no attribute 'get'`）。
  返回 `data[]`，字段：`gkxx_gsdm` 代码、`gkxx_gsjc` 简称、`gkxx_gdrq` 采取措施日期、
  `gkxx_jgcs` 措施、`hjnr`（内含 `encode-open='/UpFiles/....pdf'`）、`gkxx_sjdx` 涉及对象。
- 「限制交易与纪律处分」栏目用 `table tr` 只能抓到分页数字，
  必须取 `div/ul` 的 `innerText` 且含 `YYYY-MM-DD` 的元素。
- 注意：`/disclosure/supervision/measure/index.html` 是**无效路径**（返回 6 字节），
  正确路径是 `/disclosure/supervision/measure/measure/index.html`（监管措施）、
  `/disclosure/supervision/measure/decision/index.html`（复核决定）。
- **2026-09-29 修正后核查结论**（推翻 2026-09-28 的"命中 0"错误结论）：
  窗口内 204 条 `1800_jgxxgk` 监管措施，用 `reportdocs` 域名成功解析 268 条（含另两个栏目），
  **主体级私募命中 3 条**：
  - 杭州御硕创业投资合伙企业（有限合伙）2026-07-29 创业板监管函〔2026〕第89号 → **收录**
  - 广东汇垠澳丰股权投资基金管理有限公司 2025-11-28 公司部监管函〔2025〕第203号 → **收录（新补）**
  - 信阳市鼎新兴华/广兴股权 2026-04-21 公司部监管函〔2026〕第51号 → **剔除**（华英农业重整持股平台，非私募管理人）
  另有 16 条"仅正文含私募词"的弱命中，主体均为上市公司/自然人/中介机构，**一律不收录**。
  每次执行仍须跑完全部 CATALOGID 并把核查结论写进 `更新记录.md`；
  **若"成功解析数 / 待扫数 < 80%"，必须先修下载链路，禁止直接下"命中 0"的结论**。

### 北交所 https://www.bse.cn
- **bse.cn 有 JS 反爬**（返回一段设置 `C3VK` cookie 的脚本，urllib 直连会 302 死循环），
  必须用无头浏览器；PDF 附件也要用浏览器上下文下载（`ctx.request.get(url)` 共享 cookie），
  否则 urllib 直取 PDF 同样 302。
- 栏目（都要查）：
  - `https://www.bse.cn/disclosure/trade_supervise.html` — **交易监管，私募产品/私募机构案例在这里**
  - `https://www.bse.cn/disclosure/supervise.html` — 自律监管措施
  - `https://www.bse.cn/disclosure/disciplinary_aciton.html` — 纪律处分（注意拼写是 aciton）
  - `https://www.bse.cn/disclosure/securities_trader_supervise.html` — 会员及其他交易参与人监管
- 用 `eval_on_selector_all('table tr', ...)` 取行文本与 `a.href`；PDF 有文字层可直接 `fitz` 抽取。
- 收录判断：标题/正文含「私募证券投资基金」「私募基金管理有限公司」「资产管理有限公司」才收录；
  **仅自然人姓名的（如赵新、韩梅、李文浩、景娜娜）一般是个人投资者异常交易，不收录**。
- **2026-09-29 实测结论（覆盖完整，无缺口）**：
  - `trade_supervise.html` 共 10 条，窗口内 5 条：赵新/韩梅/陆建钢（2026-09-11 自然人）、
    **浙江浩坤昇发资产管理有限公司（2026-01-23，已入库）**、李文浩（2025-12-18 自然人）。
    窗口外 5 条含 2 条私募（华睿千和量化中性二号、中投风起扶摇五号，均 2025-03-17）
    —— 说明**窗口滚动后会有新私募条目，每轮必须重跑**。
  - `disciplinary_aciton.html` 19 条全部为上市公司（代码+简称），**私募 0**。
  - `securities_trader_supervise.html` 20 条，窗口内仅 1 条「情况公示」（2026-07-08）。
- **⚠ 北交所另有 WAF 层（2026-09-29）**：附件下载可能**全部返回 403**，
  响应体 2957 字节 HTML、`403 Forbidden`、图片来自 `blocksrc.haplat.net`。
  已验证 `ctx.request.get`、加 `Referer`/`User-Agent`、页内 `fetch()`、`page.goto()`
  **四路全 403** —— 是 **IP 级限流，不是 header 问题**，换 header 无效。
  对策：复用缓存 + 逐个请求 `sleep(1.5)` + 403 退避重试（3/8/13s）+ 落单的下一轮补。
  实测加节流后 **25/26 成功**（仅剩 1 条 2026-07-08 会员监管附件待补）。
  **绝不能把 403 造成的 0 条当成"北交所无私募案例"** —— 实际窗口内有浩坤昇发 1 条。
  - `supervise.html` 仅 1 条「关于对（拟）上市公司及相关责任主体采取自律监管措施的情况公示」。

## 四、通用检索技巧

- 决定书正文是分类与摘要的唯一依据，禁止凭标题猜测分类。
- 每条候选必须记录：决定日期、机构来源、措施类型、被处罚机构/人、文号、原文链接。
- 同一案例被多个层级处罚（先警示函后行政处罚）属不同条目，分别收录。
- 无头浏览器统一用本机 Chrome：
  `C:\Program Files (x86)\Google\Chrome\Application\chrome.exe`
  （playwright 未装自带内核时用 `p.chromium.launch(executable_path=CHROME)`）。


## 四·补、配套法规依据信源（06_以上各类所有案例的配套法规依据）

权威源优先级：**gov.cn（国务院/国办）＞ csrc.gov.cn（证监会）＞ neris.csrc.gov.cn（证监会法规库）＞ fg.amac.org.cn / amac.org.cn（中基协）＞ flk.npc.gov.cn（法律）**。禁止用第三方站点、百科、转载页充当来源。

### 1. 证监会法规库 NERIS（法律/规章/规范性文件全文，最通用）

- 法规正文页：`https://neris.csrc.gov.cn/falvfagui/rdqsHeader/mainbody?navbarId=<0法律|1|2>&secFutrsLawId=<32位id>`
- **正文由 JS 渲染**：直接 `requests` 抓 HTML 只有页面框架（约 7KB）。**必须用无头浏览器** `wait_until="networkidle"` + `sleep(2~3)` 后取 `body.innerText`。
- 页头要裁掉的行：站名/导航/「文件名称…」「效力状态…」「发文单位…」「发文日期…」「历次修订信息」「正 文」「显示注释」；页脚从 `联系我们 | 法律声明` / `主办单位：` / `网站识别码：` 起全部丢弃。**注意"相关案例"可能是正文内的行内标记（如「第五条 相关案例(1) 证券的…」），不能用"包含相关案例"就截断**，否则《证券法》只剩 492 字。
- 已验证可直接取到全文的 secFutrsLawId：

| 法规 | 文号 | secFutrsLawId |
|---|---|---|
| 中华人民共和国证券法 | 主席令第37号 | `0fc431a2a10b47909beef058f6ac3335` |
| 中华人民共和国证券投资基金法（2015修正） | 主席令第23号 | `e5cc920d0af7475ca44f1e69c5378244` |
| 证券经纪业务管理办法 | 证监会令第204号 | `c285dc2210a044e0a79b9dcbbfffb062` |
| 证券期货市场监督管理措施实施办法 | 证监会令第231号 | `b76e3a3118824358a6ffc8e56f25c168` |
| 证券期货经营机构私募资产管理业务运作管理暂行规定 | 证监会公告〔2016〕13号 | `abd0853fb3694e79ab4178c581ad76e7` |
| 证券经纪人管理暂行规定（2020年修订） | 证监会公告〔2020〕20号 | `78f1bd1210084994a9edca91d7820b84` |

- 坑：个别条目（如《私募投资基金管理人登记和基金备案办法（试行）》`c0a86116f402474e9113dfa13ea2a7c3`）**正文为空**，只有元信息 → 改从中基协法规库取。NERIS 的「效力状态」字段也可能滞后（已废止仍标现行有效），效力以发布机关公告为准。
- 找 id 的办法：用网络检索 `法规名 + neris 法规正文`（WebSearch 加 `allowed_domains: ["neris.csrc.gov.cn"]` 命中率最高）；站内综合查询页（`multipleFindController/indexJsp`）关键词检索经常返回 0 条，不可靠。

### 2. 证监会规章栏目（官方 PDF）

- 规章列表 channelId `d5c69b3f67334ed984a43e864f42f1ab`（89 部现行规章）：
  `https://www.csrc.gov.cn/searchList/<channelId>?_isAgg=true&_isJson=true&_pageSize=20&_template=index&_rangeTimeGte=&_channelName=&page=N`
- **坑：`_pageSize` 只能设 20；设 50 时第一页不足 50 条会让翻页逻辑误判结束，只拿到 20 条。**
- 正文页 `https://www.csrc.gov.cn/csrc/c106256/cN/content.shtml` 内的 `.pdf` 附件是相对路径，拼成 `正文页目录 + '/' + 附件名` 后**必须对中文文件名百分号编码**：`quote(path, safe="/%=&?~#")`。
  **不编码会拿到同一个网页打印件——多个文件都是约 216KB、MD5 相同，务必用正文首行或 MD5 校验。**
- 部分规章（如《证券经纪业务管理办法》证监会令第204号、《私募资产管理计划运作管理规定》公告〔2023〕2号）**不在该 89 部列表里**，需从 NERIS 或政府网/国务院公报取。

### 3. 中基协法规库（fg.amac.org.cn）

- 目录路径（前缀 `https://fg.amac.org.cn/governmentrules_3854/zcgz_zlgz/`）：
  - `zlgz_smjj/zlgz_smjj_cpl/` — 私募产品类自律规则
  - `zlgz_smjj/zlgz_smjj_glrdj/` — 管理人登记类（登记指引1/2/3号在 `202303/t20230327_24588|24589|24590.html`）
  - `zlgl/` — 自律管理类（自律管理和纪律处分措施实施办法等）
- 列表页为 Vue 渲染，`requests` 也能拿到 `./YYYYMM/tYYYYMMDD_NNNNN.html` 链接（正则 `href="([^"]*t\d{8}_\d+\.html)"`）；要完整清单用无头浏览器渲染更稳。
- **坑：公告类页面（如《关于发布〈私募基金管理人登记指引第1号〉的公告》）只有 1 页、正文只有公告本身，附件全文在法规库独立条目页**——入库前务必检查页数/字符数，只有「发布公告」页的不算数。

### 4. 法律（flk.npc.gov.cn / gov.cn）

- 人大法规库 `flk.npc.gov.cn` 的《证券投资基金法》PDF 为**扫描件、无文字层**，不要存；改用 NERIS 文字版。
- 国务院/国办文件（条例、指导意见、若干规定）用 gov.cn，`html.parser` 解析、正文取 `#UCAP-CONTENT`（**lxml 会把 TRS 正文解析丢成几十个字符**）。

## 五、覆盖度自检（每次执行结束前必须跑，结果写进 `更新记录.md`）

1. 中基协：两个栏目各自翻完 `createPageHTML` 给出的全部页，条数 - 台账中基协条数 = 0。
2. 证监局：36 家 × 2 栏目全部翻完，窗口内私募命中数 - 台账证监局条数 = 0。
3. 交易所：上交所静态页 + 深交所 `1800_jgxxgk` 接口 + 北交所 4 个栏目全部跑完，
   逐条正文关键词扫描，命中数 - 台账交易所条数 = 0。
4. 台账条数 = 各分类目录 PDF 实存数量（允许 `05_重点典型案例` 的额外副本）。
5. 任一项差值不为 0，必须在回复中明确写出缺口数量与未完成原因，**不得声称「已完成」**。
6. **法规依据联动（硬性）**：从全部决定书正文抽取 `《…》` 法规引用清单，与
   `06_以上各类所有案例的配套法规依据` 实际文件比对——**引用数 ≥3 的缺失必须为 0**；
   且每个法规文件都要有内容（PDF 有文字层、Word 段落数合理），
   发现无文字层扫描件或只有「发布公告」1 页的附件必须替换。
   只报案例缺口而漏报法规缺口 = 任务未完成。
