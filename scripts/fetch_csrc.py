# -*- coding: utf-8 -*-
"""证监会 36 家派出机构「行政监管措施 + 行政处罚决定」全量抓取（v2 固化，2026-10-26）。

用法：
    python fetch_csrc.py --start 2026-09-26 --end 2026-10-26 --out csrc.json \
        [--cache-dir .cache_csrc] [--bureaus guangdong,beijing] [--only-col 行政监管措施]
        [--attach-dir dl_csrc] [--text-dir texts_csrc]

实现要点（实测，详见 SKILL.md 0.4.43 / 本轮探测）：
  1. 列表：`https://www.csrc.gov.cn/searchList/{channelId}?_isAgg=true&_isJson=true&_pageSize=20&...&page=N`
     返回 data.total + data.results[]（title/url/publishedTimeStr/manuscriptId/domainMetaList[fwrq,wh]）。
     **每局每栏目 channelId 独立**，绝不共享（sources.md 已给全 36 局）。
  2. 列表接口 content/contentHtml **为空**——正文在附件里：
     取 manuscriptId → POST `/getManuscriptData`（mId=manuscriptId, status=4）
     → _source.resList[] 里找 .docx / .pdf 附件（filePath 拼 https://www.csrc.gov.cn）。
  3. 附件是 .docx（占多数）→ zipfile 读 word/document.xml 抽文本；是 .pdf → fitz 抽，
     无文字层再 R 通道 OCR（与 verify_library 同缓存键）。
  4. 落款/决定日期不在列表里，必须从正文抽；fwrq（发文日期）只在 domainMetaList，
     仅作发布口径，**决定日期一律从正文取**（SKILL 0.4.46）。
  5. 正文抽出来后渲染成 PDF（fitz insert_text）留存——库目录只放 PDF，docx 不入库。
  6. 私募筛选：正文含「私募」为强命中；其余按 SUBJ_KEY 弱命中，交由入库环节逐条通读定夺。

输出 JSON：{window, bureaus_checked, stats, items:[...]}
item = {bureau, bureau_cn, col, title, url, date_published, date_fwrq, wh,
        pdf(本地路径), text, hit_kind, attach_ext}
"""
import sys, os, re, json, ssl, time, io, zipfile, argparse, shutil
import urllib.request
import urllib.parse

sys.stdout.reconfigure(encoding="utf-8", line_buffering=True)

CTX = ssl.create_default_context(); CTX.check_hostname = False; CTX.verify_mode = ssl.CERT_NONE
UA = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
                    "(KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"}

# (pinyin, 中文名, 行政监管措施 channelId, 行政处罚决定 channelId)
BUREAUS = [
    ("shanghai", "上海", "9ebc4198232e496e8bebf1b1bb1778ef", "c8318fc200764e38b30116c2d5f72b4b"),
    ("yunnan", "云南", "ebb6ce034bb84594a0f0c06fccecde9b", "99481e931fc94d13b8592449acb386a4"),
    ("neimenggu", "内蒙古", "4c96602a020c4398adfc7123a48f2ee1", "e4f2dbcdd85c4bac9a9e9a6f2d093d19"),
    ("beijing", "北京", "ea82afe8502d46aebea7346f0729dd85", "53be9cf2273744cda5c5780e0dded972"),
    ("xiamen", "厦门", "e3feac23205440e5b10ed81efa9dd94b", "0c1b68108fca4de0b240caaa32f901ed"),
    ("jilin", "吉林", "366f3072cbe143bebd533ef37f04eff5", "9a3520fd5ea644c2ad6e05acd0a05401"),
    ("sichuan", "四川", "c8ec4969173148c19b6d52afc6500b49", "6cd6f0ccbd2f49c6af32878d34c54ae6"),
    ("dalian", "大连", "fdf5c88757e04212a6dd66f396b21168", "fd0d6ccfc90640ec8af49b99d4af14f2"),
    ("tianjin", "天津", "f374ea9e12af484691ef888847fb6615", "8dce98784326429e9c97515b634723fe"),
    ("ningxia", "宁夏", "90ada2ec65a44e3797cbb858e62ba748", "45da89683b9d4a4c9657ebc12da2a73c"),
    ("ningbo", "宁波", "fd693136e4cd4936a2ddb2255e5c1257", "8e2d5a7bf989489fb9caa99a39bfaa70"),
    ("anhui", "安徽", "5efe954d7f8049128e04a5141fc23da6", "83fb5344bde849ea8973f3715a680a15"),
    ("shandong", "山东", "f8979e6c0b544e1992fea51d3c8739e3", "3a1acab4996548faa7532b957ebf4f8e"),
    ("shanxi", "山西", "7de444553c114e0f9dd0b5de7733cf5f", "085ed9d1e1dc437cb9945e961fb2f5a0"),
    ("guangdong", "广东", "29a1328dc824498b89eae8e63c38837f", "02a93424320e46dea2631da827f96174"),
    ("guangxi", "广西", "9b57abde8c1541239b97224e26ba1da3", "f12807ebb0f948afaf7070627fadad7e"),
    ("xinjiang", "新疆", "95e320eada3b4b2a8a0b319f991d392c", "3d0077b1d74e4965ade6bdf997d3fce3"),
    ("jiangsu", "江苏", "9d6dfc15aa6144c8aea00f9878bdb1ce", "b335834574a34e43874ef3dba68d5be5"),
    ("jiangxi", "江西", "9520f36bb312424bbeffb08c4542d83c", "97691c60bc9a4b96af8b79df53b13b74"),
    ("hebei", "河北", "4091867fdc664e169267aefc763c6d1d", "e838879760e84c668062433e2cdbc389"),
    ("henan", "河南", "92dfd6c412c2470bb93ed753c3404bc4", "9724ed53f22a4333b91c65c349edaf48"),
    ("zhejiang", "浙江", "83e130eb74f74d458c81027715955391", "441d2ae10ef240cbbcfca758a73356f2"),
    ("hainan", "海南", "c42cd21f50b84d16a8614fdf1cc3e478", "6d27bb42929c46ae8487a14f41fab43b"),
    ("shenzhen", "深圳", "58959eb1bd68458088cac63f46a5fa40", "a3d847e5be4d40a5baaf387be4a56e9b"),
    ("hubei", "湖北", "5a4e3e893a5f45d3a3f450cb8f96c635", "f0a7388893fa4c15a2d6482510e2fb31"),
    ("hunan", "湖南", "2f087f7010fb43788a9015f21448447a", "9182876b7fe841d3a2990998f5b756f3"),
    ("gansu", "甘肃", "ab8c66c4e8b24a169ad3b8f93d6d7fab", "8e13b6a296324a60bc247a4a8a1df6a6"),
    ("fujian", "福建", "e0d6aa56b3564735930b1cc078646d49", "3df69151384a46cf8c8d11584cff5e94"),
    ("xizang", "西藏", "5860856427aa48a98445caad530c723d", "10f9fd2824c444ffafeaed5192682d55"),
    ("guizhou", "贵州", "7f610f00f3b6435f9e7b730f125db5e6", "8bcfadbce6b74f178c37e6eafa9438b0"),
    ("liaoning", "辽宁", "58df751f822a42a0acb1abd7471c71f8", "94cc260b27d2430b8c39c21e4627cf2e"),
    ("chongqing", "重庆", "55dbc14f9bea476bb09743d5f1c8c842", "febe5cf9074b4ce6a52fd3d34d7a5cba"),
    ("shaanxi", "陕西", "f6c680055f9e43e7addf9493680a6112", "f0dad1ecc157416ea412e4a0608e04bd"),
    ("qingdao", "青岛", "453f88a0ab9642f88908e41ceb78003d", "e3e00f2b72414f1c9ea951eeb010a6cf"),
    ("qinghai", "青海", "989b26dfc95a4188b64b3f9d4895ddd9", "439a663f89484376be17d4dcae953254"),
    ("heilongjiang", "黑龙江", "e41f4e2fc9b34f5bab55489864313654", "370916bda2524acdb187191fbba13f44"),
]
SUBJ_KEY = ["私募基金", "私募证券", "私募股权", "创业投资", "股权投资", "投资管理",
            "资产管理", "投资基金", "资本管理", "合伙企业"]


# ---------------------------------------------------------- HTTP
def http_get(url, binary=False, timeout=40, retry=3):
    for i in range(retry):
        try:
            b = urllib.request.urlopen(urllib.request.Request(url, headers=UA),
                                       timeout=timeout, context=CTX).read()
            return b if binary else b.decode("utf-8", "ignore")
        except Exception:
            time.sleep(1.5 * (i + 1))
    return None


def http_post(url, data, timeout=40, retry=4):
    """POST 并重试。2026-10-28 修：原实现只在抛异常时重试，
    但证监局内容接口限流时会返回 **200 + 空响应**，被当成"这条没有正文"→ 条目被丢弃。
    实测：广东局 155 条里 68 条这样丢掉（同一 mId 过几分钟单独请求又正常）。
    现在把「空响应 / 非 JSON / 无 results」也纳入重试，退避 2/4/8/16 秒并加抖动。"""
    import random
    body = urllib.parse.urlencode(data).encode()
    for i in range(retry):
        try:
            req = urllib.request.Request(url, data=body,
                                         headers={**UA, "Content-Type": "application/x-www-form-urlencoded"})
            raw = urllib.request.urlopen(req, timeout=timeout, context=CTX).read().decode("utf-8", "ignore")
            if raw and raw.strip() and raw.lstrip()[:1] in "{[":
                return raw
        except Exception:
            pass
        if i < retry - 1:
            time.sleep((2 ** (i + 1)) + random.random())
    return None


def http_get_json(url, timeout=40, retry=3):
    """取 JSON 并就地解析重试。返回 (obj|None, 失败现场前 200 字)。

    2026-10-28 实测：新疆/江苏/江西 6 个栏目在同一轮里齐刷刷 terminated_by=json_error、
    platform_total=0，而同一 URL 几分钟后单独请求完全正常（新疆 251/65、江苏 1512/140）。
    证监局搜索接口偶发返回**被截断的 JSON** 或 **200 + HTML 反爬页**，
    旧实现一次 json.loads 失败就 break → 整个栏目归零 → 断言只能看到"平台总 0"，
    既无法区分"真没数据"也无从查因。改为就地退避重试，并把失败现场记进 meta。"""
    import random
    head = ""
    for i in range(retry):
        raw = http_get(url, timeout=timeout, retry=2)
        if raw:
            try:
                return json.loads(raw), ""
            except Exception:
                head = raw[:200].replace("\n", " ")
        if i < retry - 1:
            time.sleep(2.0 * (i + 1) + random.random())
    return None, head


# ---------------------------------------------------------- 列表分页
def fetch_list(channel_id, start, end):
    out, page = [], 1
    # 2026-10-28：为"抓取完整性自证"记录三个事实（不做判断，判断交给 assert_crawl_complete.py）：
    #   platform_total 平台声明的总条数；pages_fetched 实际翻了第几页；
    #   terminated_by 终止原因（正常=start_boundary/short_page，异常=http_fail/json_error/empty_results/max_pages）。
    platform_total, term = 0, "unknown"
    term_detail = ""
    latest_seen = ""          # 平台最新一条的挂网日（不限窗口），用于判定"窗口内 0 条"是否正常
    while True:
        u = ("https://www.csrc.gov.cn/searchList/%s?_isAgg=true&_isJson=true&_pageSize=20"
             "&_template=index&_rangeTimeGte=&_channelName=&page=%d" % (channel_id, page))
        j, head = http_get_json(u)
        if j is None:
            term = "json_error" if head else "http_fail"
            term_detail = head
            break
        d = j.get("data") or {}
        if not platform_total:
            try:
                platform_total = int(d.get("total") or 0)
            except Exception:
                platform_total = 0
        res = d.get("results") or []
        if not res:
            term = "empty_results"
            break
        stop = False
        for r in res:
            pub = (r.get("publishedTimeStr") or "")[:10]
            if pub and pub > latest_seen:
                latest_seen = pub
            if pub < start:                      # 列表按日期倒序，越过起点即停
                stop = True
                continue
            if pub > end:
                continue
            fwrq = ""
            for dm in (r.get("domainMetaList") or []):
                for it in (dm.get("resultList") or []):
                    if it.get("key") == "fwrq":
                        fwrq = (it.get("value") or "")[:10]
            wh = ""
            for dm in (r.get("domainMetaList") or []):
                for it in (dm.get("resultList") or []):
                    if it.get("key") == "wh":
                        wh = it.get("value") or ""
            url = r.get("url") or ""
            if url.startswith("//"):
                url = "https:" + url
            out.append({"title": r.get("title", ""),
                        "url": url,
                        "manuscriptId": str(r.get("manuscriptId") or ""),
                        "date_published": pub,
                        "date_fwrq": fwrq,
                        "wh": wh})
        if stop:
            term = "start_boundary"          # 正常：列表按日期倒序，已越过窗口起点
            break
        if len(res) < 20:
            term = "short_page"              # 正常：最后一页不足 20 条
            break
        page += 1
    return out, {"platform_total": platform_total, "pages_fetched": page,
                 "terminated_by": term, "latest_seen": latest_seen,
                 "terminated_detail": term_detail}


# ---------------------------------------------------------- 正文（附件 docx/pdf 或内联 HTML）
def _strip_html(html):
    t = re.sub(r"<[^>]+>", " ", html or "")
    t = re.sub(r"&#(\d+);", lambda m: chr(int(m.group(1))), t)
    return re.sub(r"\s+", " ", t).strip()


def get_content(manuscript_id, page_url=""):
    """返回 (kind, payload, ext)：
       kind='attach' → payload=附件URL，ext=扩展名（后续下载抽文本）；
       kind='inline' → payload=正文纯文本（来自 contentHtml，无附件）；
       否则 (None,None,None)。
    2026-10-26 实测：证监局约 9 成决定书是内联 HTML（contentHtml 有全文），
       仅少数是 .docx/.pdf 附件；原实现只找附件导致 170+ 条「无附件」漏收（含西安财帮基金等私募）。
    2026-10-28 补：接口返回空时**回落到公开页面解析**（page_url 即列表给的 content.shtml），
       从页面里找 .doc/.docx/.pdf 附件链接——限流只影响 JSON 接口，静态页通常仍可读。
    """
    raw = http_post("https://www.csrc.gov.cn/getManuscriptData",
                    {"mId": manuscript_id, "status": "4"})
    got = None
    if raw:
        try:
            j = json.loads(raw)
        except Exception:
            j = {}
        arr = (j.get("results", {}).get("data", {}).get("data")) or []
        for e in arr:
            s = e.get("_source", {})
            if str(s.get("isCarbonCopy")) == "0":
                for att in (s.get("resList") or []):
                    fp = att.get("filePath") or ""
                    if not fp:
                        continue
                    if fp.lower().endswith((".docx", ".doc", ".pdf", ".wps")):
                        url = "https://www.csrc.gov.cn" + fp if fp.startswith("/") else fp
                        return "attach", url, fp.rsplit(".", 1)[-1].lower()
                html = s.get("contentHtml") or s.get("content") or ""
                txt = _strip_html(html)
                if len(txt) > 100:
                    return "inline", txt, "html"
        got = arr
    # 回落：直接抓公开页面（接口限流 / 结构变化时的兜底）
    if page_url:
        page = http_get(page_url)
        if page:
            m = re.findall(r'href="([^"]+\.(?:docx?|pdf|wps))"', page, re.I)
            if not m:
                m = re.findall(r'href="([^"]+)"[^>]*>[^<]*\.(?:docx?|pdf|wps)', page, re.I)
            if m:
                u = m[0]
                u = urllib.parse.urljoin(page_url, u)
                return "attach", u, u.rsplit(".", 1)[-1].lower()
            txt = _strip_html(page)
            # 页面兜底取正文时要求出现文书特征词，避免把导航框架当成决定书正文
            if len(txt) > 300 and ("决定" in txt or "监管" in txt or "当事人" in txt):
                return "inline", txt, "html"
    return None, None, None


_OCR = None
OCR_MAX_PX = 2000


def _ocr_pdf(path, dpi=200):
    global _OCR
    import numpy as np
    import fitz
    if _OCR is None:
        from rapidocr_onnxruntime import RapidOCR
        _OCR = RapidOCR()
    out = []
    d = fitz.open(path)
    try:
        for pg in d:
            zoom = dpi / 72.0
            if pg.rect.width * zoom > OCR_MAX_PX:
                zoom *= OCR_MAX_PX / (pg.rect.width * zoom)
            pix = pg.get_pixmap(matrix=fitz.Matrix(zoom, zoom))
            img = np.frombuffer(pix.samples, dtype=np.uint8).reshape(pix.height, pix.width, pix.n)
            if img.ndim == 3 and img.shape[2] >= 3:
                R = np.ascontiguousarray(img[:, :, 0])
                img = np.stack([R, R, R], axis=2)
            r, _ = _OCR(img)
            if r:
                out.append(" ".join(t[1] for t in r))
    finally:
        d.close()
    return "\n".join(out)


def docx_text(b):
    try:
        z = zipfile.ZipFile(io.BytesIO(b))
        xml = z.read("word/document.xml").decode("utf-8", "ignore")
    except Exception:
        return ""
    # 2026-10-28 修：原实现把所有标签替换成空格，Word 会把一句话拆成多个 <w:t> run，
    # 于是「〔2026〕53号」变成「〔20 26 〕 53 号」、「2026年5月26日」变成「2026 年5 月 26 日」，
    # 落款日期与条款号抽取全被污染。改法：段落/表格行结束符→换行，其余标签**直接删除**。
    xml = re.sub(r"</w:p>|</w:tr>|</w:tc>", "\n", xml)
    xml = re.sub(r"<w:tab/>|<w:br/>|<w:cr/>", " ", xml)
    txt = re.sub(r"<[^>]+>", "", xml)
    txt = re.sub(r"&#(\d+);", lambda m: chr(int(m.group(1))), txt)
    txt = (txt.replace("&amp;", "&").replace("&lt;", "<").replace("&gt;", ">")
              .replace("&quot;", '"').replace("&apos;", "'").replace("&nbsp;", " "))
    lines = [re.sub(r"[ \t\u3000]+", " ", ln).strip() for ln in txt.split("\n")]
    return "\n".join(ln for ln in lines if ln)


def make_pdf_from_text(text, out_path, title="", meta_lines=None):
    """把纯文本渲染成 PDF 留存（**中文必须可检索**）。

    三代实现的教训（2026-10-27 回写）：
      ① 最早用 fitz.insert_text(fontfile=...) → 中文取回乱码，verify B7 判「空壳」。
      ② 改用 reportlab + STSong-Light CID → 抽取正常，**但 reportlab 不是默认依赖**；
         换机器后 ModuleNotFoundError 被调用方的 except 吞掉
         ⇒ **证监局决定书 PDF 一份都没生成**，而抓取统计仍然显示"成功"，静默失败。
      ③ 现在：优先 reportlab（可用时），**否则用 cjk_pdf（pymupdf + 系统 CJK 字体）**，
         pymupdf 是本技能硬依赖，故这条路在任何机器上都能走通。
         两条路都保证有完整文字层；都失败则抛异常，绝不静默产出空/乱码 PDF。
    """
    try:
        from reportlab.pdfbase import pdfmetrics
        from reportlab.pdfbase.cidfonts import UnicodeCIDFont
        from reportlab.pdfgen import canvas
    except ImportError:
        import cjk_pdf
        return cjk_pdf.text_to_pdf(text, out_path, title=title, meta_lines=meta_lines)

    try:
        pdfmetrics.registerFont(UnicodeCIDFont('STSong-Light'))
        font = 'STSong-Light'
    except Exception:
        raise RuntimeError("reportlab 无法注册 CJK 字体；请改用 cjk_pdf 路径")
    os.makedirs(os.path.dirname(os.path.abspath(out_path)) or ".", exist_ok=True)
    c = canvas.Canvas(out_path)
    c.setFont(font, 11)
    x0, y0, lh, cpl = 50, 800, 16, 36   # 每行约 36 个 CJK 字宽
    if title:
        c.setFont(font, 15)
        for i, ln in enumerate(_wrap_cjk(title, cpl - 4)):
            c.drawString(x0, y0 + 40 - i * 20, ln)
        c.setFont(font, 11)
        y0 -= 20
    lines = []
    # 2026-10-28 修（真实缺陷）：原实现只在「遇到标点且当前行已≥32字」时才断行，
    # 中文决定书常有整段无标点的长句 → 一行被拉成 60+ 字，drawString 画出页面被裁掉，
    # 实测温润决定书 PDF 只剩 92% 正文（丢 109 字），事由/依据抽取跟着缺条。
    # 改为**按宽度硬换行**，优先在行末附近的中文标点处断，保证任何情况下不超宽。
    def _hard(s, w=36):
        out = []
        for para in (s or "").split("\n"):
            para = para.strip()
            while len(para) > w:
                cut = max(para.rfind(ch, 0, w) for ch in "。；，、,;：:）)】")
                if cut < w // 2:
                    cut = w - 1
                out.append(para[:cut + 1])
                para = para[cut + 1:].lstrip()
            if para:
                out.append(para)
        return out

    lines = _hard(text)
    y = y0
    for ln in lines:
        if y < 60:
            c.showPage(); c.setFont(font, 11); y = y0
        c.drawString(x0, y, ln); y -= lh
    c.save()
    return out_path


def _wrap_cjk(s, cpl):
    lines, cur = [], ""
    for ch in s:
        cur += ch
        if len(cur) >= cpl:
            lines.append(cur); cur = ""
    if cur:
        lines.append(cur)
    return lines


def save_attach(url, ext, referer, attach_dir, ocr_cache):
    os.makedirs(attach_dir, exist_ok=True)
    fn = os.path.join(attach_dir, hashlib_sha1(url) + "." + ext)
    if os.path.exists(fn) and os.path.getsize(fn) > 1024:
        b = open(fn, "rb").read()
    else:
        # 2026-10-28 修：广东局附件 URL 里直接带中文文件名（…/files/关于对XX…docx），
        # urllib 传非 ASCII 路径会直接失败，表现为「附件解析为空」→ 条目被丢弃（34 条）。
        # 请求前必须做百分号编码；哈希与落盘名仍用原始 URL，保证幂等。
        uq = urllib.parse.quote(url, safe=":/?&=%#+,;@!$'()*[]~")
        b = http_get(uq, binary=True, retry=4)
        if not b or len(b) < 500:
            return None, ""
        open(fn, "wb").write(b)
    if ext.startswith(("docx", "doc", "wps")):
        text = docx_text(b)
        # 2026-10-28 修：旧二进制 .doc 用 zipfile 解不开（BadZipFile→""），
        # 69 条广东局决定书因此被判「附件解析为空」丢弃，其中 9 条是私募案（真实漏收）。
        # 走 Word COM 转换；没有 Word/win32com 时退化为二进制 UTF-16 抓串，再不济交 R8 阻断。
        if len(re.sub(r"\s", "", text or "")) < 80:
            t2 = doc_text(fn)
            if len(re.sub(r"\s", "", t2 or "")) > len(re.sub(r"\s", "", text or "")):
                text = t2
        return fn, text
    # pdf
    import fitz
    try:
        d = fitz.open(fn)
        t = "".join(p.get_text() for p in d)
        d.close()
    except Exception:
        t = ""
    if len(re.sub(r"\s", "", t)) < 200:
        t = _ocr_pdf(fn)
    return fn, t


_WORD = None
_WORD_TRIED = False


def _word_app():
    """复用同一个 Word 实例（每条新开一次要 1~2 秒，69 条就多花两分钟）。"""
    global _WORD, _WORD_TRIED
    if _WORD_TRIED:
        return _WORD
    _WORD_TRIED = True
    try:
        import win32com.client as win32
        app = win32.DispatchEx("Word.Application")
        app.Visible = False
        app.DisplayAlerts = 0
        _WORD = app
        import atexit
        atexit.register(lambda: _quit_word(app))
    except Exception:
        _WORD = None
    return _WORD


def _quit_word(app):
    try:
        app.Quit()
    except Exception:
        pass


def doc_to_pdf(src, out_pdf):
    """把 .docx/.doc/.wps 附件用 Word 导出为 PDF（**保留官方版式**）。

    2026-10-30 新增（严重缺陷修复）：证监局约 1 成决定书正文不在网页内联 HTML 里，
    而在网页挂的 .docx/.doc/.wps 附件里；网页本身只有「索引号/发文日期/名称/文号」
    元信息卡片 + 一个附件链接。原实现无差别 `print_to_pdf(网页)`，产出「只有导航框架、
    正文零内容的空壳 PDF」——实测 69 份，逐份比对正文确实全丢（旧版归档同批案例是
    「附件转 PDF 纯正文」，两页、文字层近千字）。
    正确形态 = 附件本体转 PDF：正文、落款、文号与官方原件完全一致。
    返回 (ok, how)：how ∈ {"word", "text"}，失败抛异常由调用方兜底。
    """
    app = _word_app()
    if app is not None:
        try:
            d = app.Documents.Open(os.path.abspath(src), ReadOnly=True,
                                   AddToRecentFiles=False, Visible=False)
            try:
                d.ExportAsFixedFormat(os.path.abspath(out_pdf), 17)  # 17 = wdExportFormatPDF
            finally:
                d.Close(False)
            if os.path.exists(out_pdf) and os.path.getsize(out_pdf) > 1024:
                return True, "word"
        except Exception:
            pass
    # Word 不可用（未装 Office / 无 pywin32）→ 文本重建兜底，保证正文不丢
    t = docx_text(open(src, "rb").read()) if src.lower().endswith(("docx", "wps")) else ""
    if len(re.sub(r"\s", "", t or "")) < 80:
        t = doc_text(src)
    if len(re.sub(r"\s", "", t or "")) < 80:
        raise RuntimeError("附件正文抽取为空（Word 与文本兜底均失败）")
    make_pdf_from_text(t, out_pdf)
    return True, "text"


def doc_text(path):
    """旧版 .doc / .wps 取正文：Word COM 优先，失败退化为二进制串抓取。"""
    app = _word_app()
    if app is not None:
        try:
            d = app.Documents.Open(os.path.abspath(path), ReadOnly=True,
                                   AddToRecentFiles=False, Visible=False)
            try:
                t = d.Content.Text
            finally:
                d.Close(False)
            if t and len(re.sub(r"\s", "", t)) > 50:
                return t
        except Exception:
            pass
    try:
        s = open(path, "rb").read().decode("utf-16-le", "ignore")
        segs = re.findall(r"[\u4e00-\u9fa5A-Za-z0-9，。、；：（）《》〈〉〔〕【】%\-—·\s]{6,}", s)
        return "\n".join(x.strip() for x in segs if len(x.strip()) >= 6)
    except Exception:
        return ""


def hashlib_sha1(s):
    import hashlib
    return hashlib.sha1(s.encode()).hexdigest()[:20]


BODY_MARK = ("经查", "经查明", "当事人", "违反", "违规事实", "我局决定",
             "申请人", "本会", "复核")


def pdf_has_body(path):
    """落盘前自检：PDF 里必须真有决定书正文（防「空壳 PDF」静默入库）。

    2026-10-30 教训：附件型决定书被 print_to_pdf(网页) 打成了只有导航框架的空壳，
    抓取统计照样 +1、verify 才在事后报 B7。这里把断言前移，宁可当场报错也不许静默。
    """
    try:
        import fitz
        d = fitz.open(path)
        t = "".join(p.get_text() for p in d)
        n = d.page_count
        d.close()
    except Exception:
        return False
    if n < 1 or os.path.getsize(path) < 3000:
        return False
    body = re.sub(r"\s+", "", t)
    # 无文字层（扫描件）不在此判空——那是图片型决定书，正文靠 OCR 复核
    if len(body) < 60:
        return True
    return any(k in body for k in BODY_MARK)


# ---------------------------------------------------------- main
def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--start", required=True)
    ap.add_argument("--end", required=True)
    ap.add_argument("--out", required=True)
    ap.add_argument("--cache-dir", default="./.cache_csrc")
    ap.add_argument("--attach-dir", default="./dl_csrc")
    ap.add_argument("--text-dir", default="./texts_csrc")
    ap.add_argument("--bureaus", default="", help="逗号分隔拼音过滤，默认全部 36 家")
    ap.add_argument("--only-col", default="", help="行政监管措施 / 行政处罚决定 / 空=两者")
    a = ap.parse_args()

    sel = set(x.strip() for x in a.bureaus.split(",") if x.strip())
    items, stats = [], {"bureaus": len(BUREAUS), "list_items": 0, "window_items": 0,
                        "ok": 0, "err": 0, "strong": 0, "weak": 0}
    for pinyin, cn, cid_meas, cid_pen in BUREAUS:
        if sel and pinyin not in sel:
            continue
        cols = [("行政监管措施", cid_meas), ("行政处罚决定", cid_pen)]
        for col, cid in cols:
            if a.only_col and col != a.only_col:
                continue
            lst, meta = fetch_list(cid, a.start, a.end)
            stats["list_items"] += len(lst)
            stats["window_items"] += len(lst)
            stats.setdefault("channels", []).append(
                {"bureau": pinyin, "bureau_cn": cn, "col": col, "channel_id": cid,
                 "items_window": len(lst), **meta})
            print("  [%s·%s] 窗口内 %d 条（平台总 %d，翻 %d 页，终止：%s）%s"
                  % (cn, col, len(lst), meta["platform_total"],
                     meta["pages_fetched"], meta["terminated_by"],
                     ("  现场：%s" % meta["terminated_detail"][:80]) if meta.get("terminated_detail") else ""))
            for r in lst:
                referer = r["url"] or "https://www.csrc.gov.cn/%s/" % pinyin
                # 2026-10-28 修：正文抓取失败必须"重试 + 留痕"，绝不能一行 WARN 后静默丢弃。
                # 实测事故：广东局 9 条私募决定书因正文接口连续失败被 continue 丢掉，
                # 库里凭空少了 9 个案例，而控制台只有 9 行 WARN、退出码仍是 0。
                kind = payload = ext = None
                for attempt in (1, 2):
                    kind, payload, ext = get_content(r["manuscriptId"], r.get("url") or "")
                    if kind:
                        break
                    if attempt < 2:
                        time.sleep(2.5 * attempt)
                # 节流：155 条连发会被内容接口限流（2026-10-28 实测 44% 丢条），
                # 每条之间留一点间隔，比事后补抓便宜得多。
                time.sleep(0.25)
                if not kind:
                    stats["err"] += 1
                    stats.setdefault("failed_items", []).append(
                        {"bureau": pinyin, "bureau_cn": cn, "col": col, "reason": "正文接口无返回",
                         "title": r["title"], "url": r["url"],
                         "date_published": r.get("date_published", "")})
                    print("    [DROPPED] 无正文（已重试 3 次，已记入 failed_items）：%s"
                          % r["title"][:40], file=sys.stderr)
                    continue
                if kind == "attach":
                    fn, text = save_attach(payload, ext, referer, a.attach_dir, a.cache_dir)
                    if not text:
                        stats["err"] += 1
                        stats.setdefault("failed_items", []).append(
                            {"bureau": pinyin, "bureau_cn": cn, "col": col,
                             "reason": "附件解析为空", "title": r["title"], "url": r["url"],
                             "date_published": r.get("date_published", "")})
                        print("    [DROPPED] 正文空（已记入 failed_items）：%s"
                              % r["title"][:40], file=sys.stderr)
                        continue
                else:  # inline HTML
                    fn, text = None, payload
                # 落库 PDF（库目录只放 PDF）
                pdf_name = os.path.join(a.attach_dir, hashlib_sha1(r["url"]) + ".pdf")
                pdf_how = "web"
                try:
                    # 2026-10-30 修复（严重，实测 69 份正文全丢）：
                    #   kind=="attach" 时正文在 **附件** 里，网页只有「索引号/发文日期/
                    #   名称/文号」元信息卡片 + 一个附件链接；此时再 print_to_pdf(网页)
                    #   得到的必然是「只有导航框架、正文零内容」的空壳 PDF。
                    #   正确形态（与旧格式归档同批案例一致）：附件本体转 PDF 留存——
                    #     · .pdf 附件 → 直接取附件本体；
                    #     · .docx/.doc/.wps → Word 导出 PDF（保留官方版式）。
                    if kind == "attach":
                        if ext == "pdf" and fn and fn.lower().endswith(".pdf"):
                            shutil.copy2(fn, pdf_name)
                            pdf_how = "attach-pdf"
                        else:
                            _ok, pdf_how = doc_to_pdf(fn, pdf_name)
                        if not (os.path.exists(pdf_name) and os.path.getsize(pdf_name) > 1024):
                            raise RuntimeError("附件转 PDF 未产出有效文件")
                        if not pdf_has_body(pdf_name):
                            raise RuntimeError("附件转 PDF 后仍无决定书正文")
                    else:
                        # 内联 HTML 正文：用无头浏览器把官网原页面打印为 PDF
                        # （2026-10-29 修复：不做"文本重建"——重建会混入「原文链接：」与
                        #  &ensp; 等 HTML 实体垃圾、丢失官方版式，且常被误认为 OCR 乱码）。
                        import importlib.util as _ilu
                        _sd = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                                           "save_decision_pdf.py")
                        _spec = _ilu.spec_from_file_location("save_decision_pdf", _sd)
                        _sdm = _ilu.module_from_spec(_spec); _spec.loader.exec_module(_sdm)
                        _br = _sdm.find_browser()
                        try:
                            _sdm.print_to_pdf(_br, r["url"], pdf_name)
                        except SystemExit:
                            raise RuntimeError("html2pdf 进程退出（无输出）")
                        if not (os.path.exists(pdf_name) and os.path.getsize(pdf_name) > 1024):
                            raise RuntimeError("html2pdf 未产出有效文件")
                        if not pdf_has_body(pdf_name):
                            raise RuntimeError("网页打印 PDF 无决定书正文（疑似正文在附件里）")
                except Exception as e:
                    # 兜底：附件/浏览器都不可用（如无 Office、无限流）时退回文本重建，
                    # 保证决定书正文不丢失；但产物无官方版式，需在 collect 后跑
                    # recover_originals.py / repair_attach_pdf.py 修正。
                    stats.setdefault("pdf_fallback", 0); stats["pdf_fallback"] += 1
                    print("    [WARN] 正文 PDF 生成失败，退回文本重建（后续需 recover）：%s（%s）"
                          % (r["title"][:40], e), file=sys.stderr)
                    try:
                        make_pdf_from_text(text, pdf_name, title=r["title"][:80],
                                           meta_lines=["原文链接：" + r["url"]])
                        pdf_how = "text"
                    except Exception as e2:
                        # 不许静默（2026-10-27）：PDF 生成失败曾被调用方 except 吞掉，
                        # 结果整轮"抓取成功"但决定书零留存。这里必须计数 + 打印。
                        stats["pdf_err"] = stats.get("pdf_err", 0) + 1
                        print("    [ERROR] PDF 生成失败，该条将无决定书留存：%s（%s）"
                              % (r["title"][:40], e2), file=sys.stderr)
                        pdf_name = fn if (fn and fn.lower().endswith(".pdf")) else ""
                        pdf_how = "none"
                hit = "strong" if "私募" in text else (
                    "weak" if [k for k in SUBJ_KEY if k in text[:800]] else "")
                if hit == "strong":
                    stats["strong"] += 1
                elif hit == "weak":
                    stats["weak"] += 1
                stats["ok"] += 1
                stats["pdf_" + pdf_how] = stats.get("pdf_" + pdf_how, 0) + 1
                items.append({"bureau": pinyin, "bureau_cn": cn, "col": col,
                              "title": r["title"], "url": r["url"],
                              "date_published": r["date_published"], "date_fwrq": r["date_fwrq"],
                              "wh": r["wh"], "pdf": pdf_name, "text": text[:6000],
                              "hit_kind": hit, "attach_ext": ext, "pdf_how": pdf_how})
    json.dump({"window": [a.start, a.end], "stats": stats, "items": items},
              open(a.out, "w", encoding="utf-8"), ensure_ascii=False, indent=1)
    print("\n== 证监局抓取汇总 ==", json.dumps(stats, ensure_ascii=False))
    print("强命中(含私募) %d，弱命中 %d，正文失败 %d" % (stats["strong"], stats["weak"], stats["err"]))
    fl = stats.get("failed_items") or []
    if fl:
        print("\n[!] 有 %d 条决定书正文抓取失败，**未入库**（已写入 %s 的 stats.failed_items）："
              % (len(fl), a.out))
        for x in fl[:20]:
            print("    - %s·%s %s %s" % (x["bureau_cn"], x["col"], x["date_published"],
                                         x["title"][:48]))
        if len(fl) > 20:
            print("    …另有 %d 条" % (len(fl) - 20))
        print("    ⇒ 这些条目属于「抓到了列表却拿不到文书」，不补抓就是漏案；"
              "assert_crawl_complete.py 的 R8 会判 FAIL 阻断入库。")


if __name__ == "__main__":
    main()
