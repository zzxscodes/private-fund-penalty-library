# -*- coding: utf-8 -*-
"""06_法规依据 重做：
1) 有官方 PDF 的只留 PDF，删除同名的多余 Word；
2) 只能网页阅读的，抓取权威源正文 → 清洗导航噪声 → 按法规体例规范排版为 Word。

用法: python fetch_laws3.py <06_法规依据目录>
"""
import re
import ssl
import sys
import tempfile
import urllib.request
import fitz
from pathlib import Path

from bs4 import BeautifulSoup
from docx import Document
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.oxml.ns import qn
from docx.shared import Pt, RGBColor, Cm

UA = ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
      "(KHTML, like Gecko) Chrome/126.0.0.0 Safari/537.36")
CTX = ssl.create_default_context()
CTX.check_hostname = False
CTX.verify_mode = ssl.CERT_NONE

# 部分老旧 HTTPS 站点（如 ambers.amac.org.cn）需要「不安全旧版重协商」，
# 默认上下文会报 UNSAFE_LEGACY_RENEGOTIATION_DISABLED，故单独放开。
CTX_LEGACY = ssl.create_default_context()
CTX_LEGACY.check_hostname = False
CTX_LEGACY.verify_mode = ssl.CERT_NONE
CTX_LEGACY.options |= 0x00040000  # OP_ALLOW_UNSAFE_LEGACY_RENEGOTIATION

def _parse_args():
    """显式参数解析（2026-10-28 修，见 SKILL 0.4.94）。

    原实现是 `ROOT = Path(sys.argv[1])` 裸取位置参数：
      * 没有 argparse → `--help` 不被识别，**被当成输出目录名**，
        实测有人在技能根目录执行 `fetch_laws.py --help`，
        直接生成 `--help/01_法律与行政法规/…` 等 16 份下载文件；
      * 无参数时 `sys.argv[1]` 直接 IndexError，看不出该传什么。
    本脚本会**联网下载并写盘**，参数错了代价很大，所以必须显式解析 + 强校验：
    不给目录就报错退出，绝不用当前目录兜底。
    """
    import argparse
    ap = argparse.ArgumentParser(
        description="06_配套法规依据：下载官方 PDF/Word；只能网页阅读的抓正文排版为 Word")
    ap.add_argument("root", nargs="?", default="",
                    help="06 法规依据目录（<案例库根>/06_以上各类所有案例的配套法规依据）")
    # 默认缓存目录必须落在**系统临时目录**，绝不能再是脚本同级目录：
    # 2026-10-28 实测——旧默认值 `_lawsrc3` 让技能包 scripts/ 里长出一个 692KB、
    # 28 个 txt 的缓存目录，被一起打进技能包（污染交付物 + 违反"技能包可迁移"）。
    ap.add_argument("--tmp", default="", help="网页法规正文缓存目录（默认系统临时目录下的 _lawsrc3）")
    a = ap.parse_args()
    if not a.root or a.root.startswith("-"):
        ap.error("必须显式给出 06 法规依据目录路径（本脚本会联网下载并写盘，绝不猜目录、绝不用当前目录兜底）")
    return a


_ARGS = _parse_args()
ROOT = Path(_ARGS.root)
if not ROOT.name.startswith("06"):
    print("  [提示] 传入目录名不是以 06 开头，请确认是否为「06_…配套法规依据」目录：%s" % ROOT)
TMP = Path(_ARGS.tmp) if _ARGS.tmp else Path(tempfile.gettempdir()) / "_lawsrc3"
TMP.mkdir(parents=True, exist_ok=True)

# 已有官方文件（PDF/Word）的法规：直接下载权威源并按规范文件名留存，
# 不再走「HTML→重新排版 Word」流程（保留官方原貌、避免重复抓取）。
# 2026-10-27 修复：原 PDF_ONLY 仅删除多余 Word、从未真正下载官方 PDF，
# 导致 finalize --laws 把它们当成「缺失」。现改为直接下载官方文件。
DOWNLOAD_LAWS = [
    dict(sub="01_法律与行政法规", name="中华人民共和国证券投资基金法",
         url="https://wb.flk.npc.gov.cn/flfg/PDF/d9a88f22ce024b0282b3496bf58c9c8a.pdf", ext="pdf"),
    dict(sub="02_证监会部门规章与规范性文件", name="证券期货经营机构私募资产管理业务管理办法",
         url="https://www.csrc.gov.cn/csrc/c101953/c6987995/6987995/files/%E9%99%84%E4%BB%B61%EF%BC%9A%E3%80%8A%E8%AF%81%E5%88%B8%E6%9C%9F%E8%B4%A7%E7%BB%8F%E8%90%A5%E6%9C%BA%E6%9E%84%E7%A7%81%E5%8B%9F%E8%B5%84%E4%BA%A7%E7%AE%A1%E7%90%86%E4%B8%9A%E5%8A%A1%E7%AE%A1%E7%90%86%E5%8A%9E%E6%B3%95%E3%80%8B.pdf", ext="pdf"),
    dict(sub="02_证监会部门规章与规范性文件", name="上市公司股东减持股份管理暂行办法",
         url="http://www.csrc.gov.cn/csrc/c101953/c7483190/7483190/files/%E9%99%84%E4%BB%B61%EF%BC%9A%E4%B8%8A%E5%B8%82%E5%85%AC%E5%8F%B8%E8%82%A1%E4%B8%9C%E5%87%8F%E6%8C%81%E8%82%A1%E4%BB%BD%E7%AE%A1%E7%90%86%E6%9A%82%E8%A1%8C%E5%8A%9E%E6%B3%95.pdf", ext="pdf"),
    dict(sub="02_证监会部门规章与规范性文件", name="证券投资基金托管业务管理办法",
         url="http://www.csrc.gov.cn/csrc/c101877/c1029534/1029534/files/%E9%99%84%E4%BB%B61%EF%BC%9A%E8%AF%81%E5%88%B8%E6%8A%95%E8%B5%84%E5%9F%BA%E9%87%91%E6%89%98%E7%AE%A1%E4%B8%9A%E5%8A%A1%E7%AE%A1%E7%90%86%E5%8A%9E%E6%B3%95.pdf", ext="pdf"),
    dict(sub="02_证监会部门规章与规范性文件", name="期货公司互联网营销管理暂行规定",
         url="https://www.csrc.gov.cn/jiangsu/c105409/c7589261/7589261/files/%E6%9C%9F%E8%B4%A7%E5%85%AC%E5%8F%B8%E4%BA%92%E8%81%94%E7%BD%91%E8%90%A5%E9%94%80%E7%AE%A1%E7%90%86%E6%9A%82%E8%A1%8C%E8%A7%84%E5%AE%9A%EF%BC%88%E6%9D%A5%E6%BA%90%EF%BC%9A%E8%AF%81%E7%9B%91%E4%BC%9A%E5%AE%98%E7%BD%91%EF%BC%89.pdf", ext="pdf"),
    dict(sub="02_证监会部门规章与规范性文件", name="上市公司收购管理办法",
         url="https://www.csrc.gov.cn/csrc/c106256/c1653983/1653983/files/166-7%E3%80%8A%E4%B8%8A%E5%B8%82%E5%85%AC%E5%8F%B8%E6%94%B6%E8%B4%AD%E7%AE%A1%E7%90%86%E5%8A%9E%E6%B3%95%E3%80%8B.docx", ext="docx"),
    dict(sub="03_中基协自律规则", name="私募投资基金登记备案办法",
         url="https://ambers.amac.org.cn/web/app/static/template/managerRegisterFilingMeasures.pdf", ext="pdf"),
    dict(sub="03_中基协自律规则", name="中国证券投资基金业协会自律管理和纪律处分措施实施办法",
         url="https://www.amac.org.cn/xwfb/xhyw/202212/P020231126355298387129.pdf", ext="pdf"),
    dict(sub="03_中基协自律规则", name="私募证券投资基金运作指引",
         url="https://fg.amac.org.cn//governmentrules_3854/zcgz_zlgz/zlgz_smjj/zlgz_smjj_cpl/202404/P020240430697829777505.pdf", ext="pdf"),
    dict(sub="03_中基协自律规则", name="私募基金管理人失联处理指引",
         url="https://fg.amac.org.cn/governmentrules_3854/zcgz_zlgz/zlgz_smjj/zlgz_smjj_cpl/202307/P020231126438073714449.pdf", ext="pdf"),
    dict(sub="04_交易所业务规则（沪深北交易所）", name="上海证券交易所首次公开发行证券发行与承销业务实施细则",
         url="https://www.sse.com.cn/lawandrules/sselawsrules2025/repeal/rules/c/10785058/files/3b35a11d21844816955e4843fdb63e85.docx", ext="docx"),
    dict(sub="04_交易所业务规则（沪深北交易所）", name="上海证券交易所交易规则",
         url="https://www.sse.com.cn/lawandrules/sselawsrules2025/trade/universal/c/10816492/files/58bd5ce4dfc94b57a5391f60eb2719ab.docx", ext="docx"),
    # 2026-10-28 补：D3 校验点名引用（交易所案例）但 06 缺失的交易所业务规则
    dict(sub="04_交易所业务规则（沪深北交易所）", name="深圳证券交易所股票上市规则（2025年修订）",
         url="https://docs.static.szse.cn/www/lawrules/rule/repeal/rules/W020250425730780478873.pdf", ext="pdf"),
]

# 仅网页发布、需复制留存为 Word 的法规
LAWS = [
    dict(sub="01_法律与行政法规", name="私募投资基金监督管理条例",
         no="国务院令第762号", org="中华人民共和国国务院", date="2023-07-09", eff="2023-09-01",
         status="现行有效", site="中国政府网",
         url="https://www.gov.cn/zhengce/content/202307/content_6890738.htm", kind="gov"),
    dict(sub="02_证监会部门规章与规范性文件", name="私募投资基金监督管理暂行办法",
         no="证监会令第105号", org="中国证券监督管理委员会", date="2014-08-21", eff="2014-08-21",
         status="现行有效", site="中国证券监督管理委员会",
         url="https://www.csrc.gov.cn/csrc/c106256/c1653981/content.shtml", kind="gov"),
    dict(sub="02_证监会部门规章与规范性文件", name="关于加强私募投资基金监管的若干规定",
         no="证监会公告〔2020〕71号", org="中国证券监督管理委员会", date="2020-12-30", eff="2020-12-30",
         status="现行有效", site="中国政府网·国务院公报",
         url="https://www.gov.cn/gongbao/content/2021/content_5593454.htm", kind="gov"),
    dict(sub="02_证监会部门规章与规范性文件",
         name="国务院办公厅关于加强监管防范风险促进私募投资基金高质量发展的指导意见",
         no="国办函〔2026〕54号", org="国务院办公厅", date="2026-06-03", eff="2026-06-03",
         status="现行有效", site="中国政府网",
         url="https://www.gov.cn/zhengce/zhengceku/202606/content_7071205.htm", kind="gov"),
    dict(sub="03_中基协自律规则", name="私募投资基金募集行为管理办法",
         no="中基协发〔2016〕7号", org="中国证券投资基金业协会", date="2016-04-15", eff="2016-07-15",
         status="现行有效", site="中国证券投资基金业协会法规库",
         url="https://fg.amac.org.cn/governmentrules_3854/zcgz_zlgz/zlgz_smjj/zlgz_smjj_cpl/201912/t20191222_24544.html", kind="amac"),
    dict(sub="03_中基协自律规则", name="私募投资基金信息披露管理办法",
         no="中基协（2016年2月4日发布）", org="中国证券投资基金业协会", date="2016-02-04", eff="2016-02-04",
         status="现行有效", site="中国证券投资基金业协会法规库",
         url="https://fg.amac.org.cn/governmentrules_3854/zcgz_zlgz/zlgz_smjj/zlgz_smjj_cpl/201912/t20191222_24542.html", kind="amac"),
    dict(sub="03_中基协自律规则", name="私募投资基金管理人内部控制指引",
         no="中基协（2016年2月1日发布）", org="中国证券投资基金业协会", date="2016-02-01", eff="2016-02-01",
         status="现行有效", site="中国证券投资基金业协会法规库",
         url="https://fg.amac.org.cn/governmentrules_3854/zcgz_zlgz/zlgz_smjj/zlgz_smjj_cpl/201912/t20191222_24540.html", kind="amac"),
    dict(sub="03_中基协自律规则", name="关于进一步加强私募基金行业自律管理的决定",
         no="中基协（2018年3月27日发布，未编号）", org="中国证券投资基金业协会", date="2018-03-27", eff="2018-03-27",
         status="现行有效", site="中国证券投资基金业协会法规库",
         url="https://fg.amac.org.cn/governmentrules_3854/zcgz_zlgz/zlgz_smjj/zlgz_smjj_cpl/201912/t20191222_24536.html", kind="amac"),
    dict(sub="03_中基协自律规则", name="关于私募基金管理人在异常经营情形下提交专项法律意见书的公告",
         no="中基协公告（2018年3月27日发布，未编号）", org="中国证券投资基金业协会", date="2018-03-27", eff="2018-03-27",
         status="现行有效", site="中国证券投资基金业协会法规库",
         url="https://fg.amac.org.cn/governmentrules_3854/zcgz_zlgz/zlgz_smjj/zlgz_smjj_cpl/201912/t20191222_24538.html", kind="amac"),
    # ---- 2026-10-27 一年库重建：D3 校验点名引用 ≥3 次但未入库的配套法规（决定书正文 ∪ 核心处罚依据）----
    dict(sub="01_法律与行政法规", name="中华人民共和国证券法",
         no="主席令第37号", org="全国人民代表大会常务委员会", date="2019-12-28", eff="2020-03-01",
         status="现行有效", site="中国人大网",
         url="http://www.npc.gov.cn/c2/c30834/201912/t20191231_304436.html", kind="gov"),
    dict(sub="01_法律与行政法规", name="证券公司监督管理条例",
         no="国务院令第588号", org="中华人民共和国国务院", date="2014-07-29", eff="2014-07-29",
         status="现行有效", site="中国政府网·国务院公报",
         url="https://www.gov.cn/gongbao/content/2016/content_5139684.htm", kind="gov"),
    dict(sub="02_证监会部门规章与规范性文件", name="期货公司监督管理办法",
         no="证监会令第155号", org="中国证券监督管理委员会", date="2019-06-04", eff="2020-03-01",
         status="现行有效", site="中国证券监督管理委员会",
         url="https://www.csrc.gov.cn/csrc/c106256/c1654011/content.shtml", kind="gov"),
    dict(sub="02_证监会部门规章与规范性文件", name="证券期货市场诚信监督管理办法",
         no="证监会令第139号", org="中国证券监督管理委员会", date="2018-07-01", eff="2018-07-01",
         status="现行有效", site="中国证券监督管理委员会",
         url="https://www.csrc.gov.cn/csrc/c106256/c1653941/content.shtml", kind="gov"),
    dict(sub="02_证监会部门规章与规范性文件", name="证券期货经营机构及其工作人员廉洁从业规定",
         no="证监会令第145号", org="中国证券监督管理委员会", date="2018-06-27", eff="2018-06-27",
         status="现行有效", site="中国证券监督管理委员会",
         url="https://www.csrc.gov.cn/csrc/c106256/c1654026/content.shtml", kind="gov"),
    dict(sub="02_证监会部门规章与规范性文件", name="上市公司信息披露管理办法",
         no="证监会令第182号", org="中国证券监督管理委员会", date="2021-03-18", eff="2021-05-01",
         status="现行有效", site="中国证券监督管理委员会",
         url="https://www.csrc.gov.cn/csrc/c106256/c1653948/content.shtml", kind="gov"),
    dict(sub="02_证监会部门规章与规范性文件", name="证券市场禁入规定",
         no="证监会令第185号", org="中国证券监督管理委员会", date="2021-06-15", eff="2021-07-19",
         status="现行有效", site="中国政府网·国务院公报",
         url="https://www.gov.cn/zhengce/2021-12/16/content_5724536.htm", kind="gov"),
    dict(sub="02_证监会部门规章与规范性文件", name="证券期货经营机构私募资产管理业务运作管理暂行规定",
         no="证监会公告〔2016〕13号", org="中国证券监督管理委员会", date="2016-07-14", eff="2016-07-18",
         status="现行有效", site="中国政府网·国务院公报",
         url="https://www.gov.cn/gongbao/content/2017/content_5189214.htm", kind="gov"),
    dict(sub="02_证监会部门规章与规范性文件", name="公开募集证券投资基金销售机构监督管理办法",
         no="证监会令第175号", org="中国证券监督管理委员会", date="2020-08-28", eff="2020-10-01",
         status="现行有效", site="中国证券监督管理委员会",
         url="https://www.csrc.gov.cn/csrc/c106256/c1653806/content.shtml", kind="gov"),
    dict(sub="02_证监会部门规章与规范性文件", name="证券公司和证券投资基金管理公司合规管理办法",
         no="证监会令第133号", org="中国证券监督管理委员会", date="2017-06-06", eff="2017-06-06",
         status="现行有效", site="中国证券监督管理委员会",
         url="https://www.csrc.gov.cn/csrc/c106256/c1653949/content.shtml", kind="gov"),
    dict(sub="02_证监会部门规章与规范性文件", name="期货公司董事、监事和高级管理人员任职管理办法",
         no="证监会令第179号", org="中国证券监督管理委员会", date="2021-01-15", eff="2021-01-15",
         status="现行有效", site="中国证券监督管理委员会",
         url="https://www.csrc.gov.cn/csrc/c106256/c1653888/content.shtml", kind="gov"),
    dict(sub="02_证监会部门规章与规范性文件", name="证券期货市场监督管理措施实施办法",
         no="证监会令第231号", org="中国证券监督管理委员会", date="2025-12-31", eff="2026-06-30",
         status="现行有效", site="中国政府网·国务院公报",
         url="https://www.gov.cn/gongbao/2026/issue_12626/202603/content_7063196.html", kind="gov"),
    dict(sub="02_证监会部门规章与规范性文件", name="证券经纪业务管理办法",
         no="证监会令第204号", org="中国证券监督管理委员会", date="2023-01-13", eff="2023-02-28",
         status="现行有效", site="中国政府网·国务院公报",
         url="https://www.gov.cn/gongbao/content/2023/content_5750630.htm", kind="gov"),
    dict(sub="03_中基协自律规则", name="基金募集机构投资者适当性管理实施指引（试行）",
         no="中基协发〔2017〕4号", org="中国证券投资基金业协会", date="2017-06-28", eff="2017-07-01",
         status="现行有效", site="中国证券投资基金业协会",
         url="https://www.amac.org.cn/fwdt/wyb/jgdjhcpbeian/zcglcpba/xgzc/201912/t20191222_19886.html", kind="gov"),
    dict(sub="03_中基协自律规则", name="中国证券投资基金业协会自律检查规则",
         no="中基协发〔2023〕16号", org="中国证券投资基金业协会", date="2023-07-14", eff="2023-07-14",
         status="现行有效", site="上海证券报（官方PDF）",
         url="https://www.cnstock.com/image/202307/15/20230715174050334.pdf", kind="amac"),
    dict(sub="03_中基协自律规则", name="基金从业人员管理规则",
         no="中基协发〔2023〕23号", org="中国证券投资基金业协会", date="2023-11-24", eff="2023-11-24",
         status="现行有效", site="中国证券投资基金业协会（官方PDF）",
         url="https://www.amac.org.cn/xwfb/xhyw/202205/P020231126355300536293.pdf", kind="amac"),
    # ---- 2026-10-27 续接：D3 校验剩余缺失、用国务院公报/证监会官网 HTML 权威源补全 ----
    dict(sub="02_证监会部门规章与规范性文件", name="证券期货投资者适当性管理办法",
         no="证监会令第130号", org="中国证券监督管理委员会", date="2016-12-12", eff="2017-07-01",
         status="现行有效", site="中国政府网·国务院公报",
         url="https://www.gov.cn/gongbao/content/2017/content_5213202.htm", kind="gov"),
    dict(sub="02_证监会部门规章与规范性文件",
         name="证券基金经营机构董事、监事、高级管理人员及从业人员监督管理办法",
         no="证监会令第195号", org="中国证券监督管理委员会", date="2022-02-18", eff="2022-04-01",
         status="现行有效", site="中国政府网·国务院公报",
         url="https://www.gov.cn/gongbao/content/2022/content_5691004.htm", kind="gov"),
    # ---- 2026-10-27 续接第二轮：补齐 finalize --laws 仍报缺失的剩余法规 ----
    dict(sub="02_证监会部门规章与规范性文件", name="期货从业人员管理办法",
         no="证监会令第48号", org="中国证券监督管理委员会", date="2007-07-04", eff="2007-07-04",
         status="现行有效", site="中国证券监督管理委员会",
         url="https://www.csrc.gov.cn/csrc/c101902/c1039261/content.shtml", kind="gov"),
    # 2026-10-28 补：D3 校验点名引用（交易所/债券类案例）但 06 缺失的部门规章
    dict(sub="02_证监会部门规章与规范性文件", name="公司债券发行与交易管理办法",
         no="证监会令第222号", org="中国证券监督管理委员会", date="2023-10-20", eff="2023-10-20",
         status="现行有效", site="中国政府网·国务院公报",
         url="https://www.gov.cn/gongbao/2023/issue_10886/202312/content_6921380.html", kind="gov"),
]

# 噪声行（网页UI/导航/版权等，与法规正文无关）
NOISE_EXACT = {
    "首页", ">", "|", "字号：", "默认", "大", "超大", "打印", "关闭", "分享到", "扫一扫",
    "当前位置：", "投资者之家", "微信", "微博", "综合类", "政策法规", "私募基金", "工作动态",
    "相关稿件", "相关链接", "上一篇", "下一篇", "返回顶部", "联系我们", "法律声明", "网站地图",
    "主办单位", "网站标识码", "联系我们|", "（本文有删减）", "意见征集", "找错", "英文版",
}
NOISE_PAT = [
    r"^https?://", r"^www\.", r"^\d{1,4}$",                       # 纯URL、纯页码
    r"^Copyright", r"版权所有", r"ICP备", r"公网安备", r"京ICP",
    r"^来源[:：]", r"^责任编辑", r"^浏览次数", r"^发布时间[:：]?$",
    r"^扫一扫", r"^分享", r"^\(function", r"^var\s", r"^document\.",
    r"^中国政府网$", r"^国务院公报$", r"^字号$", r"^关闭窗口$",
]


def fetch(url):
    req = urllib.request.Request(url, headers={
        "User-Agent": UA, "Accept-Language": "zh-CN,zh;q=0.9",
        "Accept": "text/html,application/xhtml+xml,*/*"})
    return urllib.request.urlopen(req, timeout=60, context=CTX).read().decode("utf-8", "ignore")


def pick_main(soup):
    """选取正文容器：优先已知选择器，其次取包含法规条文特征的最长文本块。"""
    cands = []
    for sel in ["#UCAP-CONTENT", "div.content", ".pages_content", ".TRS_Editor", "#zoom",
                ".article", ".content-con", ".detail-content"]:
        for n in soup.select(sel):
            cands.append(n)
    scored = []
    for n in cands:
        txt = n.get_text("\n", strip=True)
        score = len(txt) + 3000 * (1 if re.search(r"第[一二三四五六七八九十]+条", txt) else 0)
        scored.append((score, n))
    if scored:
        scored.sort(key=lambda x: x[0], reverse=True)
        return scored[0][1]
    # 退化：全文最长文本块
    best, bl = soup.body or soup, 0
    for tag in soup.find_all(["div", "td", "article"]):
        t = tag.get_text("\n", strip=True)
        if len(t) > bl:
            best, bl = tag, len(t)
    return best


def clean_lines(raw_lines):
    """raw_lines：已按行切分的字符串列表（HTML 取文本或 PDF 取文本均可）。"""
    lines = []
    for ln in raw_lines:
        # 剔除混入的控制字符（csrc 等页面偶发），否则 python-docx 报 XML 不兼容
        ln = re.sub(r"[\x00-\x08\x0b\x0c\x0e-\x1f]", "", ln)
        ln = re.sub(r"[\u3000\xa0]+", " ", ln).strip()
        if not ln:
            continue
        if ln in NOISE_EXACT:
            continue
        if any(re.search(p, ln) for p in NOISE_PAT):
            continue
        if len(ln) <= 2 and not re.match(r"^[一二三四五六七八九十]+$", ln):
            continue
        if lines and lines[-1] == ln:
            continue
        lines.append(ln)
    return lines


def fetch_pdf_text(url):
    req = urllib.request.Request(url, headers={
        "User-Agent": UA, "Accept-Language": "zh-CN,zh;q=0.9"})
    data = urllib.request.urlopen(req, timeout=90, context=CTX).read()
    doc = fitz.open(stream=data, filetype="pdf")
    txt = "\n".join(p.get_text() for p in doc)
    doc.close()
    return txt


META_KEYS = ["文件名称", "效力状态", "失效日期", "发文单位", "文号", "发文日期", "实施日期"]


DATE_RE = re.compile(r"^\d{4}[-/年]\d{1,2}([-/月]\d{1,2})?日?$")


def strip_amac_meta(lines):
    """中基协法规库正文开头是元信息键值串，解析出来并从正文剥离。

    注意该页面元信息字段常缺值、且偶把文号填进「发文单位」，
    因此按字段类型做校验：日期字段只接受日期格式，疑似文号的值归入文号。
    """
    if not lines or lines[0] != "文件名称":
        return lines, {}
    meta, i, n = {}, 0, len(lines)
    while i < n and lines[i] in META_KEYS:
        key = lines[i]
        nxt = lines[i + 1] if i + 1 < n else ""
        val = "" if nxt in META_KEYS else nxt
        if key in ("发文日期", "实施日期", "失效日期"):
            if val and DATE_RE.match(val):
                meta[key] = val
                i += 2
            else:
                i += 1
        elif key == "发文单位":
            if val and ("号" in val) and len(val) < 40:
                meta.setdefault("文号", val)      # 页面把文号误填在发文单位
                i += 2
            else:
                if val:
                    meta["发文单位"] = val
                i += 2 if val else 1
        else:
            if val:
                meta[key] = val
            i += 2 if val else 1
        if key == "实施日期":
            break
    return lines[i:], meta


RE_CHAP = re.compile(r"^第[一二三四五六七八九十百]+章")
RE_ART = re.compile(r"^第[一二三四五六七八九十百]+条$")
RE_ITEM = re.compile(r"^[（(][一二三四五六七八九十]+[)）]")
RE_SEC = re.compile(r"^[一二三四五六七八九十]+、")


def structure(lines):
    """把清洗后的行组织成 (类型, 文本) 列表：chap 章 / art 条 / item 项 / p 普通。"""
    out, pending_art = [], None
    for ln in lines:
        if RE_CHAP.match(ln):
            if pending_art:
                out.append(("art", pending_art)); pending_art = None
            out.append(("chap", ln))
        elif RE_ART.match(ln):
            if pending_art:
                out.append(("art", pending_art))
            pending_art = ln
        elif pending_art is not None:
            out.append(("art", f"{pending_art}　{ln}"))
            pending_art = None
        elif RE_SEC.match(ln) and len(ln) < 24:
            out.append(("chap", ln))
        elif RE_ITEM.match(ln):
            out.append(("item", ln))
        else:
            out.append(("p", ln))
    if pending_art:
        out.append(("art", pending_art))
    return out


def set_font(run, name="宋体", size=12, bold=False, color=None):
    run.font.name = name
    run.font.size = Pt(size)
    run.bold = bold
    if color:
        run.font.color.rgb = color
    run._element.rPr.rFonts.set(qn("w:eastAsia"), name)


def build_docx(path, law, blocks):
    doc = Document()
    sec = doc.sections[0]
    sec.top_margin = Cm(2.54); sec.bottom_margin = Cm(2.54)
    sec.left_margin = Cm(3.17); sec.right_margin = Cm(3.17)

    st = doc.styles["Normal"]
    st.font.name = "宋体"; st.font.size = Pt(12)
    st.element.rPr.rFonts.set(qn("w:eastAsia"), "宋体")

    # 标题
    h = doc.add_paragraph()
    h.alignment = WD_ALIGN_PARAGRAPH.CENTER
    h.paragraph_format.space_after = Pt(10)
    set_font(h.add_run(law["name"]), "方正小标宋简体", 18, bold=True)

    # 元信息表（法规要素，非杂乱信息）
    rows = [("发布机关", law["org"]), ("文　　号", law["no"]),
            ("发布日期", law["date"]), ("施行日期", law["eff"]),
            ("效力状态", law["status"]), ("权威来源", law["site"])]
    tbl = doc.add_table(rows=0, cols=2)
    tbl.style = "Table Grid"
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    for k, v in rows:
        cells = tbl.add_row().cells
        cells[0].width = Cm(2.4); cells[1].width = Cm(12.0)
        p0 = cells[0].paragraphs[0]
        set_font(p0.add_run(k), "楷体", 10.5)
        p1 = cells[1].paragraphs[0]
        set_font(p1.add_run(v), "楷体", 10.5)
    doc.add_paragraph().paragraph_format.space_after = Pt(4)

    # 正文
    for typ, text in blocks:
        p = doc.add_paragraph()
        pf = p.paragraph_format
        pf.line_spacing_rule = WD_LINE_SPACING.MULTIPLE
        pf.line_spacing = 1.5
        pf.space_after = Pt(3)
        if typ == "chap":
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            pf.space_before = Pt(10); pf.space_after = Pt(6)
            set_font(p.add_run(text), "黑体", 14, bold=True)
        elif typ == "art":
            # 条号加粗，与条文同段，首行缩进 2 字符
            pf.first_line_indent = Pt(24)
            m = re.match(r"^(第[一二三四五六七八九十百]+条)　?(.*)$", text)
            if m:
                set_font(p.add_run(m.group(1) + "　"), "黑体", 12, bold=True)
                set_font(p.add_run(m.group(2)), "宋体", 12)
            else:
                set_font(p.add_run(text), "宋体", 12)
        elif typ == "item":
            pf.first_line_indent = Pt(24)
            set_font(p.add_run(text), "宋体", 12)
        else:
            pf.first_line_indent = Pt(24)
            set_font(p.add_run(text), "宋体", 12)

    # 溯源（单行小字）
    src = doc.add_paragraph()
    src.paragraph_format.space_before = Pt(14)
    set_font(src.add_run("原文来源（权威网站）：" + law["url"]), "楷体", 9, color=RGBColor(0x66, 0x66, 0x66))
    doc.save(str(path))


def main():
    # 1) 直接下载「已有官方文件」的法规（权威源 PDF/Word）
    print("== 直接下载官方文件 ==")
    ok_d, skip_d, fail_d = [], [], []
    for i, law in enumerate(DOWNLOAD_LAWS, 1):
        outdir = ROOT / law["sub"]
        outdir.mkdir(parents=True, exist_ok=True)
        dest = outdir / f"{law['name']}.{law['ext']}"
        print(f"[{i}/{len(DOWNLOAD_LAWS)}] {law['name'][:34]} ...", end=" ", flush=True)
        if dest.exists() and dest.stat().st_size > 1024:
            print("已存在，跳过")
            skip_d.append(law["name"])
            continue
        try:
            req = urllib.request.Request(
                law["url"], headers={"User-Agent": UA, "Accept-Language": "zh-CN,zh;q=0.9"})
            data = urllib.request.urlopen(req, timeout=120, context=CTX_LEGACY).read()
            if len(data) < 1024:
                raise RuntimeError("下载内容过小")
            dest.write_bytes(data)
            print("OK %d 字节" % len(data))
            ok_d.append(law["name"])
        except Exception as e:
            print("失败:", type(e).__name__, str(e)[:80])
            fail_d.append(law["name"])
    print("官方文件：新增 %d，跳过 %d，失败 %d\n" % (len(ok_d), len(skip_d), len(fail_d)))

    # 2) 重新生成 HTML→Word（已存在则跳过，避免重复抓取）
    ok, fail = [], []
    for i, law in enumerate(LAWS, 1):
        outdir = ROOT / law["sub"]
        outdir.mkdir(parents=True, exist_ok=True)
        dest = outdir / f"{law['name']}.docx"
        if dest.exists() and dest.stat().st_size > 1024:
            continue
        print(f"[{i}/{len(LAWS)}] {law['name'][:34]} ...", end=" ", flush=True)
        try:
            if law["url"].lower().endswith(".pdf"):
                raw = fetch_pdf_text(law["url"])
                lines = clean_lines(raw.split("\n"))
                if len("".join(lines)) <= 900:
                    raise RuntimeError("PDF正文提取过短")
            else:
                html = fetch(law["url"])
                soup = None
                for parser in ("html.parser", "lxml"):
                    s = BeautifulSoup(html, parser)
                    for t in s(["script", "style", "noscript", "nav", "header", "footer"]):
                        t.decompose()
                    # 去掉面包屑/分享等 UI 节点
                    for t in s.find_all(attrs={"class": re.compile(r"bread|crumb|share|print|related|nav|footer|header", re.I)}):
                        t.decompose()
                    node = pick_main(s)
                    lines = clean_lines(node.get_text("\n", strip=True).split("\n"))
                    if len("".join(lines)) > 900:
                        soup = s
                        break
                if soup is None:
                    raise RuntimeError("正文提取过短")
            if law["kind"] == "amac" and not law["url"].lower().endswith(".pdf"):
                lines, meta = strip_amac_meta(lines)
                # 正文首行若与法规名重复则去掉
                while lines and lines[0] in (law["name"], meta.get("文件名称")):
                    lines = lines[1:]
                # 以权威页面元信息为准覆盖手工填值（仅取校验通过的值）
                if meta.get("文号"):
                    law["no"] = meta["文号"]
                if meta.get("发文日期"):
                    law["date"] = meta["发文日期"]
                if meta.get("实施日期"):
                    law["eff"] = meta["实施日期"]
                if meta.get("发文单位") and len(meta["发文单位"]) < 20:
                    law["org"] = {"中基协": "中国证券投资基金业协会",
                                  "证监会": "中国证券监督管理委员会"}.get(
                                      meta["发文单位"], meta["发文单位"])
                if meta.get("效力状态"):
                    law["status"] = meta["效力状态"]
            blocks = structure(lines)
            body_len = sum(len(t) for _, t in blocks)
            build_docx(outdir / f"{law['name']}.docx", law, blocks)
            print(f"OK 正文{body_len}字 / {len(blocks)}段")
            ok.append(law["name"])
            (TMP / (re.sub(r'[\\/:*?"<>|]', "_", law["name"]) + ".txt")).write_text(
                "\n".join(t for _, t in blocks), encoding="utf-8")
        except Exception as e:
            print("失败:", type(e).__name__, str(e)[:80])
            fail.append(law["name"])
    print(f"\n生成 {len(ok)} 个规范Word，失败 {len(fail)}")
    for n in fail:
        print("  ×", n)


if __name__ == "__main__":
    main()
