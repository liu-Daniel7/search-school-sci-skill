#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""把流水线结果写成工作簿（看板 + 明细 + 图例 + 配色）。"""
import re, collections
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.chart import BarChart, Reference

HDR = PatternFill("solid", fgColor="1F3864")
THIN = Border(*[Side(style="thin", color="BFBFBF")] * 4)
WARN = PatternFill("solid", fgColor="FFF2CC")
SUB = PatternFill("solid", fgColor="D6DCE4")

LEVEL = {
 "1区(Top)": ("C00000", "FFFFFF"), "1区": ("E06666", "FFFFFF"),
 "2区(Top)": ("ED7D31", "FFFFFF"), "2区": ("F4B183", "000000"),
 "3区": ("FFD966", "000000"), "4区": ("D9D9D9", "000000"),
 "非SCI期刊": ("BDD7EE", "000000"), "SCI收录(未入分区表)": ("9DC3E6", "000000"),
 "SCI/ESCI收录(未入分区表)": ("9DC3E6", "000000"),
 "CCF A类会议": ("7030A0", "FFFFFF"), "CCF B类会议": ("9E5FC8", "FFFFFF"),
 "CCF C类会议": ("C5A3DE", "000000"), "会议论文": ("A9D08E", "000000"),
 "预印本": ("D0CECE", "000000"), "其他(书章/社论等)": ("FFE699", "000000"),
 # 中文期刊（来自知网）
 "EI(中文)": ("1F6F8B", "FFFFFF"), "CSCD": ("4EA6C7", "FFFFFF"),
 "中文核心": ("7FB3D5", "000000"), "CSSCI": ("A9CCE3", "000000"), "中文普刊": ("E8F1F8", "000000"),
}
GROUP = [("1区 Top", lambda v: v.startswith("1区")), ("2区 Top", lambda v: v == "2区(Top)"),
         ("2区", lambda v: v in ("2区", "1区")), ("3区", lambda v: v == "3区"), ("4区", lambda v: v == "4区"),
         ("非SCI期刊", lambda v: v == "非SCI期刊"),
         ("会议", lambda v: "会议" in v),
         ("中文核心/EI/CSCD", lambda v: v in ("中文核心", "EI(中文)", "CSCD", "CSSCI")),
         ("中文普刊", lambda v: v == "中文普刊"),
         ("预印本/其他", lambda v: v in ("预印本", "其他(书章/社论等)", "SCI收录(未入分区表)", "SCI/ESCI收录(未入分区表)"))]
GROUP_FILL = {"1区 Top": "C00000", "2区 Top": "ED7D31", "2区": "F4B183", "3区": "FFD966",
              "4区": "D9D9D9", "非SCI期刊": "BDD7EE", "会议": "A9D08E",
              "中文核心/EI/CSCD": "4EA6C7", "中文普刊": "E8F1F8", "预印本/其他": "D0CECE"}
IDENT_ORDER = ["教师一作", "学生(推定)", "校内他院一作", "外单位一作"]
TWO_PLUS = ("1区(Top)", "1区", "2区(Top)", "2区")   # 中科院 2 区及以上（按层级判定，不依赖 src 字段）
RELIABILITY = {"教师一作": "已核实（学院师资页）", "学生(推定)": "推定 · 需学籍/人事名单核对",
               "校内他院一作": "已核实（单位明确）", "外单位一作": "已核实（单位明确）"}
SUR = {"wang": "王", "li": "李", "zhang": "张", "liu": "刘", "chen": "陈", "yang": "杨", "zhao": "赵",
       "huang": "黄", "wu": "吴", "zhou": "周", "xu": "徐/许", "sun": "孙", "ma": "马", "zhu": "朱",
       "hu": "胡", "guo": "郭", "he": "何", "gao": "高", "lin": "林", "luo": "罗", "zheng": "郑",
       "liang": "梁", "xie": "谢", "song": "宋", "tang": "唐", "han": "韩", "feng": "冯", "deng": "邓",
       "cao": "曹", "peng": "彭", "zeng": "曾", "xiao": "肖/萧", "tian": "田", "dong": "董", "yuan": "袁",
       "pan": "潘", "yu": "于/余/俞", "jiang": "江/蒋", "cai": "蔡", "du": "杜", "ye": "叶", "cheng": "程",
       "su": "苏", "wei": "魏/韦", "lu": "卢/陆/鲁", "ding": "丁", "ren": "任", "shen": "沈", "yao": "姚",
       "cui": "崔", "zhong": "钟", "tan": "谭", "fan": "范", "jin": "金", "shi": "石/史", "fang": "方",
       "xia": "夏", "qin": "秦", "hou": "侯", "zou": "邹", "meng": "孟", "xue": "薛", "bai": "白",
       "duan": "段", "lei": "雷", "qian": "钱", "long": "龙", "gu": "顾", "niu": "牛", "qi": "齐",
       "lv": "吕", "lü": "吕", "zhan": "詹", "kang": "康", "mao": "毛", "qiu": "邱", "wan": "万"}


def paint(cell, lv, bold=True):
    st = LEVEL.get(lv)
    if not st: return
    cell.fill = PatternFill("solid", fgColor=st[0])
    cell.font = Font(color=st[1], bold=bold, size=9)
    cell.alignment = Alignment(horizontal="center", vertical="center")


def header(ws, row, n, heights=30):
    for c in range(1, n + 1):
        cell = ws.cell(row=row, column=c)
        cell.font = Font(bold=True, color="FFFFFF", size=10)
        cell.fill = HDR
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = THIN
    ws.row_dimensions[row].height = max(ws.row_dimensions[row].height or 0, heights)


def widths(ws, ww):
    for i, w in enumerate(ww, 1):
        ws.column_dimensions[get_column_letter(i)].width = w


def surname(n, cn):
    if cn: return cn[0]
    p = [x for x in re.split(r"[\s\-\.]+", str(n)) if x]
    if len(p) < 2: return ""
    return SUR.get(p[-1].lower(), "")


def write_workbook(rows, cfg, path, dropped=0):
    cn2en_title = {f[0]: (f[2] if len(f) > 2 else "") for f in cfg["faculty"]}
    verified = cfg.get("verified_names", {})          # 英文名 → 中文名（已核实）
    fac_cn = {f[0] for f in cfg["faculty"]}
    core = [r for r in rows if r["scope"] == "我院主导"]
    _cnt = collections.Counter(r["dir"] for r in core)
    _tail = ("教学研究", "其他/交叉")
    dirs = sorted({r["dir"] for r in rows},
                  key=lambda d: (d in _tail, -_cnt.get(d, 0), d))
    stud = [r for r in rows if r["ident"] == "学生(推定)"]
    _cn = cfg["college"].get("cn")
    college = cfg["college"].get("display") or (_cn if isinstance(_cn, str) else (_cn or [""])[0])
    inst = cfg["institution"]["name"]

    CJK = re.compile(r"[\u4e00-\u9fff]")

    def cname(n):
        """中文署名直接返回；英文署名查已核实对照表。"""
        if CJK.search(str(n)): return str(n)
        return verified.get(n, "")

    def cname_or_todo(n):
        return cname(n) or "待补"

    wb = Workbook()
    # ---------------- 看板 ----------------
    ws = wb.active; ws.title = "① 看板"; ws.sheet_view.showGridLines = False
    gn = [g[0] for g in GROUP]
    widths(ws, [24, 10, 9, 10] + [10] * (len(gn) - 1) + [14])
    ws.merge_cells("A1:L1")
    ws["A1"] = f"{inst} · {college} 科研产出看板"
    ws["A1"].font = Font(bold=True, size=16, color="FFFFFF"); ws["A1"].fill = HDR
    ws["A1"].alignment = Alignment(horizontal="center", vertical="center"); ws.row_dimensions[1].height = 32
    ws.merge_cells("A2:L2")
    ws["A2"] = (f"口径：本院为第一作者或通讯作者单位 ｜ 期刊分级：中科院分区表（本地缓存） ｜ "
                f"数据源：OpenAlex 截止 {cfg.get('as_of', '')} ｜ 已剔除撤稿及排除名单人员论文")
    ws["A2"].font = Font(size=9, color="595959"); ws["A2"].alignment = Alignment(horizontal="center")
    ws.merge_cells("A3:L3")
    ws["A3"] = ("⚠ “学生”为推定身份（第一作者不在本院在编教师名单内），也可能是博士后、"
                "未公布的兼职教授或新聘讲师，须经学籍/人事核对后使用")
    ws["A3"].font = Font(size=9, bold=True, color="9C0006"); ws["A3"].fill = WARN
    ws["A3"].alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
    ws.row_dimensions[3].height = 22
    kpi = [("本院主导(篇)", len(core), "C00000"), ("本院参与(篇)", len(rows), "1F3864"),
           ("教师一作", sum(1 for r in rows if r["ident"] == "教师一作"), "2E75B6"),
           ("学生(推定)", len(stud), "2E75B6"),
           ("1区 Top(篇)", sum(1 for r in core if r["level"].startswith("1区")), "C00000"),
           ("2区及以上(篇)", sum(1 for r in core if r["level"] in TWO_PLUS), "C00000"),
           ("主导总被引", sum(r["work"]["cited_by_count"] for r in core), "7030A0")]
    for i, (k, v, col) in enumerate(kpi):
        c0 = 1 + i * 2 if i < 6 else 1; rr = 5 if i < 6 else 7; span = 1 if i < 6 else 3
        ws.merge_cells(start_row=rr, start_column=c0, end_row=rr, end_column=c0 + span)
        cell = ws.cell(row=rr, column=c0, value=k); cell.font = Font(size=9, color="595959")
        cell.alignment = Alignment(horizontal="center"); cell.fill = PatternFill("solid", fgColor="EDEDED")
        ws.merge_cells(start_row=rr + 1, start_column=c0, end_row=rr + 1, end_column=c0 + span)
        cell = ws.cell(row=rr + 1, column=c0, value=v)
        cell.font = Font(bold=True, size=15, color=col); cell.alignment = Alignment(horizontal="center")

    r = 10
    ws.cell(row=r, column=1, value="一、学科方向 × 期刊层级（本院主导口径）").font = Font(bold=True, size=12, color="1F3864")
    r += 1
    hdr = ["学科方向", "主导篇数", "占比", "主导被引"] + gn
    for i, h in enumerate(hdr, 1): ws.cell(row=r, column=i, value=h)
    header(ws, r, len(hdr)); r += 1
    for d in dirs:
        sub = [x for x in core if x["dir"] == d]
        if not sub: continue
        ws.cell(row=r, column=1, value=d).font = Font(bold=True, size=10)
        ws.cell(row=r, column=2, value=len(sub)).font = Font(bold=True, size=10)
        ws.cell(row=r, column=3, value=len(sub) / max(1, len(core))).number_format = "0.0%"
        ws.cell(row=r, column=4, value=sum(x["work"]["cited_by_count"] for x in sub))
        for j, (nm, f) in enumerate(GROUP):
            n = sum(1 for x in sub if f(x["level"]))
            cell = ws.cell(row=r, column=5 + j, value=n or None)
            if n: cell.fill = PatternFill("solid", fgColor=GROUP_FILL[nm])
        for i in range(1, len(hdr) + 1):
            cell = ws.cell(row=r, column=i); cell.border = THIN
            if i >= 2: cell.alignment = Alignment(horizontal="center")
        r += 1
    ws.cell(row=r, column=1, value="合计").font = Font(bold=True)
    ws.cell(row=r, column=2, value=len(core)).font = Font(bold=True)
    c = ws.cell(row=r, column=3, value=1.0); c.number_format = "0.0%"; c.font = Font(bold=True)
    ws.cell(row=r, column=4, value=sum(x["work"]["cited_by_count"] for x in core)).font = Font(bold=True)
    for j, (nm, f) in enumerate(GROUP):
        ws.cell(row=r, column=5 + j, value=sum(1 for x in core if f(x["level"]))).font = Font(bold=True)
    for i in range(1, len(hdr) + 1):
        cell = ws.cell(row=r, column=i); cell.fill = SUB; cell.border = THIN
        if i >= 2: cell.alignment = Alignment(horizontal="center")
    total_row = r

    r += 3
    ws.cell(row=r, column=1, value="二、学科方向 × 作者身份").font = Font(bold=True, size=12, color="1F3864")
    r += 1
    hdr2 = ["学科方向"] + IDENT_ORDER + ["合计", "其中1区 Top", "其中2区及以上", "主导占比"]
    for i, h in enumerate(hdr2, 1): ws.cell(row=r, column=i, value=h)
    header(ws, r, len(hdr2)); r += 1
    for d in dirs:
        sub = [x for x in rows if x["dir"] == d]
        if not sub: continue
        lead = [x for x in sub if x["scope"] == "我院主导"]
        row = [d] + [sum(1 for x in sub if x["ident"] == k) for k in IDENT_ORDER] + [len(sub),
               sum(1 for x in lead if x["level"].startswith("1区")),
               sum(1 for x in lead if x["level"] in TWO_PLUS),
               len(lead) / len(sub)]
        for i, v in enumerate(row, 1):
            cell = ws.cell(row=r, column=i, value=v); cell.border = THIN
            if i == 1: cell.font = Font(bold=True, size=10)
            elif i == len(row): cell.number_format = "0.0%"; cell.alignment = Alignment(horizontal="center")
            else: cell.alignment = Alignment(horizontal="center")
            if i == 6: cell.font = Font(bold=True)
        r += 1
    row = ["合计"] + [sum(1 for x in rows if x["ident"] == k) for k in IDENT_ORDER] + [len(rows),
           sum(1 for x in core if x["level"].startswith("1区")),
           sum(1 for x in core if x["level"] in TWO_PLUS), len(core) / max(1, len(rows))]
    for i, v in enumerate(row, 1):
        cell = ws.cell(row=r, column=i, value=v); cell.font = Font(bold=True); cell.fill = SUB; cell.border = THIN
        if i >= 2: cell.alignment = Alignment(horizontal="center")
        if i == len(row): cell.number_format = "0.0%"

    lr = r + 3
    ws.cell(row=lr, column=1, value="图例：期刊/成果层级配色").font = Font(bold=True, size=11, color="1F3864")
    lr += 1
    for i, h in enumerate(["层级", "主导篇数", "合计篇数"], 1): ws.cell(row=lr, column=i, value=h)
    header(ws, lr, 3); lr += 1
    _present = {r["level"] for r in rows}
    for lv in [x for x in LEVEL if x in _present]:
        c1 = ws.cell(row=lr, column=1, value=lv); paint(c1, lv); c1.border = THIN
        a = ws.cell(row=lr, column=2, value=sum(1 for x in core if x["level"] == lv))
        b = ws.cell(row=lr, column=3, value=sum(1 for x in rows if x["level"] == lv))
        for c in (a, b): c.border = THIN; c.alignment = Alignment(horizontal="center")
        lr += 1
    ch = BarChart(); ch.type = "bar"; ch.title = "各方向主导成果数"
    ch.height = 9; ch.width = 13; ch.legend = None
    ch.add_data(Reference(ws, min_col=2, min_row=11, max_row=10 + len(dirs)), titles_from_data=True)
    ch.set_categories(Reference(ws, min_col=1, min_row=12, max_row=11 + len(dirs)))
    ws.add_chart(ch, f"E{lr + 2}")   # 图例有多少行，图就往后让多少，避免压住

    # ---------------- 论文清单 ----------------
    ws = wb.create_sheet("② 论文清单(可筛选)")
    cols = ["年份", "学科方向", "统计口径", "身份", "身份可信度", "第一作者(英)", "第一作者(中)", "层级", "题名",
            "期刊/会议", "中科院大类", "分区", "Top", "IF", "被引", "DOI", "第一单位", "本院通讯", "我方作者"]
    widths(ws, [7, 20, 10, 12, 26, 17, 13, 13, 60, 28, 12, 7, 6, 7, 7, 24, 14, 10, 18])
    ws.append(cols); header(ws, 1, len(cols)); ws.freeze_panes = "C2"
    for r0 in sorted(rows, key=lambda x: (dirs.index(x["dir"]), x["ord"], -(x["work"]["publication_year"] or 0), -x["work"]["cited_by_count"])):
        w = r0["work"]; fn = w["authorships"][0]["author"]["display_name"] if w["authorships"] else ""
        ws.append([w["publication_year"], r0["dir"], r0["scope"], r0["ident"], RELIABILITY[r0["ident"]], fn,
                   cname_or_todo(fn), r0["level"], w["title"], r0["src"]["期刊"], r0["src"]["中科院大类"],
                   r0["src"]["中科院分区"], r0["src"]["Top"], r0["src"]["IF"], w["cited_by_count"],
                   (w.get("doi") or "").replace("https://doi.org/", ""), r0["fu"],
                   "是" if r0["corr"] else "", ", ".join(r0["facs"])])
    for rr in range(2, ws.max_row + 1): paint(ws.cell(row=rr, column=8), ws.cell(row=rr, column=8).value)
    ws.auto_filter.ref = f"A1:S{ws.max_row}"

    # ---------------- 姓名对照表 ----------------
    ws = wb.create_sheet("③ 姓名对照表(待填)")
    ws.append(["第一作者(英)", "第一作者(中)", "姓氏参考", "主导篇数", "参与篇数", "身份", "建议核实方式"])
    header(ws, 1, 7); widths(ws, [22, 16, 12, 10, 10, 22, 26]); ws.freeze_panes = "A2"
    agg = collections.defaultdict(lambda: [0, 0, set()])
    for r0 in rows:
        w = r0["work"]
        if not w["authorships"]: continue
        n = w["authorships"][0]["author"]["display_name"]
        agg[n][1] += 1
        if r0["scope"] == "我院主导": agg[n][0] += 1
        agg[n][2].add(r0["ident"])
    for n, (lead, part, ids) in sorted(agg.items(), key=lambda x: (-x[1][0], -x[1][1])):
        ident = "/".join(sorted(ids))
        hint = "学号/研究生院名单" if "学生(推定)" in ident else (
            "学院师资页" if "教师一作" in ident else "署名单位已明确")
        ws.append([n, cname(n) or "", surname(n, cname(n)), lead, part, ident, hint])
    ws.auto_filter.ref = f"A1:G{ws.max_row}"

    # ---------------- 教师-方向矩阵 ----------------
    ws = wb.create_sheet("④ 教师-方向矩阵")
    col_cn = {f[0]: f[1].lower() for f in cfg["faculty"]}
    keys = {f[1].lower(): f[0] for f in cfg["faculty"] if f[1]}
    title_of = {f[0]: (f[2] if len(f) > 2 else "") for f in cfg["faculty"]}
    per = collections.defaultdict(collections.Counter); tot = collections.Counter()
    for r0 in core:
        for a in r0["work"]["authorships"]:
            k = None
            p = [x for x in re.split(r"[\s\-\.]+", a["author"]["display_name"]) if x]
            if len(p) >= 2:
                for cand in ("".join(p).lower(), (p[-1] + "".join(p[:-1])).lower()):
                    cand = re.sub(r"[^a-z]", "", cand)
                    if cand in keys: k = keys[cand]; break
            if k: per[k][r0["dir"]] += 1; tot[k] += 1
    ws.append(["姓名", "职称", "主导合计"] + dirs); header(ws, 1, 3 + len(dirs))
    widths(ws, [10, 12, 11] + [15] * len(dirs)); ws.freeze_panes = "B2"
    for k, c in sorted(tot.items(), key=lambda x: -x[1]):
        ws.append([k, title_of.get(k, ""), c] + [per[k][d] for d in dirs])
        rr = ws.max_row
        for i in range(4, 4 + len(dirs)):
            cell = ws.cell(row=rr, column=i)
            if cell.value:
                cell.alignment = Alignment(horizontal="center")
                cell.fill = PatternFill("solid", fgColor="DDEBF7" if cell.value < 5 else ("9DC3E6" if cell.value < 15 else "2E75B6"))
                if cell.value >= 15: cell.font = Font(color="FFFFFF", bold=True)
    ws.auto_filter.ref = f"A1:{get_column_letter(3 + len(dirs))}{ws.max_row}"

    # ---------------- 教师产出 ----------------
    ws = wb.create_sheet("⑤ 教师产出")
    head = ["姓名", "职称", "英文署名", "主导篇数", "署名总篇数", "一作", "通讯", "1区Top", "2区Top", "2区",
            "3区", "4区", "非SCI", "会议", "主导被引"]
    ws.append(head); header(ws, 1, len(head))
    widths(ws, [10, 12, 24, 10, 11, 8, 8, 8, 8, 8, 8, 8, 8, 8, 10]); ws.freeze_panes = "B2"
    for i, lv in [(8, "1区(Top)"), (9, "2区(Top)"), (10, "2区"), (11, "3区"), (12, "4区"),
                  (13, "非SCI期刊"), (14, "会议论文")]:
        paint(ws.cell(row=1, column=i), lv)
    stats = collections.defaultdict(lambda: dict(sig=0, lead=0, first=0, corr=0, cite=0,
                                                 lv=collections.Counter(), nm=collections.Counter()))
    def _key(n):
        p = [x for x in re.split(r"[\s\-\.]+", n) if x]
        if len(p) < 2: return None
        for cand in ("".join(p).lower(), (p[-1] + "".join(p[:-1])).lower()):
            cand = re.sub(r"[^a-z]", "", cand)
            if cand in keys: return keys[cand]
        return None
    for r0 in rows:
        w = r0["work"]
        for a in w["authorships"]:
            k = _key(a["author"]["display_name"])
            if not k: continue
            stats[k]["sig"] += 1; stats[k]["nm"][a["author"]["display_name"]] += 1
        for k in {_key(a["author"]["display_name"]) for a in w["authorships"]}:
            if not k: continue
            s = stats[k]
            if r0["scope"] == "我院主导":
                s["lead"] += 1; s["cite"] += w["cited_by_count"]; s["lv"][r0["level"]] += 1
            if r0["ident"] == "教师一作" and _key(w["authorships"][0]["author"]["display_name"]) == k:
                s["first"] += 1
            if any(a.get("is_corresponding") and _key(a["author"]["display_name"]) == k for a in w["authorships"]):
                s["corr"] += 1
    body = []
    for k, s in stats.items():
        lv = s["lv"]
        body.append([k, title_of.get(k, ""), s["nm"].most_common(1)[0][0], s["lead"], s["sig"], s["first"],
                     s["corr"], lv.get("1区(Top)", 0), lv.get("2区(Top)", 0), lv.get("2区", 0) + lv.get("1区", 0),
                     lv.get("3区", 0), lv.get("4区", 0), lv.get("非SCI期刊", 0),
                     sum(v for kk, v in lv.items() if "会议" in kk), s["cite"]])
    for x in sorted(body, key=lambda x: (-x[3], -x[14])): ws.append(x)
    ws.auto_filter.ref = f"A1:O{ws.max_row}"

    # ---------------- 学生(推定) ----------------
    ws = wb.create_sheet("⑥ 学生(推定)")
    head = ["年份", "学科方向", "第一作者(英)", "第一作者(中)", "身份可信度", "题名", "期刊/会议", "分区", "IF",
            "被引", "本院通讯", "DOI"]
    ws.append(head); header(ws, 1, len(head))
    widths(ws, [7, 20, 17, 13, 26, 58, 28, 12, 7, 7, 18, 24]); ws.freeze_panes = "C2"
    for r0 in sorted(stud, key=lambda x: (dirs.index(x["dir"]), -(x["work"]["publication_year"] or 0), -x["work"]["cited_by_count"])):
        w = r0["work"]; fn = w["authorships"][0]["author"]["display_name"]
        ws.append([w["publication_year"], r0["dir"], fn, cname_or_todo(fn), RELIABILITY["学生(推定)"],
                   w["title"], r0["src"]["期刊"], r0["level"], r0["src"]["IF"], w["cited_by_count"],
                   ", ".join(r0["facs"]), (w.get("doi") or "").replace("https://doi.org/", "")])
    for rr in range(2, ws.max_row + 1): paint(ws.cell(row=rr, column=8), ws.cell(row=rr, column=8).value)
    ws.auto_filter.ref = f"A1:L{ws.max_row}"

    # ---------------- 说明 ----------------
    ws = wb.create_sheet("⑦ 说明")
    ws.column_dimensions["A"].width = 24; ws.column_dimensions["B"].width = 104
    info = [("项目", "内容"), ("学院", f"{inst} · {college}"),
            ("数据来源", "OpenAlex 题录 + 中科院《期刊分区表》+ JCR + CCF 推荐目录（本地缓存）"),
            ("截止", cfg.get("as_of", "")), ("", ""),
            ("⚠ 身份口径", "“学生”为推定：第一作者不在本院在编教师名单内。可能是①未公布的兼职教授 ②新聘讲师/师资博士后 ③外聘访问人员 ④已毕业但署名仍为我校者。不得作为学生成果对外使用，须经学籍/人事核对。"),
            ("教师身份", "与学院师资页名单逐一比对，并用 OpenAlex 作者ID 精确匹配以避免同名。"),
            ("中文名", "教师姓名准确；其他第一作者需人工补齐，表中标“待补”，并在“③ 姓名对照表”提供姓氏参考与核实方式。拼音→汉字一对多，不做推测。"),
            ("", ""), ("统计口径", f"本院主导 {len(core)} 篇（第一作者或通讯作者为本院）；本院参与 {len(rows)} 篇。"),
            ("清洗", f"剔除撤稿与排除名单人员论文 {dropped} 篇；同名学院误命中须在配置中加校名限定。"),
            ("局限", "分区取最新版，与发表当年可能不同；未含知网独有中文期刊；被引为 OpenAlex 口径。")]
    for x in info: ws.append(list(x))
    for rr in ws.iter_rows(min_row=1, max_row=len(info), max_col=2):
        rr[0].font = Font(bold=True); rr[0].alignment = Alignment(vertical="top", wrap_text=True)
        rr[1].alignment = Alignment(vertical="top", wrap_text=True)
    for c in list(ws[1]): c.fill = HDR; c.font = Font(bold=True, color="FFFFFF")

    wb.save(path)

    by_dir = {}
    for d in dirs:
        sub = [x for x in core if x["dir"] == d]
        by_dir[d] = [len(sub)] + [sum(1 for x in sub if f(x["level"])) for _, f in GROUP]
    return {"institution": inst, "college": college,
            "core": len(core), "core_cite": sum(x["work"]["cited_by_count"] for x in core),
            "total": len(rows),
            "teacher_first": sum(1 for x in rows if x["ident"] == "教师一作"),
            "student_first": len(stud),
            "other_college": sum(1 for x in rows if x["ident"] == "校内他院一作"),
            "external": sum(1 for x in rows if x["ident"] == "外单位一作"),
            "dropped": dropped, "by_dir": by_dir}
