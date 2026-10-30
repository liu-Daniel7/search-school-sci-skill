#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
college_sci.py — 学院级科研产出统计流水线（可直接复用）

输入：一份 config.json（机构 + 学院名/写法族 + 教师名单 + 口径开关）
输出：xlsx 工作簿（看板/清单/姓名表/矩阵/教师/学生/说明）+ summary.md + summary.json

依赖：openpyxl、pandas（DSH 环境用 load_workspace_dependencies 提供的 python）
外部：api.openalex.org、GitHub raw（中科院分区表/CCF 名单，首次下载后本地缓存）

用法：
  python college_sci.py --config ../reference/config.example.json --out /path/out.xlsx
  python college_sci.py --config cfg.json --stage census      # 只做单位串普查
"""
import argparse, hashlib, json, re, subprocess, sys, time, collections
from pathlib import Path

OPENALEX = "https://api.openalex.org"
SEL = ("id,doi,title,publication_year,cited_by_count,type,language,"
       "primary_location,authorships,topics,open_access")
UA = "Mozilla/5.0 (compatible; college-sci-audit/1.0)"

# ---------------------------------------------------------------- 网络层
class Client:
    def __init__(self, cache_dir, delay=1.0):
        self.dir = Path(cache_dir); self.dir.mkdir(parents=True, exist_ok=True)
        self.delay = delay; self.n = 0

    def get(self, url, retries=5, throttle=True):
        key = self.dir / (hashlib.sha1(url.encode()).hexdigest() + ".json")
        if key.exists():
            return json.loads(key.read_text(encoding="utf-8"))
        for k in range(retries):
            if throttle: time.sleep(self.delay)
            txt = self._curl(url)
            try:
                data = json.loads(txt)
            except Exception:
                data = None
            if isinstance(data, dict) and "results" in data:
                key.write_text(json.dumps(data, ensure_ascii=False), encoding="utf-8")
                self.n += 1
                return data
            time.sleep(3 * (k + 1))
        raise RuntimeError("请求失败: " + url)

    @staticmethod
    def _curl(url):
        r = subprocess.run(["curl", "-s", "-m", "90", "-A", UA, url],
                           capture_output=True, text=True)
        return r.stdout or ""


def oa_works(cli, filt, select=SEL, per_page=100, max_pages=30, verbose=True):
    out, page = [], 1
    while True:
        q = {"filter": filt, "per_page": str(per_page), "page": str(page), "select": select}
        try:
            d = cli.get(OPENALEX + "/works?" + urlencode(q))
        except RuntimeError:
            break
        out += d["results"]
        total = d["meta"]["count"]
        if page == 1 and verbose:
            print(f"    filter={filt[:80]}… count={total}", flush=True)
        if len(out) >= total or not d["results"] or page >= max_pages:
            break
        page += 1
    return out


def urlencode(d):
    from urllib.parse import urlencode as _u
    return _u(d)


# ---------------------------------------------------------------- 基础工具
def norm_name(s):
    """刊名归一化：去标点、去 AND。"""
    return re.sub(r"[^0-9A-Z]", "", str(s).upper()).replace("AND", "")


def norm_issn(s):
    return re.sub(r"[^0-9X]", "", str(s).upper())


def pykey(display_name, faculty_keys):
    """英文署名 → 教师拼音键（兼容 Jianwei Yang / YANG Jianwei 两种写法）。"""
    p = [x for x in re.split(r"[\s\-\.]+", str(display_name)) if x]
    if len(p) < 2:
        return None
    a = re.sub(r"[^a-z]", "", "".join(p).lower())
    b = re.sub(r"[^a-z]", "", (p[-1] + "".join(p[:-1])).lower())
    if a in faculty_keys: return a
    if b in faculty_keys: return b
    return None


# ---------------------------------------------------------------- 期刊指标
_P = "%E4%B8%AD%E7%A7%91%E9%99%A2%E5%88%86%E5%8C%BA%E8%A1%A8%E5%8F%8AJCR%E5%8E%9F%E5%A7%8B%E6%95%B0%E6%8D%AE%E6%96%87%E4%BB%B6"
MIRRORS = [f"https://cdn.jsdelivr.net/gh/hitfyd/ShowJCR@master/{_P}/",      # 通常最快
           f"https://raw.githubusercontent.com/hitfyd/ShowJCR/master/{_P}/",
           f"https://gitee.com/mirrors_github/ShowJCR/raw/master/{_P}/"]
METRIC_FILES = ["FQBJCR2025-UTF8.csv", "JCR2025-UTF8.csv", "XR2026-UTF8.csv", "CCF2026-UTF8.csv"]


def ensure_metrics(cache_dir):
    d = Path(cache_dir); d.mkdir(parents=True, exist_ok=True)
    for f in METRIC_FILES:
        p = d / f
        if p.exists() and p.stat().st_size > 1000:
            continue
        ok = False
        for base in MIRRORS:
            print(f"  下载指标文件 {f} ← {base.split('/')[2]} …", flush=True)
            subprocess.run(["curl", "-sL", "-m", "600", "--retry", "2", "-A", UA,
                            "-o", str(p), base + f], check=False)
            if p.exists() and p.stat().st_size > 1000:
                head = p.open("rb").read(200)
                if b"Journal" in head or b"\xe5\x88\x8a\xe5\x90\x8d" in head:
                    ok = True; break
        if not ok:
            raise SystemExit(f"指标文件下载失败：{f}。请检查网络，或手工放入 {p}（来源：github.com/hitfyd/ShowJCR）")
    return d


def load_metrics(cache_dir):
    import pandas as pd
    d = ensure_metrics(cache_dir)
    def rd(f, **kw):
        return pd.read_csv(d / f, encoding="utf-8-sig", low_memory=False, **kw)
    fq, jcr, xr, ccf = rd("FQBJCR2025-UTF8.csv"), rd("JCR2025-UTF8.csv"), rd("XR2026-UTF8.csv"), rd("CCF2026-UTF8.csv")
    def pick(row, i):
        v = row.get(f"IF Quartile(2025)_{i}")
        return v if isinstance(v, str) and v.startswith("Q") else None
    cas_issn, cas_name = {}, {}
    for _, r in fq.iterrows():
        rec = {"大类": r.get("大类"), "分区": lead_int(r.get("大类分区")), "top": r.get("Top"),
               "wos": r.get("Web of Science"), "rank": str(r.get("大类分区") or "")}
        for x in str(r.get("ISSN/EISSN")).split("/"):
            if norm_issn(x): cas_issn[norm_issn(x)] = rec
        cas_name[norm_name(r.get("Journal"))] = rec
    j_issn, j_name = {}, {}
    for _, r in jcr.iterrows():
        qs = [q for q in (pick(r, i) for i in range(1, 7)) if q]
        rec = {"if": num(r.get("IF(2025)")) if not pd.isna(r.get("IF(2025)")) else None,
               "q": min(qs) if qs else None, "wos": r.get("Web of Science")}
        for x in (r.get("ISSN"), r.get("EISSN")):
            if isinstance(x, str) and norm_issn(x): j_issn[norm_issn(x)] = rec
        j_name[norm_name(r.get("Journal"))] = rec
    x_issn, x_name = {}, {}
    for _, r in xr.iterrows():
        rec = {"分区": lead_int(r.get("大类新锐分区"))}
        for x in (r.get("ISSN"), r.get("EISSN")):
            if isinstance(x, str) and norm_issn(x): x_issn[norm_issn(x)] = rec
        x_name[norm_name(r.get("刊名"))] = rec
    ccf_name = {norm_name(r.get("Journal")): r.get("CCF推荐类型") for _, r in ccf.iterrows()}
    return {"cas_issn": cas_issn, "cas_name": cas_name, "j_issn": j_issn, "j_name": j_name,
            "x_issn": x_issn, "x_name": x_name, "ccf": ccf_name}


def num(v):
    """宽松数值解析：'<0.1' / '10.2' / NaN → float 或 None"""
    try:
        if v is None: return None
        t = str(v).replace("<", "").replace(">", "").strip()
        return float(t) if t and t.lower() != "nan" else None
    except Exception:
        return None


def lead_int(v):
    m = re.match(r"\s*(\d)", str(v)) if v is not None else None
    return int(m.group(1)) if m else None


def journal_info(metrics, name, issn):
    def look(issn_map, name_map):
        if issn and norm_issn(issn) in issn_map: return issn_map[norm_issn(issn)]
        n = norm_name(name) if name else ""
        return name_map.get(n) or name_map.get(n)
    a = look(metrics["cas_issn"], metrics["cas_name"])
    j = look(metrics["j_issn"], metrics["j_name"])
    x = look(metrics["x_issn"], metrics["x_name"])
    c = metrics["ccf"].get(norm_name(name)) if name else None
    return {"期刊": name or "", "ISSN": issn or "",
            "中科院大类": (a or {}).get("大类"), "中科院分区": (a or {}).get("分区"),
            "Top": (a or {}).get("top"), "排名": (a or {}).get("rank"),
            "IF": (j or {}).get("if"), "JCR分区": (j or {}).get("q"),
            "新锐分区": (x or {}).get("分区"), "CCF": c,
            "WOS": (a or j or {}).get("wos"), "中文刊名": ""}


def level_of(work, info):
    t = work["type"]; name = info["期刊"]
    if t in ("article", "review") and name:
        if info["中科院分区"]:
            lv = f"{info['中科院分区']}区" + ("(Top)" if str(info["Top"]).strip() == "是" else "")
            return lv, info["中科院分区"]
        if info["JCR分区"]: return "SCI收录(未入分区表)", 5
        if info["WOS"]: return "SCI/ESCI收录(未入分区表)", 6
        return "非SCI期刊", 7
    if t == "conference-paper":
        return (f"CCF {info['CCF']}会议" if info["CCF"] else "会议论文"), 8
    if t == "preprint": return "预印本", 9
    return "其他(书章/社论等)", 9


# ---------------------------------------------------------------- 学科方向
def tag_direction(work, dirs):
    sc = collections.Counter()
    for t in (work.get("topics") or [])[:3]:
        d = dirs["topics"].get(t["display_name"])
        if d: sc[d] += 3
    title = (work.get("title") or "").lower()
    for d, pats in dirs["keywords"].items():
        for p in pats:
            if re.search(p, title): sc[d] += 1
    if not sc: return dirs.get("fallback", "其他/交叉")
    order = dirs["priority"]
    best = max(sc.items(), key=lambda x: (x[1], -(order.index(x[0]) if x[0] in order else 99)))
    return best[0]


# ---------------------------------------------------------------- 归属判定
def _as_list(v):
    if v is None: return []
    return [v] if isinstance(v, str) else list(v)


def build_rules(cfg):
    en = _as_list(cfg["college"].get("en"))
    cn = _as_list(cfg["college"].get("cn"))
    colleges = "|".join([re.escape(x) for x in en + cn] + _as_list(cfg["college"].get("extra_patterns")))
    labs = _as_list(cfg.get("lab_patterns"))
    lab_re = "|".join(re.escape(x) for x in labs) if labs else None
    return {
        "inst": re.compile(cfg["institution"]["match_regex"], re.I),
        "inst_id": None,
        "college": re.compile(f"({colleges})", re.I),
        "lab": re.compile(f"({lab_re})", re.I) if lab_re else None,
        "other": re.compile("(" + "|".join(re.escape(x) for x in cfg.get("other_school_patterns", [])) + ")", re.I)
                 if cfg.get("other_school_patterns") else None,
    }


def unit_of(auth, R):
    """返回该署名单位串的学院归属：本院 / 校内其他学院 / 仅校名 / 外单位"""
    for s in (auth.get("raw_affiliation_strings") or []):
        if not R["inst"].search(s):
            continue
        if R["other"] and R["other"].search(s):
            return "校内其他学院"
        if R["college"].search(s) or (R["lab"] and R["lab"].search(s)):
            return "本院"
    for s in (auth.get("raw_affiliation_strings") or []):
        if R["inst"].search(s):
            return "仅校名"
    inst_ids = {i.get("id", "").split("/")[-1] for i in auth.get("institutions", [])}
    return "本校(机构ID)" if R.get("inst_id") in inst_ids else "外单位"


# ---------------------------------------------------------------- 主流程
def run(cfg, args):
    cli = Client(args.cache, delay=cfg.get("request_delay", 1.0))
    inst = cfg["institution"]
    R = build_rules(cfg)
    fac = {(f[1] or "").lower(): {"cn": f[0], "title": f[2] if len(f) > 2 else ""} for f in cfg["faculty"] if f[1]}
    FAC_KEYS = set(fac)

    # 1) 机构解析
    print("[1/7] 解析机构 …")
    d = cli.get(OPENALEX + "/institutions?" + urlencode({"search": inst["name"]}))
    target = None
    for r in d.get("results", []):
        if inst.get("ror") and inst["ror"].split("/")[-1] == r.get("ror", "").split("/")[-1]:
            target = r; break
    if not target and d.get("results"):
        target = d["results"][0]
    if not target:
        raise SystemExit("未找到机构，请检查 institution.name")
    inst_id = target["id"].split("/")[-1]
    R["inst_id"] = inst_id
    print(f"      {target['display_name']}  ror={target.get('ror')}  id={inst_id}  works={target.get('works_count')}")

    # 2) 单位串普查 + 同名机构告警
    print("[2/7] 单位串普查（同名学院检测）…")
    for phrase in _as_list(cfg["college"].get("en")) + _as_list(cfg["college"].get("cn")) + _as_list(cfg.get("lab_patterns")):
        n = count(cli, f'raw_affiliation_strings.search:"{phrase}"')
        m = count(cli, f'institutions.id:{inst_id},raw_affiliation_strings.search:"{phrase}"')
        print(f"      {phrase}: 全球 {n} 篇 / 本校 {m} 篇（差 {n - m} 为他校同名）")
    homonym = check_homonym(cli, cfg, R)
    if args.stage == "census":
        print("\n[census] 仅普查模式，到此结束。确认写法无误后去掉 --stage 参数跑全流程。")
        return None
    if homonym:
        print("      ⚠ 检测到同名院系（必须用校名限定，否则会混入他人论文）：")
        for name, n in homonym[:8]:
            print(f"        - {name}: {n} 篇")

    # 3) 抓取论文
    print("[3/7] 抓取论文 …")
    works, seen = [], set()
    guard = cfg.get("unfiltered_max_hits", 2000)
    for phrase in _as_list(cfg["college"].get("en")) + _as_list(cfg["college"].get("cn")) + _as_list(cfg.get("lab_patterns")):
        bare = f'raw_affiliation_strings.search:"{phrase}"'
        n = count(cli, bare)
        filt = bare if n <= guard else f'institutions.id:{inst_id},{bare}'
        if n > guard:
            print(f"      “{phrase}” 全局命中 {n} 篇，过大 → 自动加机构限定，避免抓取他校论文")
        for w in oa_works(cli, filt):
            if w["id"] not in seen:
                seen.add(w["id"]); works.append(w)
    for kw in cfg.get("affiliation_keywords", []):
        for w in oa_works(cli, f"institutions.id:{inst_id},raw_affiliation_strings.search:{kw}"):
            if w["id"] not in seen:
                seen.add(w["id"]); works.append(w)
    print(f"      抓到候选 {len(works)} 篇")

    # 4) 教师署名扩展
    if cfg.get("expand_by_faculty_authors", True):
        print("[4/7] 教师作者ID扩展（补全只写校名的论文）…")
        ids = set()
        for w in works:
            for a in w["authorships"]:
                if pykey(a["author"]["display_name"], FAC_KEYS) and unit_of(a, R) in ("本院", "仅校名", "本校(机构ID)"):
                    aid = (a["author"].get("id") or "").split("/")[-1]
                    if aid: ids.add(aid)
        ids = sorted(ids)
        print(f"      教师作者ID {len(ids)} 个")
        for i in range(0, len(ids), 12):
            for w in oa_works(cli, "authorships.author.id:" + "|".join(ids[i:i + 12]), per_page=100):
                if w["id"] not in seen:
                    seen.add(w["id"]); works.append(w)
        print(f"      扩展后 {len(works)} 篇")

    # 5) 归属与身份
    print("[5/7] 归属判定与身份划分 …")
    records = []
    excl = set(cfg.get("exclude_faculty", []))
    for w in works:
        au = w["authorships"]
        if not au: continue
        fu = unit_of(au[0], R)
        corr_college = any(a.get("is_corresponding") and unit_of(a, R) == "本院" for a in au)
        facs = [fac[k]["cn"] for k in {pykey(a["author"]["display_name"], FAC_KEYS) for a in au} if k]
        retracted = w["type"] == "retraction" or str(w["title"]).upper().startswith("RETRACTED")
        only_excl = bool(facs) and set(facs).issubset(excl)
        college_sig = any(unit_of(a, R) == "本院" for a in au)
        if not college_sig and not (facs and any(unit_of(a, R) in ("本院", "仅校名", "本校(机构ID)") for a in au)):
            continue
        if retracted or only_excl:
            records.append(("drop", w, fu, facs, corr_college)); continue
        fk = pykey(au[0]["author"]["display_name"], FAC_KEYS)
        # 只写校名的论文：默认视为“待核实”，不计入主导；人工核实后可开启开关提升
        promote = fu == "仅校名" and cfg.get("treat_university_only_as_college", False)
        eff = "本院" if (fu == "本院" or promote) else fu
        if eff == "本院":
            ident = "教师一作" if fk else "学生(推定)"
        elif eff == "校内其他学院":
            ident = "校内他院一作"
        else:
            ident = "外单位一作"
        scope = "我院主导" if (eff == "本院" or corr_college or ident == "教师一作") else "我院参与"
        records.append(("keep", w, fu, facs, corr_college, ident, scope))
    keep = [r for r in records if r[0] == "keep"]
    dropped = len(records) - len(keep)

    # 6) 期刊指标与方向
    print("[6/7] 期刊分级与方向标注 …")
    metrics = load_metrics(args.cache)
    dirs = json.loads(Path(cfg["directions"]).read_text(encoding="utf-8"))
    out = []
    for _, w, fu, facs, corr, ident, scope in keep:
        src = (w["primary_location"] or {}).get("source") or {}
        info = journal_info(metrics, src.get("display_name"), src.get("issn_l") or (src.get("issn") or [None])[0])
        lv, order = level_of(w, info)
        out.append({"work": w, "fu": fu, "facs": facs, "corr": corr, "ident": ident, "scope": scope,
                    "src": info, "level": lv, "ord": order, "dir": tag_direction(w, dirs)})
    print(f"      入库 {len(out)} 篇（剔除 {dropped}：撤稿/排除名单）")

    # 7) 输出
    print("[7/7] 生成工作簿 …")
    from workbook import write_workbook
    dump_records(out, Path(args.out).with_suffix(".records.json"))
    summary = write_workbook(out, cfg, args.out, dropped=dropped)
    Path(args.out).with_suffix(".summary.json").write_text(
        json.dumps(summary, ensure_ascii=False, indent=2), encoding="utf-8")
    write_markdown(summary, Path(args.out).with_suffix(".md"), cfg)
    print("完成 →", args.out)
    return summary


def dump_records(rows, path):
    """沉淀中间结果，供 cnki_import.py 合并中文期刊数据。"""
    slim = []
    for r in rows:
        w = r["work"]
        slim.append({
            "title": w.get("title"), "year": w.get("publication_year"),
            "cited": w.get("cited_by_count", 0), "doi": (w.get("doi") or "").replace("https://doi.org/", ""),
            "journal": (w.get("primary_location") or {}).get("source", {}).get("display_name") if (w.get("primary_location") or {}).get("source") else "",
            "type": w.get("type"),
            "authors": [{"name": a["author"]["display_name"],
                         "affs": a.get("raw_affiliation_strings") or [],
                         "corr": bool(a.get("is_corresponding"))} for a in w["authorships"]],
            "ident": r["ident"], "scope": r["scope"], "fu": r["fu"],
            "facs": r["facs"], "level": r["level"], "ord": r["ord"], "dir": r["dir"],
            "src": r["src"], "source_db": "OpenAlex"})
    Path(path).write_text(json.dumps(slim, ensure_ascii=False), encoding="utf-8")


def count(cli, filt):
    try:
        d = cli.get(OPENALEX + "/works?" + urlencode({"filter": filt, "per_page": "1", "select": "id"}))
        return d["meta"]["count"]
    except RuntimeError:
        return 0


def check_homonym(cli, cfg, R):
    """检测同名院系：只按院系名搜索，看命中的学校有哪些。"""
    hits = collections.Counter()
    for phrase in _as_list(cfg["college"].get("en")):
        try:
            d = cli.get(OPENALEX + "/works?" + urlencode(
                {"filter": f'raw_affiliation_strings.search:"{phrase}"', "per_page": "50", "select": "id,authorships"}))
        except RuntimeError:
            continue
        for w in d.get("results", []):
            for a in w["authorships"]:
                for s in (a.get("raw_affiliation_strings") or []):
                    if phrase.lower() in s.lower() and not R["inst"].search(s):
                        m = re.search(r"([A-Z][A-Za-z\.\&\s]{6,60}(University|College|Institute))", s)
                        if m: hits[m.group(1).strip()] += 1
    return hits.most_common()


def write_markdown(s, path, cfg):
    disp = cfg["college"].get("display") or _as_list(cfg["college"].get("cn"))[0]
    L = [f"# {cfg['institution']['name']} · {disp} 科研产出速览", "",
         f"- 我院主导：**{s['core']} 篇**（被引 {s['core_cite']}）",
         f"- 我院参与：**{s['total']} 篇**",
         f"- 教师一作 {s['teacher_first']} ｜ 学生(推定) {s['student_first']} ｜ 校内他院一作 {s['other_college']} ｜ 外单位一作 {s['external']}",
         "", "## 方向 × 层级（主导口径）", "",
         "| 方向 | 主导 | 1区Top | 2区Top | 2区 | 3区 | 4区 | 非SCI | 会议 |", "|---|---|---|---|---|---|---|---|---|"]
    for d, row in s["by_dir"].items():
        L.append("| " + " | ".join(str(x) for x in [d] + row) + " |")
    L += ["", "> 学生身份为推定（第一作者不在在编教师名单内），可能是博士后、未公布的兼职教授或新聘讲师，须经学籍/人事核对。"]
    path.write_text("\n".join(L), encoding="utf-8")


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--config", required=True)
    ap.add_argument("--out", default="college_sci.xlsx")
    ap.add_argument("--cache", default="./.college_sci_cache")
    ap.add_argument("--faculty-csv", help="教师名单 CSV（中文名,全拼,职称），覆盖配置里的 faculty")
    ap.add_argument("--stage", choices=["all", "census"], default="all",
                    help="census=只做单位串普查与同名检测，不抓论文")
    a = ap.parse_args()
    cfg = json.loads(Path(a.config).read_text(encoding="utf-8"))
    cfg["directions"] = str((Path(a.config).parent / cfg.get("directions", "directions.json")).resolve())
    if a.faculty_csv:
        import csv as _csv
        rows = []
        with open(a.faculty_csv, encoding="utf-8-sig") as f:
            for r in _csv.DictReader(f):
                cn = (r.get("中文名") or "").strip()
                if cn: rows.append([cn, (r.get("全拼") or "").strip().lower(), (r.get("职称") or "").strip()])
        cfg["faculty"] = rows
        print(f"教师名单 ← {a.faculty_csv}（{len(rows)} 人）")
    run(cfg, a)


if __name__ == "__main__":
    main()
