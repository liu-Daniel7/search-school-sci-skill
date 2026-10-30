#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
cnki_import.py — 把知网（CNKI）导出的中文期刊数据并入主数据集

分工：
  · 用户/操作者负责「登录知网 + 检索 + 导出」（见 reference/CNKI_GUIDE.md）
  · 本脚本负责「解析导出文件 + 去重 + 中文期刊分级 + 与 OpenAlex 数据合并 + 重出工作簿」

用法：
  python cnki_import.py --records out.records.json --cnki-dir ./cnki_export \
      --config config.json --out 结果_含中文.xlsx

导出文件命名约定（决定中文期刊等级，优先级高于表内字段）：
  cnki_北大核心.csv / cnki_CSSCI.csv / cnki_CSCD.csv / cnki_EI.csv / cnki_普刊.csv / cnki_任意名.csv
  也支持 .xlsx，以及表内带「来源类别」列的情况。
"""
import argparse, json, re, sys, csv
from pathlib import Path

# ---------------------------------------------------------------- 列名别名
ALIAS = {
    "title": ["题名", "篇名", "标题", "题目", "title"],
    "authors": ["作者", "全部作者", "author", "authors"],
    "affs": ["作者单位", "单位", "机构", "作者机构", "author affiliation", "affiliation"],
    "journal": ["来源", "期刊", "文献来源", "刊名", "source", "journal"],
    "year": ["年", "年份", "发表时间", "出版年", "year"],
    "cited": ["被引", "被引频次", "被引次数", "cited"],
    "doi": ["doi", "DOI"],
    "category": ["来源类别", "收录类别", "数据库", "category"],
    "first": ["第一作者", "第一责任人"],
}
CAT_LEVEL = {  # 文件名/来源类别关键词 → 层级
    "北大核心": "中文核心", "cscd": "CSCD", "cssci": "CSSCI",
    "ei": "EI(中文)", "普刊": "中文普刊", "其他": "中文普刊",
}
LEVEL_ORDER = {"1区(Top)": 1, "1区": 1, "2区(Top)": 2, "2区": 2, "3区": 3, "4区": 4,
               "EI(中文)": 5, "CSCD": 5, "中文核心": 6, "CSSCI": 6, "中文普刊": 7,
               "非SCI期刊": 7, "SCI收录(未入分区表)": 5, "SCI/ESCI收录(未入分区表)": 6,
               "CCF A类会议": 8, "CCF B类会议": 8, "CCF C类会议": 8, "会议论文": 8,
               "预印本": 9, "其他(书章/社论等)": 9}


def pick(row, key):
    for k in ALIAS[key]:
        for rk in row:
            if str(rk).strip().lower() == k.lower():
                v = row[rk]
                return "" if v is None else str(v).strip()
    return ""


def read_table(p):
    if p.suffix.lower() in (".xlsx", ".xls"):
        import pandas as pd
        df = pd.read_csv if False else pd.read_excel(p, dtype=str)
        return df.fillna("").to_dict("records")
    for enc in ("utf-8-sig", "gb18030", "utf-8"):
        try:
            with open(p, encoding=enc, newline="") as f:
                return list(csv.DictReader(f))
        except UnicodeDecodeError:
            continue
    return []


def level_from_name(p, row):
    stem = p.stem.lower()
    for kw, lv in CAT_LEVEL.items():
        if kw in stem: return lv
    cat = pick(row, "category")
    for kw, lv in CAT_LEVEL.items():
        if kw in cat.lower(): return lv
    if "sci" in cat.lower(): return "SCI/ESCI收录(未入分区表)"
    return "中文普刊"


def norm_title(t):
    return re.sub(r"[\s\W_]+", "", str(t)).lower()


def tag_direction_cn(title, dirs):
    """中文标题用 directions.json 里的 cn_keywords 匹配。"""
    t = str(title)
    best, score = None, 0
    for d, pats in (dirs.get("cn_keywords") or {}).items():
        s = sum(1 for p in pats if p in t)
        if s > score: best, score = d, s
    return best or dirs.get("fallback", "其他/交叉")


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--records", required=True, help="college_sci.py 产出的 .records.json")
    ap.add_argument("--cnki-dir", required=True, help="知网导出文件所在目录")
    ap.add_argument("--config", required=True)
    ap.add_argument("--out", required=True)
    ap.add_argument("--cache", default="./.college_sci_cache")
    a = ap.parse_args()

    cfg = json.loads(Path(a.config).read_text(encoding="utf-8"))
    dirs = json.loads(Path(cfg["directions"]).read_text(encoding="utf-8")) if Path(cfg["directions"]).exists() \
        else json.loads((Path(a.config).parent / "directions.json").read_text(encoding="utf-8"))
    base = json.loads(Path(a.records).read_text(encoding="utf-8"))
    seen_doi = {(r.get("doi") or "").lower() for r in base if r.get("doi")}
    seen_title = {norm_title(r["title"]) for r in base}

    fac_cn = {f[0] for f in cfg["faculty"]}
    college_cn = cfg["college"].get("display") or ""
    cn_list = cfg["college"].get("cn") or []
    if isinstance(cn_list, str): cn_list = [cn_list]

    files = sorted([p for p in Path(a.cnki_dir).iterdir()
                    if p.suffix.lower() in (".csv", ".xlsx", ".xls")])
    if not files:
        sys.exit(f"目录里没有 csv/xlsx：{a.cnki_dir}")

    added, dup, skipped = [], 0, 0
    for p in files:
        rows = read_table(p)
        print(f"  {p.name}: {len(rows)} 行")
        for row in rows:
            title = pick(row, "title")
            if not title: skipped += 1; continue
            doi = pick(row, "doi").lower().replace("https://doi.org/", "")
            nt = norm_title(title)
            if (doi and doi in seen_doi) or nt in seen_title:
                dup += 1; continue
            afs = pick(row, "affs")
            first = pick(row, "first") or (pick(row, "authors").split(";")[0] if pick(row, "authors") else "")
            first = re.sub(r"^\s*\d+\s*", "", first).strip()
            is_college = any(c in afs for c in cn_list + [college_cn]) if college_cn else False
            ident = "教师一作" if first in fac_cn else ("学生(推定)" if is_college else "外单位一作")
            scope = "我院主导" if (is_college or ident == "教师一作") else "我院参与"
            lv = level_from_name(p, row)
            y = re.search(r"(19|20)\d{2}", pick(row, "year") or "")
            year = int(y.group(0)) if y else None
            cited = re.search(r"\d+", pick(row, "cited") or "")
            authors = [{"name": x.strip(), "affs": [afs if i == 0 else ""], "corr": i == 0}
                       for i, x in enumerate(re.split(r"[;；]", pick(row, "authors") or first)) if x.strip()]
            if not authors: authors = [{"name": first, "affs": [afs], "corr": True}]
            added.append({
                "title": title, "year": year, "cited": int(cited.group(0)) if cited else 0,
                "doi": doi, "journal": pick(row, "journal"), "type": "article",
                "authors": authors, "ident": ident, "scope": scope,
                "fu": "本院" if is_college else "外单位", "facs": first if first in fac_cn else "",
                "level": lv, "ord": LEVEL_ORDER.get(lv, 9),
                "dir": tag_direction_cn(title, dirs),
                "src": {"期刊": pick(row, "journal"), "ISSN": "", "中科院大类": None, "中科院分区": None,
                        "Top": None, "IF": None, "JCR分区": None, "新锐分区": None, "CCF": None,
                        "WOS": None, "中文刊名": "", "来源类别": lv},
                "source_db": "CNKI"})
            seen_title.add(nt)
            if doi: seen_doi.add(doi)
    print(f"  新增 {len(added)} 篇｜重复跳过 {dup} 篇｜无题名跳过 {skipped} 行")

    if not added:
        sys.exit("没有可新增的中文记录，检查导出文件或检索式")

    # 合并 + 重出工作簿
    merged = base + added
    sys.path.insert(0, str(Path(__file__).resolve().parent))
    from workbook import write_workbook
    out = []
    for r in merged:
        work = {"title": r["title"], "publication_year": r["year"], "cited_by_count": r["cited"],
                "doi": r["doi"], "type": r["type"],
                "primary_location": {"source": {"display_name": r["journal"], "issn_l": r["src"].get("ISSN")}} if r["journal"] else {},
                "authorships": [{"author": {"display_name": x["name"]},
                                 "raw_affiliation_strings": x.get("affs") or [],
                                 "is_corresponding": x.get("corr", False),
                                 "institutions": []} for x in r["authors"]]}
        out.append({"work": work, "fu": r["fu"], "facs": r["facs"], "corr": any(x.get("corr") for x in r["authors"]),
                    "ident": r["ident"], "scope": r["scope"], "src": r["src"],
                    "level": r["level"], "ord": r["ord"], "dir": r["dir"]})
    summary = write_workbook(out, cfg, a.out, dropped=0)
    summary["cnki_added"] = len(added)
    summary["cnki_duplicates"] = dup
    summary["cnki_files"] = [p.name for p in files]
    Path(a.out).with_suffix(".summary.json").write_text(
        json.dumps(summary, ensure_ascii=False, indent=2), encoding="utf-8")
    print(f"\n完成 → {a.out}")
    print(f"  主导 {summary['core']} 篇｜参与 {summary['total']} 篇"
          f"（其中中文期刊 {len(added)} 篇）")


if __name__ == "__main__":
    main()
