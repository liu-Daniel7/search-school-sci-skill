#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
wizard.py — 交互式生成 config.json（新手入口）

用法：
  python wizard.py                 # 一步一步问答
  python wizard.py --out cfg.json  # 指定输出
  python wizard.py --answers answers.json   # 批量/非交互

生成后会自动检查必填项，并打印"下一步做什么"。
"""
import argparse, json, sys
from pathlib import Path

HERE = Path(__file__).resolve().parent
REF = HERE.parent / "reference"


def ask(prompt, default="", required=False, hint=""):
    while True:
        s = f"{prompt}"
        if default: s += f" [{default}]"
        s += "："
        if hint: print(f"    ⓘ {hint}")
        v = input(s).strip()
        if not v: v = default
        if v or not required: return v
        print("    ✗ 这一项必填，请重新输入")


def ask_list(prompt, hint=""):
    print(f"{prompt}（每行一个，空行结束）")
    if hint: print(f"    ⓘ {hint}")
    out = []
    while True:
        v = input("    > ").strip()
        if not v: break
        out.append(v)
    return out


def load_faculty_csv(p):
    import csv
    rows = []
    with open(p, encoding="utf-8-sig") as f:
        for r in csv.DictReader(f):
            cn = (r.get("中文名") or "").strip()
            py = (r.get("全拼") or "").strip().lower()
            tt = (r.get("职称") or "").strip()
            if cn: rows.append([cn, py, tt])
    return rows


def auto_faculty(school_cn, college_cn, verbose=False):
    """调用 faculty_fetch.py 自动抓学院官网教师名单。"""
    import subprocess, tempfile
    here = Path(__file__).resolve().parent
    tmp = Path(tempfile.mkdtemp()) / "faculty.csv"
    cmd = [sys.executable, str(here / "faculty_fetch.py"),
           "--school", school_cn, "--college", college_cn, "--out", str(tmp)]
    r = subprocess.run(cmd, capture_output=True, text=True)
    if verbose: print(r.stdout[-1200:] + r.stderr[-600:])
    if not tmp.exists():
        print("  ⚠ 没能自动抓到教师名单。请 agent 用 web_search 找到学院官网后执行：\n"
              f"     python faculty_fetch.py --url <学院官网> --out faculty.csv\n"
              f"     再用 --faculty-csv faculty.csv 重新生成配置")
        return []
    rows = load_faculty_csv(tmp)
    print(f"  ✓ 自动获取 {len(rows)} 人（拼音 {sum(1 for r in rows if r[1])} 人）")
    if sum(1 for r in rows if r[1]) < len(rows) * 0.5:
        print("  ⚠ 拼音覆盖率偏低，建议加 --deep 重抓，或由 agent 补齐拼音")
    return rows


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--out", default="config.json")
    ap.add_argument("--answers", help="answers.json（非交互）")
    ap.add_argument("--faculty-csv", help="教师名单 CSV（中文名,全拼,职称）")
    ap.add_argument("--no-auto-faculty", action="store_true", help="不自动抓教师名单")
    ap.add_argument("--verbose", action="store_true")
    a = ap.parse_args()

    if a.answers:
        ans = json.loads(Path(a.answers).read_text(encoding="utf-8"))
    else:
        print("=" * 62)
        print("  学院科研产出统计 · 配置向导")
        print("  全程约 5 分钟，只有 6 个问题")
        print("=" * 62)
        ans = {}
        ans["school_cn"] = ask("1/4 学校中文名", required=True, hint="例：示例大学")
        ans["school_en"] = ask("2/4 学校英文名（OpenAlex 用它定位学校，可在 OpenAlex/ROR 搜到）", required=True,
                               hint="例：Shandong University")
        ans["college_cn"] = ask("3/4 学院中文全称（简称也行，但要能唯一确定一个学院）", required=True,
                                hint="例：机械工程学院。若学校有同名学院（如分校区），请写全")
        ans["college_en"] = ask("4/4 学院英文名（可留空，跑普查后会提示补）", default="",
                                hint="例：School of Mechanical Engineering")
        ans["college_en_more"] = []
        ans["lab"] = []
        ans["other"] = []
        ans["exclude"] = []
        ans["faculty"] = []
        ans["out_xlsx"] = "学院科研产出.xlsx"
        print("\n（教师名单、实验室、他院关键词都不再问你——由 agent 自动获取）")

    school_en = ans["school_en"].strip()
    college_cn = ans["college_cn"].strip()
    college_en = [ans["college_en"].strip()] + [x for x in ans.get("college_en_more", []) if x]
    cn_list = [college_cn] + [x for x in ans.get("college_cn_more", []) if x]
    # ⚠ 校名匹配只能含“学校”名称，绝不能混入学院名，否则他校同名学院会被当成本校
    school_alts = {school_en, school_en.replace(" and ", " & ")}
    if ans.get("school_cn"): school_alts.add(ans["school_cn"])
    match = ans.get("match_regex") or "|".join(sorted(school_alts))

    cfg = {
        "institution": {"name": school_en, "ror": ans.get("ror", ""), "match_regex": match},
        "college": {"display": college_cn, "cn": cn_list, "en": college_en,
                    "extra_patterns": ans.get("extra_patterns", [])},
        "lab_patterns": ans.get("lab", []),
        "other_school_patterns": ans.get("other", []),
        "affiliation_keywords": ans.get("affiliation_keywords", ["Mechanical"]),
        "expand_by_faculty_authors": True,
        "exclude_faculty": ans.get("exclude", []),
        "verified_names": {},
        "treat_university_only_as_college": False,
        "unfiltered_max_hits": 2000,
        "request_delay": 1.0,
        "faculty": ans.get("faculty", []),
        "directions": str(REF / "directions.json"),
        "as_of": ans.get("as_of", ""),
    }
    # 教师名单：优先用 --faculty-csv，其次尝试自动抓取学院官网
    if a.faculty_csv and Path(a.faculty_csv).exists():
        cfg["faculty"] = load_faculty_csv(a.faculty_csv)
        print(f"\n教师名单 ← {a.faculty_csv}（{len(cfg['faculty'])} 人）")
    elif not cfg["faculty"]:
        print("\n正在自动获取教师名单 ……")
        cfg["faculty"] = auto_faculty(school_cn, college_cn, a.verbose)

    out = Path(a.out)
    out.write_text(json.dumps(cfg, ensure_ascii=False, indent=2), encoding="utf-8")

    # ---- 自检 ----
    warns = []
    if len(cfg["faculty"]) < 5:
        warns.append("教师名单少于 5 人 —— 几乎肯定会把教师误判成学生，请补全")
    if not ans.get("school_cn"):
        warns.append("没填学校中文名 —— 中文单位串的论文会漏检，建议补上")
    if not cfg["institution"]["ror"]:
        warns.append("没填 ROR —— 跑的时候会自动匹配，但同名学校有风险，建议补上")
    if len(college_en) < 2:
        warns.append("学院英文名只写了 1 种 —— 建议跑完第 2 步“单位串普查”后，把发现的变体补进 college.en")
    print("\n" + "=" * 62)
    print(f"✅ 已写入 {out}")
    print(f"   学校：{school_en}")
    print(f"   学院：{college_cn}（英文 {len(college_en)} 种写法）")
    print(f"   教师：{len(cfg['faculty'])} 人")
    if warns:
        print("\n⚠ 需要你确认：")
        for w in warns: print("   -", w)
    print(f"""
下一步（按顺序）：
  1) 先只跑普查，确认学院名没混入别的学校：
       python college_sci.py --config {out} --stage census
  2) 再跑全流程（首次 20-30 分钟）：
       python college_sci.py --config {out} --out {ans.get('out_xlsx','学院科研产出.xlsx')}
  3) 打开工作簿，先看“① 看板”顶部的警示条和“② 论文清单”抽查 3-5 篇
  4) 若普查发现新写法，补进 college.en 后重跑（有缓存，1-2 分钟）
""")


if __name__ == "__main__":
    main()
