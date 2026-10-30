# -*- coding: utf-8 -*-
"""docs/preview.png 的生成器：全部为虚构数据，不含任何真实学校 / 学院 / 人名。

用途：README 首屏配图。任何人改版式后可用它重新生成示意图：

    python docs/gen_preview.py
    # 再用 LibreOffice 渲染 ① 看板 A1:N43 区域导出 PNG

注意：本文件只用于产出示例图，不要接真实数据源。
"""
import sys, random, itertools
from pathlib import Path
ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "scripts"))
import workbook

random.seed(7)
CFG = {
 "institution": {"name": "EXAMPLE UNIVERSITY", "ror": "", "match_regex": "EXAMPLE"},
 "college": {"display": "机械工程学院", "cn": ["机械工程学院"], "en": ["School of Mechanical Engineering"]},
 "lab_patterns": [], "other_school_patterns": [], "exclude_faculty": [], "verified_names": {},
 "faculty": [["张伟","zhangwei","教授"],["王芳","wangfang","教授"],["李强","liqiang","副教授"],
             ["刘洋","liuyang","副教授"],["陈静","chenjing","讲师"],["杨帆","yangfan","教授"],
             ["赵敏","zhaomin","副教授"],["周涛","zhoutao","讲师"]],
 "directions": str(ROOT / "reference" / "directions.json"), "as_of": "2025-12-31",
}
DIRS = ["故障诊断与健康管理", "智能制造与机器人", "工业工程与运筹优化", "机电液控制与检测", "能源与动力工程"]
W = [0.34, 0.26, 0.16, 0.13, 0.11]
LV = [("1区(Top)",34),("2区(Top)",11),("2区",38),("3区",76),("4区",121),
      ("非SCI期刊",52),("会议论文",98),("中文核心",34),("EI(中文)",10),("中文普刊",12)]
LV2 = [("会议论文",60),("4区",40),("3区",34),("非SCI期刊",22),("2区",16),("中文普刊",14)]
FAC = ["张伟","王芳","李强","刘洋","陈静","杨帆","赵敏","周涛"]

def mk(i, level, dirn, ident, scope, cited):
    if ident == "教师一作": name, aff = random.choice(FAC), "School of Intelligent Manufacturing, EXAMPLE UNIVERSITY"
    elif ident == "学生(推定)": name, aff = random.choice(["Sun Lei","Han Yu","Xu Ming","Guo Rui","He Xin","Ma Lin"]), "School of Intelligent Manufacturing, EXAMPLE UNIVERSITY"
    elif ident == "校内他院一作": name, aff = "Wu Chen", "School of Automation, EXAMPLE UNIVERSITY"
    else: name, aff = "D. Cooper", "Department of Engineering, Example Institute of Technology"
    return {
      "work": {"title": f"示例论文 {i}: {dirn} related study on intelligent manufacturing",
               "publication_year": random.choice([2020,2021,2022,2023,2024,2025]),
               "cited_by_count": cited, "doi": "", "type": "article",
               "primary_location": {"source": {"display_name": "Example Journal", "issn_l": ""}},
               "authorships": [{"author": {"display_name": name},
                                "raw_affiliation_strings": [aff], "is_corresponding": True,
                                "institutions": []}]},
      "fu": "本院", "facs": "", "corr": True, "ident": ident, "scope": scope,
      "src": {"期刊": "Example Journal", "ISSN": "", "中科院大类": "工程技术", "中科院分区": None,
              "Top": None, "IF": None, "JCR分区": None, "新锐分区": None, "CCF": None,
              "WOS": None, "中文刊名": ""},
      "level": level, "ord": 1, "dir": dirn}

rows, i = [], 0
def pick_dir():
    return random.choices(DIRS, weights=W, k=1)[0]

for level, n in LV:
    for _ in range(n):
        i += 1
        ident = random.choices(["教师一作","学生(推定)"], weights=[0.52,0.48], k=1)[0]
        rows.append(mk(i, level, pick_dir(), ident, "我院主导", max(0, int(random.gauss(18, 22)))))

for level, n in LV2:
    for _ in range(n):
        i += 1
        ident = random.choices(["校内他院一作","外单位一作"], weights=[0.31,0.69], k=1)[0]
        rows.append(mk(i, level, pick_dir(), ident, "我院参与", max(0, int(random.gauss(9, 14)))))

out = Path(__file__).resolve().parent / "_preview_demo.xlsx"
s = workbook.write_workbook(rows, CFG, str(out), dropped=0)
print("主导", s["core"], "参与", s["total"], "教师", s["teacher_first"], "学生", s["student_first"])
print("1区Top", sum(1 for r in rows if r["scope"]=="我院主导" and r["level"].startswith("1区")))
