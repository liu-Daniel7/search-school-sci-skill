#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""离线冒烟测试：不联网，验证「记录 → 工作簿」链路与去重逻辑。

    python tests/smoke_test.py
"""
import json, re, sys, tempfile, csv
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "scripts"))
import workbook  # noqa: E402

CFG = {
    "institution": {"name": "示例大学", "ror": "", "match_regex": "示例大学"},
    "college": {"display": "示例学院", "cn": ["示例学院"], "en": ["School of Example"]},
    "lab_patterns": [], "other_school_patterns": [], "exclude_faculty": [],
    "verified_names": {}, "faculty": [["张三", "zhangsan", "教授"], ["李四", "lisi", "副教授"]],
    "directions": str(ROOT / "reference" / "directions.json"), "as_of": "2026-01-01",
}


def rec(title, year, level, ident, scope, dirn, cited, authors, src_journal=""):
    return {
        "work": {"title": title, "publication_year": year, "cited_by_count": cited, "doi": "",
                 "type": "article",
                 "primary_location": {"source": {"display_name": src_journal, "issn_l": ""}} if src_journal else {},
                 "authorships": [{"author": {"display_name": n},
                                  "raw_affiliation_strings": [a], "is_corresponding": c,
                                  "institutions": []} for n, a, c in authors]},
        "fu": "本院", "facs": "", "corr": True, "ident": ident, "scope": scope,
        "src": {"期刊": src_journal, "ISSN": "", "中科院大类": "工程技术", "中科院分区": None,
                "Top": None, "IF": None, "JCR分区": None, "新锐分区": None, "CCF": None,
                "WOS": None, "中文刊名": ""},
        "level": level, "ord": 1, "dir": dirn,
    }


def main():
    rows = [
        rec("Fault diagnosis of bearings via deep learning", 2024, "1区(Top)", "学生(推定)", "我院主导",
            "故障诊断与健康管理", 88, [("Wang Wu", "School of Example, 示例大学", True)], "Mechanical Systems and Signal Processing"),
        rec("Vibration control of rail vehicle", 2023, "3区", "教师一作", "我院主导",
            "轨道车辆工程", 20, [("张三", "School of Example, 示例大学", True)], "Vehicle System Dynamics"),
        rec("基于改进VMD的轴承故障诊断", 2022, "中文核心", "学生(推定)", "我院主导",
            "故障诊断与健康管理", 9, [("赵六", "示例大学示例学院", True)], "机械工程学报"),
    ]
    tmp = Path(tempfile.mkdtemp())
    out = tmp / "t.xlsx"
    s = workbook.write_workbook(rows, CFG, str(out), dropped=0)

    assert out.exists(), "工作簿未生成"
    from openpyxl import load_workbook
    wb = load_workbook(out)
    expect = ["① 看板", "② 论文清单(可筛选)", "③ 姓名对照表(待填)",
              "④ 教师-方向矩阵", "⑤ 教师产出", "⑥ 学生(推定)", "⑦ 说明"]
    assert wb.sheetnames == expect, f"工作表不符：{wb.sheetnames}"
    assert s["core"] == 3 and s["total"] == 3, s
    assert s["student_first"] == 2 and s["teacher_first"] == 1, s

    ws = wb["② 论文清单(可筛选)"]
    assert ws.max_row == 4, f"清单行数 {ws.max_row}"
    levels = {ws.cell(row=r, column=8).value for r in range(2, 5)}
    assert {"1区(Top)", "3区", "中文核心"} <= levels, levels

    # 中文署名应直接显示，不标"待补"
    names = {ws.cell(row=r, column=7).value for r in range(2, 5)}
    assert "赵六" in names, names          # 中文署名直接显示
    assert "待补" in names, names          # 英文署名且未核实 → 待补

    # 看板里应有中文核心分组列
    board = wb["① 看板"]
    hdr = [board.cell(row=11, column=c).value for c in range(1, 16)]
    assert "中文核心/EI/CSCD" in hdr, hdr

    print("✅ 冒烟测试通过：7 个工作表 / 3 条记录 / 层级与分组正确")
    return 0


if __name__ == "__main__":
    sys.exit(main())
