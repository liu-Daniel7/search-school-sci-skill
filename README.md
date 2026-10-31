# search-school-sci-skill

> 查任意大学任意学院的科研产出，总结成可视化 excel。

给一个"学校 + 学院"，它就还你一张能直接交材料的表。

```bash
python scripts/college_sci.py --config config.json --out 学院科研产出.xlsx
```

![看板示意（示例数据，非真实学院）](docs/preview.png?v=2)

---

## 它解决什么问题

想知道"我们学院到底什么水平"，现实是这样的：

| 你遇到的问题 | 这个工具的做法 |
|---|---|
| 知网、Web of Science 都**没有"学院"这一级机构字段** | 用作者单位字符串匹配 + 教师姓名反查来还原学院 |
| 同一个学院名，**别的学校也在用**（实测多算了 122 篇） | 强制加校名限定，并在普查阶段列出"撞名"的学校 |
| 一个学院有**十几种英文写法**（连字符、Electrical/Electronic、Automobile/Vehicle、拼错） | 自动做"写法普查"，把变体全部纳入 |
| 大量论文**只署实验室、不署学院** | 支持独立配置实验室/中心名称通道 |
| **28% 的论文只写"某某大学"**，不写学院 | 用教师 OpenAlex 作者ID 反查扩展 |
| 分不清**教师和学生** | 与学院师资名单比对，学生身份标注为"推定" |
| 中文期刊查不到 | 提供知网补数流程（人登录、agent 执行）+ 合并脚本 |

---

## 输出长什么样

一个 7 个工作表的工作簿：

| 工作表 | 内容 |
|---|---|
| **① 看板** | 核心数字 + 学科方向×期刊层级 + 方向×作者身份 + 图例 + 柱状图，一屏汇报 |
| **② 论文清单** | 全量记录，层级列按颜色区分，全列可筛选 |
| **③ 姓名对照表** | 所有第一作者 + 姓氏参考，用于补齐中文名 |
| **④ 教师-方向矩阵** | 教师×方向的产出热力图 |
| **⑤ 教师产出** | 逐人分区分布、一作/通讯、被引 |
| **⑥ 学生(推定)** | 学生一作清单 |
| **⑦ 说明** | 口径定义、清洗记录、局限声明 |

期刊层级用**中科院《期刊分区表》**（1–4 区 + Top 标注），会议按 **CCF 推荐目录**标 A/B/C，中文期刊按知网"来源类别"分为 中文核心 / CSSCI / CSCD / EI(中文) / 中文普刊。每个层级一种颜色，看板内附色卡。

---

## 安装到 DSH

本仓库同时是一个 DSH 插件包（`package.json` 声明了 `dsh.bundle`），安装后技能会自动出现在技能目录里：

```bash
dsh plugin add liu-Daniel7/search-school-sci-skill
```

装完直接说「用 search-school-sci 查某大学某学院的科研产出」即可触发。

插件本体是 `lib/index.js`——一个不到 70 行的 skill provider，把仓库根目录的 `SKILL.md` 注册为技能正文，并把仓库根声明为 `resourceBase`，因此 `reference/*.md` 与 `scripts/*.py` 的相对路径都能正常解析。Python 脚本仍是实际的执行体。

## 快速开始

### 环境

需要 Python 3.9+ 与 `pandas`、`openpyxl`：

```bash
pip install -r requirements.txt
```

### 三步

**① 生成配置**（只问 4 个问题，不用交教师名单）

```bash
python scripts/wizard.py --out config.json
```

**② 抓教师名单**（agent 用 web_search 找到学院官网后执行，或让脚本自己找）

```bash
python scripts/faculty_fetch.py --url https://你的学院官网/szdw/ --out faculty.csv --deep
```

**③ 先普查，再全跑**

```bash
python scripts/college_sci.py --config config.json --faculty-csv faculty.csv --stage census
python scripts/college_sci.py --config config.json --faculty-csv faculty.csv --out 学院科研产出.xlsx
```

`--stage census` 只做"机构解析 + 单位串普查 + 同名学院检测"，2–3 分钟，不抓论文。**这一步不能跳过**，它决定你的检索式有没有混进别的学校。

### 补中文期刊（可选但强烈建议）

中文期刊 OpenAlex 基本不收，需要知网补数。流程见 [`reference/CNKI_GUIDE.md`](reference/CNKI_GUIDE.md)：人登录 → 按来源类别分轮导出 → 脚本合并。

```bash
python scripts/cnki_import.py --records 结果.records.json \
  --cnki-dir ./cnki_export --config config.json --out 结果_含中文.xlsx
```

---

## 目录结构

```
.
├── SKILL.md                    # 技能说明（给 agent 读的行为规范）
├── README.md                   # 本文件
├── reference/
│   ├── WALKTHROUGH.md          # 新手逐步向导（含示例大学实例）
│   ├── CNKI_GUIDE.md           # 知网补数全流程
│   ├── pitfalls.md             # 六个真实踩坑案例（附数字证据）
│   ├── config.example.json     # 配置模板
│   ├── config.demo.json        # 完整实例（虚构的示例大学·智能制造学院）
│   ├── directions.json         # 学科方向关键词（中英双语，可改）
│   └── faculty.example.csv     # 教师名单格式
├── scripts/
│   ├── college_sci.py          # 主流水线：抓取→普查→归属→分级→出表
│   ├── workbook.py             # 工作簿生成（看板/配色/图例/筛选）
│   ├── faculty_fetch.py        # 学院官网教师名单抓取（含拼音）
│   ├── cnki_import.py          # 知网导出合并
│   ├── wizard.py               # 交互式配置向导
│   └── run.sh                  # 自动挑选解释器
└── tests/smoke_test.py         # 离线冒烟测试（不联网）
```

---

## 口径说明（很重要）

**两种统计口径，报告里请分开引用：**

- **本院主导** —— 第一作者或通讯作者单位为本院。对外报材料用这个数。
- **本院参与** —— 含外单位/校内其他院系主导、本院教师参与的论文。

**四条必须写进报告的限定语：**

1. 学生身份为**推定**（第一作者不在在编教师名单内），可能是博士后、未公开的兼职教授或新聘讲师，未经学籍/人事核对不得作为学生成果使用
2. 分区按**中科院分区表某年版**口径，与论文发表当年可能不同
3. **不含知网独有中文期刊**（除非已跑 `cnki_import.py`）
4. 被引数为 OpenAlex 口径，与 Google Scholar / WoS 不可直接比较

---

## 常见问题

**Q：跑出来篇数比学院公布的少一半？**
A：八成是单位串写法没覆盖全。看 `--stage census` 的输出，把漏掉的写法补进配置的 `college.en` 或 `extra_patterns`。

**Q：教师被算成了学生？**
A：教师名单的拼音不对或缺失。拼音是匹配英文署名的唯一钥匙，用 `faculty_fetch.py --deep` 重抓，或手工补 `faculty.csv` 第二列。

**Q：请求返回 429？**
A：OpenAlex 免费额度按 IP 计（约 $0.10/天，UTC 零点重置＝北京时间早 8:00），一个学院规模的项目通常一次就用完。申请免费 key 加到 URL，或分两天跑（响应有缓存，会断点续跑）。

**Q：能查整个学校吗？**
A：能，但"学院"是这套方法的最小区分粒度。查全校直接按 `institutions.id` 取数即可，不需要本工具的单位串匹配。

---

## 验证

```bash
python tests/smoke_test.py
```

离线冒烟测试，不联网，验证「记录 → 7 个工作表」链路、层级配色分组、中文署名处理。

方法有效性校验（某高校机电类二级学院，2026-10 手工核对与脚本结果对照，机构与个人均已匿名）：

| 环节 | 结果 |
|---|---|
| 候选论文 | 1052 篇（剔除撤稿 4 篇、兼职教授单独关联 25 篇、同名学院误命中 122 篇以上） |
| 本院主导 | 749 篇，被引 7386 |
| 一区 Top | 60 篇 |
| 交叉验证 | 脚本判定的某学生「10 篇一作」与其公开获奖报道「一区 TOP 10 篇」吻合，说明身份推定口径可用 |

---

## 已知边界

- **中文期刊**依赖知网，需要机构账号；本工具只做合并，不做登录
- **学生身份**无法自证，只能标"推定"
- **WoS/Scopus 口径**需要机构订阅，本工具不接入
- **学科方向标签**是关键词自动判定，交叉论文可能需人工调整
- 搜索引擎自动定位官网不稳（会被反爬），建议 agent 用自带搜索能力拿链接再传 `--url`

---

## License

MIT
