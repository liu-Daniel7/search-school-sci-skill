# 新手向导：以"示例大学机电学院"为例

> 读完这一页你就能跑出结果。全程 6 个问题、4 条命令、约 30 分钟（首次）。
> 每步都写了**要做什么 / 要问用户什么 / 怎么验证**。

---

## 第 0 步：先确认"学院到底是哪个"（最容易出错的一步）

"示例大学机电学院"不是官方名称，示例大学至少有两个候选：

| 候选 | 全称 | 官网 | 英文名 |
|---|---|---|---|
| A | 机械工程学院 | example.edu.cn/mech | School of Mechanical Engineering |
| B | 机电与信息工程学院（**分校区**） | 属示例大学（威海） | School of Mechanical, Electrical and Information Engineering |

**A 和 B 是不同机构、不同数据源，混在一起算就全错了。**

> 要问用户的话：
> "你说的机电学院，是济南本部的**机械工程学院**，还是分校区的**机电与信息工程学院**？"
> 如果对方不确定，让他把学院官网链接发过来——以官网页脚的学院全称为准。

山大本部机械工程学院的英文名可从其官网英文版确认：http://example.edu.cn/mech/ENGLISH.htm

---

## 第 1 步：生成配置（只答 4 个问题，不用交名单）

```bash
python scripts/wizard.py --out sdut_mech.json
```

只问 4 个问题，**教师名单不用你贴**：

1. 学校中文名 → `示例大学`
2. 学校英文名 → `Shandong University`
3. 学院中文全称 → `机械工程学院`
4. 学院英文名 → 可留空

之后脚本自动去抓教师名单。抓不到就由 agent 补一步：

```bash
# agent 用 web_search 找到学院官网后
python scripts/faculty_fetch.py --url http://example.edu.cn/mech/ --out faculty.csv --deep
python scripts/wizard.py --faculty-csv faculty.csv --out sdut_mech.json
```

`--deep` 会逐人翻详情页找 `faculty.example.edu.cn/<拼音>/` 这类个人主页链接来补拼音。拼音是匹配英文署名的钥匙，覆盖率低会导致教师被误判成学生。

---

## 第 2 步：先只跑普查（2–3 分钟，不抓论文）

```bash
python scripts/college_sci.py --config sdut_mech.json --stage census
```

输出长这样（下面是北建大的真实结果）：

```
School of Mechanical-Electronic and Vehicle Engineering: 全球 563 篇 / 本校 449 篇（差 114 为他校同名）
⚠ 检测到同名院系：
  - Central South University: 74 篇
  - Weifang University: 57 篇
```

> **怎么判断这一关过了？**
> - 每条"全球 / 本校"的差值都在合理范围（同名院系越常见，差值越大，正常）
> - 差值为 0 反而要警惕：可能是校名限定写错，导致本校也匹配不上
> - 若"他校"列表里出现同城的对标学校，说明该院系名很常见，`college.en` 要写得更精确

**山大要特别注意**：`School of Mechanical Engineering` 是全国最常见的学院名之一，同名学校可能有上百所。**必须**保证 `institution.match_regex` 里有 `Shandong University`，否则会把清华、上交、哈工大的机械学院一起算进来。

---

## 第 3 步：补全写法（关键，决定召回率）

打开普查日志，把发现的变体补进 `sdut_mech.json` 的 `college.en`：

```json
"en": [
  "School of Mechanical Engineering",
  "School of Mechanical and Electrical Engineering",
  "School of Mechanical Engineering, Shandong University"
]
```

再补两类容易漏的：

| 字段 | 补什么 | 为什么要补 |
|---|---|---|
| `lab_patterns` | 学院的国家级/省部级实验室英文名 | 大量论文只署实验室，不写学院 |
| `other_school_patterns` | 山大其他学院英文名（控制、材料、能源、土建…） | 避免跨院合作被误判成本院 |

山大机械工程学院有"高效洁净机械制造教育部重点实验室"等平台，其英文名要作为 `lab_patterns` 加进去。

---

## 第 4 步：跑全流程（首次 20–30 分钟，之后 1–2 分钟）

```bash
python scripts/college_sci.py --config sdut_mech.json --out 山大机械学院科研产出.xlsx
```

日志会走 7 步，重点看三行：

```
[2/7] 单位串普查 …                 ← 应该和上次普查一致
[4/7] 教师作者ID N 个 → 扩展后 M 篇  ← N 太少说明教师名单有漏
[6/7] 入库 K 篇（剔除 X：撤稿/排除名单）
```

> **怎么判断这一关过了？**
> K 的量级应该和学院/图书馆的既有数字接近。差 2 倍以上 → 回第 3 步补写法。

---

## 第 5 步：检查结果（不能跳过）

打开工作簿，按这个顺序看：

1. **① 看板顶部黄色警示条** —— 确认"学生(推定)"的限定语还在，别把推定数据当正式数据用
2. **② 论文清单** —— 随机抽 3–5 篇，点开 DOI，确认第一作者确实属于该学院
3. **③ 姓名对照表** —— 看有多少"待补"中文名，决定是否找学院要学籍名单
4. **⑦ 说明** —— 把口径和局限一起读一遍

---

## 第 5.5 步：补中文期刊（不可跳过）

OpenAlex 不收中文期刊。学院真实产出里中文权威期刊常占三分之一以上，必须另补：

1. 让用户登录知网（校园网 IP / CARSI / 漫游账号），确认右上角显示机构名
2. 按 `reference/CNKI_GUIDE.md` 跑专业检索，按**来源类别**分轮导出到 `cnki_export/`
3. 合并：

```bash
python scripts/cnki_import.py --records 结果.records.json \
  --cnki-dir ./cnki_export --config config.json --out 结果_含中文.xlsx
```

补完后看板表一会多出「中文核心/EI/CSCD」「中文普刊」两列（蓝色系）。

## 第 6 步：出报告

报材料时只引用两个数：

- **本院主导（core）** —— 第一作者或通讯作者为本院，这是能对外说的数字
- **本院参与（total）** —— 含合作，用于说明影响力

其余一律加限定语：
- "按中科院分区表 2025 版口径"
- "学生身份为推定，未经学籍核对"
- "不含知网独有中文期刊"

---

## 常见卡点

| 卡在哪 | 现象 | 怎么办 |
|---|---|---|
| 第 2 步 | 本校命中数为 0 | `match_regex` 里没有校名，或校名写法不符 |
| 第 4 步 | 教师作者ID 只有个位数 | 教师名单的拼音写错了，检查是否全拼、无空格 |
| 第 4 步 | 请求 429 | OpenAlex 免费额度用尽（每 IP 每天 $0.10，UTC 零点重置＝北京时间早 8:00）；调大 `request_delay`，或申请免费 API key |
| 第 5 步 | 期刊全是"非SCI期刊" | 指标文件没下载成功，删 `.cache/*.csv` 重跑 |
| 结果 | 篇数比预期少很多 | 中文期刊不在 OpenAlex，需另用知网补 |

---

## 一句话总结

**先定学院全称 → 交互生成配置 → 只跑普查看同名 → 补全写法 → 跑全流程 → 抽检 → 只报"主导"数。**
