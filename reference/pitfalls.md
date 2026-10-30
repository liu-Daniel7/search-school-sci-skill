# 踩坑记录（每个都附真实数字）

以下数字来自 2026-10 对**某高校机电类二级学院**的一次实测（机构名、人名均匿名）。可作为同类项目的检查清单。

## 1. 同名院系：不限定校名会多算 114 篇

`School of Mechanical-Electronic and Vehicle Engineering` 这个英文名**多所高校在用**，实测至少涉及：

- 目标学院本身
- B 校（同名，写法完全一致）
- 另有多所院校使用仅一字之差的写法

只按院系名搜索得到 563 篇，其中本校仅 449 篇——**差 114 篇全是别人的**。加上校名限定后，误命中归零。

> 规则：院系名英文检索**必须**同时限定校名，且逐条核对单位串。

## 2. 单位串写法十几种，只搜一种会漏一半

同一学院在论文中的真实写法统计（按出现次数）：

```
110  School of Mechanical-Electronic and Vehicle Engineering
 30  School of Mechanical-electronic and Vehicle Engineering      ← 大小写
 24  School of Mechanical-electronic and Automobile Engineering   ← Automobile
 17  School of Mechanical-Electrical and Vehicle Engineering      ← Electrical
 10  School of Mechanical and Vehicle Engineering
  9  School of Mechanical Electronic and Vehicle Engineering      ← 无连字符
  6  School of Mechanical-Electronic Vehicular Engineering        ← Vehicular
  5  School of Mechanical–Electronic and Vehicle Engineering      ← en dash
  4  School of Mechanical, Electrical and Vehicle Engineering
  …  还有 Mechanical—Electrical、Engneering 拼错
```

> 规则：先做"写法普查"（流水线第 2 步自动完成），把"全球命中 − 本校命中"的差值当成漏检量指标。

## 3. 只署实验室、不署学院

大量论文的作者单位是实验室或研究中心，例如：

- `Beijing Key Laboratory of ... Urban Rail Transit Vehicles, <校名>`
- 各类省部级工程技术研究中心、质量监督检验站

只匹配学院名会漏掉几百篇。**必须把学院所属实验室/中心名加入 `lab_patterns`。**

## 4. 只写"××大学"，不写学院 → 用教师姓名反查

实测 1052 篇候选里，**294 篇（28%）**的单位只到校名级。靠教师姓名 + OpenAlex 作者ID 反查补全。

注意同名：某位教师的英文署名在 OpenAlex 里被拆成 **9 个不同作者ID**，必须收集全部 ID 再合并，不能只取一个。

## 5. 兼职教授会把别院的论文带进来

一位在师资页挂"教授（兼职）"、实际单位在另一个学院的人员，其关联的 28 篇论文中：

- **25 篇**只有他一人关联本院 → 内容全是推荐系统/NLP/无线网络/负荷预测，与机电学科无关 → **必须剔除**
- **3 篇**与本院其他教师合作 → 保留

> 规则：兼职人员论文按"唯一关联则剔除、有本院合作者则保留"处理，并在表中标注。

## 6. 第一作者是本校但属其他学院

有 28 篇论文第一作者署的是校内另一个学院的单位，因有本院教师合作而进入数据集。若不看第一作者单位，会把它算成"本院学生一作"——实测这种情况让学生数字虚高 18–28 篇。

> 规则：身份判定必须**同时**看"第一作者是否在编教师"和"第一作者单位是否为本院"。

## 附加：撤稿与分区版本

- 实测有 4 篇撤稿论文（`RETRACTED:` 前缀 / `type=retraction`），必须剔除
- 中科院分区每年调整：部分期刊已移出最新版分区表，按现状会归入"非SCI期刊"——写报告时要注明"按某年版分区"
