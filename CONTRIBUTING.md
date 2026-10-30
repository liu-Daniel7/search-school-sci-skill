# 贡献指南

## 分支与 PR 流程

`main` 已开启分支保护：**禁止直接推送，所有改动必须走 Pull Request**。

```bash
# 1. 从最新 main 开分支
git checkout main && git pull
git checkout -b feat/你的改动

# 2. 改完提交
git add -A
git commit -m "feat: 简短说明"

# 3. 推送并开 PR
git push -u origin feat/你的改动
gh pr create --fill        # 或到网页上开
```

分支命名：

| 前缀 | 用途 |
|---|---|
| `feat/` | 新增能力（如支持新的数据源） |
| `fix/` | 修 bug |
| `docs/` | 只改文档 |
| `chore/` | 构建、依赖、脚本清理 |
| `refactor/` | 不改行为的重构 |

## 提交信息

```
<type>: <一句话概括>

<可选：为什么这么改、影响哪些口径>
```

`type` 用 `feat` / `fix` / `docs` / `chore` / `refactor`。

## PR 合并前检查清单

- [ ] `python tests/smoke_test.py` 通过（离线，不联网）
- [ ] 若改了统计口径、层级判定或去重逻辑，PR 描述里说明**对既有数字的影响**
- [ ] 没有提交 `.cache/`、`*.xlsx`、`*.records.json` 等产物（`.gitignore` 已覆盖，提交前仍请 `git status` 确认）
- [ ] 没有提交登录凭据、Cookie、API key
- [ ] **没有提交真实学院的完整数据或真实人名**——示例一律用虚构数据（见 `docs/preview.png` 的做法）

## 数据处理红线

本工具会产出含**教师个人产出排名**和**学生姓名**的表。对外分享时：

1. 仓库、文档、示例中**只能放虚构数据**
2. 真实机构名、真实人名一律不进入本仓库
3. 生成的工作簿属于内部材料，不要直接提交到公开仓库

## 代码结构约定

| 文件 | 职责 | 改动时注意 |
|---|---|---|
| `scripts/college_sci.py` | 主流水线：抓取→普查→归属→分级 | 改过滤逻辑会影响召回率，务必说明影响 |
| `scripts/workbook.py` | 只负责排版出表 | 不要在这里做数据判定 |
| `scripts/faculty_fetch.py` | 抓学院官网教师名单 | 反爬敏感，改动请附实测站点 |
| `scripts/cnki_import.py` | 知网导出合并 | 去重规则改动需说明 |
| `reference/directions.json` | 学科方向关键词 | 中英文关键词都要维护 |

## 版本与变更记录

不维护 `CHANGELOG`，直接看 `git log`；破坏性改动（配置文件字段改名等）在 commit 正文里用 `BREAKING CHANGE:` 标注。
