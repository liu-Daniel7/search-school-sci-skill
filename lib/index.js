/**
 * search-school-sci-skill —— 以 DSH 技能（skill）形式分发学院科研产出统计工具。
 *
 * 实现方式与官方 `@deepseek-ai/dsh-skill-badge` 一致：
 * 注册一个 skill provider，把仓库根目录的 SKILL.md 作为技能正文，
 * 并把仓库根目录声明为 resourceBase，使 SKILL.md 里引用的
 * reference/*.md 与 scripts/*.py 相对路径可解析。
 */
import { readFile } from "node:fs/promises";
import { fileURLToPath } from "node:url";
import { BUNDLED_SKILL_RANK } from "@deepseek-ai/dsh-skill";

const PROVIDER_NAME = "search-school-sci";
const ROOT = new URL("../", import.meta.url);
const SKILL_BODY_URL = new URL("SKILL.md", ROOT);
const RESOURCE_BASE = { kind: "directory", path: fileURLToPath(ROOT) };

const CANDIDATE = {
  name: "search-school-sci",
  description:
    "统计某所大学某个学院（二级单位）的科研产出：按中科院分区/JCR/CCF/中文核心分层，教师与学生分开，输出带看板的可视化 Excel。适用于学科评估材料、学院年度报告、教师档案与对标分析。",
  invocation: { modelInvocable: true, userInvocable: true },
  provider: PROVIDER_NAME,
  source: "bundled",
  resourceBase: RESOURCE_BASE,
  rank: BUNDLED_SKILL_RANK,
  locator: SKILL_BODY_URL,
};

const provider = {
  name: PROVIDER_NAME,
  list: () => Promise.resolve([CANDIDATE]),
  async get(_candidate) {
    return {
      name: CANDIDATE.name,
      description: CANDIDATE.description,
      invocation: CANDIDATE.invocation,
      provider: CANDIDATE.provider,
      source: CANDIDATE.source,
      resourceBase: RESOURCE_BASE,
      content: await readFile(SKILL_BODY_URL, "utf8"),
    };
  },
};

/** Cordis 插件名。 */
export const name = "search-school-sci";
/** 依赖 skills 服务。 */
export const inject = ["skills"];
/** 把本技能注册到 ctx.skills。 */
export function apply(ctx) {
  ctx.skills.registerProvider(() => provider);
}
