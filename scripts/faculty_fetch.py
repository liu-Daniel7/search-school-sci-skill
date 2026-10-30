#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
faculty_fetch.py — 自动抓取学院教师名单（中文名 + 拼音 + 职称）

使用者只需给一个学院官网链接，其余由脚本完成：
  1. 在首页找“师资队伍/教师名录/Faculty/People”入口
  2. 跟进“教授/副教授/讲师/博导/硕导”等子页
  3. 从锚文本提取 2–4 字中文姓名；从个人主页链接（如 faculty.xxx.edu.cn/pub/zhangsan）提取拼音
  4. 输出去重后的 faculty.csv

用法：
  python faculty_fetch.py --url https://example.edu.cn/jdxy/szdw/ --out faculty.csv
  python faculty_fetch.py --url https://example.edu.cn/mech --out faculty.csv --max-pages 15

无法自动取到拼音时会留空并提示——拼音是匹配英文署名的唯一钥匙，缺了就得手工补。
"""
import argparse, csv, re, sys, subprocess, time
from urllib.parse import urljoin, urlparse

UA = "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/125.0 Safari/537.36"
NAV_KW = ["师资", "教师", "队伍", "名录", "导师", "faculty", "people", "staff", "teacher", "team"]
SUB_KW = ["教授", "副教授", "讲师", "助教", "博导", "硕导", "研究员", "实验师", "教师", "名录"]
TITLE_KW = {"教授": "教授", "副教授": "副教授", "高级实验师": "高级实验师", "讲师": "讲师",
            "研究员": "研究员", "副研究员": "副研究员", "实验师": "实验师", "助教": "助教"}
BAD_NAMES = set("""首页 上一页 下一页 尾页 更多 详情 简介 名录 名单 设置 联系 我们 学院 大学 学校
师资 队伍 教师 教授 副教授 讲师 助教 博导 硕导 导师 研究员 实验师 中心 办公室 党委 支部 工会
招生 就业 科研 教学 人才 培养 学科 建设 校友 服务 概况 领导 机构 友情 链接 版权 所有 关于 我们
新闻 通知 公告 动态 图片 视频 下载 上传 登录 注册 搜索 网站 地图 返回 顶部""".split())
CJK = re.compile(r"[\u4e00-\u9fff]")


def get(url, timeout=30):
    r = subprocess.run(["curl", "-sL", "-m", str(timeout), "-A", UA, url], capture_output=True, text=True)
    return r.stdout or ""


def links(html, base):
    out = []
    for m in re.finditer(r'<a\s+[^>]*href=["\']([^"\']+)["\'][^>]*>(.*?)</a>', html, re.I | re.S):
        href, txt = m.group(1), re.sub(r"<[^>]+>", "", m.group(2))
        txt = re.sub(r"\s+", "", txt)
        if href.startswith(("javascript:", "#", "mailto:")): continue
        out.append((urljoin(base, href), txt))
    return out


BAD_SUFFIX = ("工作", "建设", "招生", "服务", "风采", "校友", "信息", "团队", "基地", "设置",
              "简介", "信箱", "学院", "大学", "政策", "通知", "公告", "动态", "培养", "机构",
              "沿革", "文化", "顾问", "导师", "名单", "名录", "队伍", "链接", "概况", "竞赛")
BAD_WORDS = ("学院", "大学", "教授", "导师", "名单", "更多", "首页", "党群", "学生", "学科",
             "人才", "名师", "学校", "书记", "创新", "历史", "社会", "校友", "就业", "科研",
             "教学", "党建", "工会", "支部", "党委", "办公", "实验", "中心", "基地", "平台")


def is_name(t):
    if not (2 <= len(t) <= 4): return False
    if not CJK.match(t): return False
    if t in BAD_NAMES: return False
    if any(t.endswith(x) for x in BAD_SUFFIX): return False
    if any(k in t for k in BAD_WORDS): return False
    return True


LISTING_HINT = ("教授", "讲师", "师资", "教师", "导师", "名录", "队伍", "faculty", "people", "staff")


def is_listing_page(title, url):
    t = (title or "") + " " + (url or "")
    return any(k in t for k in LISTING_HINT)


def pinyin_from_url(u):
    """个人主页 URL 里常带拼音，如 /pub/yangjianwei/ 、/teacher/zhangsan"""
    m = re.search(r"/(?:pub|people|teacher|teachers|faculty|person|staff|home)/([a-zA-Z][a-zA-Z\-\.]{3,25})/?$", u)
    if not m: return ""
    slug = m.group(1).lower().replace(".", "").replace("-", "")
    return slug if slug.isalpha() and len(slug) >= 4 else ""


def parse_names(html, page_url, title=""):
    """只从“教师名录页”的链接里取名字，避免把导航栏当成教师。"""
    if not is_listing_page(title, page_url):
        return []
    out = []
    tt = ""
    for k, v in TITLE_KW.items():
        if k in (title or ""): tt = v; break
    for href, txt in links(html, page_url):
        if not is_name(txt): continue
        # 必须是指向个人页/详情页的链接，排除栏目标签页
        if not re.search(r"(/pub/|/people/|/teacher|/faculty/|\d{3,}|[0-9a-f]{16,})", href):
            continue
        out.append((txt, pinyin_from_url(href), tt, page_url, href))
    return out


# ---------------------------------------------------------------- 站点发现
def _search(query):
    """返回搜索结果里出现的 edu.cn 域名频次。"""
    from urllib.parse import quote
    engines = [
        ("baidu", "https://www.baidu.com/s?wd=" + quote(query)),
        ("sogou", "https://www.sogou.com/web?query=" + quote(query)),
    ]
    import collections
    dom = collections.Counter()
    for name, url in engines:
        html = get(url)
        for m in re.finditer(r"https?://([a-z0-9.\-]+\.edu\.cn)", html or "", re.I):
            dom[m.group(1).lower()] += 1
        if dom: break
    return dom


def discover(school, college, verbose=True):
    """只给学校和学院名，自动定位学院官网。"""
    host_hint = ""  # 由搜索结果打分
    qs = [f"{school}{college} 师资队伍", f"{school}{college} 官网", f"{college} 师资"]
    import collections
    score = collections.Counter()
    for q in qs:
        d = _search(q)
        for h, n in d.items():
            score[h] += n
        if verbose and d: print(f"    搜索「{q}」→ {len(d)} 个 edu.cn 域名")
    if not score:
        return []
    # 打分：出现次数 + 子域给分 + faculty 前缀给分
    def rank(h):
        pts = score[h]
        labels = h.split(".")
        if len(labels) >= 4: pts += 3          # example.edu.cn/mech 这类学院子域
        if labels[0] in ("faculty", "teacher", "people", "szdw"): pts += 2
        if labels[0] in ("www", "web", "home"): pts -= 1
        return pts
    cands = sorted(score, key=rank, reverse=True)[:6]
    if verbose:
        print("    候选站点：" + "、".join(f"{h}({rank(h)})" for h in cands))
    return cands


COMMON_PATHS = ["/szdw/", "/szdw/index.htm", "/szdw/index.html", "/szdw/index.php",
                "/teacher", "/teachers", "/faculty", "/people", "/staff", "/jsdw/",
                "/index/szdw.htm", "/college/szdw.htm"]


def probe_site(host, verbose=True):
    """在候选站点上找“师资队伍”入口，返回可用起点 URL 列表。"""
    base = "http://" + host + "/"
    html = get(base)
    starts = []
    if html:
        for href, txt in links(html, base):
            if any(k in txt for k in NAV_KW):
                starts.append(href)
    for p_ in COMMON_PATHS:
        starts.append("http://" + host + p_)
    # 去重并限制数量
    seen, out = set(), []
    for u in starts:
        if u in seen: continue
        seen.add(u); out.append(u)
    return out[:8]


def crawl(start, max_pages=12, verbose=True):
    seen, pages, result = set(), [], []
    queue = [(start, "")]
    while queue and len(pages) < max_pages:
        url, title = queue.pop(0)
        if url in seen: continue
        seen.add(url)
        html = get(url)
        if not html: continue
        pages.append(url)
        page_title = title
        mt = re.search(r"<title>(.*?)</title>", html, re.S | re.I)
        if mt: page_title = (title or "") + " " + re.sub(r"\s+", "", mt.group(1))
        found = parse_names(html, url, page_title)
        if found:
            result += found
            if verbose: print(f"    ✓ {url}  抓到 {len(found)} 人")
        # 找师资入口 / 子页
        for href, txt in links(html, url):
            if href in seen or urlparse(href).netloc != urlparse(start).netloc: continue
            if any(k in txt for k in NAV_KW) or any(k in txt for k in SUB_KW):
                queue.append((href, txt))
        time.sleep(0.3)
    return result, pages


def deep_pinyin(detail_url, base_host, verbose=False):
    """访问教师详情页，找 faculty.xxx.edu.cn/<拼音>/ 形式的个人主页链接。"""
    html = get(detail_url)
    if not html: return ""
    for m in re.finditer(r"https?://faculty\.[a-z0-9.\-]+/([a-zA-Z][a-zA-Z0-9\-]{2,25})/", html):
        return m.group(1).lower()
    for href, _ in links(html, detail_url):
        py = pinyin_from_url(href)
        if py: return py
    return ""


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--url", help="学院官网或“师资队伍”页地址；不给则用 --school/--college 自动发现")
    ap.add_argument("--school", help="学校中文名，如 示例大学")
    ap.add_argument("--college", help="学院中文名，如 机械工程学院")
    ap.add_argument("--out", default="faculty.csv")
    ap.add_argument("--max-pages", type=int, default=12)
    ap.add_argument("--deep", action="store_true", help="访问每位教师详情页，尝试从个人主页链接提取拼音（较慢但拼音覆盖率高）")
    a = ap.parse_args()

    rows, pages = [], []
    if a.url:
        print(f"抓取起点：{a.url}")
        rows, pages = crawl(a.url, a.max_pages)
    else:
        if not (a.school and a.college):
            sys.exit("请给 --url，或同时给 --school 与 --college")
        print(f"自动发现站点：{a.school} · {a.college}")
        cands = discover(a.school, a.college)
        if not cands:
            sys.exit("没能通过搜索引擎定位学院官网，请改用 --url 指定")
        for host in cands:
            for u in probe_site(host):
                r, pg = crawl(u, max_pages=max(4, a.max_pages // 3))
                rows += r; pages += pg
                if len({x[0] for x in rows}) >= 20:
                    break
            if len({x[0] for x in rows}) >= 20:
                print(f"    ✓ 命中站点 {host}")
                break
    print(f"访问页面 {len(pages)} 个，候选记录 {len(rows)} 条")

    best = {}
    for cn, py, tt, src, detail in rows:
        cur = best.get(cn)
        if not cur: best[cn] = {"py": py, "tt": tt, "src": src, "detail": detail, "n": 1}
        else:
            cur["n"] += 1
            if py and not cur["py"]: cur["py"] = py
            if tt and not cur["tt"]: cur["tt"] = tt
            if detail and not cur.get("detail"): cur["detail"] = detail

    clean = dict(best)
    if a.deep:
        todo = [(k, v) for k, v in clean.items() if not v["py"] and v.get("detail")]
        print(f"\n深挖拼音：{len(todo)} 人需要访问详情页（约 {len(todo)*0.4:.0f} 秒）")
        for i, (k, v) in enumerate(todo, 1):
            py = deep_pinyin(v["detail"], "")
            if py: v["py"] = py
            if i % 20 == 0: print(f"    …{i}/{len(todo)}")
            time.sleep(0.25)
    with open(a.out, "w", encoding="utf-8-sig", newline="") as f:
        w = csv.writer(f); w.writerow(["中文名", "全拼", "职称", "来源页"])
        for cn, v in sorted(clean.items(), key=lambda x: (x[1]["tt"] or "", x[0])):
            w.writerow([cn, v["py"], v["tt"], v["src"]])

    with_py = sum(1 for v in clean.values() if v["py"])
    print(f"\n✅ 写入 {a.out}：{len(clean)} 人，其中有拼音 {with_py} 人")
    if with_py < len(clean) * 0.6:
        print("""
⚠ 拼音覆盖率偏低。拼音是匹配英文署名的唯一钥匙，缺了会导致：
   · 该教师的英文论文被误判为“学生(推定)”
   建议：
   1) 打开某位教师的个人主页，看 URL 是否含拼音（/pub/zhangsan、/teacher/zhangsan）
   2) 若学院站点没有，改用知网/谷歌学术确认拼音，手工补进 faculty.csv 第二列
""")
    return 0 if clean else 1


if __name__ == "__main__":
    sys.exit(main())
