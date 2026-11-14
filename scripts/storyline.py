#!/usr/bin/env python3
"""汇报大纲的分层检查和可读视图。只用标准库，兼容 Python 3.9。

用法：
    python3 storyline.py check   outline.json [--tree]
    python3 storyline.py view    outline.json [--tree]
    python3 storyline.py approve outline.json --by user|none

--tree 只看前三层（听众与定义卡、整份、章），用于论证树确认之前；
此时各页只需要 id、role、title，不需要内容块。

approve 把当前的论证树登记为已确认的版本，写回 outline.json。用户确认后用
--by user；无法与用户交互时用 --by none。之后的 check 会列出相对这个版本的改动。

检查按层输出，从上往下改：
    第 0 层 听众与定义卡    第 1 层 整份    第 2 层 章    第 3 层 页    第 4 层 块

错误（E）是必填项缺失，必须改。提示（如 L2-3）是结构上的疑点：改掉，或在文件
顶层 "waivers" 里写明保留理由，例如
    {"rule": "L2-2", "where": "2", "reason": "八个场景同级，客户要求逐个过"}
"""
import datetime
import json
import re
import sys

LAYERS = ["第 0 层 听众与定义卡", "第 1 层 整份", "第 2 层 章", "第 3 层 页", "第 4 层 块"]
SCENARIOS = ["upward-report", "decision", "client-proposal", "training", "custom"]
ROLES = ["cover", "agenda", "opener", "overview", "content", "summary", "closing", "appendix"]
BODY_ROLES = ("content", "overview")  # 带论点的页：要有导语和证据
LAYOUTS = ["单块", "左主右辅", "上主下辅", "左右对照", "上下两段", "三栏"]

# 关系 -> 候选图示的关键词。component 里含任一关键词即视为匹配。
RELATIONS = {
    "顺序": ["流程", "时间轴", "泳道", "阶梯", "路线"],
    "分支": ["判断树", "决策树", "分支", "流程图"],
    "循环": ["闭环", "循环", "环形"],
    "层级": ["分层", "树", "嵌套", "分组", "金字塔", "架构", "地图"],
    "对比": ["对照", "对比"],
    "交叉": ["矩阵", "象限", "表"],
    "因果": ["因果", "对策", "鱼骨"],
    "汇聚": ["辐射", "汇聚", "发散"],
    "数量": ["图", "表"],
    "实物": ["截图", "原件", "照片"],
    "强调": ["大数字", "引言"],
    "并列": ["卡片", "图标"],
}
# 项数由事实决定的关系：步骤、表格行、数据点、截图不为凑数而合并，上限放宽
LONG_OK = ("顺序", "分支", "交叉", "数量", "实物")
# 能表达两层内容的图示关键词
TWO_LEVEL = ["分组", "树", "分层", "嵌套", "泳道", "表", "矩阵", "阶段", "架构", "地图"]
# 本来就要画全的图：架构图、能力地图、流程图。项数由事实决定，不按一般的上限提示
DENSE = ["架构", "地图", "流程图", "泳道"]
MAX_DENSE = 16           # 这类图一层里的项数、流程图的节点数超过它，一页上看不清

# 以下阈值均为初始建议值
MAX_CHILDREN = 5         # 一组并列的子项超过它，通常藏着一个中间层
MAX_ITEMS = 6            # 一个内容块里并列的项超过它，先看能不能分组
MAX_LONG = 9             # LONG_OK 关系的项数上限
MAX_SECTION_SLIDES = 6   # 一章的正文页超过它，考虑分组或拆章
PARALLEL_SHARE = 0.4     # 并列块占比超过它，逐个复核
MIN_CLAIM_LEN = 10       # 短于它的“论点”多半是标签
MIN_LEAD_LEN = 10        # 短于它的导语多半是标签
MAX_TITLE_LEN = 38       # 超过它的标题排出来会超过两行
MIN_PER_PAGE = 1.5       # 现场讲解每页正文至少需要的分钟数
MIN_DENSE_CHARS = 100    # 密度为中高或高的正文页，内容块里的字数少于它，排出来会发空
LABEL_LEAD = re.compile(r"^(本章|本页|本节|以下|下面|摘要[:：])|(回顾|如下|一览|概览)$")


class Report:
    def __init__(self, waivers):
        self.items = {i: [] for i in range(5)}
        self.n_err = 0
        self.n_warn = 0
        self.waived = []
        self._waivers = {(w.get("rule"), str(w.get("where", ""))): w.get("reason", "") for w in waivers}

    def error(self, layer, where, msg):
        self.n_err += 1
        self.items[layer].append("E     [%s] %s" % (where, msg))

    def warn(self, rule, where, msg, fix):
        reason = self._waivers.get((rule, str(where)))
        if reason:
            self.waived.append("%s [%s] 已说明保留：%s" % (rule, where, reason))
            return
        self.n_warn += 1
        self.items[int(rule[1])].append("%s  [%s] %s\n        改法：%s" % (rule, where, msg, fix))


def text(v):
    return v.strip() if isinstance(v, str) else ""


def relation_ok(rel):
    rel = text(rel)
    return rel in RELATIONS or rel.startswith("其他")


def all_slides(doc):
    """按放映顺序产出 (章或 None, 页)，不含附录。"""
    for s in doc.get("opening") or []:
        yield None, s
    for sec in doc.get("sections") or []:
        for s in sec.get("slides") or []:
            yield sec, s
    for s in doc.get("closing") or []:
        yield None, s


def flat_items(items):
    for it in items or []:
        if isinstance(it, dict) and "group" in it:
            yield text(it.get("group"))
            for sub in flat_items(it.get("items")):
                yield sub
        elif isinstance(it, dict):            # 表格行、流程节点、卡片
            for k in ("name", "title", "tag", "sub", "desc", "body"):
                yield text(it.get(k))
            for sub in it.get("cells") or []:
                yield text(sub)
        elif isinstance(it, (list, tuple)):   # 大数字：(数字, 含义)
            for sub in it:
                yield text(sub)
        else:
            yield text(it)


def visible_text(doc):
    """产出 (页编号, 这页上听众看得到的全部文字)。"""
    for _, s in all_slides(doc):
        parts = [text(s.get("title")), text(s.get("lead")), text(s.get("conclusion"))]
        for b in s.get("blocks") or []:
            parts.append(text(b.get("says")))
            parts.extend(flat_items(b.get("items")))
        yield text(str(s.get("id", ""))), "\n".join(parts)


# ---------------------------------------------------------------- 第 0 层
def check_brief(doc, rep):
    brief = doc.get("brief") or {}
    if text(brief.get("scenario")) not in SCENARIOS:
        rep.error(0, "brief.scenario", "场景要从 %s 里选" % "、".join(SCENARIOS))
    if not text(brief.get("occasion")):
        rep.error(0, "brief.occasion", "缺场合")
    aud = brief.get("audience")
    if not isinstance(aud, dict):
        rep.error(0, "brief.audience", "听众要写成四项：who、background、cares、angle，见 references/L0-audience.md")
    else:
        for key, name in (("who", "谁在听、谁拍板"), ("background", "专业背景，已知什么、不知道什么"),
                          ("cares", "用什么标准判断这件事"), ("angle", "用什么角度讲他们听得懂")):
            if not text(aud.get(key)):
                rep.error(0, "brief.audience.%s" % key, "缺听众的%s；从材料里看不出就问用户" % name)
    takeaways = brief.get("takeaways") or []
    if not takeaways:
        rep.error(0, "brief.takeaways", "缺目的：听众听完要记住或决定的事，写一到三条")
    elif len(takeaways) > 3:
        rep.warn("L0-1", "brief.takeaways", "写了 %d 条，听众记不住这么多" % len(takeaways),
                 "合并到三条以内；服务不了这三条的内容进附录")
    if text(brief.get("scenario")) == "custom" and not text(brief.get("deviations")):
        rep.error(0, "brief.deviations", "自定义场景要在 deviations 里写明章节顺序和这样排的理由")
    terms = brief.get("terms")
    if terms is None:
        rep.warn("L0-2", "brief.terms", "没有登记术语",
                 "把标题、导语、图示里听众可能不熟的词列出来并写处理办法；确实没有就写空列表 []")
    for t in terms or []:
        if not isinstance(t, dict) or not text(t.get("term")):
            continue
        term = text(t.get("term"))
        if re.search(r"[/、，,]", term):
            rep.warn("L0-3", "brief.terms", "“%s”把几个词登记成了一条" % term, "一个词一条，脚本才能逐个核对")
        if text(t.get("treatment")) == "换说法":
            hits = [pid for pid, blob in visible_text(doc) if term in blob]
            if hits:
                rep.warn("L0-4", "brief.terms", "听众没听过的词“%s”仍出现在 %s" % (term, "、".join(hits[:6])),
                         "换成登记的说法“%s”；必须保留原词的地方先解释再用；固定名称里含有这个词时可说明后保留" % text(t.get("plain")))


# ---------------------------------------------------------------- 第 1 层
def check_deck(doc, rep, tree_only):
    brief = doc.get("brief") or {}
    thesis = text(doc.get("thesis"))
    if not thesis:
        rep.error(1, "thesis", "缺主旨句")
    elif len(thesis) < MIN_CLAIM_LEN + 2:
        rep.warn("L1-1", "thesis", "主旨句只有 %d 个字，像题目" % len(thesis),
                 "写成一句完整的话：听众听完带走的判断或结论")
    sections = doc.get("sections") or []
    if not sections:
        rep.error(1, "sections", "没有章")
    elif not (2 <= len(sections) <= MAX_CHILDREN):
        rep.warn("L1-2", "sections", "共 %d 章" % len(sections),
                 "章数宜在 2 到 %d 之间；过多时把相近的章归到一个更上位的论点下" % MAX_CHILDREN)
    guide = [s for s in doc.get("opening") or [] if s.get("role") in ("agenda", "overview", "content")]
    if sections and not guide:
        rep.warn("L1-3", "opening", "开头没有交代全稿结构的页", "按场景档加一页：提纲、摘要或全景")
    if sections and not (doc.get("closing") or []):
        rep.warn("L1-4", "closing", "没有收尾页", "结尾回到主旨句：总结、下一步安排或请求事项")
    written = [text(str(x.get("id", ""))) for _, x in all_slides(doc) if x.get("blocks")]
    if tree_only and written and not isinstance(doc.get("tree_confirmed"), dict):
        rep.warn("L1-7", "tree_confirmed", "论证树还没有确认，已经有 %d 页写了内容" % len(written),
                 "听众或章论点一改，这些内容都要重写。先停下，把论证树交给用户确认，确认后再写页面内容")
    if not tree_only and not isinstance(doc.get("tree_confirmed"), dict):
        rep.warn("L1-6", "tree_confirmed", "论证树还没有登记确认",
                 "用户确认后运行 approve --by user；无法与用户交互时运行 approve --by none，并在交付时说明")
    minutes = brief.get("duration_min")
    talk = [s for _, s in all_slides(doc) if s.get("role") in BODY_ROLES]
    if isinstance(minutes, (int, float)) and minutes > 0 and len(talk) * MIN_PER_PAGE > minutes * 1.1:
        rep.warn("L1-5", "全稿", "正文 %d 页，现场讲 %d 分钟，按每页 %.1f 分钟算讲不完" % (len(talk), minutes, MIN_PER_PAGE),
                 "把支撑性的页移到附录，或合并内容相近的页；合并会让内容变差时说明后保留")


# ---------------------------------------------------------------- 第 2 层
def check_sections(doc, rep):
    sections = doc.get("sections") or []
    summaries = []
    for sec in sections:
        sid = text(str(sec.get("id", ""))) or "?"
        claim = text(sec.get("claim"))
        if not claim:
            rep.error(2, sid, "章缺论点 claim")
        elif len(claim) < MIN_CLAIM_LEN:
            rep.warn("L2-1", sid, "章论点“%s”像章名" % claim,
                     "章名写进 label；claim 写成一句概括本章各页的话，各页内容变了它就不成立")
        slides = sec.get("slides") or []
        body = [s for s in slides if s.get("role") in BODY_ROLES]
        if not slides:
            rep.error(2, sid, "章下没有页")
        if len(body) >= 2 and not relation_ok(sec.get("relation")):
            rep.error(2, sid, "缺 relation：本章各页之间是什么关系，从 %s 里选，或写“其他：……”" % "、".join(RELATIONS))
        if len(body) == 1 and len(sections) > 2:
            rep.warn("L2-2", sid, "本章只有一页正文", "并入相邻的章，或确认它确实要单独成章")
        if len(body) > MAX_SECTION_SLIDES:
            rep.warn("L2-2", sid, "本章有 %d 页正文" % len(body),
                     "找出页与页之间的分组：拆成两章，或用一页总览把它们分成两三组")
        if len(body) >= 3:
            if not text(sec.get("opens")):
                rep.warn("L2-3", sid, "缺章首预告 opens", "写一句：本章回答什么问题、分几部分讲")
            if not text(sec.get("closes")):
                rep.warn("L2-3", sid, "缺章末收束 closes", "写一句：本章讲完得到什么结论")
            if not text(sec.get("pattern")):
                rep.warn("L2-4", sid, "没有写章内组织模式 pattern",
                         "从 references/L2-section.md 第一节选一种，或写“其他：……”")
        for s in slides:
            if s.get("role") == "summary":
                sig = tuple((text(b.get("relation")), text(b.get("component")), len(b.get("items") or []))
                            for b in s.get("blocks") or [])
                summaries.append((text(str(s.get("id", ""))), sig))
    sigs = [sig for _, sig in summaries]
    same = [pid for pid, sig in summaries if sig and sigs.count(sig) >= 2]
    if len(same) >= 2:
        rep.warn("L2-5", "全稿", "小结页 %s 的写法一样" % "、".join(same),
                 "小结不重复罗列要点：回到本章总览图标出已讲部分，加一句收束；短的章改用结论条收束")


# ---------------------------------------------------------------- 第 3、4 层
def check_slides(doc, rep, tree_only):
    brief = doc.get("brief") or {}
    seen_ids, concept_home = set(), {}
    blocks_total, blocks_parallel, component_run = 0, [], []
    content_pages, both_pages, concl_pages = 0, [], []
    for _, s in all_slides(doc):
        pid = text(str(s.get("id", ""))) or "?"
        if pid in seen_ids:
            rep.error(3, pid, "页编号重复")
        seen_ids.add(pid)
        role = s.get("role")
        if role not in ROLES:
            rep.error(3, pid, "role 要从 %s 里选" % "、".join(ROLES))
        title = text(s.get("title"))
        if not title and role != "cover":
            rep.error(3, pid, "缺标题")
        if len(title) > MAX_TITLE_LEN:
            rep.warn("L3-1", pid, "标题 %d 个字，排出来会超过两行" % len(title), "压缩到两行以内，限定条件移到导语")
        if tree_only or role not in BODY_ROLES:
            continue
        lead = text(s.get("lead"))
        if not lead and "主题" in text(brief.get("title_style")):
            rep.warn("L3-2", pid, "标题是主题式的，这页没有导语", "用导语说出这页的论点；标题本身已是完整论断的页可以说明后保留")
        elif lead and (len(lead) < MIN_LEAD_LEN or LABEL_LEAD.search(lead)):
            rep.warn("L3-3", pid, "导语“%s”像标签" % lead,
                     "写成一句有内容的话；用主题式标题时，导语要说出这页的论点")
        if role == "content" and not pid.startswith("0."):
            content_pages += 1
            if text(s.get("conclusion")):
                concl_pages.append(pid)
                if lead:
                    both_pages.append(pid)
        blocks = s.get("blocks") or []
        if not blocks:
            rep.error(3, pid, "正文页没有内容块 blocks")
        charts = any(text(b.get("relation")) in ("数量", "实物") or any(k in text(b.get("component")) for k in DENSE)
                     for b in blocks)        # 图表、截图和画全貌的图不按字数判断
        if "高" in text(brief.get("density")) and role == "content" and blocks and not charts and not pid.startswith("0."):
            chars = sum(len(text(it)) for b in blocks for it in flat_items(b.get("items") or []))
            if chars < MIN_DENSE_CHARS:
                rep.warn("L3-8", pid, "密度定为“%s”，这页内容块里只有 %d 个字，排出来会发空" % (text(brief.get("density")), chars),
                         "给每一项补第二层信息：说明、数据或例子；材料里确实没有更多内容的，与相邻页合并，或说明这页按低密度做")
        if len(blocks) > 3:
            rep.warn("L3-4", pid, "一页有 %d 个内容块" % len(blocks), "留一个主块，其余合并、拆页或进附录")
        layout = text(s.get("layout"))
        if len(blocks) >= 2 and not (layout in LAYOUTS or layout.startswith("其他")):
            rep.warn("L3-7", pid, "这页有 %d 个内容块，没有写版式 layout" % len(blocks),
                     "从 %s 里选，或写“其他：……”；见 references/layout.md 第二节" % "、".join(LAYOUTS))
        has_evidence = False
        for i, b in enumerate(blocks, 1):
            bid = "%s#%d" % (pid, i)
            items = b.get("items") or []
            rel, comp = text(b.get("relation")), text(b.get("component"))
            if not text(b.get("says")):
                rep.error(4, bid, "内容块缺 says：这一块要表达什么，一句话")
            if len(items) >= 2 and not relation_ok(rel):
                rep.error(4, bid, "缺 relation：这几项之间是什么关系")
            limit = MAX_DENSE if any(k in comp for k in DENSE) else (MAX_LONG if rel in LONG_OK else MAX_ITEMS)
            if len(items) > limit:
                rep.warn("L4-1", bid, "列了 %d 项" % len(items),
                         "看能不能分成两到四组，写成 {\"group\": \"组名\", \"items\": [...]}，组内保留每一项；"
                         "来源本身就是一份固定清单、分组会显得生硬时，说明后保留")
            grouped = False
            for it in items:
                if isinstance(it, dict) and "group" in it:
                    grouped = True
                    if len(it.get("items") or []) > limit:
                        rep.warn("L4-1", bid, "组“%s”有 %d 项" % (text(it.get("group")), len(it.get("items"))),
                                 "组内再分，或拆成两页")
            if comp and not comp.startswith("其他"):
                if rel in RELATIONS and not any(k in comp for k in RELATIONS[rel]):
                    rep.warn("L4-2", bid, "关系是“%s”，图示却是“%s”" % (rel, comp),
                             "按 references/L4-block.md 重新判断关系；关系没错就换成该关系的候选图示，或写“其他：……”说明画法")
                if grouped and not any(k in comp for k in TWO_LEVEL):
                    rep.warn("L4-3", bid, "内容分了组，图示“%s”只能表达一层" % comp,
                             "关系不用改，在图示前加“分组”即可：分组卡片、分阶段流程、分组对照栏、分组因果链；或用树、分层架构")
            blocks_total += 1
            if rel == "并列":
                blocks_parallel.append(bid)
            for ev in b.get("evidence") or []:
                has_evidence = True
                if brief.get("external") and isinstance(ev, dict) and not text(ev.get("source")):
                    rep.warn("L4-4", bid, "对外材料里的证据没有出处：%s" % text(ev.get("fact"))[:30],
                             "补出处；没有出处的数字删掉，或标为初始建议值")
            component_run.append((pid, comp))
        if blocks and not has_evidence:
            rep.warn("L3-5", pid, "这页没有挂任何证据",
                     "补上支撑论点的事实、数据、案例或截图；定义页、安排页，或依据是推断并已在页面注明的页，可说明后保留")
        for c in s.get("expands") or []:
            if c in concept_home and concept_home[c] != pid:
                rep.warn("L3-6", pid, "概念“%s”已在第 %s 页展开" % (c, concept_home[c]),
                         "同一概念只在一页讲透，其他页只引用并注明见哪页")
            concept_home.setdefault(c, pid)
    if tree_only:
        return
    topic = "主题" in text(brief.get("title_style"))
    many = concl_pages if topic else both_pages
    if content_pages >= 4 and len(many) * 2 > content_pages:
        if topic:
            rep.warn("L3-9", "全稿", "%d 张正文页里有 %d 张写了结论条" % (content_pages, len(many)),
                     "导语已经说出论点的页，结论条只在推出了导语之外的一句话时写，例如下一步的动作、对听众的请求；逐页重新判断，见 references/page-plan.md 第三节")
        else:
            rep.warn("L3-9", "全稿", "%d 张正文页里有 %d 张同时写了导语和结论条" % (content_pages, len(many)),
                     "标题已是完整论断的页不写导语；结论条只在推出了标题之外的一句话时写。逐页重新判断，见 references/page-plan.md 第三节")
    if blocks_total and len(blocks_parallel) / float(blocks_total) > PARALLEL_SHARE:
        rep.warn("L4-5", "全稿", "并列块占 %d/%d：%s" % (len(blocks_parallel), blocks_total, "、".join(blocks_parallel)),
                 "按 references/L4-block.md 的判定顺序逐个重问；都不成立才留并列")
    run = 1
    for i in range(1, len(component_run)):
        same = component_run[i][1] and component_run[i][1] == component_run[i - 1][1]
        run = run + 1 if same else 1
        if run == 3:
            rep.warn("L4-6", component_run[i][0], "连续三个内容块都用“%s”" % component_run[i][1],
                     "回到内容关系重新判断；关系确实相同也可以换一种候选图示")


def tree_snapshot(doc):
    brief = doc.get("brief") or {}
    aud = brief.get("audience") if isinstance(brief.get("audience"), dict) else {}
    return {
        "scenario": text(brief.get("scenario")),
        "audience": {k: text(aud.get(k)) for k in ("who", "background", "cares", "angle")},
        "thesis": text(doc.get("thesis")),
        "sections": [{"id": text(str(sec.get("id", ""))), "claim": text(sec.get("claim")),
                      "pattern": text(sec.get("pattern")),
                      "slides": [[text(str(x.get("id", ""))), text(x.get("title"))] for x in sec.get("slides") or []]}
                     for sec in doc.get("sections") or []],
    }


def tree_diff(old, new):
    out = []
    if old.get("scenario") != new["scenario"]:
        out.append("场景：%s 改为 %s" % (old.get("scenario"), new["scenario"]))
    names = {"who": "听众", "background": "听众背景", "cares": "听众的判断标准", "angle": "讲述角度"}
    for k, name in names.items():
        if (old.get("audience") or {}).get(k) != new["audience"][k]:
            out.append("%s改了" % name)
    if old.get("thesis") != new["thesis"]:
        out.append("主旨句改了")
    old_secs = {x["id"]: x for x in old.get("sections") or []}
    new_secs = {x["id"]: x for x in new["sections"]}
    for sid in old_secs:
        if sid not in new_secs:
            out.append("删去第 %s 章：%s" % (sid, old_secs[sid]["claim"]))
    for sid, sec in new_secs.items():
        if sid not in old_secs:
            out.append("新增第 %s 章：%s" % (sid, sec["claim"]))
            continue
        o = old_secs[sid]
        if o["claim"] != sec["claim"]:
            out.append("第 %s 章的论点改了" % sid)
        if o.get("pattern") != sec["pattern"]:
            out.append("第 %s 章的组织模式改了" % sid)
        o_pages, n_pages = dict(o["slides"]), dict(sec["slides"])
        for pid in o_pages:
            if pid not in n_pages:
                out.append("删去第 %s 页：%s" % (pid, o_pages[pid]))
        for pid, title in n_pages.items():
            if pid not in o_pages:
                out.append("新增第 %s 页：%s" % (pid, title))
            elif o_pages[pid] != title:
                out.append("第 %s 页的标题改了" % pid)
    return out


def check(doc, tree_only):
    rep = Report(doc.get("waivers") or [])
    check_brief(doc, rep)
    check_deck(doc, rep, tree_only)
    check_sections(doc, rep)
    check_slides(doc, rep, tree_only)
    return rep


# ---------------------------------------------------------------- 输出
def readthrough(doc):
    lines = ["主旨：%s" % text(doc.get("thesis"))]
    for sec in doc.get("sections") or []:
        lines.append("%s %s" % (sec.get("id", ""), text(sec.get("claim"))))
        for s in sec.get("slides") or []:
            if s.get("role") in BODY_ROLES + ("summary",):
                lead = text(s.get("lead"))
                lines.append("    %s %s%s" % (s.get("id", ""), text(s.get("title")), ("：" + lead) if lead else ""))
    return "\n".join(lines)


def english_words(doc):
    brief = doc.get("brief") or {}
    known = " ".join(text(t.get("term")) for t in brief.get("terms") or []
                     if isinstance(t, dict) and text(t.get("treatment")) != "换说法")
    words = []
    for _, blob in visible_text(doc):
        for w in re.findall(r"[A-Za-z][A-Za-z\-\.]{2,}(?: [A-Za-z][A-Za-z\-\.]{2,})*", blob):
            if w not in words and w not in known:
                words.append(w)
    return words


NUMBER = re.compile(r"\d+(?:\.\d+)?\s*(?:%|％|[个条项类页步章层种处倍成天周月年人次元万亿]|小时|分钟)?"
                    r"|[一二三四五六七八九十两百千半]+(?:个|条|项|类|步|章|层|种|处|倍|成|件|道|批|轮|天|周|个月|年|小时)")


def claim_numbers(doc):
    """主旨句、章论点、页标题里出现的数字和数量说法，供人工核对出处。"""
    texts = [text(doc.get("thesis"))]
    for sec in doc.get("sections") or []:
        texts.append(text(sec.get("claim")))
    for _, s in all_slides(doc):
        texts.append(text(s.get("title")))
    found = []
    for t in texts:
        for m in NUMBER.findall(t):
            m = m.strip()
            if m and m not in found:
                found.append(m)
    return found


def stats(doc):
    body = [s for _, s in all_slides(doc) if s.get("role") in BODY_ROLES]
    rels = {}
    for s in body:
        for b in s.get("blocks") or []:
            r = text(b.get("relation")) or "未标"
            rels[r] = rels.get(r, 0) + 1
    total = sum(1 for _ in all_slides(doc)) + len(doc.get("appendix") or [])
    secs = doc.get("sections") or []
    framed = sum(1 for x in secs if text(x.get("opens")) and text(x.get("closes")))
    return "页数 %d（正文 %d，附录 %d）；章 %d，其中 %d 章写了章首预告和章末收束；内容块关系分布 %s" % (
        total, len(body), len(doc.get("appendix") or []), len(secs), framed,
        "、".join("%s %d" % kv for kv in sorted(rels.items(), key=lambda kv: -kv[1])) or "无")


def cmd_check(doc, tree_only):
    rep = check(doc, tree_only)
    print("== 统计%s\n%s" % ("（论证树阶段，只查前三层）" if tree_only else "", stats(doc)))
    words = english_words(doc)
    if words:
        print("页面上的英文词（逐个确认听众认识；不认识的登记到 terms）：%s" % "、".join(words[:20]))
    nums = claim_numbers(doc)
    if nums:
        print("主旨句、章论点和标题里的数字（逐个确认材料里有出处；没有出处的删掉或写待定）：%s" % "、".join(nums[:30]))
    print("\n== 标题串读（只读这些，能不能复述出主旨？再换成听众的身份读一遍）\n" + readthrough(doc))
    for i, name in enumerate(LAYERS):
        if rep.items[i]:
            print("\n== %s" % name)
            print("\n".join(rep.items[i]))
    if rep.waived:
        print("\n== 已说明保留 %d 条" % len(rep.waived))
        print("\n".join(rep.waived))
    conf = doc.get("tree_confirmed")
    if isinstance(conf, dict) and not tree_only:
        who = "用户已确认" if conf.get("by") == "user" else "未经用户确认，交付时要说明"
        changes = tree_diff(conf.get("snapshot") or {}, tree_snapshot(doc))
        print("\n== 论证树登记于 %s（%s）" % (conf.get("date", "?"), who))
        if changes:
            print("之后有 %d 处改动，交付时逐条向用户说明：" % len(changes))
            print("\n".join("  - " + c for c in changes))
        else:
            print("之后没有改动。")
    if rep.n_err or rep.n_warn:
        print("\n错误 %d 条（必须改），提示 %d 条（改掉，或写进 waivers 说明保留理由）。" % (rep.n_err, rep.n_warn))
        print("从上层改起：上层改动后，下层的提示可能随之消失或变化。")
    else:
        print("\n结构检查通过。机器只查形状，内容是否成立仍要读一遍标题串读。")
    return 1 if rep.n_err else 0


def cmd_view(doc, tree_only):
    b = doc.get("brief") or {}
    aud = b.get("audience") if isinstance(b.get("audience"), dict) else {}
    out = ["# 汇报大纲%s" % ("（论证树）" if tree_only else ""), "", "## 听众与定义卡", ""]
    out.append("- 场景：%s" % b.get("scenario", ""))
    for key, name in (("who", "听众"), ("background", "听众背景"), ("cares", "听众的判断标准"), ("angle", "讲述角度")):
        if text(aud.get(key)):
            out.append("- %s：%s" % (name, aud.get(key)))
    if b.get("duration_min"):
        out.append("- 时长：%s 分钟" % b.get("duration_min"))
    for key, name in (("occasion", "场合"), ("title_style", "标题写法"), ("density", "信息密度"),
                      ("deviations", "偏离默认做法之处")):
        if text(str(b.get(key, ""))):
            out.append("- %s：%s" % (name, b.get(key)))
    if b.get("terms"):
        out.append("- 术语处理：")
        for t in b.get("terms"):
            if isinstance(t, dict):
                plain = "：%s" % text(t.get("plain")) if text(t.get("plain")) else ""
                out.append("  - %s（%s）%s" % (text(t.get("term")), text(t.get("treatment")), plain))
    out.append("- 听众听完要记住或决定的事：")
    out.extend("  %d. %s" % (i, t) for i, t in enumerate(b.get("takeaways") or [], 1))
    out += ["", "## 主旨句", "", text(doc.get("thesis")), "", "## 论证树", ""]
    for sec in doc.get("sections") or []:
        tags = [x for x in ("模式：" + text(sec.get("pattern")) if text(sec.get("pattern")) else "",
                            "各页关系：" + text(sec.get("relation")) if text(sec.get("relation")) else "") if x]
        out.append("- **%s %s**%s" % (sec.get("id", ""), text(sec.get("claim")), "（%s）" % "；".join(tags) if tags else ""))
        if text(sec.get("opens")):
            out.append("  - 章首预告：%s" % text(sec.get("opens")))
        for s in sec.get("slides") or []:
            if s.get("role") in BODY_ROLES + ("summary",):
                out.append("  - %s %s" % (s.get("id", ""), text(s.get("title"))))
        if text(sec.get("closes")):
            out.append("  - 章末收束：%s" % text(sec.get("closes")))
    if tree_only:
        print("\n".join(out))
        return 0
    out += ["", "## 页级大纲", ""]

    def slide_md(s):
        out.append("### %s %s　〔%s〕" % (s.get("id", ""), text(s.get("title")) or "（封面）", s.get("role", "")))
        if text(s.get("lead")):
            out.append("导语：%s" % text(s.get("lead")))
        for blk in s.get("blocks") or []:
            comp = "，%s" % text(blk.get("component")) if text(blk.get("component")) else ""
            out.append("- 〔%s%s〕%s" % (text(blk.get("relation")) or "—", comp, text(blk.get("says"))))
            cols = blk.get("columns") or []
            for it in blk.get("items") or []:
                if isinstance(it, dict) and "name" in it:
                    out.append("  - %s%s" % (text(it.get("name")), "　〔突出〕" if it.get("mark") else ""))
                    for k, cell in enumerate(it.get("cells") or []):
                        out.append("    - %s%s" % ("%s：" % cols[k] if k < len(cols) else "", text(cell)))
                elif isinstance(it, dict):
                    out.append("  - %s" % text(it.get("group")))
                    out.extend("    - %s" % sub for sub in it.get("items") or [])
                else:
                    out.append("  - %s" % it)
            for ev in blk.get("evidence") or []:
                if isinstance(ev, dict):
                    src = "（%s）" % text(ev.get("source")) if text(ev.get("source")) else ""
                    out.append("  - 证据：%s%s" % (text(ev.get("fact")), src))
                else:
                    out.append("  - 证据：%s" % ev)
        if text(s.get("conclusion")):
            out.append("结论条：%s" % text(s.get("conclusion")))
        if text(s.get("notes")):
            out.append("讲述要点：%s" % text(s.get("notes")))
        out.append("")

    for s in doc.get("opening") or []:
        slide_md(s)
    for sec in doc.get("sections") or []:
        out += ["## %s %s" % (sec.get("id", ""), text(sec.get("label")) or text(sec.get("claim"))), ""]
        for s in sec.get("slides") or []:
            slide_md(s)
    if doc.get("closing"):
        out += ["## 收尾", ""]
        for s in doc.get("closing"):
            slide_md(s)
    if doc.get("appendix"):
        out += ["## 附录", ""]
        for s in doc.get("appendix"):
            slide_md(s)
    print("\n".join(out))
    return 0


def cmd_approve(doc, path, by):
    rep = check(doc, True)
    if rep.n_err:
        print("论证树检查还有 %d 条错误，先运行 check --tree 改掉，再登记。" % rep.n_err)
        return 1
    doc["tree_confirmed"] = {"by": by, "date": datetime.date.today().isoformat(), "snapshot": tree_snapshot(doc)}
    with open(path, "w", encoding="utf-8") as f:
        json.dump(doc, f, ensure_ascii=False, indent=2)
        f.write("\n")
    if by == "user":
        print("已登记论证树（用户已确认）。这一项只在用户已经回复确认之后使用。")
    else:
        print("已登记论证树（未经用户确认）。")
        print("注意：--by none 只用于两种情况：用户明确要求一次出完；或者任务由另一个程序或 agent 派来，"
              "并且对方说明了不会有人回复。")
        print("不属于这两种情况时，现在应当停下：把听众理解与讲述角度、场景与主旨句、各章论点、"
              "偏离默认做法之处写给用户，结束这一轮，等用户回复后再写页面内容。")
    return 0


def main(argv):
    by = None
    if "--by" in argv:
        i = argv.index("--by")
        by = argv[i + 1] if i + 1 < len(argv) else None
        argv = argv[:i] + argv[i + 2:]
    args = [a for a in argv[1:] if a != "--tree"]
    tree_only = "--tree" in argv
    if len(args) != 2 or args[0] not in ("check", "view", "approve") or (args[0] == "approve" and by not in ("user", "none")):
        print(__doc__)
        return 2
    try:
        with open(args[1], encoding="utf-8") as f:
            doc = json.load(f)
    except ValueError as e:
        print("E  JSON 解析失败：%s\n   改法：检查这一行附近的逗号、引号和括号；正文里的英文双引号要写成 \\\"" % e)
        return 1
    if args[0] == "approve":
        return cmd_approve(doc, args[1], by)
    return cmd_check(doc, tree_only) if args[0] == "check" else cmd_view(doc, tree_only)


if __name__ == "__main__":
    sys.exit(main(sys.argv))
