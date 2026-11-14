---
name: one-prompt-ppt
description: 为汇报、方案交流、培训类 PPT 设计有层级的内容结构。按听众、整份、章、页、块五层自上而下做：先弄清听众背景和汇报场景，再写主旨句和各章论点，然后逐页写标题、导语、内容关系和图示建议，每层用脚本检查。用于“做 PPT、写汇报、做方案交流材料、做培训课件、PPT 内容太平没有逻辑、重排 PPT 结构、出 PPT 大纲”这类请求，在生成或修改幻灯片文件之前使用。产出大纲；用户要成品时，在自带模板或用户给的模板上生成可编辑的 pptx。
---

# PPT 内容结构

这个 skill 解决一个问题：做出来的 PPT 内容都在同一层，章只是容器名，页是罗列，页内是并列要点。做法是把内容分成五层，自上而下逐层定下来，每层有自己的规则和检查。

## 五层

| 层 | 定什么 | 规则文件 |
|---|---|---|
| 0 听众 | 听众是谁、什么背景、用什么角度讲、术语怎么处理、选哪个场景档 | [references/L0-audience.md](references/L0-audience.md) |
| 1 整份 | 主旨句；分几章、每章的论点；开头和收尾；正文与附录 | [references/L1-deck.md](references/L1-deck.md) |
| 2 章 | 章内组织模式；章首预告和章末收束；页的清单 | [references/L2-section.md](references/L2-section.md) |
| 3 页 | 标题、导语、内容块的主次、结论条 | [references/L3-slide.md](references/L3-slide.md) |
| 4 块 | 块内各项的关系、图示、分组、证据 | [references/L4-block.md](references/L4-block.md) |

另有四份不分层的文件：

- [references/wording.md](references/wording.md)：各层通用的用词规则。
- [references/review.md](references/review.md)：大纲完成后的独立评审。
- [references/layout.md](references/layout.md)：生成 pptx 时的页面版式。
- [references/components.md](references/components.md)：生成 pptx 时每种图示的可编辑画法。

四条原则：

- 上层没定，不做下层。听众的角度错了，下面每一页都要重写。
- 层级和顺序分开。上一层概括下一层，所有场景都要求；先讲结论还是先讲依据，由场景档决定。
- 概括句不是标签。“平台的六大能力”是标签；组内内容换了就不成立的那句话才是概括。
- 关系先于图示。先判定几项内容之间是什么关系，再选用什么图。

所有内容写在一个 `outline.json` 里，放在用户的工作目录。格式见 [assets/example-outline.json](assets/example-outline.json)；范例只示意字段和粒度，内容结构按当前材料来定。检查脚本的用法都写在本文件里，不必读它的源码。用户要成品文件时见“生成 pptx”一节。

## 场景档

| 这次汇报的目的 | scenario | 文件 |
|---|---|---|
| 向领导或上级说明一项工作的由来、思路、进展 | `upward-report` | [references/scenarios/upward-report.md](references/scenarios/upward-report.md) |
| 要对方批准、选择或拍板 | `decision` | [references/scenarios/decision.md](references/scenarios/decision.md) |
| 向客户说明我方理解其需求并有能力交付 | `client-proposal` | [references/scenarios/client-proposal.md](references/scenarios/client-proposal.md) |
| 让听众学会一个方法、系统或概念 | `training` | [references/scenarios/training.md](references/scenarios/training.md) |
| 以上都不贴切 | `custom` | 在 `deviations` 里写明章节顺序和理由，可参照最接近的一档 |

只读选中的那一个文件。场景档按五层给出默认值；它是起点，与听众、材料或用户的要求冲突时按实际情况改，并在 `deviations` 里写明。兼有两种目的时，按最终要达成的那一个选。

## 流程

### 1. 听众和定义卡（第 0 层）

通读材料和请求，读 `L0-audience.md`，选场景档并读对应文件，填 `brief`：

```json
"brief": {
  "scenario": "decision",
  "audience": {
    "who": "谁在听，谁拍板",
    "background": "专业和岗位；对这件事已经知道什么、不知道什么",
    "cares": "他们用什么标准判断这件事",
    "angle": "用什么角度讲他们听得懂：他们熟悉的业务、说法、参照物"
  },
  "terms": [
    {"term": "听众可能不熟的词，一个词一条", "treatment": "首现解释 / 换说法", "plain": "解释或替换成的说法"}
  ],
  "takeaways": ["听众听完要记住或决定的事，一到三条"],
  "occasion": "现场讲还是会前读，会后是否留档",
  "duration_min": 15,
  "title_style": "主题式 / 论断式 / 问题式，可以写明哪类页用哪种",
  "density": "低 / 中 / 中高 / 高",
  "external": false,
  "deviations": "偏离场景档默认做法之处和理由，没有就留空"
}
```

听众是谁、什么专业背景、对这件事知道多少，从请求和材料里看不出来时，必须先问用户，一次问完，问得具体。目的和场合不清楚、并且不同答案会导致不同结构时，也一并问。其余项用场景档的默认值。

问了就停在这里，等用户回答后再往下做。不靠猜测的听众把大纲做完：听众猜错，后面每一层都要重做。当前得不到即时回复时，把问题写在回复里，结束这一轮。

- `terms`：没有需要处理的词时写空列表。
- `duration_min`：现场讲的时长。只供会前阅读的材料不写。
- `external` 为 true 表示给单位以外的人看。

### 2. 论证树（第 1、2 层）

读 `L1-deck.md` 和 `L2-section.md`，写出：

- 主旨句 `thesis`。
- 每章的 `label`、`claim`、`pattern`、`relation`、`opens`、`closes`。
- 每章的页清单，每页只写 `id`、`role`、`title`。
- `opening`、`closing`、`appendix` 里有哪些页。

章论点和页标题从听众的角度写，里面的数字都要在材料里有出处。这一步不写页面内容。

论证树指主旨句、各章论点和页清单。检查并生成论证树视图：

```bash
python3 <本 skill 的目录>/scripts/storyline.py check outline.json --tree
python3 <本 skill 的目录>/scripts/storyline.py view  outline.json --tree > 大纲.md
```

### 3. 请用户确认论证树

把 `大纲.md` 的路径给用户，在回复里用白话写出四样东西请用户确认：

1. 对听众背景的理解和选定的讲述角度。
2. 场景和主旨句。
3. 各章论点，以及每章打算怎么展开。
4. 偏离默认做法之处。

用户确认或提出修改后，登记已确认的版本，再做下一步：

```bash
python3 <本 skill 的目录>/scripts/storyline.py approve outline.json --by user
```

默认做法是停下来：把这四样东西写在回复里，结束这一轮，等用户回复。只有两种情况不停：用户明确要求一次出完；或者这次任务是由另一个程序或 agent 派来的，并且对方说明了不会有人回复。这两种情况改用 `--by none` 登记，并在最后的回复里把这四样东西交代清楚。

登记之后论证树又有改动的，整体检查会逐条列出来。

### 4. 页和块（第 3、4 层）

先读 [references/page-plan.md](references/page-plan.md)，按它的四步为每页规划内容：定对象，定每个对象要说的两三个方面，回材料逐格取内容，再看要不要加支撑块。这一步决定成品是丰富还是发空，要对着材料做，不凭大纲的标题发挥。材料里讲到系统的组成、业务的流转、一批能力或事项的全貌时，按该文件第四节画成一张完整的架构图、流程图或能力地图，不拆成几页表格。

再读 `L3-slide.md`、`L4-block.md` 和 `wording.md`，把规划写进每页：内容块写明要表达什么、块内关系、各项内容、建议的图示、证据和出处；导语和结论条只在需要时写；现场讲的材料可加讲述要点。一页有两个以上内容块时，在 `layout` 里写版式，选项见 `layout.md` 第二节。位置和底色能说明的事不再用文字重复，不写次要的标签和说明。

### 5. 整体检查

```bash
python3 <本 skill 的目录>/scripts/storyline.py check outline.json
```

- 检查结果按层输出。从上层改起，上层改动后下层的提示可能随之消失或变化。
- 错误（E）是必填项缺失，全部改掉。
- 提示（如 L2-3）是结构上的疑点，每条都写了改法。逐条判断：该改的改；有理由保留的，在文件顶层的 `waivers` 里写明理由。改一处可能引出新的提示，这是正常的。合并或拆分会让内容变差时，保留原样并写明理由，比硬改更好。
- 脚本输出的“标题串读”读两遍：一遍按各层规则文件末尾的自检问题判断，一遍换成听众的身份读。
- 脚本列出的英文词逐个确认听众认识。
- 脚本列出的数字逐个回到材料核对出处；材料里没有的删掉或写待定，不估算、不编造。

### 6. 独立评审

```bash
python3 <本 skill 的目录>/scripts/storyline.py view outline.json > 大纲.md
```

生成 `大纲.md` 后，按 [references/review.md](references/review.md) 请一位没参与制作的评审者以听众的身份读一遍。成立的意见从上层改起，改完重跑整体检查并重新生成 `大纲.md`。

### 7. 交给用户

回复里写明：大纲文件的路径；相对已确认的论证树改了什么；评审提了哪些问题、改了哪些、没改哪些及理由；保留了哪些提示和理由；哪些内容因为材料里没有而写成了待定。用白话写，不用字段名和规则编号。

## 按用户的意见调整

先判断意见落在哪一层，改那一层，再检查它下面的各层是否仍然成立，然后重跑检查、重新生成 `大纲.md`。

| 用户说的 | 层 | 改哪里 |
|---|---|---|
| 听众不对、听不懂、太技术、角度不对 | 0 | 听众四项、术语处理；必要时换场景档。下面各层都要复查 |
| 整体逻辑不顺、重点不对、章分得散 | 1 | 主旨句、分章办法、章论点、章的先后 |
| 某一章太散、太长、前后不搭、缺一块 | 2 | 该章的组织模式、页的增删与合并、章首章末 |
| 某页不知道想说什么、标题不对、太空或太满 | 3 | 该页的标题、导语、内容块的主次、拆页或合并 |
| 图不合适、全是并列、缺依据 | 4 | 该块的关系、图示、分组、证据 |

改动只影响到的那一层和它下面的层。上面的层不动。

## 生成 pptx

大纲经用户认可后，用户要成品文件时做这一步。读 [references/render.md](references/render.md)，按它来做：

1. 选模板。用户给了模板就用用户的，先用 `scripts/template_info.py` 读出它的版式、颜色和字体，写一个风格文件；没给就从自带的三套里选一套并告诉用户。
2. 用 `scripts/pptkit.py` 写生成脚本。固定位置的文字交给模板，图示用 pptkit 的方法画，颜色只写角色名。
3. 处理保存时打印的超出提醒，再用 `scripts/preview.py` 渲染成图片逐页检查。
4. 交付时说明用了哪套模板、哪些图示是另画的、哪些没有验证过。

各图示的构成和容量见 [references/components.md](references/components.md)，字号下限和放不下时的处理见 [references/layout.md](references/layout.md)。用户要调整某个图示时，按 `components.md` 第三节处理；改动涉及内容关系或页的增减时，先改大纲再改页面。

## 修改已有 PPT 的结构

用户给的是一份已有的 PPT、要求重排结构时：把原稿每页的标题和主要内容读出来当作素材，从听众和定义卡做起，写论证树时按 `L1-deck.md` 第四节从下往上归组。原稿的章节划分只作参考。重排后向用户说明哪些页合并了、哪些页移走了、哪些内容进了附录。
