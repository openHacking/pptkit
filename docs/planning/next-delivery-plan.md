# 近期执行计划：从可信导入到高级幻灯片

本计划将 [主任务](implementation-backlog.md) 与 [具体 TODO](advanced-capability-todos.md) 组合为可领取的工作包。完成状态只在原任务和执行记录维护，此处不重复勾选。顺序按依赖和可验证价值安排，不承诺工期或完整 Office 等价能力。

## 规划入口：完善幻灯片能力的领取顺序

任务状态以 [TODO 清单](advanced-capability-todos.md) 为准；本文负责排期。以下覆盖完整能力方向，不要求一次实现，也不把研究候选当作已支持。每次领取一个可验证叶子任务，登记输入、允许操作、拒绝边界及证据。

| 批次 | 目标与任务入口 | 进入下一阶段的条件 |
| --- | --- | --- |
| A：可信导入与 round-trip | IMP-10.1.1.1–.1.3 先统一操作矩阵；收口 IMP-09.2.1.1–.1.3 的文字 owner/拓扑；并推进 IMP-07/08 的未知闭包、独立比较与三轮验证。 | no-op、fresh export、源保留编辑分别报告；不支持的操作显式阻止，未改内容有独立证明。 |
| B：日常编辑闭环 | 表格 IMP-09.2.5/09.3.4，图表 IMP-09.2.6.1–.6.4，属性/动作 IMP-09.2.7/.2.8；几何、图片与结构操作继续沿原 IMP-09/AUT 队列。 | 一种 owner/操作验收后再扩范围；图表 cache/workbook 同步，结构变更的跨对象引用可维护。 |
| C：原生 SmartArt | 先领取 ADV-02.2.1.1 差分定位与 ADV-02.1.3.1 几何基准；接最小生成、布局/缓存决策、原生显示与全操作验收，随后逐布局进入 ADV-03。 | Process/Hierarchy/Cycle 分别准入；新生成输出真正显示且能原生编辑，数据/展示/缓存不矛盾。 |
| D：动画与媒体 | ADV-06.1 的时间/事件/目标契约与 ADV-06.2.1 的原生播放取证；逐项开放 Appear/Fade/Disappear，再强调/路径/文字/图表/SmartArt 子目标。媒体按 ADV-04，切换按 ADV-05。 | 逐效果记录初态/中间态/终态、点击/重入与目标维护；headless 重开不代替实际播放。 |
| E：完整创作与播放 | 文本/国际排版、几何/图片、表格/图表、主题/多母版/模板按 AUT；预览/导航 ADV-08/12，评论/无障碍 ADV-09/10，工具与输出 ECO。 | 创建、导入、编辑、布局、预览、原生输出、保留保存逐操作验收；混合文档接 IMP-10.4。 |
| F：扩展与稳定发布 | 更多五类 SmartArt 研究 ADV-02.3，特殊内容 ADV-11/13–15；兼容/性能 REL 与 v1 范围 STB。 | 原生、保留、转换、暂缓有逐项决策；选择正式/实验/后续范围，未验证项不写入支持承诺。 |

A 的证据契约先行；C/D 的独立研究可同时准备，不必等待 B 全部完成。具体实现仍遵守原父项依赖。完整对象覆盖见下方“全部幻灯片能力的覆盖地图”。

### 当前 SmartArt 研究的明确接续

[原生几何基准 ADR0024](../architecture/decisions/0024-smartart-native-geometry-observation.md)继续推进 ADV-02.1.3.1：24 个生成候选均有独立 FODP 几何、PDF span/path、原生保存文字与图片摘要，缺/空缓存分别观察。Cycle 原生箭头存在重复名称，不能据此建立稳定节点身份；长文字样本实际裁切、PDF 中文提取有缺项，不能沿用标准样本通过结论。接续文字区域/字体/溢出策略、宽深树/不同画布输入与确定性布局；PowerPoint 已解锁；后续生成候选的实际编辑证据见 ADR0025，与此几何基准分列。

新增[七张最小原生输出与布局卡](advanced-capability-todos.md#下一步执行卡smartart-最小原生输出与布局准入)，不重复拆分已有家族实现任务。[ADR0023](../architecture/decisions/0023-native-smartart-generation-research.md)已推进最小生成与差分研究：此前的“文字未显示”结论来自 PDF 单词被拆行造成的词组误判，现已纠正。移除普通文字误设的 custT 后，[六个标准候选](../testing/smartart-native-presentation-producer-evidence.json)保存重开均保留数据清单且匹配文字词组；[Cycle 单变量/交叉外壳对照](../testing/smartart-native-differential-evidence.json)区分直接打开与另存重开，并保留图片。[ADR0025](../architecture/decisions/0025-generated-smartart-native-edit-evidence.md)已记录三布局标准缺缓存候选的 PowerPoint 识别、改文字、加节点、重排、换主题、重布局及重开；GUID Cycle 三轮文字编辑保留内容 ID。接下来领取下方新增身份/原生矩阵切片，补原生删除、结构身份维护、Process/Hierarchy 三轮及空/单/长文字/空缓存，继续布局与缓存决策。任务仍未完成。

## 本次排期：先领取哪一批

规划复核日：**2026-10-25**。先把候选实现验收清楚，再扩对象和操作。下表是本次的执行入口；后文保留各方向的完整任务链和样本要求。每批可按现有稳定 ID 单独提交，不要求一次实现整个能力面。

同日实施进展：`text-runs-v1` 已通过完整构建/类型/测试，补修活动准备期间会话释放的同步监听器清理，浏览器 ESM/global 与安装包消费者也通过。详见[执行记录](execution-record.md#direct-text-run-preservation-saves-2026-10-25)。下列第一批的自动化/文档子集已交付，接续完整源引用和原生交互闸门；不需要无变更重复运行相同检查来代替实现后续任务。

| 批次 / 优先级 | 领取的 TODO / 责任 | 第一份交付物与退出条件 |
| --- | --- | --- |
| 1. 当前文字保存收口 / P1 | IMP-09.2.1.1；关联 IMP-09.1.1–.1.3、IMP-09.5.1/.5.2 / Parser、QA | 已交付当前 profile 的允许/阻止契约、完整自动化、监听器清理修复和文档。接续任意引用边界、具名原生交互/外观证据；保持混合非法编辑全量阻止、未知 part 保留、基线/过期计划和三轮保存回归。缺原生环境保持 untested，完整父项仍开放。 |
| 2. 常用对象读取与创作闭环 / P1–P2，沿用各任务优先级 | AUT-07.1.1 → .1.2；IMP-05；AUT-05.1；AUT-06.2 / Core、Parser、Layout、格式层 | 多母版/覆盖 ADR、一个表格样式切片和一个图表格式切片分别交付，不重新开发已有网格/数据导入。继承来源可追踪、原生表格编辑和图表 Edit Data 分别留证；完整交付仍满足各自父项依赖。 |
| 3. 扩大非结构编辑 / P1 | IMP-09.2.1.2/.1.3、09.2.2–09.2.8 / Parser、相关语义责任包 | 每次只开放一种 owner 或操作，先形状/备注文字、颜色/几何/图片，再表格、图表、属性和动作。每种操作交最小源差异、未知内容/共享引用证明、负例与公开消费者；文字拓扑及动画范围无法安全维护时阻止。 |
| 4. 结构与复用 / P1–P2，沿用各任务优先级 | IMP-09.3.1 → .3.2/.3.3；09.3.4、AUT-09.2 / Core、Parser、格式层 | 页重排/删除、复制/组合、表格行列与跨文档分别验收。交身份/关系重映射表、失败原子性和多轮资源增长记录；不以新建导出代替源保留。 |
| 5. SmartArt / P2 保留、P3 研究 | ADV-01.1.1–.1.3、ADV-02.1.1–.1.3、02.2.1/.2.2 → ADV-03.4.1–.4.4 / Parser、Core、Layout、Renderer | 立即可先收集 Process/Hierarchy/Cycle 各一个具名布局与许可来源，交组件/共享资源/缓存清单。原生增删节点和重布局实验决定首个家族准入；先保留，再实现一个 Process 布局，Hierarchy/Cycle 和其余五类逐项接续。无关编辑、内部编辑、缓存预览分别给结论。 |
| 6. 媒体、切换、动画与放映 / P2 实现、P3 研究 | ADV-04/05；ADV-06.1.1/.1.2、06.2.1 → ADV-07.1.1–.1.4 → ADV-08 / Core、格式层、Renderer/适配器 | 立即可先交时间/目标 ADR、最小原生 timing 样本和独立虚拟时钟期望；原生效果按 appear、fade/disappear、组合、段落/路径/子目标推进。音视频与 fade/push/wipe 在自身依赖满足后交付，不等待全部 SmartArt。预览 seek/pause/replay、静态帧与实际原生播放分别验收。 |

贯穿各批：IMP-10.1.1 的版本化操作矩阵从第一批开始登记实际结果；IMP-10.4 的混合文档回归在对应能力具备后补齐。基础研究、源样本收集和契约设计可以提前，功能完成仍满足原任务依赖。**规划与实现进展分别留证；局部自动化通过不关闭完整父项，也不承诺日历工期。**

## 下一批具体领取项

本轮补充十张[执行卡](advanced-capability-todos.md#下一批执行卡清单播放证据与操作矩阵)，全部未完成，不另设状态台账：

| 顺序 | 任务 | 第一份可审阅成果 / 接续 |
| --- | --- | --- |
| 1 / P1 | IMP-10.1.1.1 → .1.2 → .1.3 / QA、工具层 | 先交 schema 和已有证据登记，再生成/校验兼容表，最后连接独立比较与回归。可提前准备数据契约，完整完成继承 IMP-10 闸门。 |
| 2 / P2 | ADV-01.1.1.1 → .1.2 → .1.3 → .1.4 / Parser、QA | 先交嵌套组内 diagram 对象清单，再关联四类组件与缓存/资源，最后交公开报告、双方言/共享/缺失反例。之后接 ADV-01.1.2 缓存预览，不能直接宣布节点编辑。 |
| 3 / P3 研究 | ADV-06.2.1.1 → .1.2 → .1.3 / 格式层、QA | 复用四类现有样本补观察协议，实际应用播放/编辑后作逐效果准入 ADR。缺原生环境保持 untested；可继续样本与未知分支研究，不能关闭依赖播放证据的卡。 |

以上是对顶部批次 5/6 和贯穿矩阵的进一步拆分，不替代文字保存、表格/图表语义、结构编辑的既定队列。目标仍是全能力覆盖图；每次实现选择一张有界卡。

## 接续领取顺序：数据图、图表保存与时间线

以下是对既有批次的继续拆分，具体勾选只维护在[十二张执行卡](advanced-capability-todos.md#接续执行卡数据图图表保存与时间线)。完整清单现为 **346 个唯一层级任务 ID**，包含父项，不能理解为 346 个独立交付。本次只排任务，不验收工作区候选运行时代码。

| 顺序 / 轨道 | 先领取 / 责任 | 第一份可审阅交付 | 接续与暂不放开的边界 |
| --- | --- | --- | --- |
| 1 / 跨能力证据 | IMP-10.1.1.1 / QA、工具层 | 复用已有实验建立版本化操作矩阵 schema 和首批证据登记；缺证据明确 untested。 | 接生成器/校验和独立回归；所有后续卡沿用相同操作维度。 |
| 2 / SmartArt | ADV-02.1.2.1 / Parser | 对候选源图读取交契约、独立节点/连接/文本期望和反例。复用已有关联清单、缓存预览和源样本。 | .2.2 显式映射、.2.3 累计预算 → .2.4 三家族证据/ADR → 原生编辑实验 → 首个 Process。候选读取不等于完整语义模型或可编辑图。 |
| 3 / 编辑保存 | IMP-09.2.6.1 / Parser、Core | 先交 chart/cache/workbook 源绑定和允许操作 ADR；原型可提前，写回完成仍须满足父项依赖。 | .6.2 既有数值 → .6.3 系列/类别结构；.6.4 分别验证原生和三轮结果。样式/几何/图片保存、表格样式仍沿原有任务推进，不被替代。 |
| 4 / 动画 | ADV-06.1.1.1、ADV-06.2.1.1 / Core、格式层、QA | 收口已有有限时间实验，并为四种原生样本补观察协议；纯契约与原生取证各交一份成果。 | 事件重入 → 公共身份 → 契约/原生对照 → Appear/Fade/Disappear → 可控预览。原生环境缺失不阻塞独立契约准备，也不能补签播放结论。 |

贯穿验收：创建/读取/修改/预览/fresh 输出/no-op/源编辑保存分列；正例之外必须覆盖身份歧义、共享引用、预算/取消和失败后重试。SmartArt 缓存存在不证明新鲜，workbook 与缓存都写出不证明原生 Edit Data，时间线可求值不证明播放。

本次规划反向审查覆盖五项：重复开发、未知/共享引用破坏、候选代码误报完成、headless 代替原生交互、父子任务重复计数。新增任务均未勾选；保留原有 ID 和完成记录，不新增第二份状态台账。验证结果以实际文档检查为准，不借用此前完整功能测试结果。

验证：`pnpm lint` 通过基础 lint、包依赖边界、文档链接/契约及 TypeScript 示例检查；规划检查核对 346 个唯一层级 ID、十二个新增叶子未勾选、父项/显式依赖存在及新增表格列数。此次仅编辑四份规划文档，未运行功能 build/完整测试，未新增运行时支持声明。

## 当前起点

ADV-02.1.2.4 已推进[内容森林研究 ADR0022](../architecture/decisions/0022-smartart-content-model-research.md)：精确匹配 process3/orgChart1/cycle2，保留有序多根、助手、段落与布局生成闭环；7 组离线测试含 100 个随机森林和 1,500 层深树。PowerPoint 16.104/macOS 26.5 已实际编辑 Process/Hierarchy 源副本并新增、重排节点，保存后分别生成 8/10 个缓存形状；Process 已重开，Hierarchy 重开时 Mac 锁定，Cycle 原生操作未开始。候选 ID 仅属单次投影、文字仅普通段落，不构成稳定 Core 创作 API。接续具名布局的最小生成输出、确定性几何/缓存和余下原生闸门；不能以既有源可编辑代替 PPTKit 新建原生输出准入。见[具名原生记录](../testing/smartart-native-editing.md)。

ADV-02.1.2.2 已接续[显式映射切片](../api/diagram-mapping.md)：展示节点的属性/连接证据按 data part 归一，缓存引用只按实际关联的 data/cache pair 解析；共享 owner 不重复复制，原生 process3 的两个不同目标案例不强行挑文字归属。七个固定源样本独立验证 103 个展示映射和十个缓存 shape。下一项进入 .2.4 的具名家族语义模型与布局输入决策；完整所有权/布局算法、缓存新鲜度及原生编辑仍未交付，继承闸门保持开放。详见[执行记录](execution-record.md#explicit-smartart-source-mappings-2026-10-25)。

接续实施进展：ADV-02.1.2.1/.2.3 已推进[源图检查](../api/diagram-data.md)，包括一份 data part 一份图、点/连接/直接文字、显式类型引用、raw 缓存 modelId、累计工作/诊断预算及文件/保存候选配置传递。七个固定源文件经独立 Python ZIP/XML 核对；这不是语义 Core 图或原生节点编辑。下一步优先 .2.2 的内容/展示/缓存显式映射，再 .2.4 的家族模型决策；源样本许可、完整继承依赖及原生闸门继续保持开放，不从零重写当前读取器。验证命令和最终结果见[执行记录](execution-record.md#smartart-source-data-graphs-2026-10-25)。

现有 [结构导入会话](../api/import-session.md) 已支持尺寸、部分元数据、页序、隐藏状态、原生分节成员名称的页标签投影及原字节 no-op 保存；原生分节集合/空分节身份仍未映射到 Core。结构投影尚不包含完整绘图内容；不得用它生成空白页面再称为语义 round-trip。另有显式[内容投影](../api/content-import.md)支持直接富文本、七种预设形状、嵌套组、内嵌 PNG/JPEG/GIF、有界直线连接关系与点击动作，以及[直接备注正文](../api/notes-import.md)，包含来源/未映射属性报告。更广图片/连接器/动作语义、备注页及完整继承仍待完成。ZIP/XML/OPC/MCE、文档属性与原生分节检查可以复用，后续工作聚焦剩余语义转换、编辑保存和原生验证。

2026-10-25 规划复核：新增三级 TODO 均保持未勾选。[直接表格切片](../api/table-import.md)已有定向/跨包/浏览器/安装包及原生生产器清单证据；原生表格样式、完整继承、交互应用验证与保留式编辑保存仍未交付，父子项依赖按原闸门验收。本文安排工作包与覆盖范围；任务状态仍只在 [TODO 表](advanced-capability-todos.md) 和[主 Backlog](implementation-backlog.md)维护。

[有界图表导入](../api/chart-import.md)已有单家族 2D bar/line/pie、缓存/workbook 优先级、三轮独立数据、浏览器/安装包及原生生产器证据；完整数据/格式、交互 Edit Data、继承和保留保存仍未交付。后续细分已覆盖页/对象复制删除与组合、跨文档迁移、音视频/字幕生命周期、五类后续 SmartArt 研究、逐效果动画与子目标、国际排版、新图表族、评论/无障碍和输出验收。先领取下表中的最小切片；后续子项是候选排队范围，不要求在一次交付中全部实现。

## 全部幻灯片能力的覆盖地图

[占位符几何切片](../api/placeholder-inheritance.md)现已按源 index/type 匹配多母版链，记录逐属性 EMU 来源并投影本地坐标；[文字样式切片](../api/text-style-inheritance.md)现在覆盖有界 frame/段落/列表/run 属性、主题色板映射与 Latin major/minor 字体，记录来源/解析依赖并物化为本地 Core 值。二者都不复制母版提示文字。测试覆盖歧义、缺省、零值、MCE/方言、九级列表、未知高优先级阻断、多主题、映射覆盖、三轮 fresh 导出和公开消费者。接续工作仍是 AUT-07 的 Core 创作契约、原生非中性变换证据、主题色修饰符、非 Latin 脚本字体、背景/资源和保存重绑定；当前切片不关闭父子 TODO，也不提供模板重绑定或保留式编辑保存。

“完善”按对象、操作和交付链路衡量：同一对象的创建、导入、修改、布局、预览、原生导出、no-op、编辑保存分别验收。下面是覆盖检查入口，不是所有能力已实现或全部必须进入 v1 的承诺。

| 能力面 | 待补范围 | 任务入口 / 关键完成证据 |
| --- | --- | --- |
| 文档结构与属性 | 元数据、自定义属性、分节身份/空分节、隐藏页、备注及备注母版 | IMP-03.3–03.6、ADV-12；区分页标签与原生集合，保留完整源身份。 |
| 文本与国际排版 | 列表/制表位、字符段落度量、RTL/竖排、自动适配、跨框/跨页续排 | AUT-01.1–01.5、AUT-01.1.1/01.1.2、AUT-01.2.1、AUT-01.4.1；固定字体与缺字体用例，文字不丢不重复，原生可编辑。 |
| 形状、连接与组合 | 更多预设/调节点、自由路径/孔洞、直折曲连接器、组坐标和动作 | AUT-02.1–02.3、IMP-04.3；移动目标仍附着，嵌套变换只应用一次。 |
| 填充、描边与效果 | 渐变/图案/阴影、透明度、裁切、效果叠加 | AUT-03.1/03.2；组合绘制顺序及可见/命中边界，近似结果有诊断。 |
| 图片与资产 | 格式/签名策略、字节来源/解析、裁切替换、共享资源/清理 | AUT-04.1–04.4；浏览器/Node 等价、预算/取消、未知引用不误删。 |
| 表格 | 网格/合并导入、富文本、逐边边框/主题、行列编辑、测量/分页 | IMP-06.1.1–06.1.5、AUT-05.1/05.2；复用已有 LAY-07/08，非法合并原子失败。 |
| 图表 | 已有 bar/line/pie 导入，数据同步、轴/格式、scatter/bubble/area/doughnut/combo，专项图表研究 | IMP-06.2.1–06.2.5、AUT-06.1.1–06.1.3、AUT-06.2/06.3；独立读 workbook/缓存，实际编辑数据重开。 |
| 主题、母版与模板 | 多母版、占位匹配、部分文字/主题属性继承、颜色字体、背景、模板替换及动态字段 | AUT-07/08、IMP-05.1.1–05.1.4；已有主题色映射与 Latin 字体子集见 API 契约，完整本地覆盖可追踪、无重复继承对象，字段求值无隐式时钟。 |
| 布局与文本流 | 共享文字排布、国际文本度量、通用文本流与已有表格分页衔接 | AUT-01.3–01.5、LAY-01–08；保留已完成布局基线，不把测量等同原生排版一致。 |
| 导入与 round-trip | 方言/MCE、对象报告、保留闭包、比较器、逐类编辑保存、有损重建、复制/删除/跨文档 | IMP-01–10、IMP-09.3.1–09.3.3、AUT-09.2.1/09.2.2；逐操作损失报告、独立验证、三轮重开与资源增长检查。 |
| SmartArt | 识别/缓存/保留，具名 Process/Hierarchy/Cycle 原生编辑，更多五类候选研究 | ADV-01–03、ADV-02.3.1–02.3.5；数据/布局/缓存一致，原生增删节点与重布局，形状组转换另报损失。 |
| 音视频与字幕 | 封装、封面、开始/裁剪/音量/循环、浏览器播放、字幕/语言、故障清理 | ADV-04.1–04.5 及其三级切片；原生实际播放、编解码器/权限失败、离页释放与共享资源验证。 |
| 切换与动画 | fade/push/wipe、时间线、进入/强调/退出/路径、段落/图表/SmartArt 目标、媒体同步 | ADV-05–07、ADV-07.2.1–07.2.5、ADV-07.6.1/07.10.1/07.4.1；逐效果参数与目标身份，原生播放和编辑后引用维护。 |
| 放映与可控预览 | 时钟/seek/pause/replay、动作事件、导航、自定义放映、计时、讲义/打印 | ADV-08、ADV-12.1/12.3–12.5；虚拟时钟状态机与实际放映分别验收。 |
| 评论与无障碍 | 评论方言/作者/回复/锚点、隐私清理、阅读顺序、对比度/描述、键盘/字幕/reduced-motion | ADV-09.1.1/09.2.1、ADV-10.1.1/10.1.2/10.2.1；删除/复制不串锚点，消费者与辅助技术验证。 |
| 特殊 Office 内容 | OLE/控件/宏、嵌入字体、加密/限制/签名、公式、3D、墨迹 | ADV-11、ADV-12.2、ADV-13–15；先研究识别/保留/拒绝和格式策略，不执行嵌入内容。 |
| 输出与工具集成 | CLI、SVG 导入、portable SVG/图片/PDF、命中测试/选择、撤销重做/事件、插件与示例 | ECO-01–09；安装包消费者、独立查看器与边界失败测试，UI/协同网络归应用层。 |
| 兼容性与稳定发布 | 原生应用矩阵、性能/体积预算、API/IR 迁移、v1 范围与发布演练 | REL-02、STB-01–04；具名环境证据，选定 v1 必需/实验/后续范围。 |

## 开工前的证据收口

2026-10-25 已完成当前主题色/Latin 字体切片的自动化收口：独立核对源主题/输出色板后，fresh export 按有效颜色比较，保留保存继续验证原字节身份；新增自定义色板、错误颜色和其他样式变化反例。此前 `.tmp/theme-text-import.log` 中的两项表示差异已修正，完整 `pnpm test` 通过，包含 Parser 331、浏览器 bundle 23、导入集成 19 及真实安装包消费者。当前记录见[主题解析验收](execution-record.md#theme-color-and-latin-font-projection-2026-10-25)，最终日志为 `.tmp/theme-final-full-test.log`。

部分 themeOverride 也已按目标 slide/layout/master 链分别继承颜色与字体方案，并验证空组件、未知高优先级阻断、多母版共享、MCE 和错误关系。新增[原生实验](../testing/theme-text-import-evidence.json)记录了当前 LibreOfficeDev 对该具名样本的颜色/字体改写与源/fresh PDF 样式差异，因此不关闭原生保真闸门；PowerPoint 验证仍未完成。接续 IMP-08.1.1 的完整比较规则与 IMP-08.1.2/.1.3 包图/破坏验证，以及 Core 多母版创作契约。色彩变换、脚本字体和背景资源仍由 IMP-05.1.3.3.1–.3.3 单独验收。

## 下一轮可领取的最小工作

IMP-08.1.1–.1.3 已推进[独立源包比较器](../api/package-comparison.md)：精确载荷为默认，保守 XML 结构比较显式选择，未知/MCE 原分支、命名空间绑定、注释/处理指令及全部二进制内容参加比较。已有独立 ZIP/SHA256、浏览器 DOM、安装消费者与三轮 no-op 证据；源包比较不等于 fresh 有效值相等或原生编辑验证。接续完整有效颜色/字体/数值、身份双射、对象/workbook 破坏矩阵及原生闸门，这三个任务继续未完成。

IMP-07.1.1 已推进[原始关系闭包](../api/preservation-closure.md)：包含未选中分支、所属关系 part、共享/循环目标、外部入站引用与不可达清单，未知/缺目标/潜在身份冲突影响分方向完整性。已有独立生成图模型、DOM、图片/图表/备注和 no-op 证据。接续保留状态及混合 XML 最小修改边界，不能用关系清单代替全部 OOXML 引用证明，也不能据不可达清理资源；该任务仍未完成。

IMP-07.2.1 已推进[私有字符修改基础](../architecture/decisions/0007-xml-character-edit-boundaries.md)：以原始元素序号、展开名称与预期文字定位字符区域，保留周围词法 XML 和编码，拒绝目标内部混合子节点。独立字符串/编码、64 组生成样本和双方言三轮重打包检验已有测试。[`text-runs-v1`](../api/text-edit-save.md)现已接入 source→Core run、范围引用/共享源保护、修改分类和候选写包/会话，并通过完整自动化与文档同步；默认 edit 与 rebuild 仍受阻止。接续 IMP-09.2.1.1 的完整引用/原生闸门，更广 owner/拓扑按 .1.2/.1.3 独立推进；不重复开发现有实现或把局部证据当作整个 IMP-09 已交付。

IMP-05.1.3.3.1 已推进[文字颜色变换与 alpha 切片](../api/text-color-transforms.md)：支持具名 11 种变换、透明文字创作/预览/导出、来源预算和三轮 fresh 往返。[原生九例](../testing/color-transform-import-evidence.json)中灰色 shade 有差异，保留写回也未交付，因此该项及父项继续待完成。接续任务包括 shade 跨应用色彩空间/精度调查、非 Latin 字体、背景和多母版契约；不重复开发当前文字子集。

下表细化本次批次入口，任务之间可做独立样本与契约准备。实现可以逐切片进行；最终验收仍满足上级依赖。P0/P1/P2 沿用主表含义，表内先后顺序不能代替依赖闸门。

| 顺序 | 切片 / 优先级 / 责任 | 本轮可审阅成果 | 开工准备与验收闸门 |
| --- | --- | --- | --- |
| 1 | IMP-09.2.1.1 / P1 / Parser、QA | 已完成当前文字 profile 的完整自动化、生命周期修复与文档；接续完整引用/原生证据 | 完整父项闸门仍开放，按顶部第一批收口；扩 owner/操作时新增对应验证。 |
| 2 | IMP-07.1.1、IMP-07.2.1 / P1 / Parser | 未知资源闭包、混合 XML 修改边界与破坏注入样本 | 复用已有分析器/候选 writer；接续完整源引用证明和操作级原生证据，不能仅复制 ZIP part。 |
| 3 | IMP-08.1–08.3 / P1 / Parser、QA | 可独立发现缺资源/错关系/重复 ID 的比较与重开证据 | no-op 与 fresh export 分开报告，保存结果不能只由同一 writer/reader 自证。 |
| 4 | IMP-05.1.1–05.1.3 / P2 / Parser、Core | [源匹配/坐标来源](../api/placeholder-inheritance.md)与[有界文字/主题级联](../api/text-style-inheritance.md)已落地；接续 Core 多母版/覆盖契约、原生变换样本、主题变换/非 Latin 字体/背景及完整重绑定 | Core 创作契约由 AUT-07 先明确；当前 Parser 投影只生成有限本地有效值，不闭合 AUT-07 或 IMP-05 父子项。原始 XML 只由 Parser 持有，导入重绑定在后续切片验证。 |
| 5 | IMP-06.1.4/06.1.5、AUT-05.1 / P2 / Parser、Core | 在已有直接表格切片上补主题样式/继承、逐边边框和原生交互验证 | 网格/合并/文字不重新开发；样式模型扩展遵守 AUT-05.1，继承接 IMP-05；关闭父项仍满足 IMP-04、IMP-08.3。 |
| 6 | IMP-06.2.1/06.2.2 / P2 / Parser | [图表清单/优先级及有界直接导入](../api/chart-import.md)已有证据；接续 AUT-06.2 格式/标签、原生交互编辑与完整继承 | 保留未完成状态；不先扩图表族，后续数据映射按 AUT-06.2 和 IMP-06.2.3–06.2.5 验收。 |
| 7 | ADV-01.1.1–01.1.3 / P2 / Parser / Renderer | SmartArt 组件清单、缓存边界及三类具名源样本 | 样本研究可提前；保留声明等待 IMP-07/08，无关编辑声明另待 IMP-09.3；家族原生编辑需 ADV-02 决策。 |
| 8 | ADV-06.1.1/06.1.2、ADV-06.2.1 / P3 研究 / Core / 格式层 | 时间线/目标 ADR、最小原生样本与播放证据 | 先明确模型与保存边界；实现按 appear→fade/disappear→组合→文本/路径/子目标扩展，预览时钟随后验收。 |

上述大项的下一张任务卡按以下入口领取；不额外维护完成勾选，也不跳过父项依赖。研究/样本准备可以提前进行，功能完成声明仍等待全部闸门。

| 任务入口 | 先交付的材料 | 验收结果 / 接续任务 |
| --- | --- | --- |
| IMP-08.1.1 | 有效颜色相等/不等与三轮往返已有定向证据；接续完整比较规则及未知属性反例 | 不把当前测试辅助函数当作通用比较器；随后 IMP-08.1.2/.1.3 验证完整包图及破坏检测。 |
| AUT-07.1.1 | 多母版/布局/主题身份 ADR、旧 layout API 兼容样本 | 不先扩 Core XML；接 .1.2 覆盖求值、.1.3 artwork 与 .1.4 原生验证。 |
| IMP-09.2.1.1 | 现有候选的修改分类、source→Core 关联、同值/还原、多计划/取消和字节保留证据清单 | 沿 IMP-09.1/.5 复验最新实现；通过后接 .1.2/.1.3 与 .2.2–.2.8，不从零重做已接线部分。 |
| IMP-09.5.1 | 复核候选中多计划/过期/跨会话/取消状态表与故障注入点 | 接 .5.2 完整候选验证、.5.3 Node 原子替换；损失绑定接 IMP-09.6.1/.6.2。 |
| ADV-02.1.1 | Process/Hierarchy/Cycle 各一个准确布局标识及源组件清单 | 接 .1.2 数据/视觉图、.1.3 缓存实验、ADV-02.2.1/.2.2 原生实验与家族准入。 |
| ADV-07.7.1 | 独立时间数值期望与边界采样；研究可先行 | 基础效果完成后验收 .7.2 重复/终态、.7.3 同目标合成；触发与复制删除走 ADV-07.3.1–.3.3。 |

输入策略与页面语义的下一张任务卡也已拆出，适合在主队列之外先准备契约与样本。以下全部为待办入口，不增加一份完成状态，也不把页面语义排在首次文字编辑保存之前作为额外整批闸门；具体完成依赖仍看原任务。

| 方向 | 首张任务卡 → 后续切片 | 首次可审阅成果 |
| --- | --- | --- |
| 方言与文件类型 | IMP-02.3.1 → 02.3.2/02.3.3 | 按对象/操作列 Strict/Transitional 与模板/放映/宏启用输入策略；只完成分类不能宣称可编辑或可保留。 |
| 背景与母版 | AUT-07.1.3.1 → 07.1.3.2/07.1.3.3 → IMP-05.1.3.3.3 | 背景继承/清除、artwork 可见性和共享资源；预览与原生归属一致，不把继承图形复制成页内对象。 |
| 模板替换与字段 | AUT-07.2.1 → 07.2.2/07.2.3；AUT-08.1.1 → 08.1.2/08.1.3 | 占位重绑定、逐属性覆盖、原子回滚；页码/日期的显式求值上下文和原生字段分别验收。 |
| 备注页与分节 | IMP-03.3.1 → 03.3.2/03.3.3；IMP-03.5.1 → 03.5.2/03.5.3 | 正文以外的备注页几何/母版，以及原生集合/空分节/边界编辑；各有 fresh 和独立重开证据。 |
| 属性与自定义值 | IMP-03.4.1 → 03.4.2/03.4.3 | 源词法/标准化边界、自定义属性准入决策与公开消费者；复杂值和精度损失明确报告。 |

媒体、基础切换、常用创作能力可在自身依赖满足后推进，不必等待所有 SmartArt 家族完成。研究结果只解锁对应实现切片；没有原生测试环境时先交样本与契约，将 native 验收保留为 untested。

以下是首批闭环之后的接续队列；责任和优先级继承父项，各包内部仍按依赖推进。

| 方向 | 首个切片 → 接续切片 | 首次验收边界 |
| --- | --- | --- |
| 结构编辑与复用 | IMP-09.3.1 → 09.3.2 → 09.3.3；Core 跨文档走 AUT-09.2.1/09.2.2 | 页重排/删除、复制、组合分别开放；未知引用不能安全重映射时阻止操作。 |
| 媒体 | ADV-04.1.1 → 04.1.2/04.1.3 → 04.2.1/04.2.2 → 04.3.1/04.5.1 | 首个音频格式/编码、视频与封面、播放设置、适配器及释放；字幕决策接 ADV-04.4.1/04.4.2。 |
| SmartArt | ADV-01.1.1–01.1.3 → ADV-02 → ADV-03.4.1–03.4.4 | 首个具名 Process 布局通过后接 Hierarchy/Cycle；ADV-02.3.1–02.3.5 是后续五家族研究。 |
| 原生动画 | ADV-07.1.1–07.1.4 → ADV-07.2.1–07.2.5 | 基础出现/消失后逐效果加入旋转、缩放、路径、飞入/擦除；段落、图表和 SmartArt 子目标分别验收。 |
| 常用创作 | AUT-01.1.1/01.1.2、AUT-01.2.1/01.4.1；AUT-06.2 → AUT-06.1.1–06.1.3 | 国际排版与图表数据各自闭环，再扩更多图表；不影响已有确定性默认行为。 |
| 审阅与交付 | ADV-09.1.1 → 09.2.1；ADV-10.1.1/10.1.2 → 10.2.1；ECO-04.2.1 → 04.3.1/04.4.1 | 评论锚点、阅读顺序、可访问播放和 SVG/图片/PDF 输出各有独立消费者证据。 |

## 下一次实施的启动卡

以下是各方向的启动材料，不另设勾选状态。实际批次优先级见顶部，本表的顺位表示材料阅读顺序。每次选择一张；优先级与责任继承主 Backlog。第一份材料可以先准备，功能关闭仍须满足原任务的全部依赖。日期为规划复核日 **2026-10-25**，不承诺交付日。

| 顺位 / 稳定任务 ID | 要解决的问题 / 第一份交付物 | 必含反例与完成证据 | 下一张卡 |
| --- | --- | --- | --- |
| 1 / IMP-08.1.1 | 已有[精确载荷/保守 XML 源包比较](../api/package-comparison.md)，接续数值单位、有效颜色/字体和 fresh 输出规则；不把源相等与有效值相等混用 | 相同外观但主题来源不同、错误 alpha、错 workbook 值、截断未知节点；已有源包破坏证据，接续完整语义/身份变换和原生矩阵 | IMP-08.1.2 剩余身份/摘要证明 → IMP-08.1.3 完整破坏注入 |
| 2 / IMP-07.1.1 | 已有[全部源分支闭包与入站引用分析](../api/preservation-closure.md)，接续保留状态、任意源引用边界及操作级证明 | 未知对象引用已知共享图片、互斥/可能共存 ID、未选分支缺目标与超预算已有反例；继续验证 XML 修改区域与原生保存 | IMP-07.2.1 混合 XML 修改边界 |
| 3 / IMP-09.2.1.1 | 核对已有候选如何定位最小保存区域；先提交当前操作分类、source→Core 与剩余验收差距 | 重复设置同值、改后还原、替换整段、未知样式兄弟节点，最新并发取消/释放；不能只凭文档整体脏标记重建整页 | 完整验证后接 IMP-09.2.1.2/.1.3 与 .2.2–.2.8 |
| 4 / ADV-02.1.1 | 明确首批 SmartArt 究竟是哪一种布局；先收集 Process/Hierarchy/Cycle 各一个准确布局标识及有许可来源的组件清单 | 缺缓存、旧缓存、一对多数据/绘图身份；没有原生编辑环境则记录 untested，继续独立研究 | ADV-02.1.2/.1.3 → ADV-02.2.1/.2.2；保留实施走 ADV-01 |
| 5 / ADV-06.1.1 | 动画的时间与目标语义必须先确定；先交时间线 ADR、虚拟时钟输入和独立期望状态 | 零时长、嵌套串并行、同一时点点击、无限重复、非法目标；不以浏览器回调顺序定义行为 | ADV-06.1.2 → ADV-06.2.1 → ADV-07.1.1 |
| 6 / AUT-01.3.1 | 补齐日常文本度量的单位与覆盖规则；先交字段/缺省/继承/兼容表和正反例 | 混合字号、上下标、绝对/比例行距、清除覆盖、缺字体；保持当前默认输出兼容 | AUT-01.3.2 共享排布 → AUT-01.3.3 断行 |
| 7 / AUT-05.2.1 | 表格插删必须有确定内容归属；先交事务规则和跨合并区域的输入/输出样本 | 插入穿过合并区、删除锚点、最后一行/列、操作失败；文字/span/尺寸一起提交或一起回滚 | AUT-05.2.2 合并拆分 → AUT-05.2.3 测量/引用 |

新增的[编辑保存对象任务](advanced-capability-todos.md#编辑保存的对象覆盖补齐)补足以下接续链：文字 owner/拓扑 → 表格非结构保存 IMP-09.2.5 → 表格结构保存 IMP-09.3.4；图表缓存/workbook 保存 IMP-09.2.6、属性保存 IMP-09.2.7、动作/描述保存 IMP-09.2.8 各按自己的依赖领取。IMP-10.1.1 为所有方向记录实际操作矩阵，避免把 fresh export 或一个文本框成功泛化为全面 round-trip。

## 新增切片的接续安排

新增的 27 张卡在 [具体 TODO](advanced-capability-todos.md#继续拆分日常创作的完整链路) 与[高级内容一致性与回归](advanced-capability-todos.md#继续拆分高级内容一致性与回归)维护。每个方向已有父项，下面只安排顺序和退出物。

| 方向 | 排队顺序 | 退出物 / 边界 |
| --- | --- | --- |
| 文字、几何与效果 | AUT-01.3.1 → .3.2 → .3.3；AUT-02.1.1 → .1.2；AUT-03.1.1 → .1.2/.1.3 | 共享排布、断行和有限几何/效果子集；复杂文字、孔洞、透明/模糊越界有独立样本，跨效果组合接 AUT-03.2。 |
| 图片与表格 | AUT-04.3.1/.3.2/.3.3；AUT-05.2.1 → .2.2 → .2.3 | SVG/动画图片/旧矢量格式各自决定策略；行列/合并编辑原子化并复用测量分页。不会把已有 PNG/JPEG/GIF 路径重新列为从零实现。 |
| SmartArt 一致性 | ADV-03.7.1 → .7.2 → .7.3，与首个准入家族共同交付 | 修改后数据、原生组件和缓存同步或明确失效；生成失败可重试，原生重布局后仍可再导入。缓存事务不依赖尚未完成的整个家族父项，避免互相等待。 |
| 放映组合 | ADV-12.5.1 → .5.2 → .5.3 | 事件优先级、中断/页面重入和资源释放；最终用同一事件脚本分别验证虚拟时钟与原生播放。基础效果本身不等待这组集成。 |
| 混合往返与预算 | IMP-10.4.1 → .4.2 → .4.3 → .4.4 | 先定义每个操作的预期，再验非结构/结构编辑、多轮与性能；没有安全引用映射的操作明确拒绝。只对选入子集承诺原生内部编辑。 |
| 检查与复现工具 | ECO-09.1.1 → .1.2 → .1.3 | 版本化报告、安装后 inspect/compare、最小复现与脱敏边界决策；CLI 复用公开解析/比较入口。 |

## 端到端交付阶段

上表选择下一轮的最小切片；下表说明这些切片最终如何组成完整能力链路，供跨阶段审阅使用。

| 顺序 / 工作包 | 关联 TODO | 责任与依赖闸门 | 可审阅的最小交付 |
| --- | --- | --- | --- |
| 1. 明确输入与保存边界 | IMP-02.3、IMP-01.4、IMP-03.6 | Parser / 格式层；沿用 IMP-01 ADR，落实父任务剩余验收 | 一份方言/扩展支持表及公开消费者测试；明确结构模式、语义模式与不可保存原因，覆盖释放、取消、失败重试。 |
| 2. 首个可编辑导入切片 | IMP-04.1、IMP-04.2、IMP-04.4、AUT-04.1 | Parser / Core；IMP-03 与各子项依赖满足后验收 | 一个含富文本、预设形状、图片与嵌套组的具名样本；可编辑对象有来源映射，未知属性逐项报告，Node/浏览器一致。直接文字/形状/组和内嵌栅格切片已有实现，后续补继承、图像格式/效果及跨生产器证据，不重复开发。 |
| 3. 保留闭包与比较器 | IMP-07.1–07.3、IMP-08.1–08.3 | 格式层 / QA；保留闭包研究可与语义研究交错开展，完整关闭仍遵守主表依赖 | 未知节点与共享资源保留；比较器能抓出人为注入的破坏；no-op 同时验证字节或允许的语义差异和原生重开。 |
| 4. 第一种编辑保存 | IMP-09.1、IMP-09.2、IMP-09.4–09.6 | Parser / Exporter；完成 IMP-08、IMP-05/06 及相关子项闸门 | 首先仅改一个文本框文字；列清可编辑子集和禁止操作，生成绑定快照的计划与损失报告；失败后源文件和会话可重试。后续独立添加颜色/位置/图片替换。 |
| 5. 结构编辑闭环 | AUT-09.1/09.2、IMP-09.3、IMP-10.1–10.3 | Core / 格式层 / QA；非结构编辑通过之后 | 复制、重排、删除分别交付；链接、分节、评论和时间线重映射或显式拒绝；至少三轮保存检查 ID、内容及资源增长。 |
| 6. SmartArt 保留 | ADV-01.1–01.3 | Parser / 格式层；IMP-07/08，混合编辑另需 IMP-09.3 | 具名 process/hierarchy/cycle 输入的组件清单、缓存预览与无关编辑保存；缺缓存有占位，不能标为可编辑。 |
| 7. 首个原生 SmartArt 家族 | ADV-02.1/02.2、ADV-03.1/03.4/03.7 | Core / Layout / 格式层 / Renderer；家族可行性闸门通过 | 一个具名 process 布局，节点增删/重排、长文字与换主题；原生应用仍可重布局，缓存与数据一致。Hierarchy、cycle 后续逐个验收。 |
| 8. 媒体与基础切换 | ADV-04.1–04.5、ADV-05.1 | Core / 格式层 / Renderer；资产和语义导入依赖满足后 | 音频、视频、fade、push、wipe 分别建立样本；覆盖封面、播放参数、加载失败、清理和未知切换保留，每项独立交付。 |
| 9. 原生动画再到预览 | ADV-06、ADV-07.1–07.10、ADV-08.1–08.4 | Core / 格式层 / Renderer；先时间线 ADR，再按效果完成依赖 | 先整对象 appear/fade/disappear，再段落、强调/路径、图表/SmartArt 目标；每批实际播放留证。预览注入时钟，seek/pause/replay 与静态模式可验证。 |
| 10. 混合文档与交付工具 | IMP-10.4、ADV-12.5、ECO-02、ECO-09 | QA / CLI；对应高级功能及 CLI 父任务依赖满足后 | 含高级对象的混合文档回归、可读/机器可读能力报告、安装后的 inspect/import/save 命令；每个损失可追到对象和操作。 |

上述工作包可能跨多个提交；每次提交只选择一个可独立验收的对象、操作或效果。研究可以提前收集样本，不能越过依赖宣告功能完成。

## 同步补齐日常创作能力

- **文本与模板**：AUT-01.1–01.5、AUT-07.1/07.2、AUT-08.1，优先列表/段落度量和多母版继承；自动适配、竖排/双向、文本续排分别验收。
- **形状与图片**：AUT-02.1–02.3、AUT-03.1/03.2、AUT-04.1–04.4，区分原生连接器、自由路径、外观组合和资产生命周期。
- **表格与图表**：AUT-05.1/05.2、AUT-06.1–06.3，先结构编辑和数据一致性，再扩图表族；每类原生可编辑性独立证明。
- **审阅、无障碍与交付**：ADV-09.1/09.2、ADV-10.1/10.2、ADV-12.1–12.5、ECO-04.1–04.4，评论、键盘/阅读顺序、字幕、打印、SVG/图片/PDF 各有独立验收。
- **研究储备**：ADV-02.3、ADV-05.2、ADV-11、ADV-13–15，更多 SmartArt、Morph/Zoom、OLE/宏、嵌入字体、公式、3D、墨迹先交实验与 ADR，再拆实现，不阻塞基础版本。

## 首批样本与失败用例

样本 ID 为待建立的计划名，不能当作已有测试证据。记录生产应用及版本、来源许可、输入摘要、预期对象/操作能力；PowerPoint、LibreOffice 与 PPTKit 样本分别维护。来源不可用时标为 untested。

| 计划样本 | 必含内容 | 必验失败与边界 |
| --- | --- | --- |
| import-editable-basics | 富文本/列表、形状、图片、嵌套组、隐藏页 | 空段、混合字体、缺资产、未知同级对象、旋转/翻转、无效引用 |
| import-table-grid-merges | 普通/二维合并表、整行纵向合并、空但有样式的格、直接边框与点击 | 越界/重叠/非空 continuation、frame/grid 不一致、稀疏格、未知主题样式 |
| import-chart-data-conflict | bar/line/pie、嵌入 workbook、缓存一致/冲突/缺失、共享数据源 | 外链不读取、公式/错误值策略、坏 workbook、系列删除/重排后残留缓存 |
| import-multi-master-inheritance | 多 master/layout、主题字体配色、占位本地覆盖及普通形状 | 匹配歧义、缺主题、重复继承、清除覆盖、换布局溢出 |
| import-dialect-kind-policy | Strict/Transitional 的同语义样本、模板/放映/宏启用类型、MCE 双分支 | 类型冲突、错误命名空间、未知扩展、未选分支来源；对应 IMP-02.3.1–02.3.3 |
| template-background-rebind-fields | 两母版共享图片、继承图形开关、本地背景覆盖、换模板及页码/日期 | 背景清除、重复 artwork、歧义占位、失败回滚、隐藏/续页编号和固定时区；对应 AUT-07.1.3.1–.3.3、AUT-07.2.1–.2.3、AUT-08.1.1–.1.3 |
| notes-page-sections-properties | 共享备注母版/缩略图、空与重名分节、带时区/精度边界的自定义属性 | 多正文、缺母版、跨节移页、分节与页面删除区别、复杂属性拒绝和原生重开；对应 IMP-03.3.1–.3.3、IMP-03.5.1–.5.3、IMP-03.4.1–.4.3 |
| theme-effective-values | 两母版同主题槽位、直接 RGB/主题引用、变换组合、Latin/CJK/complex-script | 有效颜色相等/不等、未知变换阻断、缺字体、脚本选择、源语义不能被归一化比较吞掉；对应 IMP-08.1.1、IMP-05.1.3.3.1/.3.2 |
| preserve-mixed-extensions | 已知对象和未知扩展共享资源 | 缺失关系、重复身份、循环关系、删除后误清理、预算超限 |
| package-comparator-mutations | 正确包与每次只破坏一个属性/关系/资源的派生包 | 错颜色、错 workbook 值、删资源、错接关系、重复 ID、截断未知节点均被独立发现；对应 IMP-08.1.2/.1.3 |
| edit-save-text | 仅修改一个文本框，周围保留未知效果 | 计划过期、损失未授权、取消、并发准备、失败后重试 |
| save-transactions-loss-consent | 同值/改后还原、多会话多计划、候选验证和具体损失清单 | 跨计划授权、并发编辑、验证/写入/重命名失败、取消后重试、源文件不变；对应 IMP-09.1.1–.1.3、IMP-09.5.1–.5.3、IMP-09.6.1/.6.2 |
| smartart-process-native | 一个具名 process 布局及缓存 | 长文字、空图、增删节点、缺/旧缓存、无关文字修改 |
| smartart-hierarchy-native | 一个具名 hierarchy 布局、多层树、兄弟及子树移动 | 循环/多父、助手节点、深/宽树、原生重布局及主题切换 |
| smartart-cycle-native | 一个具名 cycle 布局、顺序、起点与闭合边 | 空/单节点、方向与重排、长文本、数据与缓存不同步 |
| motion-media-native | 整对象/段落动画、切换、媒体 | 目标删除、重复点击、零时长、seek、跳页、缺编解码器 |
| media-assets-lifecycle | 首个具名音频/视频格式与编码、独立封面、共享资源、嵌入/外链 | 缺封面、坏签名、超预算、权限拒绝、取消、重复挂载、离页释放及最后使用者清理 |
| media-captions-trim | 多语言字幕轨、源时间码、媒体裁剪段与循环 | 无效时间、缺轨、seek/循环/切换语言后字幕偏移、外链默认不读取 |
| animation-paragraph-targets | 多级列表、空段、混合 run、emoji/组合字符 | 插删段落、拆合 run、整段替换后目标错位、回放与 seek 不一致 |
| animation-effects-targets | spin/grow-shrink、直线路径、fly/wipe、图表系列与 SmartArt 节点 | 原始旋转/组变换、反向/重复边界、数据或节点重排、原生重布局后目标错位 |
| animation-time-composition | 同目标透明度/位移/旋转组合、嵌套串并行、触发事件和未知分支 | 零时长、小数/无限重复、反向、边界前后采样、seek 与顺序求值一致、复制删源；对应 ADV-07.3.1–.3.3、ADV-07.7.1–.7.3 |
| cross-document-theme-copy | 两份不同尺寸/主题/母版的文档、相同字节不同名称的资产 | 占位冲突、ID 冲突、未知引用、复制后删源、失败回滚、重复粘贴资源增长 |
| review-accessible-output | 评论/回复、阅读顺序、装饰对象、表格与图表描述、独立输出 | 缺作者、复制删除锚点、对比度不可判定、动画隐藏内容、键盘导航、PDF 字体/文字提取 |
| transitions-advance | fade/push/wipe、多方向、点击与定时推进 | 零时长、中断、未知切换、动画未结束与隐藏页组合 |
| morph-zoom-research | 跨页对象匹配、页面/分节/摘要导航对象和返回路径 | 改名/增删对象、删页/改分节、嵌套返回、原生播放与预览差异；对应 ADV-05.2.1/.2.2，仅研究样本 |
| mixed-round-trip | 主题、图表、SmartArt、评论、备注、媒体、动画 | 三轮保存、重排还原、复制删除、资源增长、原生重开/播放 |
| text-layout-breaks | 混合字号/上下标、行段距、CJK/emoji、软连字符及特殊空白 | 缺字体、字素拆分、长 URL 溢出、默认行为变化；对应 AUT-01.3.1–.3.3 |
| geometry-effects-images | 多子路径/孔洞、渐变/图案/阴影、SVG 配对回退与方向元数据 | 绕向错误、透明色标、模糊裁切、回退不同步、缺转换器和预算；对应 AUT-02.1、AUT-03.1、AUT-04.3 新切片 |
| table-edit-transactions | 插删行列、合并/拆分、合并锚点、测量与分页 | 最后行列、跨合并插入、内容重复/丢失、失败回滚及目标失效；对应 AUT-05.2.1–.2.3 |
| smartart-cache-transactions | 准入布局的数据图、缓存、主题和原生重布局结果 | 新数据/旧缓存、未知一致性、生成中取消与资源增长；对应 ADV-03.7.1–.7.3 |
| slideshow-interruption | 点击/定时/动画/媒体同时结束，切换中跳转与页面重入 | 连点、后台暂停、媒体未加载即离页、重复事件和已释放资源复活；对应 ADV-12.5.1–.5.3 |
| mixed-round-trip-budgets | 小型边界包、大型混合包及逐操作差异清单 | 导入/保存/播放求值超限、取消释放、峰值内存与输出增长；对应 IMP-10.4.1–.4.4 |

## 统一退出条件

1. 公开 API、责任包和支持子集明确；Core 不携带 OOXML 保留结构。
2. 对象的创建、导入、编辑、预览、导出、原样保存、编辑后保存分别有结论；静态预览不能代替原生编辑/播放。
3. 正例、边界、负例、失败重试和相关消费者测试通过；原生声明必须记录具名应用版本及实际结果。
4. 同步更新能力矩阵、API/示例和执行记录，链接父子任务；部分实现保持未完成。
5. 审查假无损、假可编辑、引用破坏、资源泄漏、环境差异五类风险，保留证据与未验证项。

## 本次规划复核记录

2026-10-25 本轮交付仅调整规划文档：沿用 315 个已有层级任务 ID，新增九项，合计 324 项；包含父项与叶子切片，不能解释为 324 个互相独立的交付。新增项全部保持未完成。

- **重复开发风险**：工作区已有文字保存候选，新增 IMP-09.2.1.1 先验收，修改了原先“尚未接线”的陈旧规划；不重新安排从零实现。
- **能力外推风险**：把表格、图表、备注/属性、动作及文字拓扑的源保存单独立项；导入、fresh export、no-op 和编辑保存不互相替代。
- **高级能力虚假完成风险**：SmartArt 缓存、形状组与原生节点编辑分别验收；动画静态截图、虚拟时钟与原生播放分别留证；原生环境缺失保持 untested。
- **任务阻塞与范围风险**：研究/样本可提前，功能验收仍继承原依赖；放开文字拓扑前阻止尚未验证的范围重映射。工作包只作导航，不增加勾选台账或冻结全部 v1 范围。
- **文档失联风险**：检查新增任务 ID、祖先 ID、显式依赖、表格结构及本地链接；路线图、主 Backlog、能力矩阵和文档入口同步指向本次批次。

验证：`pnpm lint` 通过，包含包结构/边界、文档链接和 TypeScript 示例；规划脚本核对 324 个唯一 ID、新增九项未勾选及其祖先/显式依赖，并检查相关文档空白与冲突标记。本轮未执行功能 build/完整测试，未改变运行时代码，文字保存候选的最终验证仍由 IMP-09.2.1.1 完成。

动画主线已有 [ADV-06.1.1 有限时间实验](../architecture/decisions/0013-finite-motion-time-research.md)：复用其毫秒/串并行/边界测试，接续分数和无限重复、事件目标身份与原生 timing 映射；不要把纯进度求值计为效果或播放器实现。

ADV-06.2.1 已有[四类原生 timing 实验](../architecture/decisions/0015-native-animation-mapping-experiment.md)：进入、强调、退出和路径的 headless 存储结果可复跑，但实际事件顺序/终态尚未观察。下一张任务卡应复用样本做原生播放与初终态验收，避免将保存重开计为效果准入。

### 本次补充执行卡的复核

2026-10-25：在上述 324 项基础上新增十个叶子 ID，当前清单共 **334 个唯一层级 ID**，包含父项；新增项全部未勾选。本次只调整 ROADMAP、主 Backlog、子任务清单与排期，不修改运行时代码。

交付前反向检查并收紧四个风险：多图缓存不得猜测归属；原生保存重开不得替代播放证据；任务数量不得当作独立交付量；新工作卡不得覆盖既有常用对象/编辑保存队列。任务的完整完成仍继承父项依赖，局部准备可以先行。

验证：脚本检查 334 个 ID 唯一、十个新叶子未勾选且父项存在；`pnpm lint` 通过包结构、边界、文档链接及 TypeScript 示例检查。纯规划变更未运行功能 build/完整测试，也未据此新增实现或兼容性结论。

SmartArt 执行进展：ADV-01.1.1.1/.1.2 已有[对象与四组件检查](../api/smartart-inventory.md)，复用现有 OPC/MCE 与 raw 身份索引。下一项接 ADV-01.1.1.3 的缓存/共享资源归属，再接完整公开报告验收；不要重新实现文件名扫描，也不能把当前 located 状态提升为节点可编辑。原生样本与完整父项闸门仍开放。

ADV-01.1.1.3 已推进[页作用域缓存与直接共享资源](../api/smartart-inventory.md#drawing-cache-association)：不在 data part 中错误解析缓存 ID，不把未归属候选当作已用缓存，也不推断 freshness。已有四个固定版本公开源样本的独立检查证据，接续样本许可准入、完整来源/闭包验收及 ADV-01.1.2 缓存预览；已有源清单与图片引用报告可复用。

ADV-01.1.2 已有[显式缓存预览](../api/diagram-preview.md)：复用已有绘图字段生成普通 Core 组，缺失/空缓存给占位，图片不联网，预览文字不获得保存绑定。下一步接完整缓存几何/文字变换和主题样式、具名非空源缓存视觉对照、旧缓存策略；原生 SmartArt 节点编辑仍在 ADV-02/03，不由普通形状组替代。

缓存预览已扩展独立 txXfrm 文字框：形状/文字/正文旋转组合、父级子坐标重定位和未知预设的独立文字有 [ADR0019](../architecture/decisions/0019-cached-diagram-text-transform.md) 及七个公开源文字框的独立 XML 证据。接续翻转/upright、主题/字体和自定义几何；这些结果仍不准入原生节点编辑，也不关闭 ADV-01.1.2。


## 身份与原生删除接续 — 2026-10-25

ADV-02.2.1.2.1 已有 [显式身份状态实验](../architecture/decisions/0026-smartart-identity-state-research.md)：三个家族插入/重排/重挂接/删除/恢复共 48 个生成阶段独立核对内容与引用；继续来源绑定、状态提交并发和 stateful 输出原生结构验证。此前整数生成 Cycle 的新增空白节点，已按用户本次授权原生删除，保存关闭重开后四节点/闭合可见；文件检查另验保留 ID、文字、顺序和缓存引用。该样本不代替 Process/Hierarchy 删除或 stateful 选项原生准入。


## 源身份接线接续 — 2026-10-25

[ADR0027](../architecture/decisions/0027-smartart-source-binding-research.md)推进 ADV-02.2.1.2.1：显式源范围/part/layout/虚根身份绑定，识别源 modelId 与管理 ID 分离，观测到退役后重现的源 ID 需显式恢复。七组纯测试和三家族 Parser/DOM/fresh 生成集成已验证；五组十一个真实源/原生保存版本独立核对身份、文字与拓扑。整数经原生整体重编号的对照按预期拒绝，不按标签/位置强配。接续来源基线/别名与重编号策略、并发提交、source 与 writer 两份状态的事务，以及 stateful 生成的实际原生结构矩阵；完整源编辑和父项保持开放。


## 确定性 Cycle 布局接续 — 2026-10-25

[ADR0028](../architecture/decisions/0028-cycle-layout-feasibility.md)推进 ADV-02.1.3.2：Cycle 等圆节点、闭合块箭头、双节点独立弧线、框旋转/翻转、显式文字测量/负 bearing 居中与溢出策略已有代码和六组纯/两组浏览器测试。六个保留 PowerPoint 缓存的圆节点最大偏差为 0.137–0.235 点，全部未过 0.1 点研究阈值；只比较圆盒，不代表文字或箭头通过。接续原生间距/舍入/描边归因、更多画布/方向参数、双节点原生行为、异尺寸/完整字体文本策略，再决定缓存与公开布局。Process/Hierarchy 布局仍按原任务独立实现；不因该实验关闭家族或父项。


## 缓存版本与原生偏差接续 — 2026-10-26

[ADR0029](../architecture/decisions/0029-diagram-cache-consistency-research.md)推进 ADV-03.7.1/.7.2 及 ADV-02.1.3.3：私有缓存会话捕获显式依赖/环境/生成器版本，区分未知与受管当前，候选经独立验证后按修订/发布版本提交；失败、过期、取消、跨会话和配额有测试。真实 SVG 几何缓存的 DOM 集成拒绝旧输入/旧载荷；该证明不外推到原生缓存。Cycle 原生圈间距观察约为 0.5017–0.5042；EMU 舍入仍不能通过原 0.1 点阈值，换主题的线宽变化也没有改变该五节点样本的圆盒。接完整源依赖闭包、故意陈旧的原生缓存样本、字体/布局准入与整包事务；现有父项保持待完成。

## 文本对照之后的缓存接续 — 2026-10-26

[ADR0030](../architecture/decisions/0030-diagram-text-observations.md)已交付显式源/缓存段落对照与十五个受控样本；三家族 LibreOffice 数据单改仍显示旧缓存，不能把有效 XML 当作新鲜缓存。[ADR0031](../architecture/decisions/0031-native-stale-cache-policy.md)已补 PowerPoint 15 个对照、三份重布局恢复及三家族连续三轮冲突重建：按数据重建缓存，与 LibreOffice 非空缓存行为不同。接续 ADV-03.7.3 的主题/字体/几何依赖、共享源/图片与原生节点编辑边界。ADV-03.7.1/.7.2 的完整依赖推导、原生生成/整包原子更新及各家族准入仍开放。

## CLI 检查与比较接续 — 2026-10-26

[ADR0032](../architecture/decisions/0032-read-only-cli-inspection.md)将 bootstrap 替换为可执行 inspect/compare，复用现有公开导入报告和 OPC 比较器。接续 ECO-09.1.1 的完整版本化操作/任务/证据目录、混合样本与高级预算选项；创作和源保存命令另按既有对象/操作准入，不以只读命令交付宣称已支持写入。

## 能力观察视图接续 — 2026-10-26

[ADR0033](../architecture/decisions/0033-import-capability-observations.md)补齐显式七操作、未知/冲突、任务/证据引用、筛选与完整主体截断；CLI 需显式选择能力视图。继续 IMP-04.4 的完整逐属性源清单、ECO-09.1.1 精确证据覆盖与候选保存条件；已有引用只作导航，不将具名原生样本结论套给当前输入。

## 图表数值源保存接续 — 2026-10-26

[ADR0034](../architecture/decisions/0034-chart-numeric-preservation-save.md)推进 IMP-09.2.6.1/.2：显式 existing-values-v1 绑定既有数值，cache 与内嵌 XLSX 双层词法写回、独立差异和实际候选再绑定，拒绝共享/别名/公式/外链/冲突。两方言三家族三轮、混合文字、取消失败重试与精确预算耗尽已有测试。接续完整引用/依赖策略、更多原生产器、具名实际 Edit Data 与三轮保存；结构变更由 .6.3 独立准入，本批不关闭父项。

图表原生接续补充：已有[三轮源保存独立检查与 LibreOffice 重开](../testing/chart-value-save.md)，复用独立 Python oracle 和四页对照样本。该版本原生重存删除 XLSX 数据源，数值保存资格必须继续判为 unmapped；不能由 cache 保留升级为可编辑。下一步直接做具名实际 Edit Data 和 XLSX-backed 原生产器三轮，不重复 fresh-export 数据检查。

## 有限时间算术公开契约接续 — 2026-10-26

[ADR0035](../architecture/decisions/0035-public-finite-timeline-arithmetic.md)推进 ADV-06.1.1.1：Core 纯函数编译 span/sequence/parallel 与采样，不新增 authoring/IR/Layout 动画字段；整数 repeat 独立预算、精确算术、零时长/填充/同刻排序、完整 JSON schedule 校验已有代码与独立测试。接 ADV-06.1.2.1 的事件日志/重入，再到稳定目标/传输；原生映射/实际播放与分数/无限/autoReverse 仍按原任务准入，不因数值 API 关闭父项。

## 前台事件日志与重入接续 — 2026-10-26

[ADR0036](../architecture/decisions/0036-event-timeline-replay.md)推进 ADV-06.1.2.1：公开 compile/replay 数值工具，日志按 page/visit/event ID 定义身份，提前点击不排队，同刻按数组顺序一次消费；每次显式当前对象清单阻止已删除目标/触发源，含未来阶段。重入/seek/暂停按显式 active-time/log 快照定义，不派发事件。接 ADV-06.1.2.2 的 authoring/IR/Layout/复制删除传输，以及 .2.3 原生映射；独立触发轨、原生点击完成和真实播放器仍待准入。

## 幻灯片目标身份传输接续 — 2026-10-26

[ADR0037](../architecture/decisions/0037-slide-event-target-transport.md)推进 ADV-06.1.2.2 的 Core/IR/Layout 切片：slide 绑定 owned whole-object stages，整页复制同步 map page/target/source，引用删除原子拒绝，replaceElement 保留根 ID。PPTX 默认拒绝未准入 native timing，显式 omit 才报告损失并静态生成；SVG 警告，分裂分页不猜目标归属，Parser 源保存阻止新增时间线。接 .2.3 的源 timing/native 映射与实际播放，再补 subtarget/跨文档/完整源结构与多页策略；不因 IR 传输完成而关闭父项。


## 原生 timing 清单接续 — 2026-10-26

[ADR0038](../architecture/decisions/0038-native-timing-source-inventory.md)交付显式只读 inspectPptxTiming：两方言、raw/effective、逐 part 身份候选、未知/子目标/别名保守处理、全局预算与独立 XML 对照。接续 ADV-06.2.1.3 / ADV-06.1.2.3 的逐效果 accept/defer/reject、相关编辑阻止反例和真实播放；不将 located 或 source-only 提升为 Core/native 写入权限，父项保持开放。


## 段落索引目标的源保存接续 — 2026-10-26

[ADR0039](../architecture/decisions/0039-stable-paragraph-target-preservation.md)新增显式 paragraph-stable-v1：原始 owner 唯一、直接段落来源/范围可证明时，允许既有 scalar text 修改并原样保留 timing；不放开字符/未知/结构修改。12 组 Parser、2 组浏览器和独立 Python 三轮有证据。具名 LibreOffice 文字保留但范围改写为 0..0，未编辑对照同样改写；下一步核对具名 native timing 结构/实际播放、字符偏移与段落结构策略，ADV-07.8/ADV-06 父项继续开放。


## 表格结构事务接续 — 2026-10-26

[ADR0040](../architecture/decisions/0040-atomic-table-structure-edits.md)推进 AUT-05.2.1/.2/.3 当前直接样式模型：纯 editTable 插删/合并/拆分、内容归属与 span/尺寸一并候选提交、完整 anchor mapping、测量分页接线。跨包发现并修复跨行且跨三列以上的 PPTX continuation 错位。13 组 Core、4 组跨包、安装消费者及三批独立 Python/native headless 证据见 API。后续接完整边框/主题优先级、cell 目标身份、原生 UI 编辑与 IMP-09.3.4 source writer；不把 fresh 输出当保留保存，不关闭完整父项。

### 直接逐边表格切片后的下一步 — 2026-10-26

[ADR0041](../architecture/decisions/0041-shared-table-borders.md)推进 AUT-05.1.1 / IMP-06.1.4：直接稀疏覆盖、共享 physical edges、合并锚点外框和自动测量分页接线已有实现。具名 native width/alpha/transform 仍 partial，完整任务不勾选。后续领取：

1. AUT-05.1.1 完整 table/row/column/first-last/banded/cell cascade；逐项定义冲突来源，避免把当前直接样式规则当 Office cascade。
2. AUT-05.1.2 与 IMP-06.1.4 的 table style/theme/文字方向来源和局部覆盖；两方言与未知样式反例。
3. IMP-06.1.5 原生交互、分段合并边界及 rotation/flip 对照；严格线宽/透明度与几何颜色分别验，不以 headless PDF 代替编辑。
4. IMP-09.2.5 准入既有格直接样式/行列尺寸的 source writer，再接 IMP-09.3.4；证明未知引用/扩展保留，混合不支持修改原子阻止，三轮保存重开。

每项继承 TODO 表中依赖；fresh 输出、source 保留与 native 行为分别留证。

### 创作区域样式切片后 — 2026-10-26

[ADR0042](../architecture/decisions/0042-table-region-style-materialization.md)补齐 Core 创作层 whole/band/首尾/四角/局部文字与边框的求值；8 组 Core（160 独立分区 oracle）、4 组跨包、安装消费者和三场景独立源/native fill-text 证据已落地。输出物化为直接格格式，不保留原生区域定义；不能据此关闭完整 cascade/native 任务。

接续工作：IMP-06.1.4 读取具名源 table style、主题与 flags 的可追踪 effective cascade；AUT-05.1.2 明确 native 样式身份/复用、方向及属性来源；IMP-06.1.5 验实际改样式/插行/换主题后保存重开；IMP-09.2.5 在证明未知源引用保留后开放直接 style/metric writer。原生分段边框、alpha/width/transform 差异继续分别验，不用本轮 fill/text pass 替代。

### 源表格样式直接区域切片后 — 2026-10-26

[ADR0043](../architecture/decisions/0043-source-table-style-projection.md)推进 IMP-06.1.4：引用/内联/default/flags、源主题/override/映射、直接region属性、逐cell来源与session预算已有实现。10组Parser+2组跨包/浏览器、安装消费者、三轮source scalar保存和三个独立native fill控制已有证据。native实验发现并修复双band四格的错误优先级；conditional仍fallback，不复制Core自定义求值规则去假装原生cascade。

下一批：fillRef/lnRef与theme fmtScheme/占位色依赖；源默认text appearance与脚本字体；更多重叠/合并/RTL及native语义矩阵；具名PowerPoint改样式/换主题/插行后保存重开；完成未知源引用证明后接IMP-09.2.5 style/metric保留writer。每项分别有独立来源/差异/取消/预算与native证据，父项继续开放。

### 主题矩阵引用切片之后 — 2026-10-26

[ADR0044](../architecture/decisions/0044-table-theme-matrix-references.md) 接续上批 fillRef/lnRef：独立 fmtScheme 选择、位置不移位、lazy phClr、连续 RGB/alpha 组合、逐属性来源和共享预算已实现，10组Parser+2组跨包通过。五个[独立 native 控制](../testing/theme-matrix.md)保留背景引用、变换与线宽差异，fill1000 的应用失败单独记 unavailable；不能据此关闭完整 native 闸门。

下一批按[四张具体卡](advanced-capability-todos.md#表格主题矩阵之后的具体执行卡)领取：IMP-06.1.4.1 默认文字/脚本字体 → .4.3 PowerPoint/区域对照；.4.2 更多paint/line按对应Core能力分别准入。IMP-09.2.5.1 在父项前置与未知源引用证明后开放一种直接属性/尺寸writer，再接结构保存。SmartArt原生家族、动画播放和操作矩阵继续原队列，未做的完整能力不标已完成。

### 表格文字默认色与显式字体集合之后 — 2026-10-26

[ADR0045](../architecture/decisions/0045-table-text-appearance-and-uniform-fonts.md) 推进 IMP-06.1.4.1：fontRef颜色默认、region/local覆盖、显式同family lt/ea/cs与supplemental集合、不同family阻断及来源已有实现；source保存仍只允许原有scalar文字。独立native四控制中region/local颜色通过，reference-only默认色偏差、Arial替换为LiberationSans保留为partial，见[观察记录](../testing/table-text-appearance.md)。

接续 IMP-06.1.4.1 与既有 IMP-05.1.3.3.2：先明确每个脚本的font声明/缺省/覆盖与Core格式无关传输，再实现按language/script的选择与混合run策略；实际shaping/缺字形/替代分别验收。省略默认与空段/插入样式另核对原生，具名PowerPoint复查reference色及改样式/换主题；不依赖LO三项颜色通过推断字体相同。其他矩阵paint、native区域和source writer保持各自队列。

### 显式脚本字体创作与选择之后 — 2026-10-26

[ADR0046](../architecture/decisions/0046-script-font-authoring-and-selection.md) 已交Core两独立family字段、undefined继承/null重置、共享有界Unicode选择、Layout具体测量fragment、SVG双路径和fresh PPTX独立字体槽。8组Core、3组Layout与4组跨包覆盖旧IR、脚本字符/组合/UTF16、预算、缺字体/替代、独立XML/DOM、源保留守卫及CSS字符串；不改变原始run身份。具名[原生观察](../testing/script-fonts.md)保留run5→15、字体替代、CJK视觉缺字而可提取的反例。

接续 IMP-05.1.3.3.2 / IMP-06.1.4.1：独立source字体profile应逐字段记录Latin/ea/cs来源与阻断，正确继承/局部覆盖/空slot，保留旧默认profile和source写守卫；双方言、多master、三轮source scalar保存及browser/installed分别验。随后接language/Han变体/supplemental脚本映射、真实字体coverage/shaping和具名PowerPoint外观；不要因Core已经可传输而删除当前Parser多family阻断。原生字体写回仍需原任务独立准入。

### 显式源脚本字体 profile 之后 — 2026-10-26

[ADR0047](../architecture/decisions/0047-source-script-font-profile.md) 已交显式script-families-v1：每字段继承/阻断、近层missing slot依赖、table collections/fontRef和direct notes来源、跨scope审计隔离、options捕获与候选同profile复验。14组Parser、3组跨包及[四源包独立验证](../testing/script-font-import.md)覆盖当前profile；默认旧模式不变，font/style source写回仍整批blocked。

接续 IMP-05.1.3.3.2 / IMP-06.1.4.1：先决定language/Han变体与supplemental script map的传输/求值契约，用具名源和已安装字体覆盖各分支，不能把空ea/cs猜成Latin；notes master/theme继承继续独立准入。真实coverage/shaping/bidi、PowerPoint字体选择/改样式/换主题需分别有证据。具名LO仍替代synthetic字体、run5→15及CJK不可见，保留observed而不关闭native闸门。其他表格paint/native区域与source writer保持原队列。

## 脚本字体映射接续领取

[ADR0048](../architecture/decisions/0048-script-font-maps-and-language-selection.md) 已交格式无关map/重置、显式language/Han选择、Layout/SVG与警告式fresh物化；源script-families-v1保持原三slot语义，不自动扩大权限。下一批按[叶任务](advanced-capability-todos.md#脚本字体映射的接续任务)领取：先source supplemental profile准入/逐源传输与阻断，再候选复验/三轮scalar保存，随后具名真实字体coverage/shaping及PowerPoint改字/改样式/换主题。未实现的reader、writer及native闸门继续开放；不把已存在的author map或当前文字输出当作源语义完成。

## Supplemental 源字体接续

[ADR0049](../architecture/decisions/0049-source-supplemental-font-projection.md) 已交显式v2：slot cascade后再选择supplemental map、逐key来源、空slot不猜Latin、unknown/duplicate/MCE和近层component边界；captured profile候选复验与三轮mixed scalar保存已有自动化/独立source证据。接续具名原生font优先级对照（含Latin fallback及cross-slot引用）、notes master/theme独立继承与真实字体coverage/shaping。源字体writer仍blocked，不因当前文档能携带map而开放style写权限；完整父任务继续未完成。

## Notes 字体继承接续

[ADR0050](../architecture/decisions/0050-notes-master-text-inheritance.md) 已交显式notesTextProfile：notesStyle→按type匹配的master body list→local list/paragraph/run、独立notes theme/颜色映射与font v1/v2、captured bytes/file/candidate复验。13组Parser、3组跨包与四阶段22 style property独立source证据推进.4。下一步仍需真实字体coverage/shaping与PowerPoint notes改字/换主题/保存重开对照；完整notes页面几何/fields/artwork和source style writer保持独立任务，不把文本继承等同整页支持。

### 真实字体与原生选择对照后的下一批

[具名实测](../testing/real-fonts.md) 已交8个regular真实face的cmap/HarfBuzz/Chromium实际PS/度量与负例；PowerPoint20控制完成三轮无编辑保存重开，selected声明与栅格稳定，但字体仅15/20匹配。LO通过child-only Fontconfig才发现宿主CJK字体，其主题选择偏差不能替代PP证据。下一批按[8张执行卡](advanced-capability-todos.md#真实字体证据之后的执行卡)：优先.6.1补齐5差异/coverage/缺省槽边界，再.5.2真实SVG与fresh原生外观、.6.3 source改字/样式/主题；mixed bidi、emoji、author-map编辑与notes原生各自验收。仅.5.1限定regular切片完成，完整字体父任务保持开放，source writer不扩权。

接续32成对native控制已取得20匹配/12差异：韩语短tag与ko-KR区分、实际coverage确认、absent/empty/present EA-CS、covered cross-slot、major refs和bare zh alternate map；5实际ZIP mutation controls全部拒绝，修正了oracle漏查presentation真实slide顺序的问题。下一步.6.1集中表格fontRef/notes与保守source-profile合同：v2的缺槽Latin补充规则推断已有反证，不能再当native parity；兼容旧显式profile的同时另行准入修正，完整source/candidate/双方言与原生证据缺一不可。SVG/mixed bidi/emoji与native editing仍按原卡推进。

### Declared-slot v3 source合同交付后的接续

[ADR0051](../architecture/decisions/0051-declared-slot-supplemental-projection.md) 已交独立v3：不借Latin字典填缺省EA/CS，逐source rule定位unsupported与实际selector/target；继承/explicit collection/empty generic/cross-slot、table/notes/empty insertion统一finalization，byte/file/candidate保留captured profile和完整snapshot。9组Parser、3组跨包、installed runtime/types与[四包独立证据](../testing/declared-slot-font-import.md)覆盖source合同，v2/default兼容。下一步仍是具名native表格fontRef/notes及实际样式/主题编辑、同face SVG；v3没有推断Office隐式字体或替代源family/lang。完整.6.1/.6.3和父任务继续开放。

### Actual SVG字体路径之后的接续

[ADR0052](../architecture/decisions/0052-svg-text-boundaries-and-font-names.md) 修正native nowrap/空白/连续末尾换行/跨font fragment字素边界、XHTML空行高度和native text/chart literal family quoting。[16真实SVG控制](../testing/real-font-svg.md)以唯一alias核验actual custom PS、HB宽度、可见PNG及inline/image精确栅格；无embedding alias反例和三实际PNG破坏反例拒绝。6组新增browser回归覆盖实际缺陷，baseline不换图。此切片推进.5.2，fresh-PPTX同face原生编辑/外观、mixed bidi/emoji、native表格fontRef/notes及完整父任务继续开放，不因结构上字素未拆分就宣称emoji正确显示。

### Fresh作者PPTX原生基线之后的接续

[16页fresh控制](../testing/real-font-fresh.md)通过author API直接生成，无source XML patch；7个具名regular family的8字符串分别使用直接字体槽与script-map物化。PowerPoint实际PS全部匹配，三个无编辑保存/关闭/重开轮次的文本/lang/槽/run数与精确页面栅格保持；五实际ZIP/PDF破坏反例已拒绝。下一步.6.2按SC/TC/ja/ko/ar/he补实际改字、regular/bold/italic、换主题及每操作三轮验证，记录map规则无法恢复、原生run分裂与实际替代；不以本轮无编辑稳定性代替编辑。表格fontRef/notes、mixed bidi/emoji仍按各卡推进，.5.2/.6.2父项不勾选。

### SC物化页原生操作之后的接续

[SC四阶段各三轮](../testing/real-font-edits.md)已交：追加Ω/lang与run分裂、unused Grek规则不恢复、actual W6、W3合成shear、真实Office theme链与方案变化；16份Parser v3/精确no-op、逐阶段精确页面栅格和10实际负例通过。native编辑移除段落pPr、首次重开新增dirty=0已明确记录，不将这些native规范化描述为无变化。下一步将同样验收推广至TC/ja/ko/ar/he，再补混合run和局部样式编辑；native表格fontRef/notes及mixed bidi/emoji保持独立任务。仅.6.2.1 SC切片完成，完整.6.2/.5.2与父项仍开放。

### 五脚本原生全文操作之后的接续

[TC/ja/ko/ar/he四阶段各三轮](../testing/real-font-multi.md)已交；source语言变化、2/3run、actual粗体与合成italic、Cambria/Cambria-Bold/Cambria-Italic fallback、Office主题链、阶段内256页精确栅格、11旁路控制、16份Parser精确no-op与43实际负例均验证。Arabic ToUnicode少一个source fatha、Hebrew presentation forms单独记录，不以提取文字冒充完整mark/bidi。系统fc-match Cambria返回Times New Roman而native用Office内置Cambria，同名候选核SHA/coverage但未认证PDF subset字节相同。下一步.6.2.3用作者原始多脚本跨family混合段测试fresh物化与局部编辑；native表格fontRef/notes、mixed bidi/emoji维持独立任务。只完成.6.2.2切片，不勾父项。

### 原始配对混合段局部编辑之后

[六组Latin+本脚本+Greek](../testing/real-font-mixed.md)由1作者run物化3family，局部插入X后5run；Latin范围bold/italic与Office主题各三轮，actual Arial-BoldMT/Arial-ItalicMT、相邻regular glyph序列和阶段内栅格、72实际反例、17份Parser精确no-op均通过。双击选词未覆盖X与混合bold快捷键清除样式的真实样本保留且被拒绝；RTL空格/正文同名资源分离导致的漏验已用descriptor一致性及全部资源负例修正。下一步.6.2.3同段包含完整Latin/SC/TC/ja/ko/ar/he组合，不能以六组配对替代；native表格fontRef/notes和mixed-bidi/marks/emoji继续按原卡推进，父项保持开放。

### 同段完整组合之后

[完整同段组合](../testing/real-font-combined.md)完成.6.2.3：单段八family、1物化warning、Latin局部插字/粗斜体/Office主题各三轮；首次拼写检查8→9run及dirty/err精确记录。actual PS、三Han绑定、源/trace/视觉顺序、邻居glyph序列、逐span墨迹和栅格分别核，55真实负例与17份Parser精确no-op通过。父.6.2操作测试门齐，但.5.2/.6.1依赖未闭合，继续开放；不再把同段组合列为待办。下一步回到native表格fontRef/独立notes实际字体与主题，之后source主题字体编辑；完整mixed-bidi/marks/emoji维持独立准入。

### 表格fontRef与显式集合基线之后

[36个native控制](../testing/native-table-fonts.md)完成27+9两组各三轮，144页栅格稳定，source声明仍保留；20匹配/16差异明确：minor/major短ko与ko-KR不同，显式集合空泛用槽走SimSun/Batang/Arial，filled EA/CS走ArialUnicodeMS，不能仅归为空槽回退或缺glyph。source v3运输与native actual分开，8精确no-op/79真实反例/独立DOM来源链回归通过。下一步先明确声明source与外观准入策略，保持现profile兼容并拒绝把fallback写回源，再做表格实际编辑/样式/主题与独立notes原生验证；.6.1和父项保持开放。


### 表格与备注字体的下一张领取卡

具体任务见[表格与备注原生字体执行卡](advanced-capability-todos.md#表格与备注原生字体的可领取任务)。先领取 `IMP-05.1.3.3.2.6.1.1`，用现有36控制决定 source运输与native外观的合同；随后 `.6.1.2` 在新副本追加文字，`.6.1.3/.6.1.4` 分别验证样式与主题，`.6.1.5` 汇总候选复验/负例。notes先走 `.6.4.1` 独立master/theme真实字体基线，再 `.6.4.2` 实际编辑，不等待完整表格父任务。原生操作卡要求三轮重开和独立证据。现.6.1.1策略、.6.1.2具名追加、.6.1.3具名样式及.6.1.4主题/.6.1.5汇总已完成；下一张领取notes .6.4.1独立基线。

后续[15目标native追加证据](../testing/native-table-font-edits.md)已落地：两组各三轮、108页、112真实反例；仅推进追加切片，未完成样式/主题/notes。图表Edit Data、SmartArt原生节点编辑、动画实际播放继续保留既有任务与证据状态；解锁回复不能作为这些验收已完成的证据。完整能力仍按本文“端到端交付阶段”推进。


### 表格追加完成后的实施顺序

[ADR0053](../architecture/decisions/0053-declared-fonts-and-native-appearance.md)保留v3源声明合同，native字体选择差异单独记录，不烘焙fallback。`.6.1.2`已完成15目标的真实追加；下一步`.6.1.3`在新副本逐目标粗体、纯斜体、恢复regular，每阶段三轮；再`.6.1.4`比较引用与显式集合在具名主题下的响应，`.6.1.5`汇总完整编辑回归。独立notes `.6.4.1`仍需真实字体/独立theme基线，不能用slide PDF替代。完整父项继续开放。


### 表格样式完成后的实施顺序

[十五目标bold/纯italic/恢复regular](../testing/native-table-font-styles.md)各完成三轮native重开，324页、380真实反例、20+6精确no-op、18 source样式写拒绝与54 headless scalar候选；恢复regular的36页与样式前栅格相同，源显式false另记。`.6.1.3`限定切片完成。下一张领取`.6.1.4`：从新副本对major/minor引用与两类显式集合施加具名主题，独立跟踪关系链/字典/局部覆盖和实际字体，再按`.6.1.5`汇总。notes `.6.4.1`及完整父项继续开放，原16选择差异不改为通过。


### 表格Office主题与候选汇总后 — 2026-10-26

[四组主题证据](../testing/native-table-font-themes.md)完成.6.1.4及.6.1.5具名范围：216页主题后/72页基线、实际theme链与Office资产、source局部覆盖、显式集合正文保持、具名1.2pt基线和subset GID/ToUnicode差异；484反例、16+4精确no-op、36 scalar候选。追加/样式/主题合计44条输入验证路径、108候选、976反例；复用baseline不算独立文件，headless候选不算native重开。

接续.6.4.1独立notes master/theme基线，再.6.4.2备注实际操作；.6.1及字体/表格/round-trip父项、source主题编辑、mixed bidi/emoji保持开放。已完成主题source回归覆盖8输入/144页/三消费者各三轮真实关系重绑定与local覆盖，并接入test:import；不依赖宿主专有字体。


### 独立notes字体基线后 — 2026-10-26

[两组30控制](../testing/native-notes-fonts.md)完成.6.4.1具名基线：两套实际notes master/theme、type-vs-index、local/空/坏槽与三轮备注打印，24匹配/6差异、90重开页+30源页、226反例、8+2精确no-op与18 headless scalar候选。源构造patch、主题迁移、冗余声明删除、40/16pt→PDF41/17观察、RTL ToUnicode差异和空slide-image均具名保留。

接续.6.4.2：新副本实际备注追加、局部样式、slide主题操作；核独立notes theme是否响应而非预设跟随slide，每阶段三轮准确路径重开和实际notes PDF。正文、装饰/fields与slide邻居分开验证。headless候选和CI sourcefixture不替代native；完整notes页、字体/round-trip父项与source style writer继续开放。


备注页缩略图限制另由 `IMP-03.3.2.1` 领取：先与PowerPoint自身生成的可见sldImg对照，定界作者基础包/source构造/原生解释差异，再修复所属层并验证三轮notes输出。该任务不通过降低字体基线或删除缩略图要求解决。


### 备注局部操作source回归与锁屏恢复点 — 2026-10-26

[二run/source局部样式回归](../testing/native-notes-font-edits.md)完成A/B×四阶段8输入、120源页、Node/ESM/global各三轮72 scalar/no-op路径、candidate完整相等/字符外byte精确与style拒绝恢复。解锁后已在run-t0kxZH完成原生追加/局部三样式各三轮，共360实际Notes页、24 exact no-op、88实际ZIP/PDF反例；原run/旁路稳定、恢复regular全页栅格回到append基线。新增Ω字体声明6匹配/16差异另列.6.4.2.2.1。

.6.4.2现拆为source合同（已完成）、原生追加、局部三样式、两域主题响应与候选/实际反例汇总。追加与局部样式现已完成，主题及汇总仍开放；包内两notes-master正向表述已由下方规范复核纠正：额外master仅作非合规source控制，合法隔离采用单master下逐notes override。整体目标继续active。


### Notes master基数纠正与两类source隔离 — 2026-10-26

[规范复核与隔离证据](../testing/notes-master-isolation.md)确认每包最多一个Notes Master part；保留现有“两ID登记即拒绝”守卫。此前未列出第二master的fixture是非合规容错源，已补明分类和finding，不继续把它设成标准原生创作要求。合法包内差异由共享notes master+逐notes themeOverride构造。

.6.4.2.4.1已完成两组60页、三消费者三轮A→B→A共18保存/no-op路径、独立8状态/每状态248字体字号属性、80实际ZIP反例和缺override/组件缺键/坏type/引用隔离。合法基数控制不冒充全schema验证；非合规控制的诊断及字符保存不冒充修复或native支持。旧notes-style四阶段source证据已补分类并复跑。

下一步仍是解锁后恢复明确run-t0kxZH原生追加，再做局部样式与.6.4.2.4.2的合法主题/override响应。源域基数纠正没有完成原生轮次，也没有关闭notes/字体/导入/round-trip父项。

2026-10-26解锁后接续：[备注原生追加](../testing/native-notes-font-edits.md)已完成A/B各三轮保存/关闭/准确重开/Notes打印，共90页；6 no-op、18 headless字符候选和28实际破坏反例通过。新Ω原生使用Latin槽，声明选字6/22匹配、16差异另拆.6.4.2.2.1，不完成字体等价。局部三样式随后也已完成（下段），下一步.4合法单master/逐notes override主题响应和.5汇总；完整notes页/source样式writer仍开放。此进展取代上方锁屏恢复段的当前状态。

接续.6.4.2.3也已完成：A/B×bold/pure-italic/regular各三轮，共18输入/270实际Notes页；18精确no-op、60实际ZIP/PDF破坏反例全部通过。原文/旁路trace和Ω外栅格稳定、regular全页恢复append基线，实际face/资源/矩阵独立核，无合成stroke/shear。下一张领取.6.4.2.4/.4.2主题响应；.2.1语言/槽差异、.5完整候选汇总和完整notes页面仍开放。见[原生样式证据](../testing/native-notes-font-edits.md#原生局部样式验收)。

2026-10-26接续：[三组Office主题](../testing/native-notes-font-themes.md)完成.6.4.2.4/.4.2：A/B独立notes方案与单master/逐notes独立override各三轮，180主题Notes页+90基线/30首次页；另240普通slide PDF页。15+3 no-op、45候选与73实际破坏反例通过；末字A kern/slide两类Y位移和shared override原生repair提示明确保留。三消费者source隔离增加distinct控制，27 scalar/no-op路径通过。下一步.6.4.2.5汇总各原生阶段候选/资源反例，.2.1字体语言/槽差异与完整notes页/slide-image继续开放，整体目标active。


### 备注页恢复接续状态 — 2026-10-26

[当前候选汇总](../testing/native-notes-consolidated.md)已完成37输入/三消费者与143实际反例复验；[A源rect](../testing/native-notes-pages.md)、[默认作者版式](../testing/authored-notes-pages.md)、[全空插入样式](../testing/notes-insertion-styles.md)及[B/独立override图像](../testing/notes-image-domains.md)已分别交付具名证据。不能继续把这些当作尚未开始的任务，也不能借此关闭完整notes父项。下一步领取IMP-03.3.1/.2的完整角色/notesSize/几何/artwork/fields读取与source-only契约，再接.3的对应输出；保留16字体差异和共享override准入限制，原完整待办队列不缩减。


### Notes source-only合同之后 — 2026-10-27

[ADR0056](../architecture/decisions/0056-notes-source-inventory.md)已交独立`inspectPptxNotes`及角色、notesSize、实际关系域、显式type候选、几何/artwork/非视觉属性与字段cache的source-only策略。10既有captured文件/98页/4字段/7550来源引用具独立DOM证据、完整Node/global相等与精确no-op；不重新计算为native操作。内部MCE包装分类问题在本轮审查修复并加回归。

下张领取IMP-03.3.1/.2剩余有效值合同：限定notes root/hf/占位本地值与master值的可解释子集、明确字段动态值/静态cache策略和几何/visibility继承边界，再设计对应fresh作者输出及具名native重开。不要重复开发来源清单，也不要将缓存字段直接烘焙为当前日期/页码。16字体语言/槽差异、source样式writer、完整notes native/父依赖与原全部待办仍开放。


## 多母版前置基础接续 — 2026-10-27

notes有效值/字段合同依赖的AUT-07.1.1此前没有作者模型，故先推进[ADR0057](../architecture/decisions/0057-master-theme-ownership.md)：Core init只读master定义、layout归属、独立theme/background；正常/内在/表格测量、SVG和fresh native master/theme关系已接通，notes仍独立global域。旧输入IR不增字段、空/缺registry输出bytes相等、自定义Document无新方法要求；backgroundSource新增master需要穷举消费者处理。

[当前验证](../testing/multi-master-authoring.md)有10聚焦组和Node/ESM/global全normalized/Layout/SVG/bytes相等；独立DOM核红/蓝/红/默认实际slide→layout→master→theme与全局layout ID、Parser结构及精确no-op。测量测试发现ordinary诊断仍走global theme并修正；readonly公开属性改为getter，拒绝未知parent/artwork声明。具名source PPTX与4 SVG已生成，manifest明确未做native操作。

当前完整build/typecheck/lint/test最终结果待相应进程结束记录。未提交发布。后续仍需AUT-07.1.2有效值/共享编辑/clear-reset/provenance、.1.3母版artwork、.1.4实际native换布局/主题三轮和IMP notes/fields对应输出；本foundation不关闭完整多母版、备注或字体父项。16字体差异与原全部待办目标保持active。


接续第一步：读取 `docs/testing/multi-master-authoring-evidence.json` 的当前source路径及SHA，在新副本做PowerPoint首次打开和三轮另存/关闭/准确重开；实际核master/layout/theme身份、背景/文字和unused master响应，再具名研究换布局/换主题。此为AUT-07.1.4前置研究，不提前完成artwork/整个父卡。先取native准入证据再扩大共享编辑，避免仅凭XML图关系通过就建设更大的未准入模型；不要覆盖source或把browser-preview算成native输出。随后继续AUT-07.1.2有效值、共享编辑/provenance及notes/fields原队列。


## 多母版原生准入、实际失败与ID修复 — 2026-10-27

首次用PowerPoint16.104打开上一轮作者源触发修复，已取消并保留原source、AX和截图。两因素实际打开控制：有/无unused master × 原ID/只抬高master ID，只有原ID两源触发repair，两个ID-only控制正常打开。生产Exporter改为额外master ID从全部layout ID之后分配；新公共作者输出只有presentation master ID值变化，其余payload精确保持。包测试和独立DOM都补master/layout联合ID守卫，另有24布局用例。原自动化pass不再冒充native准入，失败source永久留在native-b85qrfu0/input.pptx和原run。

[当前独立证据](../testing/multi-master-native.md)完成修复后fresh源4首次PDF页+三次SaveAs/close/准确路径reopen的12 PDF页。4master含unused、5layout、全部theme/layout/master/slide/notes内容与关系和content types按稳定身份核；native确实重排layout/theme文件名和关系ID，不按文件名猜域。仅具名noGrp/备注图像locks、-0、空guide扩展规范化，pres/view/app/core完整变化另录。后两轮仅core.xml bytes变，12页全144dpi raster/glyph trace/实际嵌入字体program与首次PDF精确一致；四页contact已查看，绿通道PDF float渲染少1byte单列，不烘焙为source颜色。

10实际PPTX破坏与2实际PDF遮字/错页全拒绝；当前fixed/round1/round2/round3共4输入Node/ESM/global完整inspection/report与12精确no-op另具manifest，不算新增native操作。完整build/typecheck及最终workspace tests/lint结果待命令结束后记入；未提交发布。当前PowerPoint窗口为 `.tmp/multi-master/native-b85qrfu0/round3.pptx`。

下一步在新副本真正换layout、换theme和改占位并每阶段三轮，区分root theme-linked属性、native覆盖和旁路未改页；从当前fixed源或对应基线派生新文件，不覆盖失败/基线。其后继续AUT-07.1.2共享编辑/clear-reset/provenance、.1.3 artwork、notes/fields与原全部待办。仅本fixture的无编辑native准入已具证据，AUT-07和notes/字体完整父项仍开放。


## 多母版实际layout/text/selected-theme操作完成 — 2026-10-27

[当前实际操作证据](../testing/multi-master-operations.md)已交三阶段：仅第一页Red→Blue layout-1；在同一body末尾实际输入` EDIT`；对仅第一页用Office主题的“应用于所选幻灯片”。每阶段3次SaveAs/close/准确路径reopen，12首次操作页+36已保存重开页。Source输入各为前阶段第三轮的精确副本；原失败、无编辑基线、输入均保留。

换layout只有第一slide edge与其本地xfrm移除，先要求本地几何与新layout继承几何精确相等；有效box仍10/20/300/60。Text语言拓扑为en-US原run、zh-CN space/altLang en-US、en-US EDIT/altLang zh-CN与具名dirty0；其他样式/notes/共享定义保持。Office选中页操作真实新增1master和12layouts，共5master/17layout；旧Master/theme/layout/旁路slide/notes不变，第一页cloned body与原blue body相同。实际face为DengXian-Regular，源family等线；palette4472C4。本地化profile与本机legacy/modern两个Office资源结构匹配，不推断app内部算法；两原资源已本地复制并绑定SHA。

旁路全144dpi栅格精确保持；共享TrueType subset的GID/program hash会因第一页改字库而变，故要求Unicode/位置/bbox/实际简单glyf字节/units/advance/lsb精确匹配，不把重编号当串改。另明确文字首次存储前3span→存储后2span，字符0索引6仅x99.1875→99.18272399902344pt（-0.0047760009765625）和所报spacewidth7→3.5；pixel与字体程序不变，保存round1/2/3 raw trace/程序精确保持。没有用一般数值容差抹掉此差异。Glyph helper仅准入这组ASCII simple TrueType，不冒充通用字体等价。

27实际坏文件（18ZIP/9PDF）全拒绝，含3个栅格完全相同但hmtx advance被改的真实PDF，校正TTF checksum后仍能被glyph/metric守卫抓住。9实际输入Node/ESM/global完整文稿/report/bytes、有效独立theme/继承box；27no-op/restore、27混改阻断和27已准入scalar候选路径通过，实际Node候选9份，仅改真实第一slide最后source text leaf，其他payload精确保持。候选未native重开，不授予source layout/theme/style writer。

本轮无生产TypeScript修改；增加独立oracle、captured runner、报告及旧oracle默认仍严格4/5的可参数化内部验证。原无编辑oracle/4输入12 no-op也已重验并刷新manifest绑定。当前操作oracle和parser manifest SHA一致；63个实际PPTX/PDF/negative/snapshot/candidate文件hash复核通过。类型、lint/边界/docs示例、4项Exporter/browser定向回归和diff通过，日志`.tmp/masters-operations-{oracle,parser-final,regression,typecheck,lint}.log`，旧默认验证日志另有baseline-oracle/baseline-parser；本轮未声称完整workspace套件重跑，最近完整118 import仍为上一轮结果。UI当前为`.tmp/multi-master/native-b85qrfu0/theme-round3.pptx`，无存活runner待续等；未提交发布。

下一步回到AUT-07.1.2：读取现有Core不可变master registry、layout定义和ElementStore替换/原子性流程，设计并实现共享定义编辑/取消覆盖/有效值provenance的统一合同与公开兼容性，再接.1.3 master artwork及notes/fields原队列。不要重复三阶段native操作，不将stock Office模板的native导入当作Core已经建模这些artwork。完整AUT/notes/字体父依赖和原全部待办保持active。


## Core共享定义编辑切片 — 2026-10-27

[新API/验证](../testing/shared-definition-edit.md)已实现ManagedPresentationDocument：工厂返回扩展接口，旧PresentationDocument不增加必需方法；define/replace master、replace layout保持身份与位置。完整替换不是patch，省略声明恢复继承，none/opacity0/false为显式值，null不是reset。新方法先capture有限plain data，拒绝accessor/cycle/symbol/sparse/未知顶层字段；不增加master parent/artwork。

ElementStore新增collection staging：候选全部root在同一excluded/pending域内物化，可在原layout roots/groups间复用ID，但不能抢其他owner。候选layout与每个绑定slide的语义/placeholder/connector检查通过才提交ID、分配器和定义；失败保持旧snapshot、资源引用和后续自动ID。纯验证接受显式数据视图，不伪装document/Proxy。无关草稿语义错误不阻止，受影响页现存错误会阻止，malformed运行时数据仍可能令验证抛错。

10组Core定向和两项多包/browser测试已通过；完整normalized/Layout/SVG/PPTX bytes在Node/ESM/global相等，独立DOM检查5页实际owner/theme、layout形状未复制到页、背景归属；隔离安装Node/types已加新接口用例。完整build/typecheck/lint已通过，现正运行最终workspace test，结束后追加结果。源码/测试已固定，不在全量运行中修改。日志前缀.tmp/definition-edit-*。

没有新的PowerPoint操作，本共享编辑输出不能借用此前native三阶段准入。下一步AUT-07.1.2逐属性有效值来源：保持单一Core cascade，先theme/background具名来源，再扩placeholder/style与完整clear/reset；.1.3 artwork、notes/fields和全部旧TODO继续。三阶段layout/text/selected-theme原生取证已完成，不重复做。未提交发布，所有旧改动与失败源保留，原目标active。

共享编辑最终验证完成：完整pnpm test退出0，Core160/Layout139/Parser696/Exporter57/Renderer19/CLI11、视觉33和导入119通过（本轮Playwright日志为33，以当前实际输出为准），隔离安装Node/browser/types通过。build/typecheck/lint已退出0；日志.tmp/definition-edit-{full-build,typecheck,lint,full-test}.log。全量后仅补强测试断言，明确接受深度128的合法customData、以RangeError拒绝129并保持旧snapshot，以及null不是reset；未改生产代码。随后完整Core+两项多包/browser测试162组全部通过（.tmp/definition-edit-core-final.log）。Docs/diff检查通过。没有native新操作/提交/发布，goal仍active；接续逐属性provenance、artwork及notes/fields原全部队列。


## Core theme/background逐属性来源切片 — 2026-10-27

[新来源API与验证](../testing/inheritance-origins.md)已实现resolvePresentationInheritance，返回普通IR+单独冻结origins。theme name/12 colors/heading-body fonts与background paint/type/color/opacity按document/master/layout/slide记录实际声明kind/id和作者路径；缺省背景和合成opacity1明确为default，none没有虚构color/opacity，零值保留声明。报告不进入IR/PPTX；共享Core resolver同时选值和记录来源，普通normalize、Layout/SVG/export不新增第二级联。Notes独立域、element/preset/run、artwork和完整reset仍待接续。

12组Core+2项多包/browser定向14项通过；包括完整IR/report/Layout/SVG/bytes三消费者相等、每个已声明leaf路径独立读值比对、同值不同声明、全角色唯一覆盖、移动页与标点ID、旧custom文档非枚举合并与name getter单读。构建/typecheck/lint通过，完整workspace test正运行（.tmp/inheritance-full-test.log）；源码测试已固定。

当前公开作者流程重生成原multiMasterDocument源与已原生准入fixed.pptx逐bytes相同，25739bytes，SHA943270e78b2fe373a20e2983c239b22211ee4c973bf9fa9d23f5b64fa0576e89（.tmp/inheritance-native-baseline-final.log）。最初误把normalized IR传给authoring-only generatePptx，失败日志保留；已改用声明合同，并单独比origin-enabled/ordinary IR相等，没有扩大Exporter入口。没有新的native操作，不借此关闭新shared-edit样本的原生验证。

接续element/placeholder/preset/run逐属性来源与本地取消覆盖，沿单一Core cascade；随后.1.3 artwork、notes/fields和原全部队列。旧三阶段native操作已完成不用重复；新共享编辑native样本待单独准入。未提交发布，目标active。最终workspace结果在结束后追加。

Theme/background来源最终验证完成：完整pnpm test退出0，Core172/Layout139/Parser696/Exporter57/Renderer19/CLI11、视觉33、导入119，以及隔离安装Node/browser/strict types通过。完整build/typecheck/lint已退出0，日志.tmp/inheritance-{full-build,typecheck,lint,full-test}.log；定向14项.tmp/inheritance-focused-final.log。源码/测试在全量前已固定，之后仅文档记录变化。final baseline bytes日志为当前生产代码，仍精确943270e78b2fe373a20e2983c239b22211ee4c973bf9fa9d23f5b64fa0576e89。

下一步从normalizeTextFallbacks / normalizePlaceholder / normalizeTextContent / normalizeTextRunStyle的实际级联入手，扩element/placeholder/preset/paragraph/run来源，必须同路记录选值，覆盖false/0/空文本及已有script-font null/barrier。master/theme/background的null规则不能覆盖element已有的脚本字体清除语义；格式源provenance仍由Parser持有。resolveNormalizedLayout接受新API的普通IR；generatePptx只接受作者document，不要再次误传IR。无待续runner，无新的native操作，没有提交发布；全部TODO目标active，当前切片不关闭AUT-07.1.2或其父项。


## 文字格式来源、原生off值与布局文字修复 — 2026-10-27

[API/验证](../testing/text-format-origins.md)实现TextBox/ShapeText/placeholder的frame、paragraph、run逐属性来源：defaults→placeholder preset→inline→element preset→local，scalar/partial margin、bullet/autoFit整项默认、false/0/空内容和script-font null/map-slot屏障；后续slot可重开。保留原PresentationInheritanceTarget interface，可继续extends；新增TextTarget和source preset，路径安全引用带点/引号preset名。metadata不进IR，表格/图表/notes来源与完整glyph选择仍待做。

真实失败两层：DOM先发现bold:false未写b=0（无该阶段native捕获）；修复共享placeholder paragraph/run默认和四种关闭值后，实际PowerPoint发现Shared仍消失——layout-owned对象带了p:ph重复prompt身份。原失败源和首PDF保留；只移除shared-text的一处p:ph后Shared恢复、旁路第二页全144dpi raster/raw trace不变。原旧multi-master源本轮12payload变化仅是4slide+4notes off值和4layout默认块，不能继续套上一轮bytes完全不变的结论。

当前prepared manifest docs/testing/text-origins-evidence.json，run-kmASsB；失败manifest text-origins-initial-failure-evidence.json，run-90YemX原source SHA361d2da63daaceda3cd04cb796fd51baa1829db265f1dcc15303704fdc0b4789。实际native目录.tmp/text-origins/native-QLLzu9；input.pptx/first.pdf是失败；fixed-input.pptx SHAef3d886dcf09d80e23b58cbbafb67cbb26be84fbb2a5c71797c2eb0a3219cd2d与当前public源bytes相同；fixed-first.pdf及fixed-round1/2/3已实际SaveAs/close/准确路径reopen，2首次+6重开页全raster/raw glyph trace/drawing geometry/实际嵌入font程序bytes精确一致。About面板证实PowerPoint16.104(25121423)。旧失败、原first、各round均保留。

独立oracle scripts/research/oracles/text-format-native.py及manifest text-origins-native-evidence.json：全master/layout/theme/slides/notes/content-type/关系按身份核；schema default fontScale100000/startAt1省略、body前导alpha0空run移除、child空run移除但完整endPara insertion样式保留，连同原noGrp/notes image locks/-0/empty guides具名规范化，其他内容精确。后两轮只core modified/lastPrinted和app TotalTime0→3→3变化。13实际坏ZIP+4实际坏PDF全拒绝，含原Shared消失和同栅格hmtx advance+1（校正checksums）控制。原四页oracle恢复原样、默认仍严格4页，没有用新样本放松旧验证。

六实际PPTX（failed/fixed/3round/list-control）Node/ESM/global完整inspection/report与18精确no-op另见text-origins-native-parser-evidence.json，已绑定最终native oracle JSON SHA；无Core完整master/artwork导入或source style writer权限扩大。Native contact与当前Chromium preview已查看。

8页list-control原生首次观察还发现前导空run alpha0会误套到SVG marker，以及负首行outdent错误。Renderer已按text-bearing run（全空用最后insertion）选marker字体/paint、忽略marker underline保留strike；首行钳制至frame左边，XHTML留hanging宽。独立DOM本样本marker local0/text local27精确。一般marker advance/换行/字体print量化仍open；不要把reported font size当图形affine，临时.tmp/text-origins-list-observation.json的source-point估算不是验收模型，最终native报告只存raw PDF coords。原16字体差异未关闭。

最终full build/typecheck/pnpm test退出0（.tmp/text-origins-final-{build,typecheck,full-test}.log），Core182/Layout139/Parser696/Exporter59/Renderer20/CLI11/visual33/import120，隔离安装Node/browser/strict types通过。全量前生产code/自动化case固定；之后仅加强Core已有case要求所有支持的有效格式leaf恰有一条来源，再跑Core/Exporter/Renderer/browser264组通过（.tmp/text-origins-complete-final-focused.log）。最终lint/docs/diff/所有artifact hash结果另追加。未提交发布，goal active。

接续：table-region/cell、chart text、notes独立theme/style的来源接同一selector/merge链；保留null barriers、source与effective分离。继续AUT-07完整reset/live-template、master artwork/images/hide/z-order、AUT-08字段与原全部TODO。3round本样本已完成不要重复；8list controls只有first-open，不能声称复杂lists/font完整准入。UI当前fixed-round3.pptx，styleNative helper保留在CUA，但恢复前先rewriteDocumentation并核真实窗口；旧mm helper仍指向旧native目录，别误覆盖。

最终收口核验：lint、包依赖边界、文档链接/契约/typed examples全部通过（.tmp/text-origins-final-lint.log）；129项产物hash及bytes、3个oracle依赖SHA、6份parser输入/快照、parser绑定native JSON SHA全部匹配（.tmp/text-origins-artifact-verification.log）。当前source SHA仍为ef3d886dcf09d80e23b58cbbafb67cbb26be84fbb2a5c71797c2eb0a3219cd2d；3round、13 ZIP/4 PDF negative与18 no-op计数一致。git diff --check通过。完整goal保持active，未提交发布。


## 图表与直接 authored notes 的格式来源 — 2026-10-27

上一轮是有实物/代码/证据变化的progress；本轮接续实际normalizer，未重做已验三轮native。为chart的common/title/legend/category/value文字五字段接同一selectStyle/StyleTrace；标题fontSize12/bold true是独立built-in，不错误归到chart common。显式false/默认相等值保留local；theme符号声明位置与归属master解析出的具体font/color分离，报告配合已有layout theme role来源读。pie无虚构axes；无title/hidden legend仍报告IR有效样式但不声称可见。

直接notes接现有normalizeTextContent，target为{kind:notes,id:slideId}，property notes.p.style/notes.p.runs.r.style，source为slide路径slides.i.notes...；默认、null/map-slot屏障、空文本/空数组/无run保留实际语义，不借slide preset/placeholder/元素样式。作者侧独立notes master/theme graph未实现，不把Parser源格式继承当作者模型。原PresentationInheritanceTarget interface未改，report unions扩新case。

8组Core测试独立读实际声明leaf/符号theme、要求全部有效leaf恰有一条来源，并测reset/immutable/两master交错及默认域/三chartfamily/layout/group。scripts/multi-master-parity新增组比较Node/ESM/global完整IR/origins/Layout/SVG/警告/bytes，独立DOM读chart title/axes与notes actual XML；原3组也通过。安装包Node/strict TS扩notes narrowing/clear和chart path，full run待最终terminal记录。审查最可能翻车5项：trace改IR、title默认误归common、符号theme来源被具体值覆盖、notes偷用slide placeholder、empty/reset遗留旧来源；均有对应定向case。

具名样本tests/fixtures/chart-notes-origins.mjs为4页11图表；scripts/prepare-chart-notes-origins.mjs输出.tmp/chart-notes-origins/run-LJku3G的10产物（726来源），manifest docs/testing/chart-notes-origins-evidence.json。3个script-map-materialized警告保留，不承诺字体map原生继承。preview已看；10产物hash/bytes、7依赖hash、726唯一target/property通过，保留旧source与.tmp/text-origins/native-QLLzu9/fixed-input.pptx bytes精确一致，SHA仍ef3d886dcf09d80e23b58cbbafb67cbb26be84fbb2a5c71797c2eb0a3219cd2d。没有为新chartnotesfixture做新PowerPoint运行；旧native证据不扩大准入。

当前build/typecheck/lint/boundary/docs/typed examples/diff通过（.tmp/chart-notes-origins-{build,typecheck,lint}.log），artifact检查见.tmp/chart-notes-origins-artifact-verification.log；full suite handle61715日志.tmp/chart-notes-origins-full-test.log还在运行，terminal/计数待追加。完整goal active，未提交发布。

下步继续table regional/cell来源：必须在table-style.ts实际layer/mergeTextRunStyleInput链保留声明，不事后把materialized style归为local；注意band row/column/edge/corner顺序、preset与cell覆盖、null whole-map后slot reopen、implicit blank无虚构声明、compact merged-grid authored索引与normalized索引不同。cell文字与margin/verticalAlign来源分别走真实selector；fill/stroke/border若纳入也需完整fieldwise/atomic契约。继续完整AUT-07 reset/live-template/artwork/images/hide/z-order、AUT-08字段及原所有待办；本轮不勾父项。

最终terminal收口：完整pnpm test handle61715退出0，Core190/Layout139/Parser696/Exporter59/Renderer20/CLI11、visual33、import121及隔离安装Node/browser/NodeNext strict declarations全部通过（.tmp/chart-notes-origins-full-test.log）；build/typecheck/lint/boundary/docs/typed examples也退出0。新10产物hash/bytes、7依赖hash、726唯一来源及保留native源bytes相等核验通过。只完成chart/direct-notes来源切片，不关闭table来源或父级goal；未提交发布。


## 表格真实合并链格式来源 — 2026-10-27

上一轮chart/direct-notes已实质交付，本轮按AUT-07.1.2接table格式来源；未重做旧PowerPoint三轮。新增StyleInputTrace承担raw fieldwise/atomic merge的声明sidecar，现有StyleTrace/normalizePaint/Stroke/Insets/Paragraph/Run在同一值selector中消费声明；metadata不进IR、无Core文件/格式依赖。table-style实际whole/border roles/column-band/row-band/edge/corner/preset/cell/content链逐步保留source，不将materialized字段假归cell本地。报告目标仍element/tableId，property为rows.r.cells.c.style与content.p.style/runs.x.style；source指作者真实rule/preset/occurrence。

覆盖fill、uniform stroke、sparse side candidate、margin/verticalAlign及完整现有paragraph/run字段。scalar margin统一指原scalar、partial分边；新paint无opacity用1且不继承低层alpha，none只type；bullet原子替换的省略字段恢复现有builtin。whole/slot null屏障及slot reopen保留state。styled隐式空格、空array/zero-run insertion延续原值语义，不虚构local位置；unstyled margin局部partial仍替换preset input后用builtin补边（保留旧行为）。逻辑边框候选与Layout共享物理边胜者分开；具名候选四边width0而导出右/下线1/3pt，独立DOM分别验证，未误判为源报告错误。

对抗式审查修掉真实别名位置bug：按cell引用Map找index会使重复对象第一次位置归到最后；table/style-grid.ts现在每anchor保留inputIndex，implicit无该字段。11组Core含重复对象来源回归、所有支持leaf恰有一条origin、独立读取实际source值/null、96个独立矩形分区的等值fill/text源owner、merge/preset/empty/reset/edit/immutable/group与两master。新scripts/table-origins-parity.test.mjs比较Node/ESM/global完整IR/origins/Layout/SVG/warnings/bytes，并用独立DOM读actual native cell边距/flags/paint/alpha/共享edge及XHTML padding/content。已接完整test:table-style。隔离安装runtime/types加implicit cell source及row/cell path契约。

具名fixture tests/fixtures/table-origins.mjs为3页，两styled master域+group内unstyled；preview发现Corner在大字体/list frame换行，改短End后已查看当前preview。scripts/prepare-table-origins.mjs生成.tmp/table-origins/run-BxpshR共11产物、2138唯一来源，manifest docs/testing/table-origins-evidence.json绑定12依赖，保留5条region/script materialization警告。旧run-fCR2zQ保留，不清理。11 hash/bytes+12 deps+unique rows已验证（.tmp/table-origins-artifact-verification.log）；旧native源EF3...和上轮chartnotes源28d95f7a061abbbdff0af3ac5bf003622eeb226f7289163bcfff2cd718c34774均actual bytes精确，chartnotes全部普通IR/726origin也精确。新tablefixture无新native app运行，不扩大原生style/appearance准入。API/测试见docs/api/table-format-origins.md、docs/testing/table-format-origins.md。

final build/typecheck/lint/boundary/docs/typed examples通过（.tmp/table-origins-final-{build,typecheck,lint}.log），diff无错误。发生occurrence修复前启动的full run handle66803已terminal0，日志另存.tmp/table-origins-before-occurrence-full-test.log，不作最终proof。最终完整full run handle87252仍在运行，日志.tmp/table-origins-final-full-test.log，待append terminal。生产代码/fixture/全部case已固定后才启动该最终run。goal active，未提交发布。

后续保持原全部TODO，不勾父项：AUT-07 master artwork/共享images/hide/z-order、完整reset/live-template及AUT-08字段；完整geometry/content origins、作者侧notes master/theme、native reusable table styles、source style/structure writes与既有字体物理差异仍open。master当前仅name/theme/background，layout已有elements；下一步先核真实registry/affected-slide校验、IR/Layout owner图、共享asset closure和native master/list IDs，再补master artwork基础及独立原生验证，避免扁平复制到每页。已验source-only notes inventory与三轮rect/default/文字源证据保留，不重复实现。

最终terminal收口：修复后完整pnpm test handle87252退出0（.tmp/table-origins-final-full-test.log），Core201/Layout139/Parser696/Exporter59/Renderer20/CLI11、visual33、table-style5（含新browser/XML case）、import121及隔离安装Node/browser/strict NodeNext declarations全部通过。11组新增Core包含96几何分区和重复cell occurrence回归。最终build/typecheck/lint/boundary/docs/typed examples均通过；11产物hash/bytes、12依赖hash、2138唯一origin及两既有source bytes相等校验通过。当前manifest目录run-BxpshR已查看preview，旧run保留。完整goal仍active，父项未勾选，未提交发布；后续按上段继续master artwork/隐藏/层次与原全待办。


## 母版 artwork 基础与原生三轮 — 2026-10-27（进行中）

上一轮table来源为progress。本轮开始AUT-07.1.3，已改Core/IR/Layout/Renderer/Exporter但尚未完成文档与最终收口，不得当整项验收。当前工作目录是权威，保留全部既有改动和旧证据。新增master.elements（optional，旧输入不加字段），全局ElementStore注册初始/新增/替换master对象，replaceMany staging/allocator rollback；masterId诊断/asset引用保护含unused/hidden；受影响layout/slide引用校验含省略layoutId的layout-blank。ManagedPresentationSlide新增setShowInheritedShapes，ManagedDocument.slides/add/insert/duplicate返回新能力；legacy PresentationSlide/Document不加必需method。layout.showMasterShapes=false只藏master，slide.showInheritedShapes=false藏layout+master，background与已物化placeholder内容不藏；省略为true，explicittrue导出1。

Layout.masters现在LayoutSlideMaster（resolved LayoutElement children），接image/group/chart/table、table sizing/intrinsic/analysis/owner theme/composition(master container)、pagination保留master IDs与图形ID。连接器layout/slide可用master几何（隐藏层仍可作几何anchor）；master内scope不跨owner。Renderer按背景→master→layout→slide，每owner一次。Exporter master own part+relationships，不复制到slides；先stage master XML/chart references再emit chart/workbooks，保持旧ZIP顺序；image/font/table/connector警告补masterId。两类flags写native showMasterSp。旧text-native准入源EF3...仍生成相同bytes。

当前新测试：Core7组（含global/nested/autogen IDs、atomic rollback、dependent connector/default layout、asset protection unused+hidden、flags/copy/setter、capture）；Layout3组（master chart/table/group、sizing/theme/composition/analysis、hidden target geometry）；Exporter2组（master charts/workbook ownership、invalid PNG failure）；scripts/master-artwork-parity.test.mjs一组Node/ESM/global全IR/origins/Layout/SVG/warnings/bytes、独立DOM owner/media/flag/z-order/preview及source-only exact no-op，已接test:import；安装包Node/types加managedSlide/master types。已有masters/definition-edit拒绝case将elements:[]换成malformed{}，不再拒绝新能力。

scripts/prepare-master-artwork.mjs生成docs/testing/master-artwork-evidence.json，当前prepared .tmp/master-artwork/run-OvUBoK，native .tmp/master-artwork/native-XYerBL，source SHA7773cbaa76a74d1bea8d714dc60ece07a2423b94b01b72aee441e061d535940b，13 artifacts含6 SVG；旧source EF3...在prepared作bytes回归。fixture tests/fixtures/master-artwork.mjs六页：red-on两层、red-on slide false两层藏、red-off master藏layout显、blue-on、red-off slide true仍master藏、red-on local bg。

实际PowerPoint首次无repair，first-open.ax/png和first.pdf已捕获；artwork-round1/2/3.pptx及pdf均实际SaveAs/close/准确路径reopen，saved/closed/reopened AX/png完整。About实际16.104(25121423)已捕获且dialog已关闭，当前UI artwork-round3.pptx。CUA artNative={...styleNative,app:artApp,base:native-XYerBL,stage:artwork}助手保留；round拒绝覆盖，三轮已做勿重跑。恢复UI前rewriteDocumentation再核窗口。native PDF实际6首次+18重开页全raster/raw trace/drawings/image instances/font bytes精确一致。源frame affine使用实际背景rectangle的sx/sy，按XML720x540校验各vector/image bbox<0.0001 PDFpt；不是拿字体size当affine。sourcePNG全图8x4，native0/1/2实例符合控制；overlap pixel查master/layout/local顺序，body/page对应验证。native-contact.png已查看。

独立scripts/research/oracles/master-artwork-native.py已terminal0，log .tmp/master-artwork-native-oracle.log，manifest docs/testing/master-artwork-native-evidence.json，14 actual ZIP+4 PDF坏文件拒绝（含同raster font hmtx advance+1有效checksums）。全图/relationships/contenttypes/notes/master/theme逐项语义，不改旧strict4page oracle。具名normalization只有media按exact bytes哈希alias（logo.png→image1.png，仍一图）、showMasterSp=1默认省略、精确720x540增加screen4x3 type、旧pPr -0/noGrp/notes locks/empty guides；其他精确，metadata首存full delta/后两轮仅clock+TotalTime。完整first6/reopen18记录、48artifact hash及3oracle依赖。Native table/chart-master appearance及native shared edit仍没验，不能用group/image子集关闭父项。

当前full pnpm test handle76119日志.tmp/master-artwork-full-test.log，原生oracle工作期间还在执行，需poll同handle确认terminal；production改动已固定后启动。新定向Core/Layout/Exporter/browser均通过；pnpm typecheck handle40992已terminal0。先前build78232通过（devapp常规>500k提示保留）；之后Exporter warning masterId补齐已package build/test通过，full command会重建全部。仍需：实际4native输入Node/ESM/global full inspection/no-op（可仿scripts/inspect-native-text-origins.mjs改4输入/6页/3master4layout，绑定newnativeJSONsha，不要改旧script）；docs/API/typed examples/architecture/ROADMAP/backlog及final build/typecheck/lint/docs/diff/hash。若新增原生oracle断言/negative后重跑native JSON，后续parser绑定也要刷新。

正读主源MicrosoftDocs：working-with-slide-layouts 的showMasterSp、VBA slide.displaymastershapes（背景对象不含background paint）；实际native六页已验证slide hides both/layout hides onlymaster。未做任何commit/publish/reset/cleanup。全goal active，原所有TODO不缩减。


最终收口 — master artwork：完整最终pnpm test handle4040 terminal0，.tmp/master-artwork-final-full-test.log，Core208/Layout142/Parser696/Exporter61/Renderer20/CLI11/visual33/import124/installed Node-browser-strict NodeNext全通过，包含新增3个public cases（baseline browser/XML、shared replacement isolation、source-write refusal/reset/no-op/stale plan），所有production/tests固定后才启动。前一full run handle76119也terminal0保留日志不覆盖。final build48109/typecheck/lint20856/boundary/docs均0；diff通过。core7/layout3/exporter2/public3当前全验。

docs/api/master-artwork.md、typed example、testing doc、ADR0057/IR-v2/document model、现有multi-master API/asset lifecycle/capability matrix/ROADMAP/backlog已同步；AUT-07.1.3.2/.3仅进度前缀，仍未勾父项。source-only guard实际测试：setShowInheritedShapes(false) source edit blocked/save-edit-unsupported，resetundefined恢复exact no-op；新建master likewiseblocked，旧plan save-snapshot-stale。shared fresh replacement移除red master logo/改shape-text仅变red master XML及rels两payload，4非显示/blue pages SVG精确，2 red visible pages更新、media仍由others保留，未扩Parser权限。

最终native oracle .tmp/master-artwork-native-oracle-final.log terminal0：6first/18reopen pages、15 actual ZIP/4 actual PDF negatives，RGB PNG的独立PIL decoded MD5与每native image实例相同、actual text colors/faces一致；84artifacts proof里13prepared+71native，Native manifest3依赖SHA；oldnegative folders保留。最后Node/ESM/global实际4输入12 no-op .tmp/master-artwork-native-parser-final.log terminal0，nativeEvidenceSHA已刷新。scripts/verify-master-artwork-evidence.mjs terminal0比对当前public Pptx bytes/IR/origins/Layout/6SVG与捕获全exact，并验证84artifact hash/bytes、39implementation deps、3oracle deps、4parser snapshots和parentSHA。prepared manifest现在带deps/native+parser链接（docs/testing/master-artwork-evidence.json），current source仍7773cbaa76a74d1bea8d714dc60ece07a2423b94b01b72aee441e061d535940b；native目录XYerBL，preparedOvUBoK，UI仍artwork-round3且About关闭。无需重做这六页三轮。

真实后续问题（非本轮新引入）已记录：Source720x540的docProps/app.xml PresentationFormat仍来自legacy常量`On-screen Show (16:9)`，native改为`全屏显示(4:3)`；实际源frame与print affine几何4:3已经严格验。下一步先修Exporter.appPropertiesXml按显式layout.size给正确presentation-format label，保留默认16:9源bytes、加4:3/custom/其他size测试；若更改当前source，旧7773源/native与当前source必须做明确delta或独立新证据，不可篡改旧files/冒充已native打开。之后继续master advanced table/chart/connector actual appearance、native共享art编辑及完整reset/rebinding/images/fields等原所有待办。core/source/notes权限界限不缩，所有parent仍open，完整goal active。未commit/publish/清理旧文件。


## 2026-10-27 — PresentationFormat 元数据修复（最终全量测试待收口）

沿上一断点完成Exporter `appPropertiesXml`按layout.size的已序列化EMU精确比值输出16:9/4:3/16:10，其余Custom；BigInt交叉相乘避免unsafe，helper非法/舍入零fallback不代表新尺寸准入。未加Core字段/preset/几何改变/Parser写权限。生产仅package-parts和build-package接size。4组Exporter、8场景Node/ESM/global完整IR/warnings/bytes+独立ChromiumDOM+结构no-op已通过；portable tests/fixtures/master-artwork-format-baseline.json保留7773源54part黄金，不依赖.tmp；默认16:9 EF3… byteexact。

新prepared .tmp/presentation-format/run-pTh2Vg，独立native .tmp/presentation-format/native-EU32W4，sourceSHA acbe40dc3140880d1da734d13af59405ad6b740c7585884dd7e89bbbc0b98770；对7773cbaa76a74d1bea8d714dc60ece07a2423b94b01b72aee441e061d535940b只有app.xml一个16:9→4:3字符串，53其余payload exact。旧7773所有native文件/原oracle严格规则保留，未冒充新source已打开。

实际PowerPoint16.104(25121423)新input首次打开、first.pdf、format-round1/2/3 SaveAs-close-exactreopen-print全部独立完成，saved/closed/reopened AX/png与About版本保存。当前CUA UI format-round3，About已关闭。formatNative={...artNative,app:formatApp,base:native-EU32W4,stage:format}助手存在；再用UI先rewriteDocumentation/核窗口。不得重跑覆盖这三轮。

scripts/research/oracles/presentation-format-native.py已terminal0：6first与旧7773first实际PDF各raster/trace/drawing/image/font逐项exact；18reopen页与newfirst exact，master/layout/body/page/visibility/geometry/关系仍调用unchanged master-artwork-native.py严格检查。newsource英语4:3/native中文全屏显示(4:3)、9144000×6858000EMU明确验；4实际corruptZIP（wrong/missing label、1EMU frame、body）拒绝。8自动尺寸不是8native尺寸验收。docs/testing/presentation-format-native-evidence.json含33native artifact/4oracle deps。scripts/inspect-native-presentation-format.mjs terminal0，4actualinputs Node/ESM/global12exactnoop，nativeJSON SHA绑定，见presentation-format-native-parser-evidence.json。

scripts/verify-presentation-format-evidence.mjs已terminal0：43prepared/native hash-byte、11producer/test deps、4oracle deps、8scene重生成、newsource/IR exact、12noop与snapshot验证。scripts/verify-master-artwork-evidence.mjs现在保留原evidence.dependencies作为历史producer，只新增currentProducer分离当前依赖/源SHA/format native证据/one-string bridge；84oldartifact+3oldoracle/39currentdeps/old12noop及IR/origins/Layout/6SVG exact仍通过。旧nativeJSON与文件未改。文档API/testing/ROADMAP/backlog已明确历史和新proof，父项不勾。

pnpm build handle31346 terminal0 .tmp/presentation-format-final-build.log；typecheck terminal0 .tmp/presentation-format-typecheck.log；lint handle75804 terminal0 .tmp/presentation-format-lint.log，含boundary/docs/typed examples；git diff --check通过。完整pnpm test handle78140 .tmp/presentation-format-full-test.log尚待poll；当前Core208/Layout142/Parser696/Exporter65/Renderer20/CLI11/visual33及其余前序已过，import正在跑既有较慢notes font edits（进程有CPU），不可先写fullgreen。所有当前production/tests在启动前固定。

对抗式审查五项：float epsilon误判、unsafe crossproducts、误改sizepreset/geometry、旧proof冒充当前、bodyoracle漏查metadata，均有精确边界/黄金/独立三轮/label负例。更广原生尺寸、invalid dimension validation、master table/chart/connector实际外观和共享编辑、完整reset/rebinding/notes/fields及所有原TODO仍open。下一步全量terminal后append收口，再接advanced master实际nativecontrols（现有仅自动master chart/table/group工作簿关系，不算原生验收），不重复普通sixpage三轮。完整goal active，未commit/publish/reset/清理。


最终terminal收口 — PresentationFormat：完整pnpm test handle78140退出0，.tmp/presentation-format-full-test.log，Core208/Layout142/Parser696/Exporter65/Renderer20/CLI11/visual33/import125、隔离安装Node/browser/strict NodeNext declarations全通过；没有skip/cancel/todo。final build31346/typecheck/lint75804/docs6186均0，git diff --check0。43newartifact/11producer deps/4oracledeps/8scene重生成/12actualnoop验证和84oldartifact/39currentproducer deps/oldoracle严格边界+metadata bridge均terminal0，日志.tmp/presentation-format-artifact-verification.log、.tmp/presentation-format-artwork-bridge-verification.log。docs/testing/presentation-format.md为本轮scope与五项adversarial审查，旧master-artwork manifest保留历史deps另设currentProducer。未改旧source/native identities。所有父项仍open，goalactive；继续advanced master table/chart/connector实际appearance/shared编辑及原完整TODO，勿重复已有基础实现/本轮三轮。未commit/publish/reset/清理。


## 2026-10-27 — advanced master 原生验证发现真实 table serialization bug（进行中）

上一轮PresentationFormat已fullgreen progress。本轮新tests/fixtures/master-advanced.mjs双master red/blue、4pages(redvisible/hidden/redvisible/blue)、group合并table、三类bar/line/pie chart、master native attached straight connector。scripts/prepare-master-advanced.mjs生成原before .tmp/master-advanced/run-NPSqNa/native-UyNU16，source80f396071c3c22f3f79e98dd909dd9e81c9b397faadf809cbe79882eb9619447，两table-style-materialized masterId warnings，未加新Core功能。实际首次无repair和advanced-round1/2/3 SaveAs-close-exactreopen-print完成，AX/png齐全；first-contact已view，charts/connector/table可见，hidden仅PAGE1。未完成独立strictoracle不得算advanced父项验收。

真实故障：无边框table实际6条white线/visiblepage。原Exporter tableCellXml在tcPr中先fill后border，违反CT_TableCellProperties sequence，native保存丢掉anchorborder恢复Office白色style。单变量源码ZIP control sequence-control.pptx仅两masterpart的tcPr border-before-fill规范化（XMLprefix serialization等价），原生sequence-control.pdf白线6→0。另margin4pt原写a:bodyPr被native忽略实际default7.2pt；cell-frame-control.pptx在sequencecontrol上只把lIns/rIns/tIns/bIns/anchor移至tcPr.marL/R/T/B/anchor，原生textx76.14444→72.83333，delta(7.2-4)*actual745/720；topbaseline差.24 nativepoint也保留。两个actualcontrol AX/png/pdf保存在原UyNU16不清理。Microsoft官方TableCellProperties页面+Primer XSD已browse，文档待加。

production packages/pptx-exporter/src/ooxml/elements.ts唯一tableCellXml变更：bodyPr空、margin/anchor属tcPr、四border在paint之前，注释解释。无Core/IR/layout/Parser改动；Parser本已优先读tcPr，兼容老bodyPr读取不改。新Exporter test/table-cell-properties.test.mjs一组27组合(3owners slide/layout/master×3align×none/solid/opacity theme)，nestedgroup/mergedanchor/ghost/zero+sparse+positive margin/solid border测试schemaorder/propertyplace。已通过。scripts/table-origins-parity.test.mjs读取tcPr而非bodyPr，scripts/inspect-table-style.mjs strict检查props.mar*。旧scripts/table-style-parity.test.mjs descendant solidFill误读borderpaint导致native-fill失败，现改directchild；定向style+origins test handle30366已terminal0 .tmp/master-advanced-table-style-fixed.log。第一次table-style23767退出1保留log .tmp/master-advanced-table-style.log。不隐瞒此失败。

原before manifest已copy到 docs/testing/master-advanced-before-cell-fix-evidence.json，current docs/testing/master-advanced-evidence.json指新source9643618a0df3f2dc10111acc645c5447e85bdcd2c5eb0c18b5cd37927176fcec，prepared run-y4xpHp，native-QRT0SF。源码old→new只有slideMaster2/3.xml的cellproperty move/order，其他ZIP exact，独立Python canonical按此transform逐项验证过。新source四firstpages，cell-fixed-round1/2/3实际SaveAs-close-exactreopen-print全部完成，saved/closed/reopened AX/png，实际About16.104(25121423)已capture且关闭。当前UI cell-fixed-round3；CUA cellNative={...advancedNative,base:QRT0SF,stage:cell-fixed,app:advancedApp}助手仍有，下一次UI先rewrite/getApp绑定。勿重复已有三轮。

nativefirstsaved tcPr保留全部mar*=50800和四noFill borders，anchor=t按schema默认省略。原advancednativechart首次变化很多，不得宽泛忽略：workbooks6只是renameMicrosoft_Excel_WorksheetN→Microsoft_Excel____*.xlsx、6payload SHA原样(各red/blue3同payload)；charts首存langzh-CN/style2/autoTitleDeleted/invert defaults/dLbls defaults/numfmt/general/float lexical/mode defaults/uniqueID扩展，table row/colID extensions。完整delta与strict后续回合尚待oracle，旧fourpage/sixpageoracles未改，不用结果冒充sourcefirstsave all exact。

当前processes：full pnpm test handle36700 .tmp/master-advanced-cell-full-test.log仍跑（在修DOM测试前启动，不能作最终全fixed proof，即使终端过也记前序run）；final full需要全部tests固定后再启动。final pnpm build handle90701 .tmp/master-advanced-cell-final-build.log未poll；pnpmtypecheck85739 .tmp/master-advanced-cell-typecheck.log未poll。Exporter单build80549已0。还没lint/docsfinal/actual-fileParser12noop/nativeoracle/hashfinal/API文档。

下一步先独立scripts/research/oracles/master-cell-properties-native.py：验old3rounds保持错误稳定+2单变量control+new3rounds每tableprops exact(仅anchor=t默认省略允许)、PDF无whiteborder/textx按actualframe4pt，fullrawtrace/drawings/raster/font/img后续全exact；fullpackage round2/3对round1 canonical仅validclocksTotalTime允许，其余allparts严比；source→firstsave完整delta记录，targetedowner/chartcache/workbook/resource/nativeconnector invariants，未验证chart首次样式不宣称全source语义不变。实际坏ZIP/PDF负例用于property和same-rasterfont。之后4actual文件Node/ESM/global fullreports+12noop、currentnativeJSONSHA绑定，publicDOM当前更改已测试。需doc/API说明tablepropsfix意外bytes/实际外观变化，old80…错误证据保留，父项不勾。全部goal active，后续原生共享编辑/reset/fields及全TODO不缩，未commit/publish/reset/清理。


后续进度 — table cell fix：原生独立oracle scripts/research/oracles/master-cell-properties-native.py terminal0(.tmp/master-cell-properties-native-oracle.log)。验old3rounds错误6白线稳定、sequencecontrol0白线、framecontroltextX恢复4pt、correctedfirst全PDF与framecontrol exact、new3rounds12页fullraster/rawtrace/drawings/images/fontbytes exact；newtableprops原生全部mar*=50800、anchor=t仅schema默认省略、border四noFill保留。源old80f…→new9643…只有两masterpart cellprops move/order，其他payload exact；后两轮完整allparts对firstnative，仅ISOmodified/lastPrinted及numericTotalTime归一，其余严格。6chart caches、6workbooks payload exact、closure/owner/page/hide/nativeconnectorXML无变化，6实际ZIP负例拒绝。source→firstnative全部58条entry delta完整记录，不忽略、不宣称chart styling等价。master-cell-properties-native-evidence.json当前94native/before/control/negative/parser artifacts（以实际manifest数为准），4oracle deps。无新PDF负例，勿将旧19negatives冒充本轮；本轮6ZIP。

新scripts/inspect-native-master-cell-properties.mjs在oracle最终更新后重新完成，handle49621terminal0，master-cell-properties-native-parser-evidence.json绑定当前nativeJSONSHA，4actualinputs Node/ESM/global12exactno-op、3masters/3layouts/4slides fullinspect/report。之前82687/58953的parser目录保留，当前以新manifest绑定为准。scripts/verify-master-cell-properties-evidence.mjs已terminal0，currentPptx/IR/origins/Layout/warnings/4SVG重生成、103hash/bytes（加入已viewnative-contact后将104）、9producer/testdeps/4oracledeps/12noop/4snapshots验证；.tmp/master-cell-properties-artifact-verification.log。JSON序列化Layout会省略pie:undefined，因此比较同JSON序列化视图，未改变布局或放松其他值。newcontact在prepared run-y4xpHp已view，manifest新增hash。

新增docs/testing/table-cell-properties.md、API exporter cellprops契约、ROADMAP/backlog进度；parent仍unchecked。明确cellfixfresh bytes/appearance改变、旧source/nativefailure保留，nativeadvanced firstsavechart58delta语义审查和actual sharedediting仍open。生产只有elements.ts表格cellXml，最终build90701/typecheck85739terminal0，lint17132terminal0含boundary/docs/typed examples，diff0。全量first handle36700 terminal0日志.tmp/master-advanced-cell-full-test.log，但在DOMtestfix前启动不作final。最终全量handle19772日志.tmp/master-advanced-cell-final-full-test.log，所有production/tests固定后启动，目前仍跑需poll；最后新增只oracle/verifier/docs不改production或CItest。需再final lint/docs/diff/hash因verifier JSON比较修正后更新；oracle已最终fixed，勿重跑造成nativeJSONSHA变化又需parser绑定（无新断言需求不重跑）。当前UIcell-fixed-round3且About关闭，cellNative helper；未commit/publish/reset/清理，goalactive。


最终terminal收口 — table cell native修复：最终full pnpm test handle19772退出0，.tmp/master-advanced-cell-final-full-test.log，Core208/Layout142/Parser696/Exporter66/Renderer20/CLI11/visual33/import125与installed Node/browser/strictNodeNext全通过，所有production/CItests在该run启动前固定。前序full36700亦0但仅保留前序日志不替代final。final build/typecheck/lint/boundary/docs/typedexamples/diff均0；104artifacts hash/bytes+9producerdeps+4oracledeps+4actualParserinputs/12exactnoops通过，manifestcurrent source9643618a0df3f2dc10111acc645c5447e85bdcd2c5eb0c18b5cd37927176fcec。新doc docs/testing/table-cell-properties.md已记五项adversarialscope与真实fail/control/repaired3rounds。普通master-artwork源acbe…及format8scenes current再生exact验证0，oldmastermanifest仅currentProducer依赖刷新不改历史source/deps/native。日志.tmp/master-cell-properties-{artifact-verification,artwork-compatibility,format-compatibility}.log。preparedrun-y4xpHp含已viewcompactnative-contact；nativeQRT0SF为corrected，UyNU16为before/isolatedcontrols，所有均保留。

此轮具体改变使真实无边框与4pt margin在PowerPoint生效，不算completeadvancednative或全goal。下步继续master advanced三类chart首存58entry差异语义（sourcecache/workbook已exact，不代表fullstyling）及实际native共享母版artwork编辑（改red影响两个visible页、hidden/blue不变），然后完整reset/rebinding/字段及原所有待办。当前UIcell-fixed-round3，About关闭，切UI先rewrite/getApp绑定；不要重做这次corrected3rounds。完整goalactive，父项仍open，未commit/publish/reset/清理。


## 2026-10-27 — master shared native heading + chart首存逐项审查（进行中）

上一轮cellprops真实bug修复fullgreen progress。本轮没有production/API改动；补原计划AUT-07.1.3的实际shared编辑及六chart具名原生审查。tests/fixtures/master-advanced.mjs新增可复用editMasterHeading helper，Core replaceSlideMaster只改red-heading的MASTER red→MASTER red 2；新scripts/master-advanced-parity.test.mjs Node/ESM/global完整IR/origins/Layout/SVG/warnings/bytes两个状态一致，blockednetwork，独立DOM验两个master3chart各自owner、mergedtable margin/order、localnativeconnectorIDs；source仅slideMaster2.xml变，SVG0/2变且1/3 exact。新1case已通过两次，加入package.json test:import；无Core生产变动，兼容保留全部旧改动。

实际native copy源为 .tmp/master-advanced/native-QRT0SF/cell-fixed-round3.pptx SHA0f06d3c164b9501a7a58795c94cae485f85a80890cd36ff472cd6e790d3ba58c，独立新dir .tmp/master-advanced/shared-native-vwggt5tr/input.pptx同bytes，input-binding.json指oldactualpath/sha，不覆盖旧文件。actualPowerPoint viaView→SlideMaster→第二个masterroot(红)→双击MASTER red标题→CmdA文字范围截图确认→pasteplain MASTER red 2→关闭master。master-selected/master-text-edited/after-edit-open AX/png保存，before/after-edit PDFs actual。shared-round1/2/3实际SaveAs-close-exactpathreopen-print，saved/closed/reopenedAX/png全，About16.104(25121423)capture并已关闭。

本轮strict scripts/research/oracles/master-shared-native.py：全package只允许redmaster paragraphtext/two dirty=0 +core revision1→2、合法modified/lastPrinted/TotalTime；rawparts仅master2、core/app变。后3rounds fullnativepackages/sourceprops/table/caches/workbooks/connector/owner/hide exact。两个visible red heading更新；all other traces/oldheadingglyphs/drawings/images exact；headingbbox外pixels exact；hidden1/blue3 fullraster/texttrace exact。发现PDFTT2ArialMT字体cmap format6映射顺序改变(added2提前)，字体其余table含glyf/hmtx/units/sidebearingexact；进一步实际检查证明stream cmap之外全bytes包括tableDirectory也exact，已加strictassert。font资源xref/名字list exact；cmap仅单subtable format6及sameglyphmapping set permutation允许。其他3fontprogram全exact。12reopenpages对immediateafteredit fullrecords/fontbytesexact。4actual ZIP坏文件(bluechange/localcopy/hideflag/targetfont)+3actual PDF(coveredtable/pageswap/same-raster repairedfontadvance)都拒绝。绝不将cmap许可推广任意fontchanges。

scripts/prepare-master-shared.mjs产新shared-run-8eOxvg，fresh before9643618a…，freshafter b4fb8dd173f387a6237b69098316584be8df45f52da987d05aeaacbdd00f94b3，prepared source/IR/layout/origins/warnings/4SVG约11文件。此helper对native fresh-core-edit.pptx现在flagwx，禁止覆盖；不要重跑prepare替代既有证据。fresh-core-edit.pptx actual已open/print/capture，4pages fullraster/trace/drawings/image/fontbytes与manualnativeafteredit全exact；仅首次fresh Core打开，不冒称另3rounds。currentUI fresh-core-edit.pptx normalview；sharedNative helperapp:sharedApp base:shared-native-vwggt5tr stage:shared；以后UI先getApp绑定/若compact rewriteDocumentation。

新chart oracle scripts/research/oracles/master-chart-first-save.py不改老oracle，review oldcurrentQRT0SF源6positive两点A/B/bar-line-pie/master图表，3actualoldround18完整chart XML比对，限定native语言enUS→zhCN/style2/autoTitleDeleted0；sourcew/h factor模式和bar overlap0省略；manual x/y/w/h转float.hex要求binary64完全相同无epsilon；General categoryfmt、disabled dLbls、pie无显示labels时leader1、bubble3D0、positivebar invertIfNegative1及精确白色inversefill、10uppercase wellformed unique GUID。其余所有chart元素/顺序/属性/值exact；workbook payload、cache、ownerclosure沿strictcellowner验。11actual坏ZIP拒绝color/font/oneULP/mode/visiblelabel/fmt/inversefill/guidbad/guidduplicate/unknownext/cache。docs/testing/master-chart-first-save-evidence.json绑oldcellnativeJSONsha，不新造native运行。这个具名recipe的观察不是全球style2/locale/negativevalues/visiblelabels样式等价认证，父项不勾。

现有docs/testing/master-shared-native-evidence.json已生成，native总文件随negatives/parser保留增加；刚加outsideCmap stream严格断言后重新运行handle（看下一工具output，日志.tmp/master-shared-native-oracle.log）还需poll terminal并重跑scripts/inspect-native-master-shared.mjs刷新parentSHA。此前27929 parser已0，5actualinputs(input/r1/r2/r3/freshCore)Node/ESM/global15exactnoops；当前新oracleJSON更改后它会stale，必须刷新。docs/testing/master-shared-native-parser-evidence.json最终以新绑定为准。scripts/verify-master-shared-evidence.mjs已经过一次，之后oracle更新还需reverify；它regenfreshbeforeafter/JSON/Layout/SVG、绑定native+chart+parser全parentSHA和artifact/deps，目前preparedmanifest docs/testing/master-shared-evidence.json新鲜current需最后重写links。添加contact PNG还未做(可做prepared不放native dirs避免oracleSHA再次变化)。

Docs已新增docs/testing/master-shared-edit.md，API/master-artwork、testing/master-artwork、table-cellprops、ROADMAP/backlog/advancedTODO进度链接父项unchecked。5项adversarial:flattening/sharedpages/fontmetricbehindpixels/过宽chartnorm/混淆freshsource与nativecopy，各有具名严格检查。第一份chartreview .tmp/master-chart-first-save-oracle.log已0，11negatives的第二个dir保留，不清理。

当前最终fullpnpmtest handle87830 .tmp/master-shared-final-full-test.log仍跑，production与新CItest helper已固定后启动；目前import已64到慢notesnativefontedits，没有发现失败，不能先写fullgreen。pnpm finalbuild57962终端0 .tmp/master-shared-final-build.log，typecheck0 .tmp/master-shared-typecheck.log；lint74541 terminal0 .tmp/master-shared-final-lint.log，但outsideCmap oracleassert/doc更新后需final lint/docs/diff。旧cell/ordinaryart/format manifest依赖fixtures/packagejson已变(新增helper/case)，需各verifier重新绑定currentdeps并确认source/output不变：verify-master-cell-properties-evidence、verify-master-artwork-evidence、verify-presentation-format-evidence；不改旧nativeJSON/sourcefiles。fullgoalactive，不commit/publish/reset/cleanup。下步收口本轮后继续native master其他普通art/image editing与advanced table/chart data/style，再完整reset/rebinding/fields及原全TODO，不能用text子集关闭全部shared gate。


最终oracle/hash收束（full test仍待）：outsideCmap断言实际通过，sharedoracle最终 .tmp/master-shared-native-oracle.log已0；新nativeevidence包含freshCore actual firstprint4pages与nativeedit fullrecords exact，全部fontstream cmap之外bytes含directory也exact。最后scripts/inspect-native-master-shared.mjs handle10051终端0，parentJSONsha已刷新，5actualinputs/15noops当前有效。prepared shared-run-8eOxvg/native-contact.png已view并加hash，12preparedartifact。scripts/verify-master-shared-evidence.mjs终端0 .tmp/master-shared-artifact-verification.log：90hash/bytes(12prepared+67native+11chart negatives)、13producer/testdeps、10oracledependencybindings(两个5项，重复共享helper非10独立文件)、3native sharedround12pages、18chartcomparison、4ZIP3PDF共享负例、11chart负例、15noop全验。当前nativeJSON严格规范不再改，勿无故重跑oracle导致parent绑定stale。current UI fresh-core-edit且About关闭。

cell/ordinaryart/format current兼容verifier也均0，日志.tmp/master-shared-{cell,artwork,format}-compatibility.log。更新的只是各preparedmanifest currentproducer deps(新增fixture helper/package CIcase)，旧nativeJSON/source/evidence files不变。final lint60849终端0含boundary/docs/typedexamples，diff0；finalbuild/typecheck已0。fullpnpmtest87830 .tmp/master-shared-final-full-test.log仍需pollterminal，最新import67+在慢notes阶段后继续。没有production更改，全部CItests/helper在run启动前固定。完成后append terminalcounts与docs/testing/master-shared-edit.md Check record；父项仍open，完整goalactive，未commit/publish/reset/清理。后续继续actual master其他objects/images shared edits、table/chart style/data实际editing、完整reset/rebinding/field和原全部TODO，不把ordinaryheading这一子集改叫全goal。


最终terminal收口 — master sharedheading/chart review：fullpnpmtest87830退出0，.tmp/master-shared-final-full-test.log，Core208/Layout142/Parser696/Exporter66/Renderer20/CLI11/visual33/import126、installedNode/browser/strictNodeNext全通过，所有production/CItests/helper在run前固定。finalbuild/typecheck/lint/boundary/docs/typedexamples/diff均0。shared/charts两独立oracle最终均0；90artifactsha/bytes、13producerdeps、两proof共10oraclebindings、5actualinputs15sourceonlynoop全验，见 .tmp/master-shared-artifact-verification.log。新docs/testing/master-shared-edit.md已记录checkscope及5项adversarial，父项未勾。freshCoreafter实际firstprint4页与manualnativeedit全records/fontbytes exact，nativeedited3round12页完整exact；hidden/blue前后fullpixeltext exact，外editedheading范围glyphmetricfontstreamcmap之外全byteexact。4ZIP3PDFsharednegatives及11chartZIP独立真文件拒绝。sourceidentitiesprepared/current/natives分别绑定保留，旧native证据未覆盖。

当前UI fresh-core-edit.pptx，nativecopiedinput0f06…/freshoriginal9643…/fresheditb4fb…不同身份不能混；helper sharedNative base shared-native-vwggt5tr stage shared可恢复，original3rounds不重复。后续按原全TODO继续native master其他ordinaryobjects/images shared edits、实际table/chart data/style编辑及完整reset/rebinding/fields；本轮只关闭具名heading/shared与positivechartreview子范围，generalparent和全goal仍active。未commit/publish/reset/清理。最新三个恢复记录为权威明文状态。


## 2026-10-27 — AUT-07.2 rebind authoring foundation（实现进行中，未验收）

上一轮sharedheading/chart proof已fullgreen，不能代替模板换布局。当前转原TODO AUT-07.2.1/.2/.3；实际检查PresentationSlideImpl.layoutId/background为creation-only readonly，无换布局API，replaceSlideLayout仅同identity共享定义。AGENTS/RTK/相关Core实现与架构已读。不spawnagents，全部保留existingdirtytree。

已新增public pure prepareLayoutRebind(data,requests,options) + ManagedPresentationDocument.applyLayoutRebind(plan)，legacy PresentationDocument/Slide无新必需methods；新types/layout-rebind.ts(index exports)，document/layout-rebind.ts。新PresentationSlideData/PresentationData是read-only authoringvalue aliases，无mutationmethods；public normalizePresentation/resolveInheritance以及validator/internalNormalizer改收data，去掉fake methods requirement，旧factorydocs仍structurallycompatible。未改IR schema/新包/依赖/Parser source permissions。

Plan opts requiredlocalAppearance preserve/adopt，unmatched默认block可detach，adopt-only reset unique geometry/text/background。Request slideId/targetlogicalLayoutId，optional bindings elementID→placeholderkey/null；optional references oldInheritedID→newInheritedID。Samekey+samekind优先，缺identity时仅uniquekind自动；多候选/同key错kind/missing默认needs-review，候选sortedkeys不按遍历猜；普通objects不自动绑定；explicit可改变role附warning，explicitnull可detach。脱离bound对象box使用oldresolvedbox防丢失。Core无nativeidx，尚未引入arrayindexidentity，需ADR/doc如实界定父项不关。

preserve materialize normalizedforeground所有7types(含nestedgroups/tablecells/line-chart marker+axes/colors/fonts/scripts)到localrawinputs，tokens按oldownertheme转concrete，background也oldresolvedpaint；table regionrules冻成cells并warning，dropstyleSource/plainText等IRderivedfields；targetmaster/layoutartwork仍新owner，不复制扁平化。adopt保留rawlocaldeclarations供newtheme/placeholder继承；显式reset仅matchedgeometry/text及pagebg，textreset保留runs文本/actions，清localstyles/preset/frame。Core deterministicoverflowestimate/font-familyavailability-unverified/shrink-policy警告，不偷偷调字号。对换掉inheritedanchor的connector要求referencesmap，preserve若map改变targetID则needsreview(不把Layoutanchor/transformresolver抄进Core)；localconnectorIDs保留。更广resolvedgeometry preservation/reset/fields/crossdoc尺寸/nativeindex仍open，不能当全AUT07.2。

Plan是frozenvalue侧车：sourceSnapshot为明确JSON authoring snapshot string(非hash/token)，requests/options/report/matches/逐effectiveleafchanges/changedSlideIds。apply只读rootinput字段descriptors不execgetters，重算stage不信forgedstatus/report/candidate；checkcurrentSnapshot在normalization前，任意selected/unselected/rawdefinition/notes/asset变化stale-plan，currentinvalid/nonfinite也stale。needs-review无afterchanges preview以免冒充已选目标；可取消不apply。atomicapply在ElementStore.replaceMany一次stage所有changedslides roots/IDs/allocator→全candidateData validator→storecommit→sameSlideImplhandles更新layout/bg/elements。selectedempty也支持；unselected/master/layout/assets/notes/timing/slideorder保留。PresentationSlideImpl内部commitLayoutRebind(不在publicSlideinterface) mutablelayout/bg，entries逐项push避大spread commit后异常。无新ids，old/rootnested ownership保持。

当前Core tsc -p已0(最近65390最初有readonlyGroupInput/linefalse type errors已修；cloneInput recursive用于readonlygroup，linefalse映射none-width0，descriptors asobject取runtimePlain字段)。新tests/fixtures/layout-rebind.mjs三页source/target/choices/empty layouts、red/blue masters/fonts/palette、3boundtitle/body/image、ordinaryshape/groupedtable+linechart/localconnector、logo资产；page1 hidden和hideinheritedtrue。Core test/layout-rebind.test.mjs8组最近全过(第一次chartcolor expectedhex缺#测试修为compareoldnormalized，不改生产)；涵盖key/role、不自动ordinary、歧义/roleconflict/missing/explicitdetach/override、preserveallstyles/background、adoptresetlongtextoverflow、inheritedconnectorrefblock/map、atomicmulti+IDrelease/stale/plaincapture/repeat/defaultimplicitlayoutblank no-op等。

当前还没有publicNode/ESM/global parity、Layout/Exporter/Renderer/native controls、完整build/fulltests、installedtypes、docs/ADR。不能标完成！新SourceData函数widen改动要测下游typebuild。需补测试stale-invaliddata/no-opauthoreddefaultlayout、timing/actions、allchartfamilies/shapes/groupedgecases及atomic失败registry allocator等。已发现并修：sameLayoutfastreturn要保留rawundefined defaultlayout而非显式blank；needsreview不可发虚假detachchanges；stalechecks必须早于normalize；commit不大spread。当前Coretestpass为前述8组，未跑fullcore其它208groups。

下一步完善adversarial单测、三端public parity+独立DOM/preview+sourcewriteguards(ImportSession.document typedbasePresentationDocument，JS实际factory可apply但types不拓权限)，build所有包后native具名freshadopt/preserve对照可选下轮完成；ADR0058尚未建，用logicalkey/uniquerole+explicitdecision声明index/crossdoc/hostfont/preserveinheritedgeometry未验边界。更新CoreAPI/typedexample/docchecks/ROADMAP/backlog/advancedTODO进度只留unchecked父项。原全goalactive，未commit/publish/reset/cleanup。所有旧proof/newsharedheading files保留；修改Core生产后旧producerdepHashes会stale，后续需current兼容verifier更新并保持native历史source不变，勿重跑原生旧rounds。


2026-10-27 — Rebind 原生字体回归定位/修复及有限证据收口（完整AUT-07.2仍未验收）：
1. 原因收敛：font-all/font-latin/font-aux均保持旧before基线，推翻仅visible-run字体具体化的猜测。后续从旧native input制作end-font（仅表格非空段末字体slots Arial）/default-font（仅pPr defRPr字体slots Arial）两控制，均实际打开并含隐藏页打印3页，完整used glyph trace与旧before相同。可复现实验说明目标theme继承到paragraph mark/default font会影响表格行度量。未用位置offset/字体容差修。
2. 生产修复packages/pptx-exporter/src/ooxml/elements.ts：textParagraphsXml第四个内部retainParagraphFont开关仅tableCellXml true；非空tableparagraph endParaRPr保留最后resolvedrun三个fontslots，script map经既有emptyRunStyle，绝不copy字号/颜色/action；全空paragraph沿用原有完整insertionstyle；ordinarytext/shape/notes不改。新增text-format-inheritance专测验证themefont、mixedlastfont、独立EA/CS、没有size/color/actions和非table旧输出。Exporter67tests0。注意该修复有意改变其他fresh table bytes，因此旧nativeproof依旧是历史记录，旧currentProducer bytebridge不能盲刷新/伪称仍exact，需按真实新输出建立兼容bridge；旧nativeJSON/files不改。
3. 新prepared .tmp/layout-rebind/run-QHSNXk；native .tmp/layout-rebind/native-OiIJHF。beforeSHA b1b03e761eb82ca0f0b1ac7e3ca94b757f60b63290b7715ff42b7511e5ea8638；after/nativeinputSHA cc14ef5dea496aa54a62edc172365e36f83006657dc763679f59a8166da393d0。原before046aa/after29e失败run/native-0OWvU8全部保留；原pendingmanifest已归档run-I6KLrT/pre-font-fix-manifest.json。新current docs/testing/layout-rebind-evidence.json绑定新目录，不要换回旧。
4. 新before/first、rebind-round1/2/3实际SaveAs→关闭→精确URL重开→打印完成，各3页，隐藏checkbox1+preview3/3。修复后before与旧before全3页完整PDFrecords（raster/fontstreams/trace/drawings/images）精确相同；修复后preserve table三处baseline完全恢复。三轮全PDFrecords精确一致；所有native package parts首次save后除ISOclock/TotalTime校验normalize外全相等。实际版本16.104(25121423)在旧nativebase/powerpoint-version AX/png，来自本轮捕获。
5. 新scripts/research/oracles/layout-rebind-native.py有限oracle：Unicode/GID pairs由实际ToUnicode bfrange+TrueType format6 cmap解码，face+allusedpairs选择resource，多候选必须usedglyph programs全部相等，拒unknown/compositeglyph而不抹gid假装等价。native cmap格式6尾部实际固定2zero经assert。preservepage0/未选page2完整glyphprogram/字宽/sidebearing/位置等精确相同；page0唯一drawings差异是owneraccent(CC2200→0044CC)其余keys全同，page2raster/drawings/images全同。真实4PDF负例coveredtext/swap/missinghidden/sameRasterfontAdvance均被拒，最后一组像素全3页完全一样但字体advance+1仍查出。旧失败本身用glyph comparison明确不等。docs/testing/layout-rebind-native-typography-evidence.json verifiedTypographyAndReopen true、completeNativeRebindAccepted FALSE，不准把它叫全native支持。首次save graph/reset语义/其余delta/ZIP负例仍未收口。
6. scripts/inspect-native-layout-rebind.mjs，5actualfiles(before/input/r1/r2/r3)，3slides/3masters/5layouts，Node/ESM/global全inspection/report一致、15exactsource-onlynoops；docs/testing/layout-rebind-native-parser-evidence.json绑定typographyproof。最近oracle加sourcebefore精确基线assert后已再跑0，随后Parser刷新handle33038仍待pollterminal，再运行scripts/verify-layout-rebind-evidence.mjs，勿留parentSHA stale。该verifier已过一次（最后assert前），regen15preparedPPTX/IR/Layout/origin/plans/warnings/SVG逐bytes相等，native/parser/artifact/depSHA逐项验，记录219currentproducerdependencies；status明确partial。
7. 新fix后pnpm build .tmp/layout-rebind-font-fix-build.log 88106终端0；typecheck0；lint .tmp/layout-rebind-font-fix-lint.log 37250终端0，最终docs后lint60838 .tmp/layout-rebind-font-fix-final-lint.log也0；完整pnpm test74027 .tmp/layout-rebind-font-fix-full-test.log终端0：Core218/Layout142/Parser696/Exporter67/Renderer20/CLI11/visual33/import128、installedNode/browser/strictNodeNext。最后diffcheck0。后续无production/CItest再改，仅nativeoracle/scripts/docs，所以不用无因再跑全部tests。
8. 当前原生CUA因round2单次30s超时kernelreset过，旧rebindNative/sharedNative等bindings已丢；实际round2saved/closed/reopened已完成，再续PDF，不要重跑覆盖。已重建u对象含app/base/state/observe/pick/capture（base native-OiIJHF）；当前UI rebind-round3文档，三轮PPTX/PDF都完成，不需要再操作。用object methods this.app避免repl函数闭包捕获被重新绑定的旧app，之前obs旧closure读printwindow而截图已savepanel的问题已查明纠正。若上下文压缩后先cua.rewriteDocumentation。
9. 继续审查准备：run-QHSNXk/first-save-unreviewed-differences.json保留逐XML diff（尚未验收，不在15preparedartifacts中）。source/native ownergraph初看每页layout/master/placeholderidx完全一致：p0/p1 Target blue title3/body2/pic1，p2 Source red title1/body2/pic3；page1 show0/showMasterSp0保存保持。三轮3张positive2pointlinecharts复用master-chart-first-save.py canonical逐比较9次均0、每轮3uniqueGUID；只是一次实际验证输出，尚未写专有graph/chartproof。其余常见firstsavechanges包括noGrp locks、indent -0→0、tablecol/rowIDs、anchor t默认省略、rels/media/workbookrename、theme lexical/defaults、notes/defaults/meta/presProps等，不能靠全XML忽略，需要逐项审图/负例。下一步补全这些bounded source/native graph+resetpage实际PDF语义/ZIP反例/首次save严格review，再nativeproof父级accepted，随后原AUT07.2 generalreset/index/crossdoc与全部backlog继续。
10. 对抗式检查5点：段末copy字号改变baseline（专测禁止sz且新before全PDF exact）；误copy超链接（专测禁止）；多同名字体资源错配（unicode/GID匹配+多候选glyphprogram相同）；仅看像素漏字体字宽（真实sameRaster fontadvance负例）；partialproof冒充fullAPI（JSON与docs explicit false；父TODO不勾）。旧Producer evidence因table字节改变的兼容未补，这是明确未验证项，不能覆盖旧source证明。全goal active，未commit/publish/reset/清理，现有大脏树保留。


本轮最终绑定收口：Parser刷新33038终端0（5inputs/15noops）；随后verify-layout-rebind-evidence终端0，.tmp/layout-rebind-artifact-verification.log：15preparedartifacts/219producerdependencies/2linkedproofs，completeNativeRebindAccepted false。以上覆盖前段所说待poll状态。后续继续first-save/resetpage/sourcegraph严格审查和旧producer兼容bridge，不重做已完成3rounds。


2026-10-27 — Rebind first-save/reset native gate + paragraph-font historical compatibility 收口（本轮progress，完整goal仍active）：
已核读当前AGENTS与前轮journal/code/proofs，未依据历史口述直接验收。未改生产/CItest；上轮fix后的完整build/typecheck/fullpnpmtest0仍适用，避免无因重跑全套。新工具/真实文件/negative evidence代码均执行过。

新增scripts/research/oracles/layout-rebind-first-save.py，直接对native-OiIJHF的input与r1/r2/r3每个packagepart严格比较：身份图校验3页/3masters/5layouts；p0/p1 Target/blue title3body2pic1，p2 Source/red title1body2pic3；隐藏page1和showMasterSp0；未复制masteranchor；notes双向backlink及Notes0/1/2；native text/imageframes手工期望EMU与targetbold/字号/超链接。真实owner-based重映射theme/media/workbook partnames与relationshipIDs；三workbook及image bytes原样。完整XML/binaryparts只准已见且逐值/结构校验的placeholder locks、table10000/20000row/colIDs、-0→0、anchor t省略、screen4x3、emptyguides、presProps三extensions、view128%/origin、metadata本机作者/locale统计字段和validclock/TotalTime、3positive2pointlinecharts有限canonical。全部原生parts三轮canonical SHA c661158232fdc6f9fce7bb3eb7724840eee721b55d46f28ac4af26d53f5ecbe9。native theme3/4不是颜色真的变了，是PP重排owner-bound主题part；精确按owner映射后全部同，不能按rawfilename误判主题损坏。

reset_pdf实际first/r1/r2/r3检查三页、全部strings与oldbefore完全同，CourierNewPSMT/CourierNewPS-BoldMT、targettitle/body/ordinary的精确observedPDFsize/origin，backgroundRGB F0F8FF、hidden masteranchor位置仍bg，ordinary/table3cells几何/opacity除red→bluefill外全同，image8x4digest原样/targetphysicalbbox精确、4palette像素探针与sourcePNG角色一致，localconnector真实新boundrect，body前景ink未遮盖。原生PDF字号是缩放打印观测29/21/19，XML另验28/20/18，不把打印值当作者字号。实际nativecontact .tmp/layout-rebind/run-QHSNXk/native-contact.png已生成并view，before/after6格三页都检查。

新first-save证据docs/testing/layout-rebind-first-save-evidence.json verified true +verifiedNamedResetPage true；29actualZIP corruptions（wronglayout/theme/index/hidden/inheritedflag/notesbacklink/danglingimage/unknownrelattr+rootattr/duplicaterid/titlefont/resetbodybold+action/paragraphindent/tablefont+rowID+border/lock/1EMUgeometry/chartcache+GUID/unknownpresExt/aspect/defaulttextsize/apphidden/unknownclocknamespace+duplicateclock/workbook/image）均被拒。另3actualPDF坏例unchangedreset/coveredpicture/coveredbody拒；连上上一proof4PDF=29ZIP+7PDF。对抗审查发现metadata必须用完整namespace+每field exactly1，已修oracle防unknownnamespace同localname被吞；没有改生产。修复Python测试helper body函数被for body变量遮蔽的问题后最终日志.tmp/layout-rebind-first-save.log0。所有失败中间negative dirs保留。

scripts/verify-layout-rebind-evidence.mjs新增3rd proof hash/dep/3roundhash/canonical+29/3negative检查，仍regen15preparedartifacts/219currentproducerdeps，最新 .tmp/layout-rebind-first-save-artifact-verification.log0。上一有限typographyproof的completeNativeRebindAccepted false原样保留，它不代表完整templateAPI，nativeApplication文本明确named3pagecontrol accepted而parentopen；Parser15noops证据仍正确绑原typographyJSON，不因新independentproof无关变动重跑。docs/testing/layout-rebind.md、APIexporter前轮doc、ROADMAP/backlog/advancedTODO进度已如实更新，不勾AUT07.2任何父项。

旧producer兼容：ordinarymasterart与PresentationFormat两verifier直接0（.tmp/layout-rebind-artwork-compatibility.log / format-compatibility.log），无table所以source未变。masteradvanced/shared fresh因paragraph-font修复字节变化，新增scripts/research/advanced-table-font-bridge.mjs纯变换断言：legacy/current ZIPentry顺序与其他所有payloadbytes全同，只master2/3内exactly2tables6nonemptyparas从emptyendmark添加3slots Arial，非table/emptycontinuation不动。scripts/prepare-master-paragraph-font-bridge.mjs仅生成新唯一dir .tmp/master-advanced/font-mark-bridge-TBvCok/before.pptx sha0ecf999b736656fe2c4ccaa28bad8f3ca93e97ac20a6681b311101bfc667824f / after.pptx sha8cd5b9e17e91418a0c0922956227652e6cc2dc6efe098b03c7802a3af4e35a73；旧before9643618a/fresheditb4fb8dd1、native3round files与原生JSON不改。

实际PowerPoint16.104(25121423)本轮分别打开bridgebefore/after精准URL，captureAX/png，普通slidesprint各4页，共8页；before.pdf全4页完整pageRecords==历史native-QRT0SF/first.pdf，after.pdf全4页完整pageRecords==历史shared-native-vwggt5tr/fresh-core-edit.pdf，包括rawglyphIDs/trace/所有embeddedfontstreambytes/raster/drawings/images全exact。没有把这两个firstprint叫新的reopenround。版本AX/png也捕获在bridgefolder。新scripts/research/oracles/master-paragraph-font-bridge.py写docs/testing/master-paragraph-font-native-evidence.json；另4actualZIP只精确单字符串改fontmark/runsize/showMasterSp/appname，不用XML重序列化制造无关前缀差异；字段negativeFixtures只声明生成，真正rejection由Nodeverifier读取bytes调用纯helper检查并记录。scripts/verify-master-paragraph-font-bridge.mjs已0 .tmp/master-paragraph-font-verification.log：2bytebridges/8exactnativepages/4actualZIPnegatives；旧nativeproofs+newnative+sourcefiles/depsSHA均检查。docs/testing/master-paragraph-font-bridge-evidence.json verified true。

scripts/verify-master-cell-properties-evidence.mjs 和verify-master-shared-evidence.mjs改为保存原historical .sourceSha/.freshBefore/.freshEdited/.dependencies，新增currentProducer带当前源hash/bridgeEvidence/当前deps，不重写历史source。既有IR/origins/Layout/warnings/SVG/Parser/chart/nativehashchecks保留，并验证newnativeproof hash/dep/artifacts以及纯bytebridge/当前生成等于实际newfile。前者 .tmp/layout-rebind-cell-compatibility.log0：104historicalartifacts/10currentdeps/原4oracledeps/12noops/currentbefore0ecf；后者 .tmp/layout-rebind-shared-compatibility.log0：90historicalartifacts/14currentdeps/原10oraclebindings/15noops/18chartcomparisons/currentbeforeafter实测eq。桥source0ecf/8cd5与原source9643/b4fb互不覆盖。新bridgeevidence没有时间戳更新字段，重验后deterministic相同不stale oldcurrentlinks。

本轮最终pnpm lint handle32117 .tmp/layout-rebind-first-save-final-lint.log终端待下一output确认（此前67970同log已0，最后advancedTODO文字更新后又跑）；gitdiffcheck0。生产/CItest从上轮fullgreen以后未改，故不要再无因跑fulltests。新oracles/Nodeverifiers已逐项过，非“未执行计划”。旧cell/shared/docs新增桥段落，普通artifact/format兼容已过；新layoutproof/bridge全hash有效，未commit/publish/reset/cleanup。

CUA当前bridgeNative对象有app/base/state/observe/pick/capture(base font-mark-bridge-TBvCok)，当前UI关于PowerPoint版本dialog仍打开，后续新UI操作要freshobserve并关闭该dialog；前轮u仍存在但旧app链接一度inactive已getApp重新激活到bridgeNative.app，别误用旧ubase。所有本轮PPTX/PDF都完成，重开三轮不需要重复，也勿再prepare生成新dir伪装旧文件opened。

下一步应继续实现未完成的真实能力，勿在已接受具名rebindfixture上无穷补同类反例：AUT07.2.1仍缺stable numeric/index identity（Core目前只有logicalkey/kind，nativeidx在Parserprovenance/Exporter自动arrayindex映射，必须先审查模型/架构决定公共contract，不能假装已实现index匹配或以遍历索引充身份）；AUT07.2.2完整propertyreset/fields/跨尺寸与resolvedgeometry仍开；跨文档AUT09.2.1、native master其他artwork/image/table/chart data/style编辑等原全部TODO继续。原goal所有范围保留，当前turn progress，未complete/blocked。


本轮最终terminal确认：lint32117已退出0（.tmp/layout-rebind-first-save-final-lint.log），覆盖上段待确认措辞；新first-save/resetproof、bridge、四个既有compatibility verifier均0，最终diffcheck0。全部目标仍active，下一步转缺失index/逐属性reset等实现。


2026-10-27 — Explicit placeholder indices implementation progressed; not delivered yet.
Production: optional index?:number on Core PlaceholderDefinitionInput/NormalizedPlaceholderDefinition, located uint32(0..4294967295) and layout-local uniqueness validation, normalization retains only authored values. Rebind priority explicit map→same key+kind→declared index+kind→unique kind; same-index role mismatch blocks. Match sourceIndex/targetIndex omitted unless declared; differing chosen explicit target index emits info. No positional/native-export ordinal becomes Core identity. Exporter new ooxml/placeholder-indices.ts reserves explicit indices, implicit entries prefer old ordinal or monotonic first-free positive; layout prompts and all local NV bindings share one context map. Unindexed legacy output remains exact.
Core6groups placeholder-index.test.mjs pass (initial test used nonexistent addShape, corrected to addElement); Exporter2groups pass including 200 dense explicit/implicit allocations. New tests/fixtures/placeholder-index.mjs: source body0/photo4294967295/title73/implicit subtitle4; target decoy17/implicit subtitle1/reservedbody2/photomax/newbody0/title73. Three pages/two master themes/ordinaryunbound, page1hidden. New third public case scripts/layout-rebind-parity.test.mjs compares adopt/preserve/reset Node/ESM/global all plans/reports/IR/origins/Layout/SVG/warnings/bytes/Parserprovenance and exact noops, plus independent DOM index maps/text/picture bindings. All3publicgroups handle83796 terminal0 .tmp/placeholder-index-public-parity.log. Installed node/types consumers added field/method tests.
Full pnpm build55695 .tmp/placeholder-index-build.log terminal0; typecheck0; lint84699 .tmp/placeholder-index-lint.log0. Fullpnpmtest60790 .tmp/placeholder-index-full-test.log lognow ends tests129/pass129/fail0, NEED poll60790 terminalbeforeclaimexit0. Expected totals Core224/Layout142/Parser696/Exporter69/Renderer20/CLI11/visual33/import129 plusinstalledconsumers; inspect actual summary counts. No production/CItest changes since run started.
Docs new ADR0059-explicit-placeholder-indices.md + testing/placeholder-index.md; updated layout-rebind API/canonicalIRv2/ADR0058 subsequent note. Nativepending wording remains until proof is built. Need architecture index/coreAPI/typedexample, TODOprogress update. Microsoft primarydoc https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.presentation.placeholdershape.index?view=openxml-3.0.1 confirms UInt32 native property. This adds authored numeric identity, not native template import or source write authority.
Old current-producer compatibility all passed without repeating old native work: verify-master-paragraph-font-bridge first (refresh currentElementXml hash), thencell/shared/artwork/format/layout-rebind in parallel. Logs .tmp/placeholder-index-{font-bridge,cell,shared,artwork,format,rebind}-compatibility.log all0. All old source PPTX/IR/Layout/SVG unchanged. Native JSON/files untouched. Layout verifier now220producerdeps due newexporthelper.
Prepared scripts/prepare-placeholder-index.mjs with unique dirs .tmp/placeholder-index/run-uKHDrz and native-5fWET2. beforeSHA3db127f3b6befde74f1e8fe9ffeca226335459b8c80fdaee96ab2a52205cce89; after/nativeinputSHA114824002c396627ad13557c60edc2d4020f246b367e825af3e2309e91370c35. Fifteen artifacts before/after PPTX/norm/Layout/SVG, origins/plans/warnings(0/0), docs/testing/placeholder-index-evidence.json pending. Sequentialp0preserve thenp1adopt+resetgeometry/text/background→target, p2sourceuntouched. Do not regenerate a new directory and relabel it as opened.
ACTUAL PowerPoint16.104(25121423) versionAX/png innewnativebase. before/input openedwithoutrepair; before.pdf/first.pdf3pages includingexplicit hidden checkbox1/preview3/3. index-round1/2/3 each SaveAs→savedAX/png→CmdWclosedAX→preciseURLreopenAX/png→ordinary3pagePDF COMPLETE! Third Open initiallydisabled raisedunique0 aftersave/closealreadydone; freshobserveenabledclickedexactfileandprinted. Do not rerun saves/overwrite files. Raw PPTX confirms sourcebodyidx0 becomes omitted on save (schema default0), image4294967295 stays exact, title73 stays, legacy1 on targetpages/4 on sourcepage stays. All before/first/r1/r2 text contains TITLEi/BODYi/LEGACYi/Ordinaryi; r3completed butindependentchecks still pending.
CUA indexNative object active with app/base/state/observe/pick/capture, base native-5fWET2; methods open(filename), finishPdf(name)(printdialogalreadyopen), saveClose(name)(overwriteguard,saves+closesonly). CurrentUI index-round3 document. fsimport persists. bridgeNative previousbasefont-mark-bridge-TBvCok and uoldbase do not confuse. All3roundscomplete, none needrepeat. Ifcontextsummary readcua.rewriteDocumentationbeforeUI.
NEXT required: create independent placeholder-index-native.py (not yet created), use existing layout-rebind-native.py read-only glyph_trace/page_records (3pages) and c.full helpers for exact postfirst nativepackages. Manual source/targetpromptmap/0default/max/implicitcollision/localtext+picture/nativeowners/notes/hidden/geometry checks; preservep0 usedglyphprogram+trace/images/drawingssameexceptmasterred→bluefill, unselectedp2raster/drawings/imagesexact with glyphsubset-canonical trace; actualresetpage targetframes/font/picture check. Three reopened PDFs exact full records, allnativeparts afterfirst exact exceptvalidclock/TotalTime. Generate focused real corruptZIP/PDFs including wrong0→decoy17/max→1/missingmax/duplicateprompt/implicitcollision/outofrange/wronglayout/hidden and covered/swapped/missingpage/fontmetric. No native acceptance until checked. New actual5file Node/ESM/global Parserinspector cancopy inspect-native-layout-rebind.mjs but use index-round1/2/3 and3slides/3masters/3layouts (NOT5), 15exactnoops. Add artifact verifier regeneration15preparedvalues/bytes with frozen source bindings; update docs/progress/finalchecks and pollfulltestterminal. Full AUT07.2 stillneeds general propertyreset/fields/crossdoc/sourceprojection, full goalactive. Noagents/commit/reset/publish/cleanup. This journal is authoritative current state; earlier nativepending statements are historical until new oracle.


Index follow-up terminal: full pnpmtest60790 exited0, actual Core224/Layout142/Parser696/Exporter69/Renderer20/CLI11/import129 and installed Node/browser/NodeNext consumers. Nativeindex-round3 exactreopen/print alreadycomplete. Latest docs correctly say captured3rounds but independent native oracle/negative/Parser/artifact proof NOT accepted yet; no proof created this turn. Architecture index added0058/0059, canonicalIR optionalfield, typedexample usesexplicit0, CorepresentationAPI/newtestingdoc updated. Lint37291 .tmp/placeholder-index-final-lint.log pending terminal after doc/typeexample changes, poll; previouslint0. No production/tests changed afterfullrun. Subsequent advancedTODO/CoreAPI prose updates add only existing relative links; run final lint/doccheck once before delivery. Next meaningful work is native-index proof (not rerunningUI or implementation). Goal active; no need to redo alreadypassedfullsuite.


Final index implementation turn checks: final lint52138 terminal0 and gitdiffcheck0. Fulltest60790 terminal0 alreadyrecorded. Independent indexed-native proof remains next required work, actual three round files unchanged. No final native acceptance claimed.


Follow-up implementation constraint for AUT-07.2.2 (investigation, not implemented): current cascade is defaults→placeholder preset/inline→element/shape-text preset→local frame/paragraph/run. Deleting a local font/color leaf alone does not resume target inheritance when its local named preset still supplies that leaf. Deleting the whole preset and materializing all remaining leaves would freeze unrelated live preset/theme behavior. Preserve existing textStylePreset:string public consumers; widening it to an object union would break consumers reading strings. A separate optional per-reference inheritance/omission selector is a candidate, not an accepted API yet. It must filter the local preset before normalization, preserve unrelated preset fields and symbolic references, clear selected local declarations, and report actual inherited origins. AutoFit/bullets currently select whole objects; scalar/side margin and script-map/null barriers need explicit tests. Cell presets/region rules have their own cascade; table geometry binding does not imply placeholder text defaults. Inspect materialization/storage paths before choosing the contract. Fields do not yet have a Core authoring model, so do not claim the existing geometry/text/background reset completes fonts/paint/field/general geometry requirements. AUT-08.1 needs declared field identity vs cached text, explicit clock/order/locale/zone, per-page shared-owner evaluation and copy/pagination policies; plain text must never be inferred as a field. Original full scope remains open.


2026-10-27 — Explicit index native acceptance and runtime admission follow-up completed. This goal turn is progress, full goal remains active.
New scripts/research/oracles/placeholder-index-native.py checks source/native manually expected Source/Target prompt maps and all local bindings including zero/default omission/max4294967295/implicit subtitle1 vs source4, unique object/prompt identities, owner master/theme/fonts, EMU frames, strings/images/notes backlinks/hidden flags and internal edges. Actual preservepage0 used glyph programs/metrics/positions+image records stay exact except masteraccentred→blue; page2 raster/drawings/images and used glyph trace exact. Resetpage uses Courier; picture moves to target with unchanged pixels/digest. All3reopened PDFs full records==first; every nativepackage part exact afterfirst save exceptvalidatedclock/TotalTime. VersionAX16.104(25121423), preciseopened/reopened URLs, allhiddenprintcheckbox1/preview3/3 allchecked. None of the existing native operations were repeated.
23actualZIP negatives and4actualPDF negatives allrejected, including wrong0→decoy17, picturemax→implicit, missingmax, duplicate/collidingprompt, overflow/negative/fraction, role/owner/theme/font, notes/hidden/geometry/imagebytes/objectID/embed/pageSize, missingprompt; PDFs coverimage/swap/missinghidden andfontadvance+1 withall3pagepixelsidentical butglyphmetricdifferent. Nativeproof docs/testing/placeholder-index-native-evidence.json verified/nativeIndexGateAccepted true; completeTemplateRebindAccepted false andfullFirstSaveEquivalence false. Allfirstsave added/removed/changedpartSHA differences recorded, no broad normalizer hides other semantics. 41native artifacts/6oracledependencybindings. Nativegraph sourcebefore3db127f3.../sourceafter114824002c... andall3actualindex-round files stayoriginalnative-5fWET2.
New scripts/inspect-native-placeholder-index.mjs inspects fiveactualfiles before/input/index-round1/2/3 withcontent projection inNode/ESM/global, full inspections/reports/normalizedCore data allsame and15exactsource-onlynoops. All12localphs/input matchuniqueparentbyindex, nativezero explicitIndexfalse whilefreshsourcestrue, maxpresent3times. Newdocs/testing/placeholder-index-native-parser-evidence.json binds468SourceTS/compiledJS/declarations/script/package/lockdependencies, sofutureCore/Parser changes require consumerreinspection. Corrected inherited copier scope text: slideorder/hiddenstate WEREalreadyprojected byexistingprofile, onlycomplete master/artwork model/inheritedartvisibility/source template-index writes remainopen. This preventsfalse absence claims.
New scripts/verify-placeholder-index-evidence.mjs regeneratesall15preparedartifacts(byte-exactPPTX/IR/Layout/origins/plans/warnings/SVG), checks2prooflinks/allnativeartifact+negative+depSHA, fiveconsumerfiles/snapshots15noops and220currentproducersourcebindings. Newnative-contact.png inrun-uKHDrz (6panels before/after) generatedandviewed; separatelyboundvisualArtifacts hashes withreadonlyrecheckonfutureverifierruns, notcountedasre-generatedsourceartifact. Latest .tmp/placeholder-index-artifact-verification.log0. VerificationScope explicitin docs/testing/placeholder-index.md/ADR0059. AUT07.2.1 matchingpolicy auditedagainstoriginalrow eachrequirement and marked[x] onlythatrowinadvanced-capability-todos; parentAUT07.2, .2/.3/fullreset/fields/sourceprojection/crossdoc NOTclosed. Roadmap/backlogprogressupdated. Nativeproof jsonfalseparentflagsremainliteral.
Adversarialreview foundrealexistingruntimebug:validateLayout didnotcheckkey/kind andunsupportedJS kindcouldmatchanotherinvalidroleandexportasbodyfallback. AddedCore PLACEHOLDER_KINDS validation andlocated invalid-placeholder-key/kind errors. A secondreproduction withtwoaliasedinvalidobjectkeys {toString:1,valueOf:2} showedduplicatesformatterthrowsTypeErrorbeforefieldvalidation; fixedkeyscollector tofiltervalidnonemptystrings beforeduplicateformatting. OneextraCoretestgroup coversmalformedprimitivekey/kind, unknown-role pair, atomicreplacementrollback, legitimate__proto__ stringkey, andtwo malformedobjectkeys returning2locatederrors+PresentationValidationError. Exportertestnowrequiresunsupportedkindreject. Legacydefinevalidationstilldeferredlikeexistingcontract; validdefinitionbehavior/oldbytesunchanged.
Allproduction/testchangesfinalbeforefullrun16311. First admissionbuild23982/typecheck/focus9tests/full90402 all0 butpre-objectkeyfix; retainhistorical .tmp/placeholder-index-admission-full-test.log (notfinal). Finalbuild96739 .tmp/placeholder-index-admission-final-build.log0; finaltypecheck0; finalfocused9tests0 .tmp/placeholder-index-admission-final-focused.log; COMPLETEpnpmtest16311 .tmp/placeholder-index-admission-final-full-test.log terminal0:Core225/Layout142/Parser696/Exporter69/Renderer20/CLI11/import129+installedNode/browser/strictNodeNext. No laterproduction/CItestchanges. Final lint34672 .tmp/placeholder-index-native-final-lint.log terminal0; diffcheck0.
AfterfinalRootfullteststopped(compiledfilesstable), reran actualParserinspector57013 terminal0 withcurrentruntime; thennewindexartifactverifier17373terminal0. Earlierinspector33121ranbeforelastobject-keyfix andwhilefirstrootbuildwaslive, soitsresultwas superseded, nottrustedforfinalruntimebindings. Refreshed oldcurrentProducercompatverifiers sequentialfontbridge thenparallelcell/shared/artwork/format/unindexed-layoutrebind; all0 in .tmp/placeholder-index-admission-{font-bridge,cell,shared,artwork,format,rebind}-compatibility.log. OldnativeJSONs/sourcefiles/roundsunchanged, onlypreparedcurrentProducer dependencies/linksrefreshedafterexactsource regeneration. Newcoreguarddoesnotchangethesewell-formedgoldens.
Next realwork: AUT07.2.2 completeper-propertyreset/inheritance, .3explicitreturn-to-template tests/dependencyclosure, AUT08.1declaredfields andfullremainingoriginalTODO. next-delivery-plan hascandidateanalysis: deletinglocalleafdoesnotresetpreset-suppliedleaf; deletingallpresetfreezesunrelatedlivebindings; preserveexistingtextStylePreset:stringAPI andconsiderseparateper-reference inheritance selectors after inspectingstorage/origin/cell paths. Fields absent fromCoreauthoringmodel, soexistingresetcannotclaimfullfieldbehavior. Do not redo acceptedindexfixtures/UI orkeepaddingequivalentindexnegatives. No agents/commit/publish/reset/cleanup; originalfullobjectiveunchanged.

## 2026-10-27 selective text reset progress

Implemented bounded AUT-07.2.2 selectors `text.fontFamily` and `text.color` for matched text/shape-text placeholders in adopt mode. `textStylePresetReset` bypasses only selected named-preset run properties so target placeholder/theme remains live and unselected preset properties are not materialized. Local declarations, including later explicit runs, continue to override. Whole-text reset clears the selector. Core covers multi-slide, untouched slide, shape text, actions, false/zero, origins, repeat, duplicate and sparse invalid inputs; Node/ESM/global parity compares authoring, normalized, origins, layout, SVG and PPTX bytes. No new native PowerPoint evidence; script-specific fonts, color opacity, other text/paint/geometry, fields and full AUT-07.2.2 remain open.

Final Core focus `.tmp/granular-reset-final-focused.log`, public parity `.tmp/granular-reset-parity.log`, final typecheck/lint `.tmp/granular-reset-final-{typecheck,lint}.log` all passed. Build `.tmp/granular-reset-build.log` passed before final one-line sparse-array validation change; the running full suite rebuilt after it. Prior full suite `.tmp/granular-reset-full-test.log` session 97913 is still running import tests; do not claim it completed until polled. Earlier package Core tests ran before the last validation change, so the later final Core focused run covers it. Source-bound indexed Parser inspection and placeholder evidence `.tmp/granular-reset-index-{parser,evidence}.log`, font bridge and five compatibility checks `.tmp/granular-reset-*-compatibility.log` all passed. Existing native files were not reopened or rewritten. Continue complete backlog; do not mark parent done.

Follow-up: full `pnpm test` session 97913 finished exit 0; `.tmp/granular-reset-full-test.log` includes 130/130 import parity checks, including the added selector parity test. Because the final sparse-array validator landed after package tests had begun, `pnpm --filter @pptkit/core test` was rerun afterward and passed (`.tmp/granular-reset-final-core-test.log`). Final typecheck and lint passed. The goal remains active: next extend independently reviewable per-property inheritance without conflating Western and script-specific fonts or opacity, and obtain real PowerPoint reopens before claiming native acceptance; fields and transformed geometry remain open.

## 2026-10-27 selective text reset expansion and native investigation

Progress, full goal remains active. AUT-07.2.2: matched text/shape-text may now reset each of the 12 run fields independently, plus 7 paragraph and 4 frame properties using `text.paragraph.*` and `text.frame.*`; `textStylePresetReset` bypasses the selected preset declarations without freezing other preset fields. Whole-map script fonts, bullet, margin and autofit are atomic; individual script keys and margin sides remain open. Tests cover every selector including null font barriers, origin paths, zero/false, actions, multiple runs, shape text, stale atomic batch, Node/ESM/global full output. API docs and decision 0060 updated. `pnpm build`, typecheck, final lint, focused Core, full `pnpm test` including 130/130 import parity all passed, logs `.tmp/style-reset-{full-build,typecheck,final-lint,focused,full-test,parity}.log`. Old indexed parser and placeholder evidence plus font bridge and five compatibility scripts passed, `.tmp/style-reset-index-{parser,evidence}.log` and `.tmp/style-reset-*-compatibility.log`.

Native controls `.tmp/style-reset-native/run-IRiqlb`: generated deterministic before/input via `scripts/prepare-style-reset-native.mjs`; six outputs byte-identical on fresh regeneration (`.tmp/style-reset-native-repro.log`). PowerPoint 16.104 actual SaveAs/close/exact-reopen round1/2/3, 3 pages with page2 hidden; final round PDF printed with hidden page (3 pages). Independent `scripts/research/oracles/style-reset-native.py` and `docs/testing/style-reset-native-evidence.json` validate actual source and native XML: selected page font/color/bold/paragraph align/wrap, preserved hidden/untouched controls and notes/layout relation; PDF shows CourierNew bold blue vs Arial red; seven actual corrupted ZIPs rejected. Only final round PDF, no assertion of three-round visual equivalence or all-property native validation.

Critical new finding: Core re-inheritance currently serializes normalized values as explicit local PPTX declarations. Omission-only actual probe (`inheritance-probe.pptx/.pdf`) rendered black Arial, not target, so cannot simply omit. A second actual probe adding target `a:lstStyle/a:lvl1pPr` to layout rendered blue Courier bold; a third copy changing only this layout default to green Georgia regular rendered selected page green Georgia while both control pages remain red Arial; selected slide XML byte-identical between probe packages. PDF geometry vertical position differs from materialized control, so only run style inheritance is demonstrated, not full geometry/wrap/paragraph equivalence or native interactive edit. Probes and hashes and source bindings in evidence JSON; old files preserved.

Next: implement and test native inherited run properties end-to-end rather than claiming Core inheritance suffices. Inspect `packages/pptx-exporter/src/generation/generate-pptx.ts`, `ooxml/elements.ts` placeholder and run writers, and Core `resolvePresentationInheritance` provenance contract; introduce explicit selective source metadata through owner boundaries, populate layout lvl1 defaults and omit only selected eligible local properties. Preserve no-reset byte compatibility or honestly update goldens/actual native proof. Verify themed references, script null barriers, unsupported host fallback, default/custom/layout-level origins, paragraph/frame separately, three real native reopens and presentation projections. The two probes are handmade ZIP changes, not a production implementation. Continue fields, geometry and all other backlog afterward.

## 2026-10-27 opt-in native placeholder run inheritance

Progress, original goal remains active. Core reset selectors now retain explicit `textStylePresetReset` intent even without a named preset. Exporter `nativePlaceholderText: "inherit"` captures authoring intent before async asset loading and permits native inheritance for selected Western font family/color/bold when each run has no subsequent local override, matches the target placeholder, has no script-map conflict, and has safe font-slot/opacity semantics. Target layout writes `a:lstStyle/a:lvl1pPr`; local run declarations are omitted only for eligible properties. Other properties and unsafe cases remain materialized with located `native-text-inheritance-materialized` warnings. Opt-in warns `native-placeholder-metrics-may-change` per affected shared layout. Default `"materialize"` retains historical PPTX output and reports materialization; unknown mode rejects. Source intent is a value captured between Core/Layout and exporter, types live under exporter types, no global registry or invalid package dependency. Shape warnings use `text.textStylePresetReset`. See `docs/architecture/decisions/0061-opt-in-native-placeholder-text.md`, API docs and `docs/testing/layout-rebind.md`.

Actual native evidence: `.tmp/style-reset-native/run-RfcLfr`, `docs/testing/native-placeholder-text-evidence.json`, verifier `scripts/research/oracles/native-placeholder-text.py`. PowerPoint 16.104 printed first and third-round PDFs with all 3 pages incl hidden and first/third raster pixels exactly equal, three SaveAs-close-exact-reopen PPTX retain intended missing run declarations and layout defaults; five independently corrupted ZIPs rejected. A package-level layout-only change to Georgia green regular changed selected page while unchanged slide XML and local controls remain Arial red. No claim of native UI interactive editing or other properties. Fresh current code reproducibly generates input bytes exactly; default materialized PPTX exactly matches historical `.tmp/style-reset-native/run-IRiqlb/input.pptx`. Warnings differ from earlier captured file because metric diagnostics were added after measurement, not native byte drift. Source hashes current.

Adversarial review: adding native list defaults shifts text in shapes sharing the target layout by ~0.96pt vs historical control; two actual mark-size/mark-bold probes failed to remedy. Thus profile is deliberately opt-in, warns, and this regression and broader run/paragraph/frame/geometry/field work remain open. Additional guards prevent null East Asian/complex barriers, preset barriers, explicit opacity even 1, opacity 0, script maps, later local overrides and mismatched values from silently losing semantics; focused tests cover all.

Checks: full `pnpm test` session 45108 exit 0, import 130/130 (`.tmp/native-inheritance-full-test.log`); it started before the final conservative guard/path/layering changes, so final Core package test 230/230, exporter package test (see `.tmp/native-inheritance-final-exporter-test.log`), final Node/ESM/global parity, typecheck/lint and exact byte reproduction were rerun after relevant final changes. Native oracle passed; source-bound old index evidence and font bridge/five compatibility verifiers passed with current sources (`.tmp/native-inheritance-*-compatibility.log`, `-index-evidence.log`, `-oracle.log`). Old style-reset-native proof is historical and not relabeled as current producer output. All existing dirty work kept.

Next investigate proper preservation of paragraph-mark and line metrics for every shape sharing a layout, with independent multi-page glyph/geometry comparisons and controls for title/body/empty/shape text and custom size/theme before considering default native inheritance. Expand supported run properties, paragraph/frame, individual script map/margin leaves and native interactive edit only with provenance-aware tests; address actual fields, transformed geometry and remaining backlog.

## 2026-10-27 minimal native placeholder defaults and metrics investigation (in progress)

Previous goal turn was progress; full goal remains active. Exporter now aggregates the actual requested eligible properties per placeholder and writes only these as `a:lstStyle/a:lvl1pPr/a:defRPr` (no unrequested font/size/paragraph defaults). Multiple pages with distinct requests union their target defaults but each run omits only its own selected local property. The shared-layout warning remains; font metrics not fixed. Focused `packages/pptx-exporter/test/native-placeholder-inheritance.test.mjs` passes including independent-property and cross-page union cases. Build/typecheck/lint passed before one test-only addition; production code final.

PowerPoint 16.104 actual production control `.tmp/native-color-bold/run-jejLfi`: first materialized and inherited, plus third reopened PDF all three pages (including hidden) identical in pixels, glyph traces, drawings and image records; three actual SaveAs/close/exact-reopens. Package-only layout edit changes only selected page color/bold, both controls unchanged. Five genuine altered ZIP negatives rejected by `scripts/research/oracles/native-minimal-color-bold.py`; evidence `docs/testing/native-minimal-color-bold-evidence.json`. Current producer reproduces PPTX and warnings byte-for-byte (`.tmp/native-color-bold-repro.log`). Separate actual font/color/bold minimal-output control `.tmp/style-reset-native/run-Igc1ns` also 3 save/reopens, first/final three-page PDFs match previous opt-in PDF pixel/glyph/draw/image; package-only layout change gives Georgia/green while controls stay red Arial. Five actual negatives; `scripts/research/oracles/native-minimal-font.py`, `docs/testing/native-minimal-font-evidence.json`. Historical default materialized PPTX exact bytes preserved. Shared-font metric shift remains.

Metrics probes `docs/testing/native-placeholder-metrics-probes.json` and `scripts/research/oracles/native-placeholder-metrics.py`: only layout font causes the named control's shift; color or bold alone preserves local control pixels. Explicit Arial end-paragraph font restored named original three-page raster, and Georgia/Times root themes also happened to match, but Courier root-theme counterexample did not. Do not promote root-font heuristic into production; marker theme semantics and broader cases unverified. Probe files are handwritten and actual PowerPoint PDFs, not production implementation.

First full `pnpm test` `.tmp/minimal-native-defaults-full-test.log` exited 1 due dev-app 5s UI timeout under concurrent native testing; standalone `pnpm --filter @pptkit/dev-app test` passed 16/16 `.tmp/minimal-native-defaults-dev-retry.log`. Full retry `.tmp/minimal-native-defaults-full-test-retry.log` live session 64734; it passed dev-app and is still running. Re-poll it; do not mark passed before terminal. Once browser builds settle, refresh old indexed proof/bridge/font shared/artwork/format/rebind verifiers; regenerate color/font proofs if any source changes. Afterwards run exporter package tests including the latest appended union test and browser selective parity, and update these journals with real terminal status. Continue entire backlog after resolving shared-font metrics, script/margin leaf resets, paragraph/frame fields and cross-document mapping.

Follow-up terminal checks: full retry session 64734 finished exit 0 (`.tmp/minimal-native-defaults-full-test-retry.log`), with 130/130 final import tests; prior dev-app timeout reproduced as passing 16/16 in isolation and in the full retry. Latest union addition was covered by separate exporter suite 79/79 (`.tmp/minimal-native-defaults-final-exporter-test.log`) and selective Node/ESM/global parity passed (`.tmp/minimal-native-defaults-final-parity.log`). Build, typecheck and final lint including docs boundaries passed. Old indexed Parser inspector and evidence, font bridge and five compatibility verifiers all passed after final builds (`.tmp/minimal-native-defaults-index-{parser,evidence}.log`, `.tmp/minimal-native-defaults-font-bridge.log`, `.tmp/minimal-native-defaults-*-compatibility.log`). Native minimal color/bold and font source/actual-asset oracles passed and both current producers regenerate their input bytes exactly. Next investigate font-only default's line-height and theme/host behavior; no general end-mark remedy is validated. Full backlog remains open.

## 2026-10-27 expanded native run inheritance (in progress)

Previous turn was progress; original goal remains active. `nativePlaceholderText: inherit` now also supports independent `fontSize`, `italic`, `underline`, `strike`, `language` in addition to Western font/color/bold, preserving later local overrides, script-map safety and shared-layout warnings. Minimal target defaults include only selected fields; no-reset serialization remains compatible. Focused tests cover each new property, per-run later overrides and explicit false/none (`.tmp/native-run-expanded-final-focused.log`); entire exporter 81/81 `.tmp/native-run-expanded-final-exporter-test.log`, Node/ESM/global parity `.tmp/native-run-expanded-parity.log`, full build/typecheck/final lint passed.

Production native control `.tmp/native-run-expanded/run-5mHxhe` generated by `scripts/prepare-native-run-expanded.mjs` includes selected page, hidden local-override page and untouched page. PowerPoint 16.104 actual SaveAs/close/exact-reopen round1/2/3; materialized, inherited first-open and final reopened 3-page PDFs match pixels, glyph traces, drawings and image records. A layout-only XML change changes selected size/italic and removes two red decoration rectangles; both controls remain pixel-identical; selected language/default values are checked in XML, not visually inferred. Seven actual corrupt ZIPs rejected. `scripts/research/oracles/native-run-expanded.py`, `docs/testing/native-run-expanded-evidence.json`; source and artifact hashes bound. Current code reproduces PPTX and warnings byte-for-byte (`.tmp/native-run-expanded-repro.log`); prior color/bold and font controls also reproduce and their oracles have current source bindings (`.tmp/native-run-expanded-{color,font}-repro.log`, `.tmp/native-run-expanded-old-{color,font}-oracle.log`). Shared font-family metrics shift, paragraph/frame, script/opacity remain open; no native UI layout editing claim.

Full `pnpm test` session 66215 live in `.tmp/native-run-expanded-full-test.log`, script-font stage; re-poll until terminal before claiming. Once final import builds settle, refresh indexed Parser inspector and index verifier, font bridge then five legacy compatibility verifiers; preserve all native evidence. Latest exporter tests were added after full suite start and separately pass. Finish by documenting actual status. Do not mark objective complete; resume broad backlog.

Follow-up terminal checks: `pnpm test` session 66215 finished exit 0 (`.tmp/native-run-expanded-full-test.log`), final import 130/130. Latest tests added after suite start separately passed in full exporter test 83/83 (`.tmp/native-run-expanded-final-exporter-test.log`); selected empty text paragraph-mark, hyperlink relationships and explicit false decoration controls all pass. Final lint/typecheck/build and Node/ESM/global parity passed. Indexed Parser inspector and index evidence, font bridge and all five old compatibility verifiers passed after final builds (`.tmp/native-run-expanded-index-{parser,evidence}.log`, `-font-bridge.log`, `-*-compatibility.log`). Three native producer fixtures reproduced their existing PPTX/warnings byte-for-byte and their current-source oracles passed. Preserve the named five-property native scope; actual interactive language behavior and general layout/font metrics are not proven. Keep full objective active; next expand remaining semantics or resolve known font metric shift and continue all backlog.

## 2026-10-27 native paragraph inheritance (in progress)

The original all-backlog goal remains active. Opt-in `nativePlaceholderText: inherit` now handles selective paragraph alignment, indent, hanging indent, line spacing, space before/after and bullet in addition to previously supported run fields. Later local overrides are preserved; list-generated indentation is conservatively materialized with warnings where local or unmasked preset bullets can affect it. Default materialization remains the compatibility path. API, architecture decision 0061 and testing documentation describe the scope.

Actual PowerPoint 16.104 control `.tmp/native-paragraph/run-XPcnYn` includes selected, hidden local-override and unselected pages. Three distinct SaveAs/close/reopen rounds retain the intended native behavior. Materialized, initial inherited and third-round PDFs match for all three pages in pixels, glyph traces, drawings and image records. A package-level layout-only edit changes the selected page while controls remain pixel-identical; this is not an interactive layout-edit test. Nine corrupted ZIP controls are rejected. Evidence: `docs/testing/native-paragraph-evidence.json`, oracle `scripts/research/oracles/native-paragraph.py`. This proves one seven-property combination, not each property independently across hosts and fonts. PowerPoint removes empty local paragraph properties on save.

Mixed paragraph/run control `.tmp/style-reset-native/run-hFxabJ` likewise underwent three native reopens; initial and final rendering matches the earlier font control. `docs/testing/native-mixed-paragraph-evidence.json` and its oracle replace the old minimal-font fixture as the current producer proof: the old fixture is historical because the generator now also resets alignment. The known shared-layout font-family metric shift remains unresolved; inheritance stays opt-in with warning. Current producer fixtures reproduce captured PPTX and warning bytes exactly. Four paragraph input/reopen files have Node/browser ESM/global Parser projection and 12 byte-preserving no-op checks; this does not prove Parser models native paragraph inheritance (`docs/testing/native-paragraph-parser-evidence.json`).

Current checks: exporter 89/89 `.tmp/native-paragraph-final-exporter-test.log`, build `.tmp/native-paragraph-full-build.log`, typecheck `.tmp/native-paragraph-typecheck.log`, parity `.tmp/native-paragraph-parity.log`; full `pnpm test` session 19514 and final lint session 51115 remain live and must be polled to terminal. After last browser builds settle refresh paragraph and indexed Parser evidence, font bridge and legacy compatibility verifiers. Continue unverified font metrics, individual native property effects, frame/fields/geometry and the remaining backlog. Preserve all existing work and actual native artifacts.

Follow-up: final lint passed (`.tmp/native-paragraph-final-lint.log`), and `git diff --check` passed after removing two trailing blank lines. After the final browser build, the paragraph Parser inspector, indexed Parser inspector and index evidence verifier, master paragraph font bridge and all five legacy compatibility verifiers completed exit 0 (`.tmp/native-paragraph-parser-final.log`, `-index-{parser,evidence}.log`, `-font-bridge.log`, `-*-compatibility.log`). Full suite session 19514 still runs its import stage; poll to terminal and record result.

Terminal result: full `pnpm test` session 19514 exited 0 (`.tmp/native-paragraph-full-test.log`), final import 130/130. The latest test-only additions also passed separately in the final exporter suite 89/89. Native and compatibility evidence was refreshed after the final browser build. This is progress, not completion of the original goal; continue the remaining backlog and unresolved native font metrics.

## 2026-10-27 margin-side rebind reset (in progress)

Previous turn was progress and the original all-backlog objective remains active. `text.frame.margin.top/right/bottom/left` now independently clear local and preset declarations while retaining other sides. Numeric shorthand expands to the three retained declarations; a later local value still overrides the reset. Whole-margin reset composes in either order. Preset scalar source attribution retains the actual `frame.margin` path rather than inventing side declarations. Applies to text and shape text; absent presets and repeated rebinds are covered. Unsupported selectors reject. Exporter materializes frame insets and warns, including with `nativePlaceholderText: inherit`; no PowerPoint live frame inheritance is claimed. API and testing docs updated; AUT-07.2.2 stays open.

Current evidence: build `.tmp/margin-side-build-all.log`, typecheck `.tmp/margin-side-typecheck.log`, final lint `.tmp/margin-side-final-lint.log`, Core suite 232/232 `.tmp/margin-side-core-all.log` (two later focused tests separately pass `.tmp/margin-side-core-test.log`), exporter suite 90/90 `.tmp/margin-side-exporter.log`, Node/ESM/global parity `.tmp/margin-side-parity.log`, `git diff --check` pass. Four previous native oracles passed against current source `.tmp/margin-side-native-{paragraph,mixed,run,color}.log`, but generated fixtures still require byte reproduction before claiming current producer continuity. Full `pnpm test` session 26285 is live `.tmp/margin-side-full-test.log`; poll to terminal. Once final import build completes, reproduce native inputs and refresh paragraph Parser/index evidence and font bridge/legacy verifiers. Preserve all existing artifacts. Continue unresolved native font metrics, script map leaf reset, frame native inheritance and entire backlog.

Terminal follow-up: full `pnpm test` session 26285 exited 0; final import 130/130 (`.tmp/margin-side-full-test.log`). Four native producers regenerated captured PPTX bytes exactly (`.tmp/margin-side-repro.log`); paragraph/run/color also reproduced warnings, mixed scope compared input and default PPTX. Paragraph Parser and indexed Parser inspectors, index verifier, font bridge and five compatibility verifiers all exited 0 after final browser build (`.tmp/margin-side-parser.log`, `-index-{parser,evidence}.log`, `-font-bridge.log`, `-*-compatibility.log`). Final lint repeated after decision update passed. This is Core progress only; live native frame inheritance, shared font metric shift and all other open backlog remain unverified.

## 2026-10-27 per-script font rebind reset (in progress)

Previous turn was progress; the complete backlog objective remains active. `text.scriptFontFamilies.<supported ISO 15924 key>` now independently resets a script in matched text and shape-text placeholders. Null whole-map local or preset clears expand to preserve every other script barrier; retained preset clear origins point to the actual whole-map source. A later local declaration wins; whole-map and key resets compose in both orders. Sparse maps, themed target values, target clear provenance, no accidental other-key changes and invalid keys are tested. Exporter still materializes script fonts with warning; no live PowerPoint script-map inheritance claim. API, ADR 0060, testing and advanced backlog updated; AUT-07.2.2 remains open.

Checks: all supported codes across text/shape and null/map barriers, Core suite 238/238 `.tmp/script-leaf-core-all.log`, exporter 91/91 `.tmp/script-leaf-exporter.log`, Node/ESM/global end-to-end parity `.tmp/script-leaf-parity.log`, build `.tmp/script-leaf-build-all.log`, typecheck `.tmp/script-leaf-typecheck.log`, lint `.tmp/script-leaf-lint.log`, diff-check all pass. Four existing actual native asset oracles now pass against current source `.tmp/script-leaf-native-{paragraph,mixed,run,color}.log`. Full `pnpm test` session 89164 is live `.tmp/script-leaf-full-test.log`; poll to terminal. After final browser builds finish reproduce native fixture bytes and refresh paragraph/index Parser evidence, font bridge and five compatibility verifiers. Existing work and actual native artifacts must be preserved. Remaining backlog, shared-layout font metric shift and native frame/script inheritance remain open.

Follow-up: four native producers reproduced captured PPTX byte-for-byte; paragraph/run/color warning files also match and mixed comparison covers input/default PPTX (`.tmp/script-leaf-repro.log`). After final browser build paragraph Parser, indexed Parser and index verifier, font bridge and all five compatibility verifiers passed (`.tmp/script-leaf-parser.log`, `-index-{parser,evidence}.log`, `-font-bridge.log`, `-*-compatibility.log`). Full test session 89164 remains live in import stage; poll to terminal before claiming success.

Terminal result: full `pnpm test` session 89164 exited 0; final import 130/130 (`.tmp/script-leaf-full-test.log`). Current producer byte reproduction and Parser/index/bridge/five compatibility verifiers passed after final browser builds. This delivers Core per-script reset only, not live native script-map inheritance or completion of AUT-07.2.2; continue the original full backlog.

## 2026-10-27 opt-in native frame inheritance (unfinished)

Previous turn was progress. This turn found a hand-patched six-attribute PowerPoint 16.104 frame probe `.tmp/native-frame-probe/run-tz0or6_i`: baseline and inheritance PDFs match three pages in pixels/text traces/drawings/images, layout-only attribute edits change selected page but not hidden local and untouched controls. Independent verifier `scripts/research/oracles/native-frame-probe.py`, evidence `docs/testing/native-frame-probe-evidence.json`; this tests a combination, not individual properties, native interactive editing or reopens.

Exporter now opt-in omits eligible wrap/verticalAlign/four insets on matched slide placeholders; target layout already contains bodyPr defaults. Later local values stay explicit, autofit and unsupported properties warn/materialize; frame-only does not add list defaults or issue the shared list-metrics warning. Changes are in `packages/pptx-exporter/src/generation/native-placeholder-inheritance.ts`, `types/native-placeholder-inheritance.ts`, `ooxml/elements.ts`; tests per selector and later local overrides passed `.tmp/native-frame-focused.log`, exporter suite `.tmp/native-frame-exporter.log`, Node/browser parity `.tmp/native-frame-parity.log`, build/typecheck/lint passed. Review remaining edge cases (whole vs side reset, shape text, preset/local barriers and interactions), improve conditional-map type, rerun focused and required checks after changes. API/ADR are updated but explicitly preliminary.

Production generator `scripts/prepare-native-frame.mjs` produced `.tmp/native-frame/run-waD1Ql`; actual PowerPoint first-open `inherit.pdf` and `materialize.pdf` three pages match pixels/text traces/drawings/images. Actual SaveAs/close/reopen round1 and round2 completed; round3 saved but NOT yet closed/reopened or printed. PowerPoint current foreground window is `round3` path in that directory; close it and explicitly reopen round3 then print PDF including hidden slide. Examine all saved XML, construct actual production layout-only edit and print, make independent source/hash-bound oracle with negative controls. Verify selected properties individually and autofit separately before making broad native claims. Current mixed font+paragraph proof is historical: its producer resets frame.wrap and output now changes; do not rebind old oracle to changed producer. Create current mixed production three-round evidence or preserve and label old proof correctly.

Full `pnpm test` session 77183 is live in `.tmp/native-frame-full-test.log`; poll to terminal, do not infer from timeout. After last builds and any source fixes refresh current-source native evidence only when its fixtures reproduce, paragraph Parser/index verifier, bridge and legacy compatibility; rerun final lint/typecheck and adversarial review. No overall goal complete; all backlog remains active.

## 2026-10-27 native frame production evidence follow-up

Previous turn was progress. Production `.tmp/native-frame/run-waD1Ql` completed three actual PowerPoint 16.104 SaveAs/close/exact-reopen rounds, final print including hidden slide. Materialized, initial inherited and round3 PDFs match all three pages in raster pixels, glyph traces, drawings and image records. A package-only layout bodyPr edit changes selected page while hidden local override and unselected pages remain pixel-identical. `scripts/research/oracles/native-frame.py` independently checks source/actual asset hashes, all four frame XML states, materialized baseline, unchanged non-layout package parts, PDFs and rejects eight mutated packages. `docs/testing/native-frame-evidence.json` records the precise scope. The manual probe is separate evidence; no individual attribute effects, native interactive layout editing, autofit, or other hosts are proven.

Adversarial additions cover whole-margin reset with partial later local override on shape text, mixed paragraph/frame and autofit warning; internal value-map type simplified. Final build `.tmp/native-frame-final-build.log`, typecheck `.tmp/native-frame-final-typecheck.log`, lint `.tmp/native-frame-final-lint.log`, exporter 94/94 `.tmp/native-frame-final-exporter.log`, Node/ESM/global parity `.tmp/native-frame-final-parity.log`, diff-check all passed. Full `pnpm test` session 77183 exited 0 `.tmp/native-frame-full-test.log`, final import 130/130; it started before final type-only simplification and test additions, covered by final separate checks. Four frame/paragraph/run/color producer inputs and warnings reproduce byte-for-byte `.tmp/native-frame-repro.log`; native oracles refreshed against current source. Frame actual-input Parser projection and twelve source-only no-op saves match Node/ESM/global (`docs/testing/native-frame-parser-evidence.json`, `.tmp/native-frame-parser-current.log`); this does not establish native frame inheritance projection. Paragraph Parser/index verifiers, font bridge and five legacy compatibility verifiers passed.

Old mixed run/paragraph proof `native-mixed-paragraph-evidence.json` is historical because its producer resets frame.wrap, now emitted as native inheritance; current mixed profile needs new actual native validation. Prior known shared-layout font default metric shift remains unresolved. Continue independent frame-property visual controls (especially wrap and right/bottom insets, with clipping, long text and zero values), native autofit, other hosts, mixed font/frame and all remaining backlog. The all-backlog goal is not complete.

## 2026-10-27 individual native frame controls

Previous turn was progress. Current production generator `scripts/prepare-native-frame-individual.mjs` created `.tmp/native-frame-individual/run-FyToIw` containing six separately reset frame properties (wrap, vertical alignment and four sides) each with a local-override control, twelve pages including a hidden page. Actual PowerPoint 16.104 materialized and inherited first-open PDFs and a third-round PDF after three distinct SaveAs/close/exact-reopen rounds match every page in pixels, full text traces, drawings and image records. All three reopened packages retain each selected omission and other local declarations. A package-level edit changes exactly one property on each of six target layouts and no other parts; all six selected pages visibly change while all control pages remain pixel-identical. Wrap uses clipped long text, right/bottom use right/bottom alignment and initial margins are zero. In one hidden wrap control raw PDF font subset glyph IDs differ when another slide changes, but Unicode/coordinates and all other text-trace fields match; pixels remain exact. Six independently corrupted ZIPs rejected. Source-bound oracle `scripts/research/oracles/native-frame-individual.py`, evidence `docs/testing/native-frame-individual-evidence.json`, reproducible edit producer and `.tmp/native-frame-individual-repro.log` (generated PPTX, warnings and case inputs byte-identical). Actual UI layout editing, other host applications and additional values remain open.

Four generated/reopened 12-page packages also passed Parser inspection/content projection parity in Node/ESM/global and twelve exact source-only no-op saves (`scripts/inspect-native-frame-individual.mjs`, `docs/testing/native-frame-individual-parser-evidence.json`, `.tmp/native-frame-individual-parser.log`); no native inheritance import or source-edit capability claimed. Final lint `.tmp/native-frame-individual-final-lint.log` passed. No production code changed this turn; previous full build/typecheck and full `pnpm test` remain the last checks for those unchanged files. Verify lint/diff check terminal and continue unfinished mixed font/frame native proof, autofit, unresolved font metrics and all backlog.

## 2026-10-27 native autofit production progress

Opt-in native placeholder inheritance now omits the whole eligible `frame.autoFit` child for none, shrink and resize; later local declarations, including values equal to the target, stay explicit; default export still materializes and warns. Text/shape tests, local override and no-marker cases, mixed paragraph/frame checks passed `.tmp/native-autofit-focused-final.log`; build/typecheck/lint passed `.tmp/native-autofit-implementation-build.log`, `.tmp/native-autofit-typecheck.log`, `.tmp/native-autofit-lint.log`. Current full `pnpm test` session 62787 is running in `.tmp/native-autofit-full-test.log`: poll to completion. `scripts/layout-rebind-parity.test.mjs` now includes autoFit and must be run after final browser builds.

Manual PowerPoint 16.104 probe `.tmp/native-autofit-probe/run-Fxqa9h` covers none/shrink/resize, three save/close/reopen rounds, package-only layout edits and controls. In actual text edits on saved copies inherited resize and explicit resize grew to 3436838 EMU while a copy with the parent switched to none remained at 889000 EMU. Those interactive copies were saved but not separately reopened; static resize layout edit alone has no visible effect. Historical report `docs/testing/native-autofit-probe-evidence.json` binds pre-implementation source hashes: do not relabel as current. Its Parser report needs regeneration after browser build because the manual report changed after the prior Parser run.

Production `scripts/prepare-native-autofit-production.mjs` generated `.tmp/native-autofit-production/run-btf6eN`; actual PowerPoint three save/close/reopen rounds completed; final PDF including hidden slide equals the byte-identical materialized probe baseline PDF in six-page pixels, full text traces, drawings and image information. All original/reopened slide and layout fit children checked; initial production PPTX is semantically equivalent part-by-part to the manual probe. Generator baseline/inherit/warnings reproduce byte-for-byte. Independent oracle `scripts/research/oracles/native-autofit-production.py`, `docs/testing/native-autofit-production-evidence.json`, `.tmp/native-autofit-production-oracle-final.log` verify all this and reject three corrupted selected-slide local fit children. Dynamic edit occurred on equivalent manual package, not the production copy; native interactive layout edits and other hosts remain unverified. Earlier source-bound native reports now have old source hashes: refresh only if their producer outputs reproduce, never relabel historical mixed evidence.

Still open: full test terminal status, updated Node/ESM/global parity, production Parser/no-op evidence, refreshed reproducible prior oracles and Parser evidence, final lint and adversarial review. Mixed font/frame native proof, shared-layout font metric shift and the entire remaining backlog remain open; this is progress, not completion of the user's goal.

Follow-up: updated Node/ESM/global layout-rebind parity including autoFit passed `.tmp/native-autofit-parity.log`; final lint `.tmp/native-autofit-final-lint.log` and `git diff --check` passed. Full test session 62787 still running at last observation.

Autofit adversarial review: (1) later local override equal to target and a matching value without reset both remain explicit, covered by `.tmp/native-autofit-focused-final.log`; (2) shrink is atomic including scale and all three modes have local override tests and independent PowerPoint controls; (3) hidden/control pages and round-trip geometry/content are checked, including three corrupted selected-slide packages rejected; (4) shared list defaults are not introduced by frame-only inheritance; mixed paragraph/frame behavior is tested; (5) manual dynamic and package-only parent edits are distinguished from production three-round evidence and from unverified interactive layout edits/other hosts. Production Parser and historical probe Parser were refreshed after browser builds, `.tmp/native-autofit-production-parser.log` and `.tmp/native-autofit-probe-parser-final.log`. Five older frame/individual/paragraph/run/color generators reproduced all outputs byte-for-byte and their native oracles refreshed against current source (`.tmp/autofit-refresh-native-*.log`); frame, individual and paragraph Parser evidence was refreshed too.

Follow-up: full `pnpm test` session 62787 exited 0; final import group 130/130, `.tmp/native-autofit-full-test.log`. Current mixed run/paragraph/frame production sample `.tmp/style-reset-native/run-uFixey` passed three actual PowerPoint 16.104 SaveAs/close/exact-reopen rounds. Initial and third-round PDFs match all three pages in pixels/text traces/drawings/images. Compared to the actual materialized PDF, selected page and same-layout local override page each show a body-text vertical shift about -0.96 pt; unrelated page matches pixel/text. This confirms the previously known shared-layout metric issue remains, not a rendering equivalence claim. `scripts/research/oracles/native-mixed-frame.py` checks XML across rounds, exact PDF comparisons, per-page shifts and rejects five corrupted properties; source-bound `docs/testing/native-mixed-frame-evidence.json`; generator outputs reproduce byte-for-byte `.tmp/style-reset-native/run-Lhv5NE`. No layout-edit or other-host evidence for this current mixed sample yet. Continue investigation/fix of the metric shift, current mixed live layout edits, and the full backlog.

Further shared-metrics investigation: manual `mark-effective.pptx` on the current mixed sample changes only selected/control paragraph-mark font slots to their actual Courier New and Arial fonts. Actual PowerPoint initial print shows local control returns to exact materialized pixels/text, but selected remains ~0.96 pt high. A second actual print with selected paragraph mark bold=1 is identical to first probe and still fails selected; both are non-production, initial-open-only controls. Reproducible producer `scripts/research/prepare-mixed-effective-mark.py`, independent source/file-hashed oracle `scripts/research/oracles/native-mixed-effective-mark.py`, evidence `docs/testing/native-mixed-effective-mark-evidence.json`. Do not implement the local-mark patch as a universal fix: it would not fix selected and could freeze insertion styles/alter later inheritance. Further isolate line layout and investigate preserving future native font changes.

Current mixed fixture now also has a package-only target-layout edit: only `slideLayout3.xml` differs, changing placeholder wrap, alignment, font, bold and color. Actual PowerPoint initial print changes selected text from Courier bold/blue to Georgia/green; local override keeps Arial/red but its body moves down ~0.96 pt, unrelated page remains pixel-identical. Reproducible producer `scripts/research/prepare-mixed-layout-edit.py` matches every ZIP part; `scripts/research/oracles/native-mixed-frame.py` checks all mutations, slide preservation, rendered outputs and records the shared metric effect, `.tmp/native-mixed-frame-oracle-final.log`. This is package editing plus native rendering, not interactive PowerPoint layout editing; shared metrics remain unresolved. Full test, focused test, build, typecheck, lint, parity and previous revalidated evidence remain passed; full backlog remains active.

Final post-evidence lint `.tmp/native-mixed-evidence-final-lint.log` and `git diff --check` passed. Continue with the unresolved selected-page font metric effect and all remaining backlog; no overall completion claim.

## 2026-10-27 paragraph mark size counterexamples

Previous goal turn was progress: production autofit plus actual three-round evidence, mixed frame/run proof, shared-layout metric counterexamples. This turn corrected stale ADR 0061 statements that autofit is unsupported. Current files were checked; historical/full backlog remains open.

Generated two isolated research packages in `.tmp/style-reset-native/run-uFixey`: `mark-size.pptx` sets only both selected/control paragraph-mark `sz=2400`; `mark-effective-size.pptx` adds the same size to previously tested effective mark fonts. PowerPoint 16.104 initial-open printed both including hidden page. Size-only PDF matches initial input for all pages in pixels, text traces, drawings and images. Font+size PDF matches font-only probe exactly; selected page remains -0.96 pt from materialized, control and unrelated pages match. Independent source/file-hashed verifier `scripts/research/oracles/native-mixed-mark-size.py`, report `docs/testing/native-mixed-mark-size-evidence.json`, `.tmp/native-mixed-mark-size-oracle.log` verifies package-only differences and rejects two corrupted size values. `scripts/research/prepare-mixed-mark-size.py` reproduced both PPTX package parts using independent fixture `run-Lhv5NE`. This is a bounded negative result, not a production fix or native save/reopen validation. Continue researching line metrics and live inheritance without locking local insertion style; user goal and all unchecked backlog remain active.

Additional native counterexample: `scripts/research/prepare-mixed-no-mark.py` removes the end-paragraph property element only from selected and same-layout control body paragraphs in an independent fixture. PowerPoint 16.104 initial print `no-mark.pdf` matches the effective-font-mark PDF across all three pages in pixels, text traces, drawings and image records; selected remains -0.96 pt vs materialized, local control and unrelated page match. `scripts/research/oracles/native-mixed-no-mark.py`, `docs/testing/native-mixed-no-mark-evidence.json` independently validate XML-only removals, PDF scope and reject two restored-mark corruptions; producer reproduces all ZIP parts on `run-Lhv5NE`. This is not a universal correction: selected remains shifted, insertion styling, multi-run behavior, interactive editing and native reopen were not tested. No production change. Continue other first-principles controls and the full backlog.

Additional control: inserting an empty `a:defRPr` under each selected/control paragraph property adds no local font and leaves PowerPoint 16.104 initial rendering identical to the unmodified inherited input across all three pages, including both ~-0.96 pt shifts vs materialized. `scripts/research/prepare-mixed-empty-paragraph-default.py` reproduces all package parts; independent oracle `scripts/research/oracles/native-mixed-empty-paragraph-default.py` and `docs/testing/native-mixed-empty-paragraph-default-evidence.json` check only two paragraph changes, both full PDF records and reject two mutated defaults. It does not establish dynamic layout edits or three reopen rounds and is not a production fix.

Post-research review: results are limited to the named static PowerPoint 16.104 fixture; manual probes do not prove production behavior, insertion-style safety, multiple fonts, interactive edits or native reopen. Both package-difference checks and corrupted-element controls passed. `pnpm lint` passed `.tmp/native-mixed-marks-research-final-lint.log`; `git diff --check` passed. No production code changed this turn and the previously passed full test remains the last production verification. Continue searching for a general font-metric solution and all remaining backlog; no completion claim.

## 2026-10-27 versioned operation evidence catalog (partial)

Previous turn was progress: bounded PowerPoint paragraph mark counterexamples and corrected stale ADR. This turn advanced IMP-10.1.1.1/.1.2 using actual retained evidence without narrowing the overall objective. `docs/testing/operation-evidence.json` has 12 exact object/owner/family/variant × operation × producer/version × application/version × profile/fixture records, independently separates capability from status, and explicitly marks one PowerPoint playback combination untested. Candidate V1/V2/V3 text preservation, Node create/replace, LibreOfficeDev Fade/Push-left/right/Wipe-left resaves, and timing inventory are scoped to their cited evidence. Original file-publication command was not recorded and is not invented. Producer/version `unrecorded` is not a wildcard.

`scripts/operation-evidence.mjs` validates exact schema/enums, uniqueness, portable local evidence paths, linked JSON SHA-256, scalar JSON-pointer assertions and recorded fixture digest; renders deterministic `docs/testing/operation-evidence.md`. `scripts/operation-evidence-cli.mjs` checks generated drift in `pnpm lint`; opt-in `--verify-local-fixtures` actually rehashes seven retained `.tmp` PPTX files and fails missing/changed/outside-root paths, `.tmp/operation-evidence-final-local.log`. Fixture bytes remain machine-local, not distributed, and default checks do not claim to reread them. `scripts/operation-evidence.test.mjs` tests duplicate, invalid enums/schema, false/missing assertions, changed/missing evidence, bad hashes, path/symlink escapes, fixture corruption, missing files and deterministic ordering; final focused `.tmp/operation-evidence-final-test.log` passed. `docs/testing/operation-evidence-guide.md` documents limits. `pnpm build`, typecheck, final lint `.tmp/operation-evidence-final2-lint.log` passed; `pnpm test` session 82312 is still live in `.tmp/operation-evidence-full-test.log` (started before the last matrix-only hardening; final focused test covers those changes). Poll to terminal and rerun lint/diff after all documentation. Do not mark IMP-10.1.1 complete: broader cases, portable fixtures and actual independent regression attribution/native/mixed gates remain open. Continue unresolved shared font metrics and the full backlog.

Adversarial review of operation catalog: (1) known producer rewriting is separated from PPTKit candidate preservation; Wipe row is specifically left, and Push left/right assertions are both checked; (2) unspecified versions are `unrecorded`, not wildcards, and the named PowerPoint playback remains untested; (3) manifest SHA, selected outcome assertions and recorded fixture hashes are checked, while the explicit local check rehashed seven actual retained fixtures; portable CI still cannot claim their availability; (4) duplicate owner/family/variant/operation/application/profile/fixture combinations, invalid status/method combinations and stale generated tables fail; (5) missing, corrupted and symlink-escaped inputs are rejected. Final focused tests also exercise actual changed evidence bytes, JSON-pointer escapes and Markdown delimiters. Node compatibility avoids `import.meta.dirname` because the repository permits Node 20.0. `pnpm test` session 82312 remains live; its first catalog checks predate final hardening, covered separately by `.tmp/operation-evidence-adversarial-test.log`.

Final operation catalog result: full `pnpm test` session 82312 exited 0, final import group 130/130 (`.tmp/operation-evidence-full-test.log`). The suite started before the final catalog schema and validation hardening, so current catalog code has separately passed final focused `.tmp/operation-evidence-shipping-test.log` (including physical evidence mutation, escaped JSON pointers, markdown delimiters, fixture corruption/missing/escaping and deterministic order). `pnpm lint` `.tmp/operation-evidence-shipping-lint.log`, `git diff --check` and actual seven-source SHA recheck `.tmp/operation-evidence-shipping-fixtures.log` passed afterward. Build/typecheck passed earlier in this turn; no package runtime source changed. Catalog source provenance review caught timing input producer unrecorded; both timing rows now state `unrecorded` rather than assuming a producer. The existing 12 records are bounded observations, not universal support; local sample paths remain undistributed and broader IMP-10.1.1/native/mixed gates stay unchecked. Overall goal remains active.

## 2026-10-27 V1 PowerPoint interactive and post-native source save

Previous turn was progress: operation evidence catalog and seven retained fixture hashes, full test passed. This turn copied V1 source/edited files only after matching `docs/testing/text-edit-save-evidence.json` hashes into `.tmp/text-edit-native/run-fstis57m`. Actual PowerPoint 16.104 opened the PPTKit candidate, user-interface text append ` POWERPOINT UI 789` was saved, followed by three distinct Save As/close/exact-reopen rounds. Independent ZIP/XML verifies selected/control text, shape IDs and boxes, notes text and named opaque extension through all native outputs; the original orphan disappears after a PowerPoint resave. Captured interactive/final PDFs have equal pixels/drawings/images and same glyph identities/styles, but raw text trace segmentation differs and maximal glyph-origin difference is 0.05859375 pt; no full-trace equivalence claimed. Control Unicode glyph positions/style match across captures. PowerPoint's PDF uses ArialMT although source asks for Liberation Sans.

`scripts/prepare-native-v1-post-edit.mjs` then reimported native `round3.pptx` under `text-runs-v1`, changed the first existing run and preserved native-added runs; report is scoped-preservation. Independent ZIP/XML confirms only one text atom in one slide payload changes and every other part remains byte-identical. Re-executing producer on a copied native input reproduces PPTX and report byte-for-byte (`post-repro-2131ylu1`). PowerPoint opened that candidate and completed three more distinct Save As/close/exact-reopen rounds; initial/final PDF pixels, complete text traces, drawings and images all equal. It does not restore the orphan already removed by the application. `scripts/research/oracles/native-v1-interactive.py`, `docs/testing/native-v1-interactive-evidence.json`, `.tmp/native-v1-interactive-oracle.log` verify scoped claims and reject six mutated in-memory cases; manual PowerPoint interactions are observations, not simulated by the oracle. `scripts/inspect-native-v1-interactive.mjs` validates ten source/candidate/native files across Node/ESM/global with thirty byte-exact no-op saves, `docs/testing/native-v1-interactive-parser-evidence.json`, `.tmp/native-v1-interactive-parser-final.log`.

Operation catalog now has 15 scoped records, including three separately attributed PowerPoint UI and PPTKit post-native observations; explicit local validation rehashed ten fixture files `.tmp/native-v1-interactive-fixtures.log`. Documentation `docs/testing/native-v1-interactive.md`, `docs/api/text-edit-save.md` and IMP-09.2.1.1 note the bounded result; full reference coverage, other fixtures/versions, layout/fidelity, arbitrary opaque dependencies and interactive notes remain open. `pnpm build`, typecheck, lint, matrix tests, focused text-save tests 19/19, local fixture check and `git diff --check` passed (`.tmp/native-v1-interactive-*.log`). No package production source changed this turn; prior complete `pnpm test` remains latest full run. Adversarial review: application vs PPTKit loss attributed separately; raw PDF glyph differences recorded; control and notes checked rather than inferred; first and post-native edit transformations independently narrowed; six corruption cases rejected. Original full-backlog goal remains active.

## 2026-10-27 V2 PowerPoint shape and notes interaction

Previous goal turn was progress: V1 actual PowerPoint interactive edit, two three-round native sequences, post-native PPTKit source save and explicit losses. This turn copied V2 source/candidate only after matching `docs/testing/text-owner-save-evidence.json` hashes into `.tmp/text-owner-native/run-0t7fofk9`. In PowerPoint 16.104 UI appended ` UI SHAPE` to first slide shape text and ` UI NOTES` to first notes paragraph, saved, then performed three distinct Save As/close/exact-reopen rounds. The original V2 candidate was separately printed as notes before UI edits: PowerPoint printed three pages before and after the edits (first note paragraph, control paragraph, second slide notes); cause is not established. Native slide PDFs before/after three rounds match raster pages, but first-page raw text/draw records differ and max glyph-origin drift ~0.0062 pt; notes first-page pixels differ and max glyph-origin drift ~0.1234 pt; remaining note and slide pages' complete records match. XML retains decomposed e + combining mark; PowerPoint PDF text extraction yields composed é. Do not claim whole-notes layout fidelity.

Actual PowerPoint resave removed orphan, renamed shared image/workbook paths but retained their bytes, chart cache values, selected shape geometry, two named opaque extensions, internal slide and external notes link targets on the correct original runs. Independent raw relationship checks prove both slide↔notes pairings and note placeholders remain body idx=1. Links were not activated. `scripts/prepare-native-v2-post-edit.mjs` imports native round3 via `text-runs-v2`, changes existing first shape and notes runs, and preserves UI-added runs. Source save changes only target slide and notes XML, each solely at selected character atom; all other parts byte-identical. Candidate/report reproducible byte-for-byte in `post-repro-5i7uhqm5`. PowerPoint opened that candidate then performed three further distinct Save As/close/exact-reopen rounds; candidate-first vs final slide PDFs (2 pages) and notes PDFs (3 pages) match raster, complete text traces, drawing and image records. The already removed orphan was not restored.

`scripts/research/oracles/native-v2-interactive.py`, `.tmp/native-v2-interactive-oracle-final.log`, `docs/testing/native-v2-interactive-evidence.json` bind inputs/results/sources and reject 12 mutated cases including wrong notes owner, link owner/target, slide geometry, chart/workbook/image and opaque/orphan loss. These checks inspect captured files; they do not simulate or prove that the UI actions occurred. `scripts/inspect-native-v2-interactive.mjs`, `.tmp/native-v2-interactive-parser-final.log`, `docs/testing/native-v2-interactive-parser-evidence.json` verify 10 source/candidate/native packages across Node/ESM/global and 30 byte-identical no-op saves. `docs/testing/native-v2-interactive.md`, API docs and IMP-09.2.1.2 document limited evidence; V2 full reference/other owners/host versions, interactive link activation, notes-page layout and parent gates remain open. Operation matrix has 21 scoped records; local recheck verified 13 fixture files `.tmp/native-v2-interactive-fixtures-final.log`. `pnpm build`, typecheck, lint and 27 focused owner/matrix tests passed (`.tmp/native-v2-interactive-*.log`); no package runtime source changed. Previous full `pnpm test` remains last full suite. Adversarial review distinguished application loss from PPTKit, preserved Unicode/source vs PDF normalization, checked raw page pairing and links rather than inferred visuals, retained failed first/final comparisons, and attacked object/resource/geometry/owner references. Full goal remains active.

## 2026-10-27 V3 table native interaction and interleaved resaves

Previous goal turn was progress, not backlog completion. After the Mac was unlocked, PowerPoint 16.104 opened the hash-verified V3 table candidate in `.tmp/table-save-native/run-fm88km7y`; actual UI edits appended ` UIA` to the merged anchor and ` UIB` to physical row 2 column 3. Three distinct Save As/close/exact-reopen rounds retained those cells, the control text, grid, notes pairing, chart cache and image/workbook bytes. PowerPoint renamed resource paths, removed the orphan and the injected unknown cell attribute, and added lexical span attributes to merge continuations. Initial vs final PDF pages have equal pixels, drawings and image records; first-page text traces differ, with maximum glyph-origin drift 0.0198059082 pt, so complete trace equality is not claimed. Links were not activated or independently checked here.

`scripts/prepare-native-v3-post-edit.mjs` reimported the native round 3 and changed two existing table text runs; independent XML/ZIP inspection establishes that only those two text atoms in one slide part changed and all other payloads remained byte-identical. Candidate and report reproduced byte-for-byte on a copied source. PowerPoint opened that candidate and completed another three distinct Save As/close/exact-reopen rounds; initial/final PDFs have equal pixels, full text traces, drawings and image records. `scripts/research/oracles/native-v3-interactive.py` and `docs/testing/native-v3-interactive-evidence.json` check captured artifacts and reject 13 mutated cases; they do not independently witness the UI actions. `scripts/inspect-native-v3-interactive.mjs` checks 10 files in Node, browser ESM and browser global, with 30 exact no-op saves. See `docs/testing/native-v3-interactive.md` for scope and limits.

A separate LibreOfficeDev 26.8.0.0.alpha0 headless experiment starts from a verified prior native output and alternates three PPTKit table edits with native resaves, reopening each native result. Each PPTKit save changes only the two selected text atoms, but LibreOffice shrinks the unedited shared image by 360 EMU in each dimension per native resave; page 1 raster comparisons also differ. The earlier native source had already lost a workbook, orphan and opaque extension, as checked against the prior PPTKit candidate; these are not attributed to the subsequent PPTKit saves. `scripts/research/oracles/table-interleaved-native.py`, `docs/testing/table-interleaved-native-evidence.json` and `docs/testing/table-interleaved-native.md` document the bounded observations and 11 rejected mutations. A separate run reproduced the first PPTKit candidate and all reports byte-for-byte and all candidate/native slide trees semantically; native ZIP identity is not claimed. Seven packages across Node/ESM/global yielded 21 exact no-op saves.

The operation catalog now contains 26 scoped records; local verification rehashed 17 retained fixture files. Final build, typecheck, lint, 25 focused tests, local evidence verification and `git diff --check` passed; final parser reruns exited successfully. Build reported the existing example bundle size warning. No package runtime source changed in this turn. The last full suite was the preceding goal turn; its initial catalog checks predated later catalog hardening, which was separately tested. Adversarial review distinguished native application rewrites from PPTKit changes, checked both physical table coordinates and controls, retained negative PDF comparisons, checked note pairing and resource content, and exercised corrupted inputs. Other fixtures and versions, link activation, structural operations, broader visual fidelity and IMP-09.2.5 remain unverified; the full goal remains active.

## 2026-10-27 chart PowerPoint observation

Previous goal turn was progress: V3 table interaction and interleaved native evidence, not full completion. The prior chart-save round-3 candidate was copied after matching its manifest SHA into `.tmp/chart-value-save-native/run-2afntf6a`; PowerPoint 16.104 opened it. The selected chart’s Edit Data in Excel command was disabled; no interactive data edit or native save/reopen occurred. Excel was not found at the conventional `/Applications/Microsoft Excel.app` path, but the reason for the disabled command remains unproven. Attempting to launch the bundled LibreOffice headless app graphically timed out; no process remained. These do not establish application-wide support or a persistent blocker.

The opened first chart visibly clips the selected bar at the fixed axis top. Independent raw chart XML shows the original value 5 and candidate value 37.25, while both retain a fixed 0–9 axis. The previous independent oracle separately checked chart cache and workbook; this turn’s [observation](../testing/chart-value-powerpoint-observation.json) binds source/candidate/screenshot hashes and records manual UI observations, not automated visual assertions. `scripts/research/verify-chart-powerpoint-observation.py` rechecks those hashes and XML facts. API/testing docs now explicitly warn that fixed axis bounds remain unchanged and can clip edited values. This is a documented limitation of the current numeric-only contract, not proof of a numerical save failure or universal native behavior. PowerPoint Edit Data, LibreOffice interactive editing, other chart families and three native-produced editing rounds remain unverified; IMP-09.2.6.4 and full goal stay open.

Verification: retained observation hashes and XML check passed; `pnpm test:chart-save-evidence` passed 6/6, `pnpm lint` and `git diff --check` passed. Adversarial review separated fixed-axis preservation from numeric correctness, manual screenshot observation from machine-verified XML, and missing conventional Excel path from an established cause. No package runtime source changed.

## 2026-10-27 SVG line series visibility

Previous goal turn was progress: named PowerPoint chart clipping and disabled data-edit observation, with independent retained-file checks; no native interactive edit was claimed. Code review found that the SVG renderer suppressed markers when a resolved line series had `line: false`, and used a bar-like legend square when a line series had `marker: false`. The renderer now treats series line and marker independently for both plot and legend. Public authoring can request `marker: false`; the `line: false` case was exercised through the existing resolved-layout contract because the ordinary authoring input normalizes lines to stroke records. No new authoring capability or native chart fidelity is claimed.

`packages/svg-renderer/test/smoke.test.mjs` checks all four line/marker combinations, two data points, legend glyphs and labels, including neither visible; the SVG package suite passed 21/21 and typecheck passed. Adversarial review checked that absent line/marker does not suppress its counterpart, that bar/pie swatches retain their original branch, and that unsupported native/data-edit gates stay open. Full goal remains active.

Final checks: repository `pnpm build`, `pnpm typecheck`, `pnpm lint`, SVG package tests 21/21 and `git diff --check` passed. Full-suite results from earlier turns were not rerun for this focused renderer change.

## 2026-10-27 SVG definition identity collisions

Previous turn was progress: line-series/legend visibility fixed and tested; no broader native chart parity claimed. Review found chart clip IDs collapsed distinct element IDs such as `chart/a` and `chart?a` to the same string, and identical clip IDs could also collide when slides were embedded together. SVG rendering now allocates deterministic, per-slide definition IDs using a shared traversal counter for charts, images and connector markers, including inherited artwork. The slide prefix includes render position and encoded slide identity. These identifiers are internal and can change with traversal order. Independently rendered copies with the same slide identity and position may still collide when placed in one DOM; documentation explicitly requires isolation or namespacing for that case.

Chromium integration exercises two simultaneously inline slides with colliding sanitized chart names, different plot widths, an image clip and two arrow markers per slide. It checks unique definitions, every reference resolves inside its own SVG, distinct plot clip geometry, and deterministic repeated serialization. The original test fixture reused object IDs across slides and was rejected by Core; it was corrected to preserve its global identity rule before the passing run. Full Chromium SVG/measurement suite passed 34/34, SVG package tests 21/21, repository build, typecheck, lint and `git diff --check` passed. Adversarial review covered shared allocator state through inherited contexts, separate image/marker/chart definitions, escaping of untrusted IDs via encoding, unchanged visual snapshots, and the remaining independently rendered-copy limitation. Full backlog remains open.

## 2026-10-27 SVG caller-controlled namespaces

Previous goal turn was progress: deterministic internal definition allocation and multi-slide Chromium checks, with independently rendered copies still potentially colliding. Both public SVG entry points now accept optional `idNamespace`; nonempty strings are injectively encoded into a prefix for root, title and all definition IDs. Different namespaces isolate independently rendered copies in one DOM, while the absent option preserves prior default bytes. Invalid empty/non-string values reject with `TypeError`; caller-selected namespaces must be distinct, and collision with arbitrary surrounding-page IDs is not ruled out. No random state, hidden global allocation or changed Core identities were introduced.

SVG package tests cover deterministic output, default compatibility, both public APIs, special-character namespaces, distinct identifiers and invalid values. Chromium checks two independently rendered copies in one document, with local clip and accessible-title resolution. Installed package Node and TypeScript consumer checks cover the new public option. Adversarial review considered SVG injection, namespace encoding collisions, default compatibility, aria-labelledby and clip references, repeated namespace reuse and untrusted surrounding DOM. The overall backlog remains open.

Verification: SVG package tests 22/22, Chromium SVG/measurement tests 35/35, installed package tests 5/5, repository build, typecheck, lint and `git diff --check` passed. Full monorepo suite was not rerun in this turn.

## 2026-10-27 bounded GIF structure inspection

Previous goal turn was progress: explicit SVG namespaces and installed/browser validation. AUT-04.3.2 remains open: current exporter preserves GIF bytes and renderer hands URLs to the browser without a playback or poster policy. The exporter now exposes pure `inspectGif` from default/browser/Node entries, with byte/frame/block budgets and pre-abort support. It reports logical screen, image frame rectangles, delays, disposal, transparency, interlace and recognized loop counts. It does not decompress pixels, guarantee playback, fetch data, automatically inspect exports or issue animation-loss warnings. Unknown extensions are skipped; plain text consumes graphic control without becoming an image frame. Synchronous work cannot be interrupted by a later event-loop callback.

Tests cover a Pillow 12.1.1 independently encoded/decoded 2x2 two-frame GIF with 0/70 ms delays and repeat count 2, as well as handcrafted local-table/plain-text/transparent/interlaced controls, every truncated prefix, malformed blocks, duplicates, geometry, detached input/output, offset views, limits and pre-abort. Node and browser global results match; installed Node and strict TypeScript consumers exercise public exports. Exporter package tests passed 101/101; installed package tests 5/5 before the final invalid-null budget hardening, which was subsequently covered by focused exporter tests. Typecheck, lint and browser parity passed. Adversarial review considered decompression bombs (no decompression; bounded bytes/blocks), malformed/truncated inputs, loops and zero delays, control scope, accidental native playback claims, and synchronous cancellation limits. Native playback, poster conversion, other formats and automatic diagnostics remain open; full goal remains active.

## 2026-10-27 opt-in GIF export diagnostics

Previous goal turn was progress: bounded GIF structural inspection with independently encoded fixture and public runtime parity. Export now accepts explicit `gifDiagnostics: "warn"`; for each loaded GIF asset, a detected multi-frame stream reports application-dependent playback and timing, while malformed/unsupported/budget-exhausted inspection reports unknown. Both cases retain original bytes. The default is off and produces byte-identical output and prior warnings; shared references produce one asset-level warning. All registered assets are considered, including unused; loaded resolver/path results are inspected without extra fetch. Cancellation remains an abort, not an inspection warning. No static poster conversion, actual loss claim or native playback assurance is implied.

Exporter tests cover animated/static/truncated GIFs, original bytes, Node/default parity, shared uses, resolver inputs, invalid modes and 10,001-frame inspection limit. Chromium global and Node generated bytes/warnings match. Documentation describes default limits, opt-in behavior and remaining native/static-adapter gaps. Adversarial review checked preservation on failed inspection, abort propagation, duplicate/shared warnings, unknown-versus-static classification and default backward compatibility. AUT-04.3.2 and the full goal remain open.

Verification: exporter package tests 104/104, installed consumers 5/5, Chromium global/Node parity, repository build, typecheck, lint and `git diff --check` passed. Full monorepo suite was not rerun.

## 2026-10-27 resolver-reported animated first-frame fallback

Previous goal turn was progress: opt-in animated GIF export diagnostics without changing media bytes. A resolver may now return `animationFallback: "first-frame"` with PNG/JPEG bytes; a successfully packaged asset emits one `asset-animation-flattened` warning even when GIF diagnostics are off. The resolver attests to the transformation; the exporter neither decodes the source nor proves the output is its first frame. Missing reports cannot be inferred. Invalid declarations and incompatible GIF output fail as asset reads and do not falsely report successful fallback. Existing source/output MIME consistency remains in force; a declared GIF MIME cannot be resolved to PNG without changing that declaration.

Tests use a Pillow 12.1.1-decoded PNG of the first frame of the independent two-frame GIF, compare original/output package bytes and shared-use warning counts, check default/no-report compatibility, resolver and Node parity, invalid declarations, output/MIME mismatch and abort during deferred resolution. Chromium global/Node output and diagnostics agree; installed consumer types exercise the new field. Adversarial review considered false source attribution, APNG/full-content validation (not claimed), animation losses other than first frame (not inferred), asset omission, shared assets and cancellation. Native playback, automatic conversion and broader media policies remain unverified; AUT-04.3.2 and the full goal remain open.

Verification: exporter tests 107/107, installed consumer checks 5/5, Chromium/Node parity, repository build, typecheck, lint and `git diff --check` passed. Full monorepo suite was not rerun.

## 2026-10-27 GIF uninspected graphics and extensions

Previous turn was progress: resolver-attested first-frame loss reporting and package/browser checks. Adversarial review found a one-image GIF can also contain plain-text graphics or unknown extensions; the previous opt-in diagnostic could return no warning although their visual/timing effects were not inspected. `inspectGif` now reports optional `plainTextBlocks` and `unknownExtensionBlocks` counts; comments are recognized and not counted. Plain-text graphic control is still consumed, not projected as an image frame. Opt-in export emits `gif-inspection-incomplete` for these inputs, in addition to the existing playback warning when multiple image frames are present. Unrecognized application and extension payloads remain preserved; malformed multi-subblock recognized loop extensions reject inspection. No GIF pixel decoding or native playback assertion was added.

Tests cover one image plus plain text, unknown application, unknown extension, ordinary comment, coexistence with two image frames, byte preservation, control scoping, malformed loop data and Chromium/Node observation parity. The exporter package passed 109/109, typecheck and lint passed. Adversarial review considered unknown visual effects, false static classification, duplicate warning purposes, untrusted block budgets, and compatibility of absent optional count fields. Full animation behavior, poster conversion/native playback and the original full backlog remain open.

Final verification: exporter tests 109/109, installed consumer tests 5/5, Chromium/Node inspection parity, repository build, typecheck, lint and `git diff --check` passed. The full monorepo suite was not rerun.

## 2026-10-27 GIF PowerPoint and Chromium playback observation

Previous goal turn was progress: the inspector now reports uninspected GIF graphics/extensions. This turn generated a Pillow 12.1.1 red/blue, two-frame, 1,500 ms-per-frame, infinite-loop GIF and a PPTKit one-slide deck with two image objects sharing that asset. On macOS 26.5 (25F71), PowerPoint 16.104 opened the candidate and a slideshow was manually started. Native screenshot batches captured a red→blue change and then, in an uninterrupted batch, red→blue→red; both sample points agreed in every captured image. Headless Chromium 149.0.7827.55 displayed the generated SVG, with network blocked, and captured red→blue→red at both sample points. These samples demonstrate visible changes and a return to red, not precise frame duration, continuous synchronization, indefinite looping or native save/reopen. Manual UI provenance is an observation, not established by the oracle.

`scripts/research/oracles/gif-native-playback.py` independently checks Pillow frame pixels/delays/disposal, the two PPTX image relationships and exact shared GIF bytes, preview data URLs, screenshots/timestamps and five corrupted controls. `docs/testing/gif-powerpoint-playback-evidence.json` hashes the inputs, 228 screenshots and source scripts; `docs/testing/gif-native-playback.md` records methods/limits. A second run of the source and producer reproduced GIF, PPTX, SVG and producer report byte-for-byte. GIF inspector focused tests, oracle, lint, script syntax and `git diff --check` passed; no package runtime source changed. Adversarial review separated manual UI provenance from file checks, ignored the gap between native batches, avoided timing/indefinite-loop claims, checked shared resources and rejected static/replaced/corrupt controls. Other GIF constructs, finite/zero-duration timing, native edit/save/reopen, fallback and applications remain open; full goal remains active.

## 2026-10-27 GIF native three-round reopen

Previous goal turn was progress: PowerPoint and Chromium observed red-blue-red playback of a named generated GIF sample. After matching candidate and GIF hashes to that evidence, PowerPoint 16.104 performed three distinct Save As/close/exact-path-reopen rounds. On the third reopened presentation, 90 screenshots sampled blue→red→blue→red→blue, with both sampled shared-image locations matching. Playback of the first two reopened files was not separately sampled; no exact timing, indefinite-loop or general native compatibility claim is made.

Independent `scripts/research/oracles/gif-native-reopen.py` checks the source plus three native files, original GIF bytes, both picture identities/descriptions/boxes, title and shared internal media reference. PowerPoint renamed the media part and consolidated two image relationships into one; it retained the same GIF bytes and selected geometry. Four corrupted package/screenshot cases reject. `docs/testing/gif-powerpoint-reopen-evidence.json` records package and screenshot hashes, while the oracle cannot independently witness the UI actions. `scripts/inspect-gif-native-reopen.mjs` imports all four via Node, browser ESM and browser global, checks one shared GIF asset/two image uses and twelve byte-exact no-op saves, with source/build hashes in `docs/testing/gif-powerpoint-reopen-parser-evidence.json`. The parser manifest's linked native evidence digest was rechecked after final oracle rerun.

GIF inspector focused tests, lint and `git diff --check` passed; this turn changed evidence scripts/docs only, not package runtime source. Adversarial review distinguished source versus native resave, app URI renaming from byte loss, retained transitions from exact timing, manual UI observations from file checks, and import/no-op from source editing. Other GIFs/versions, zero/finite loop behavior, editing, poster conversion and full AUT-04.3.2 remain open; the original goal remains active.

## 2026-10-27 bounded JPEG EXIF orientation inspection

Previous goal turn was progress: three PowerPoint GIF save/reopen rounds and final playback observation, with independent source and parser checks. AUT-04.3.3 remains open. The exporter now exposes `inspectJpegExifOrientation` through default/Node/browser entries: it reads EXIF APP1 TIFF 0th-IFD orientation 1–8 before the first scan, handles little/big endian and reports absent EXIF/tag separately. It rejects malformed EXIF identifiers, offsets, duplicates, invalid tag type/count/value and segment lengths. Byte/segment/IFD budgets and pre-abort are explicit; it does not decode, rotate, read ICC, inspect later scans or claim native/preview orientation agreement. See `docs/api/byte-assets.md` and the referenced CIPA Exif standard.

Tests cover all eight orientations in both endiannesses, a Pillow 12.1.1-generated asymmetric 3×2 JPEG with independently read orientation 6, source immutability/offset views, missing tag, non-EXIF APP1, pre-scan boundary, every truncated prefix, corrupt identifiers/offsets/duplicates, limits and cancellation. The initial corruption test modified a byte already zero and was corrected before successful rerun. Chromium global/Node parity and installed TypeScript consumer validation exercise public distribution. Adversarial review distinguished early-scan absence from full-file absence, excluded thumbnail/unknown tag and color claims, rejected ambiguous duplicate sources, and kept original export untouched. Native display/transform behavior, broader metadata and full AUT-04.3.3 remain open; the full objective stays active.

Final verification: exporter tests 114/114, installed consumers 5/5, Chromium global/Node parity, repository build, typecheck, lint and `git diff --check` passed after malformed-identifier hardening. Full monorepo suite was not rerun.

## 2026-10-27 opt-in JPEG orientation diagnostics

Previous goal turn was progress: explicit bounded JPEG EXIF orientation inventory, not native orientation parity. Exporter now accepts `jpegOrientationDiagnostics: "warn"`; each loaded JPEG with pre-scan EXIF orientation 2–8 emits `jpeg-exif-orientation-unverified`, and malformed/unsupported/budget-exceeded inspection emits `jpeg-exif-inspection-incomplete`. Both preserve loaded bytes, and shared uses produce one asset-level warning. Default off preserves existing output and warnings; registered unused assets are included. The absence of a warning does not prove pixel validity or rule out metadata after the first scan; no automatic rotation, ICC processing or native/preview agreement is claimed.

Tests use the independently generated asymmetric orientation-6 JPEG plus all tag values 1–8, malformed identifier, resolver input, shared asset, Node/default byte equality, invalid option and asynchronous cancellation. Chromium global/Node output and warning parity and installed TypeScript consumer coverage check public boundaries. Adversarial review separated preservation from visual orientation, unknown/malformed from default orientation, cancellation from warnings, application-owned resolution from hidden I/O, and pre-scan scope from complete image validation. AUT-04.3.3 and the full objective remain open.

Verification: exporter package tests 117/117, installed consumer checks 5/5, Chromium/Node generated output and warning parity, repository build, typecheck, lint and `git diff --check` passed. Full monorepo suite was not rerun.

## 2026-10-27 EXIF orientation 6 native/preview mismatch

Previous goal turn was progress: opt-in JPEG orientation warnings without modifying media. Generated an asymmetric 300×200 JPEG with EXIF orientation 6 and a 200×300 PNG control derived from Pillow 12.1.1's decoded EXIF-transposed JPEG. PPTKit placed both in equal stretched boxes, preserved exact source bytes and emitted its orientation warning. On macOS 26.5 (25F71), PowerPoint 16.104 opened the PPTX and slideshow displayed the JPEG's unrotated 3×2 pixel grid, while the normalized PNG showed the rotated 2×3 grid. Chromium 149.0.7827.55 preview showed both as the rotated grid with network blocked. This is an actual mismatch for this sample and mode, not a universal Office result; no automatic normalization was added.

`scripts/research/oracles/jpeg-orientation-native.py` independently verifies Pillow source metadata/transpose, exact PNG pixels, PPTX relationships/payloads/boxes, SVG payloads, sampled native/browser colors and three corrupted controls. `docs/testing/jpeg-orientation-native-evidence.json` binds fixture, screenshot and producer hashes; manual PowerPoint UI provenance is an observation, not independently witnessed by the oracle. Source JPEG, PNG, PPTX, SVG and producer report reproduced byte-for-byte in a second directory. Focused JPEG tests, oracle, lint and `git diff --check` passed; no runtime source changed. Adversarial review distinguished actual raw-vs-normalized display from byte equality, sampled colors from full visual fidelity, screenshot evidence from timing/UI proof, and stretch mode from crop/contain behavior. Other orientation values, ICC, native save/reopen, application versions and automatic normalization remain unverified. Full AUT-04.3.3 and overall objective remain active.

## 2026-10-27 caller-attested JPEG orientation normalization

Previous goal turn was progress: named EXIF 6 JPEG showed raw PowerPoint pixels versus EXIF-oriented Chromium preview. Export resolvers now may declare `orientationNormalization: "pixels"` and supply PNG/JPEG output, including PNG for a source declared JPEG. The exporter emits `asset-orientation-normalized`; it does not decode or verify conversion, or alter authored frame, crop or dimensions. Output MIME/signature must agree; an explicitly non-JPEG source MIME, GIF output, combined animation-fallback declaration, invalid flag or normalized JPEG still carrying a recognized non-default pre-scan EXIF value fails asset loading. Bounded JPEG inspection and asset size budgets apply. Default resolver behavior and MIME mismatch rules remain unchanged without the explicit declaration.

For the preceding asymmetric orientation-6 stretched sample, a Pillow-transposed PNG was returned by the resolver. PowerPoint 16.104 opened the new candidate and displayed both image regions in the same rotated 2×3 arrangement. `scripts/research/oracles/jpeg-orientation-normalized.py` checks source EXIF, pixel-exact Pillow transpose, both PNG payloads, frame boxes, sampled screenshot cells and three negative controls; `docs/testing/jpeg-orientation-normalized-evidence.json` binds it to prior evidence and records hashes. The native action is a manual observation. Candidate and producer report reproduced byte-for-byte in a second directory. Exporter tests 121/121, installed consumers 5/5, Chromium/Node output parity, build, typecheck, lint and diff checks passed. Adversarial review checked MIME conversion authority, untouched crop/dimensions, contradictory JPEG EXIF, false resolver attestation, cancellation and distinction between sampled and full native fidelity. Automatic decoding, EXIF orientations other than the named native sample, crop/contain, color/metadata policy and full backlog remain open.


## 2026-10-27 explicit browser JPEG orientation normalization

Previous turn added caller-attested normalized resources, without decoding. The optional `@pptkit/pptx-exporter/dom` adapter now snapshots JPEG bytes, inspects bounded pre-scan EXIF and frame dimensions, decodes with the browser's orientation option and returns a PNG plus oriented dimensions and resolver attestation. Input, declared pixels and output have configurable limits; cancellation stops waiting and closes late bitmaps. The browser decoder may allocate before dimensions can be rechecked, PNG is encoded before its output limit check, and cancellation cannot forcibly stop underlying browser work. No network access, automatic crop/frame rewrite, ICC fidelity or preserved metadata is promised.

Chromium 149 tested all eight orientations against independently permuted raw RGBA pixels, including mirrored cases, limits, snapshot and cancellation. The named orientation-6 browser PNG exactly matches Pillow's independently transposed RGB output; PowerPoint 16.104 displayed its PPTX and the control with the same sampled 2×3 color grid in slideshow. `docs/testing/jpeg-browser-normalized-evidence.json` and its oracle bind inputs, package relationships, geometry, screenshot and three negative controls; the native UI action is manual, and screenshot sampling is not complete visual proof. PNG, report and PPTX reproduced byte-for-byte in a second directory.

Exporter tests 122/122, installed consumer tests 5/5 including browser entry, Chromium orientation test, repository build, typecheck and lint passed. Build emitted an existing large-chunk advisory. Adversarial review examined orientation/geometry, deceptive or malformed image sizes and decoder allocations, late cancellation cleanup, output allocation limits, color/metadata loss, and scope of native evidence. Firefox/Safari, other native applications, crop/fit, ICC, save/reopen and remaining AUT-04.3.3 and overall backlog remain open.


## 2026-10-27 installed DOM module graph verification

Previous turn delivered the explicit browser JPEG adapter and marked actual package entry validation as passed, but its browser installation test loaded the separately generated `dist/browser/dom.js` bundle through a data URL, rather than the package's declared `./dom` default target. That evidence did not establish that the declared module and its relative imports worked in a browser. The isolated packed npm consumer now reads the installed manifest, routes only its actual `dist` JavaScript modules through a local virtual origin, imports the declared target in Chromium, and normalizes the orientation-6 fixture. It checks PNG pixels against the independent normalized fixture, output dimensions/attestation, the actual entry and EXIF inspector requests, and rejects all other network routes. No external service is contacted. The original standalone-bundle check remains separate.

The installed consumer suite passes 6/6. This establishes the tested Chromium and packed module graph only; it does not extend native, crop, color-profile, other-browser or save/reopen evidence. Adversarial review checked that the test uses the installed manifest rather than workspace source, verifies output instead of symbol presence, tests transitive module loading, blocks unexpected requests and distinguishes the separately built bundle. AUT-04.3.3 and the overall backlog remain open.


## 2026-10-27 raw JPEG crop mapping

Previous turn strengthened actual installed DOM module-graph evidence. Layout resolves contain/cover using registered dimensions before asset resolution; replacing JPEG bytes later does not recompute geometry. Added explicit pure `orientJpegCrop` to exporter default, browser, Node and DOM entries. It maps a complete raw-pixel-coordinate crop to the same region after EXIF orientation 1–8, rejects nonfinite, out-of-range and empty regions or invalid orientation, and leaves inputs untouched. It does not decide a frame, swap dimensions, rewrite a presentation, retain screen orientation, or infer whether a crop is already oriented. Callers must register oriented dimensions and choose their own fit/frame policy.

Tests check asymmetric edge mappings and inverse orientations, zero and invalid boundaries, entry parity and detached output. Chromium checks each of the eight mappings against the actual transformed pixels, while an integration test confirms a mapped orientation-6 crop reaches Layout and the serialized `a:srcRect` and that normalized PNG bytes are packaged. The installed consumer checks the real DOM module and types. Initial integration test assumed a fixed media filename and failed; it was corrected to inspect actual media entries. Exporter suite 125/125, installed suite 6/6, browser pixels, typecheck, build and lint passed. The separate orientation-6 native oracle was rerun after the browser test changed.

Adversarial review separated raw and already-oriented coordinates, same region from identical on-screen presentation, oriented dimensions from resolver output, malformed/degenerate crop handling, and mathematical/package evidence from native rendering. Native crop, crop/fit across applications, other browsers, ICC, save/reopen and the overall backlog remain open.


## 2026-10-27 EXIF 6 crop native comparison

Previous turn delivered explicit raw-pixel crop mapping and pixel/package tests. For a 300×200 EXIF-6 JPEG, raw crop excluding the rightmost third maps to bottom one-third of the browser-normalized 200×300 PNG. A separately prepared Pillow control physically crops the raw 200×200 region and rotates it clockwise. PowerPoint 16.104 opened the generated PPTX and slideshow showed both images with yellow/red and magenta/green cells; `crop-powerpoint.png` records the manual UI observation.

`scripts/research/oracles/jpeg-crop-native.py` independently checks decoded source and transformed crop pixels, PPTX image relationships, frame geometry, serialized crop, four screenshot sample cells per picture and four corrupted substitutions. `docs/testing/jpeg-crop-native-evidence.json` binds artifacts and source hashes, but cannot independently witness the UI action or prove whole-image/native roundtrip fidelity. The control PNG, candidate PPTX and producer report reproduce byte-for-byte in a second directory. Adversarial review distinguishes a physical crop from repeated use of the same crop API, confirms no native save/reopen occurred, rejects substituted artifacts and screenshot, and confines the result to one orientation, crop, box and application. Other native crops, editability, ICC, save/reopen and full backlog remain open.


## 2026-10-27 EXIF 6 crop native save/reopen

Previous turn provided a named orientation-6 native crop display and independent physical-crop control. A separate candidate copy was opened in PowerPoint 16.104, its title was edited with `R1`, saved, closed and reopened, then displayed in slideshow. Both sampled four-color grids still matched. The actual native file contains the edit marker and differs from its original; PowerPoint renamed media parts and omitted zero-valued crop attributes. An independent oracle checks semantic crop values with zero defaults, exact PNG resources, picture relationships and frames, screenshot cells, and rejects original/no-edit, crop-mutated and wrong-screenshot controls. Evidence binds file and script hashes. The UI sequence is manual, and screenshot sampling does not establish full visual identity.

PPTKit content projection recovers the crop 0/0/0/0.33333, both frame boxes, resource dimensions and exact PNG bytes from the native file. The no-op save is byte-identical to that file; `scripts/inspect-jpeg-crop-reopen.mjs` reproduces these checks and its report is bound to `docs/testing/jpeg-crop-reopen-evidence.json`. This is one title edit and one save/reopen; crop editing, other orientations and fit, ICC, other apps and the overall backlog remain open. Adversarial review checked saved-file identity, zero-default XML serialization, media relationship remapping, source bytes, UI provenance and false inference from title edit to crop edit.


## 2026-10-27 native crop-handle edit and preview

Previous turn established one title edit and native save/reopen without editing crop. A separate copy was opened in PowerPoint 16.104; its left crop handle was dragged inward, applied, saved, closed and reopened. PowerPoint changed the source crop from left 0 to 25326/100000 while retaining bottom 33333/100000, and narrowed/repositioned the image frame. The untouched control and original resource bytes remained present. The native slideshow shows the narrower left image; this UI sequence is a manual observation.

PPTKit content import recovered crop, changed frame and both exact PNG bytes, and no-op save was byte-identical to the edited native file. Its SVG preview rendered in network-blocked Chromium; sampled colors agree with the reopened native slideshow and measured left-image horizontal spans differ by one pixel. `scripts/research/oracles/jpeg-crop-edit.py` checks native package relationships and geometry, both screenshots, parser report and three substituted controls; `docs/testing/jpeg-crop-edit-evidence.json` binds hashes. This is sampled evidence for one handle, orientation, image and application, not full pixel fidelity or general image editing. Adversarial review verified a real crop change rather than a title-only edit, frame changes separately from source crop, unchanged control/resource bytes, UI provenance and negative controls. Other handles, rotations, applications, image replacement and full backlog remain open.


## 2026-10-27 per-image replacement plan

Previous turn established a named native crop-handle edit/import/preview observation. Core now provides pure `planImageReplacement` with explicit preserve, stretch, contain, cover or caller-specified size policies. It returns deeply frozen detached replacement and previous image snapshots for existing element replacement/undo; it does not mutate assets or the document, register a target, infer pixel-to-slide size, change other users of a shared asset, remove assets or write imported package parts. Preserve means crop fractions and authored frame, not identical content or resolved visible geometry. Contain/cover validate target/frame dimensions and numerical aspect safety; size requires positive finite presentation-unit dimensions.

Core tests cover policies, shared-use isolation, snapshots/undo, invalid fields, crop and dimension failures, extreme aspect underflow, and atomic no-mutation. A package integration test retargets one shared image from a 2×3 PNG to a 3×2 JPEG, verifies cover layout and PPTX crop serialization, two independent media resources and restoration. An initial exact floating-point equality and incorrect media extension in the test were corrected. Browser installed consumer executes the public plan, and installed TypeScript checks its contract. Core tests 241/241, exporter 126/126, installed consumers 6/6, build, typecheck and lint passed; build has a large-chunk advisory.

Adversarial review checked shared identity and stale undo, declaration versus decoded image dimensions, normalized versus raw EXIF, numerical limits, authored versus resolved visible frame, and separation from preservation-aware package writing. Native image replacement, general source crop policies, transparency, unknown references, importer edit-aware save and full backlog remain open.


## 2026-10-27 transparent replacement boundary test

Previous turn implemented explicit per-image replacement plans. Added a transparent 4×2 PNG fixture and a cross-package browser test for replacing one of two users of a shared 2×3 PNG, with `contain` in a 200×200 authored frame. Chromium sampled opaque red, transparent-to-white and letterbox pixels; the other image's sampled pixels stay unchanged. Fresh PPTX generation and content import preserve both independent media bytes and the resolved frame. Replacing the imported image blocks preservation edit even with `text-runs-v1` enabled; undo plus removal of the newly registered asset restores byte-identical no-op save. Undo on the original authoring document restores sampled preview pixels. The test is part of `test:browser-bundles`.

This verifies a controlled fresh-export/import path and rejection/undo on a preservation session, not a preservation writer for image edits or PowerPoint transparency fidelity. Adversarial review distinguished alpha compositing from a white opaque image, shared source isolation, authored versus resolved frame, edit-profile opt-in from image write permission, and no-op undo from general transactional editing. Other fits, native image replacement, ICC, unknown references and the full backlog remain open.

Verification: browser bundle suite 43/43 (including actual PNG alpha decoding and transparent replacement parity), browser package builds, lint and `git diff --check` passed. Full monorepo tests were not rerun for this test-only change.


## 2026-10-27 resolved visible-frame replacement plan

Previous turn added transparent replacement cross-package evidence, while source-preserving image writes remained blocked. Layout now exposes `planVisibleImageReplacement`: it resolves the current document with default layout at invocation, finds a top-level slide image and registered target, then creates a per-image replacement plan using the old resolved box and an explicit preserve/stretch/contain/cover fit. Preserve carries resolved crop fractions; reset policies clear them. The detached previous snapshot retains authored state for undo. It rejects unknown IDs, unresolved contain/cover and zero-area frames. It does not apply plans, guarantee stale-plan safety, interpret EXIF, use optional composition or provide imported-package writeback.

Layout tests exercise contain and cover, resolved versus authored frames, policy choices, source mutation before planning, shared asset isolation, undo and invalid identities/geometry. Exporter integration verifies a 2×3 source contained in a square becomes an 80×120 visible frame which stays in PPTX after replacing it with a transparent 4×2 image. Installed Chromium and TypeScript consumers exercise the public entry. Layout tests 145/145, exporter tests 127/127, installed consumers 6/6, repository build, typecheck and lint passed; build retained a large-chunk advisory.

Adversarial review checked no stale external layout input, zero-area stretch images, unresolved intrinsic sizes, preserving crop fractions versus visual content, contain inside the selected frame, and authored-state undo. Default-layout top-level only is intentional; composed layouts, nested/master images, native replacement, ICC and source-preserving image writes remain unverified or unimplemented. Full backlog remains active.


## 2026-10-27 composed visible-frame replacement

Previous turn added default-layout visible-frame replacement planning. The Layout API now accepts an optional explicit composition snapshot, runs it before image fitting and uses the resulting top-level image frame. Any resulting layout diagnostic blocks the plan, so failed rules or other diagnostics are not silently ignored. It reads the composition option once, clones it before reading the document and preserves authored undo state. Other Layout options, pagination, nested objects and master/layout-owned images remain unsupported. A plan captures geometry only; reapplying the same composition after committing can move or resize the image again.

Tests cover composed alignment followed by contain fitting, detached rules, single-read capture, failed/missing targets, unsupported options, unchanged document and undo. Export integration checks a composed resolved frame in generated PPTX with no implicit recomposition. Installed browser/type consumers exercise the option. Layout tests 147/147, exporter tests 128/128, installed consumers 6/6, build, typecheck and lint passed; build emitted the existing large-chunk advisory. Adversarial review examined double application, partial composition, stale plan, unrelated diagnostics, numerical geometry and undo of authored versus resolved state. Native replacement, further layout stages, nested/master ownership and source-preserving image writeback remain open; full backlog remains active.


## 2026-10-27 nested visible-frame image replacement

Previous turn added composition-aware visible-frame planning for top-level images. Layout now also exposes `planNestedVisibleImageReplacement` for a slide-owned group subtree. It resolves the selected image in its immediate parent's local coordinates and returns detached deeply frozen whole-root-group replacement and previous snapshots for the slide's existing atomic `replaceElement`. Siblings, group IDs, transforms and shared asset uses remain intact. Optional composition may place the child; if a composed ancestor group frame or transform differs from default layout, the operation rejects rather than implying that the uncommitted parent movement was preserved. Other sibling composition outputs are not included, and applying composition again may move content. No source-package edit permission is added.

Layout tests cover two group levels, parent scaling/rotation/flip, local contain geometry, shared sibling isolation, detached snapshots and undo, local composition and ancestor-change refusal, missing identity/assets and unresolved fit. Cross-package tests verify fresh PPTX export, content reimport, nested frame, exact independent media and undo. Installed browser/type consumers cover the entry. Layout tests 150/150, cross-package parity 2/2, installed consumers 6/6, build, typecheck and lint passed; build retains a large-chunk advisory. Adversarial review checked local vs slide-space bounds, ancestor changes, sibling preservation, stale root snapshots, group identity/timeline validation through existing commit, and distinction between fresh export and source-preserving writeback. Master/layout-owned groups, coupled ancestor transformations, image replacement in imported source packages and full backlog remain open.


## 2026-10-27 shared master/layout visible image replacement

Previous turn delivered slide-group image plans. Layout now provides `planSharedVisibleImageReplacement` for ordinary images directly or recursively owned by an authored master or layout. The plan resolves the selected owner domain, returns detached frozen complete definition snapshots and uses the owner's existing atomic replacement operation. The change is shared by every page displaying that definition, not a slide-local override. Optional composition and ancestor-frame guards match nested slide planning; altered sibling geometry is not implicitly committed. Unresolved or empty frames, wrong owner/image/asset and diagnostic-producing composition reject. No source-package preservation writing is implied.

Layout tests cover direct master artwork, nested layout artwork, definition metadata and visibility, shared references, owner isolation, local composition, ancestor movement rejection, failure atomicity and undo. Export integration follows both slide → layout → master relationships, confirms a single master picture references the new media and no per-slide picture copies, checks its resolved frame and source bytes. Installed Chromium and TypeScript consumers cover the new entry. Layout tests 154/154, exporter tests 129/129, installed consumers 6/6, repository build, typecheck, lint and diff checks passed; the build retains a large-chunk advisory.

Adversarial review addressed shared-versus-local scope, complete definition preservation, hidden master artwork, grouped local transforms, media relationships and stale snapshots. Native master-image editing/visibility, unrelated layout stages, EXIF/ICC and source-preserving image replacement remain unverified or incomplete; the overall backlog stays active.


## 2026-10-27 shared master transparent replacement native display

Previous turn implemented shared master/layout image plans. A generated four-page deck used the plan to replace a contained 2×3 master PNG with a transparent 4×2 PNG over a blue shape. PowerPoint 16.104 first-open slideshow manually showed red/blue on both shared-master pages, white when inherited artwork was suppressed, and the green bystander on an independent master. This establishes only named initial display, not a native image edit, source-preserving save or reopen.

`scripts/research/oracles/master-image-replacement-native.py` independently follows slide/layout/master/image relationships, verifies picture geometry, exact transparent resource alpha, source page text and visibility, samples screenshot colors and checks distinct title crops. Three substituted controls reject. Screenshot title crops are not OCR; UI/page correspondence remains manual. `docs/testing/master-image-replacement-native-evidence.json` binds files and hashes. PPTX, report and four SVGs reproduced byte-for-byte in a second directory. Adversarial review distinguished transparent pixels from blue image paint, shared owner from per-slide copies, inherited visibility from separate master, screenshot identity from UI provenance and fresh output from source preservation. Native editing/reopen and full backlog remain open.

Verification: independent oracle and three negative controls, eight focused layout/export tests, lint and `git diff --check` passed. No production source changed in this turn; the full monorepo suite was not rerun.


## 2026-10-27 shared master image native crop edit

Previous turn recorded first-open slideshow of a generated shared-master transparent image. In PowerPoint 16.104, a separate copy was opened in shared master view; the image's left crop handle was dragged inward, saved, closed and reopened. The two shared pages then showed blue/red/blue, the hidden-inheritance page remained white and the independent master green. This is one manual native crop edit and reopen, not a native image-resource replacement.

The independent oracle checks page/layout/master relationships, title text, visibility, picture identity, left crop 20870/100000, changed frame, exact PNG bytes, screenshot title-region consistency and color samples, and three substituted controls. `docs/testing/master-image-native-edit-evidence.json` binds the edited package, screenshots, previous evidence and parser report; neither screenshot sampling nor matching title regions independently proves UI provenance. PowerPoint renamed the media and removed the sample's unused old resource; unknown source references cannot be assumed safe to delete.

PPTKit content import reports no authored master objects for this named native file; no-op save retains every byte. `scripts/inspect-master-image-native-edit.mjs` reproduces this boundary. Adversarial review distinguished native crop from resource replacement, semantic import from byte preservation, shared pages from the hidden and alternate master controls, media-byte identity from arbitrary resource cleanup, and manual UI observation from oracle evidence. Source-preserving image writes, semantic master import, further applications/edits and the full backlog remain open.

Verification: independent edit oracle and three negative controls, parser/no-op reproduction, lint and `git diff --check` passed. No production source changed this turn; full monorepo tests were not rerun.


## 2026-10-27 effective master-picture source inventory

Previous turn recorded a PowerPoint master crop edit and found that content projection did not reconstruct its master artwork. Added explicit `inspectPptxMasterPictures` in Parser root/browser/global entries. It inventories pictures in effective drawing trees of listed slide masters: raw element and group indices, lexical ID/name/placeholder, embed/link relationship classification, explicit transform/crop attributes and listed-layout slide membership. It does not project artwork, follow external URLs, interpret geometry/EXIF/ICC, inspect inactive MCE branches or prove unknown reference closure or edit/delete permission. Unlisted or ambiguous layout ownership is excluded from slide membership. Limits bound inventory work, records and string characters in addition to existing OPC/ZIP/XML limits; shared-memory input rejects.

Tests cover Strict and Transitional, MCE selection, nested groups, raw identities, wrong/missing/external references and missing OPC targets, unlisted layouts, budgets, byte snapshot and cancellation. The Parser suite passed 700 tests before the final additional unlisted-layout test, which also passed in the five focused tests. Chromium/Node parity and installed browser/types consumers pass. A named PowerPoint-edited file was independently checked against raw ZIP/XML for source index 33, crop 20870, frame, media target and three slide identities; `docs/testing/master-pictures-native-evidence.json` binds the source and three corrupted controls. Repository build, typecheck, lint and diff checks passed, with an existing large-chunk advisory.

Adversarial review distinguished effective versus raw-only MCE, declared relationship from validated target, source membership from visibility, lexical geometry from rendered coordinates, bounded inventory from whole-package memory, and source identity from edit permission. Initial tests exposed lower-layer rejection of missing internal targets and shared-memory race risk; the public classification and input guard were corrected. Semantic master artwork projection, complete hidden-reference closure, source-preserving image writes and the full backlog remain open.
