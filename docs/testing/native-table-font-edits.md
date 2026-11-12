# 原生表格字体：追加文字验证

2026-10-26：完成15个具名表格单元格追加 Ω、两组各三轮PowerPoint保存/关闭/准确路径重开与PDF导出。仅完成 IMP-05.1.3.3.2.6.1.2 的追加切片，regular/bold/italic、换主题、独立notes和完整父项仍开放。

## 控制与原生操作

输入来自[36控制无编辑基线](native-table-fonts.md)的第三轮native包，复制到新目录后操作；原始source与native基线未修改。该基线是作者有效包加具名source patch，不是fresh author-map round-trip。环境为PowerPoint for Mac16.104（16.104.25121423）。

| 组 | 追加目标 | 未改控制 |
| --- | --- | --- |
| reference，27页 | minor-tc/ko/ko-KR/ar；major-tc/ko/ko-KR；collection-tc/ar/he，共10格 | 其余17页 |
| filled，9页 | collection-generic-tc/ko/ko-KR/ar/he，共5格 | 其余4页 |

通过CUA进入单元格文本，折叠选区后移动光标，确认原文完整再追加。首个尝试在RTL视觉左端插入，XML显示 Ω 被放在原文开头；该真实失败副本 `filled/append-round1.pptx` 保留，并被顺序oracle拒绝。改从未修改baseline重新操作，使用视觉右端，独立XML确认Arabic/Hebrew的原始逻辑文本与全部组合mark之后才是 Ω。没有用GUI删除或替换文字来修正失败样本。

两组正式文件均为 `append-right-round1..3.pptx/.pdf`；每轮都另存、关闭、用准确路径重开，再原生导出PDF。三轮保存不是同一输入的程序复制。

## 实际结果与合同

全部15个目标由一个run变成两个run。原run保留文字/lang/48pt，新增 `dirty=0`；Ω run为 `lang=el-GR`、48pt、`dirty=0`，CJK/韩语另有原语言的 `altLang`，Arabic/Hebrew没有该属性。两run均无local Latin/EA/CS字体槽。尾段属性新增 `dirty=0`，编辑对象新增具名p14 modId，三轮保持。

原正文的实际字体、glyph、Unicode trace与位置逐项和native baseline相同，Ω均实际使用Georgia；未改21页逐页栅格和baseline一致，全部36页在后两轮与第一轮栅格完全一致，共108个native PDF页面记录。source规则与实际字体仍分开：[基线的16差异](native-table-fonts.md#实际字体结果)不会因为追加成功变成通过。Latin/Greek在显式集合中的实际字体与generic Latin相同，不能由此证明应用使用了supplemental规则。

已实际查看27页和9页contact sheet，确认文字与追加内容可见。Arabic/Hebrew的ToUnicode结果与源文本分开记录；例如filled Arabic trace含 `*` 的映射，不能将其当作源文字已变或通过全局删mark来求相同。既有glyph和栅格稳定不证明完整bidi或mark轮廓正确。

[ADR0053](../architecture/decisions/0053-declared-fonts-and-native-appearance.md)选择保留现v3声明运输合同，不把Office实际fallback写入源family/lang，也不凭此Mac样本引入通用native resolver或自动替代字体。现default/v1/v2/v3和source style writer边界不变。

## 独立验证与源保存

[编辑专用oracle](../../scripts/research/oracles/native-table-font-edits.py)先用原严格baseline reader核对输入hash、真实presentation页序、tableStyles与layout/master/theme关系，再验证实际编辑包。基线oracle没有放宽单run、原文或无local字体条件。编辑oracle从baseline构造只允许的run/语言/dirty/modId变更，逐目标比较完整slide树；所有其他part字节精确相同，仅docProps的编辑时间、修改时间、revision与精确word-count增量单独记录。theme、tableStyles、关系、未改slide、notes及其他资源均在字节比较内。

PDF分别检查页面标签顺序、字体资源与descriptor一致性、实际face/glyph、每个非空span可见dark ink、原glyph/坐标、未改页和轮次栅格。能提取文字本身不算外观通过。

[跨包runner](../../scripts/inspect-native-table-font-edits.mjs)对两组baseline加三轮共8包导入v3，核原run和新增run的完整map、16或8条定位规则与实际选择，再做8次精确no-op；两份第三轮输入另做browser global全document/report一致性及精确no-op，网络请求为空。每个native编辑输入还运行三轮scalar Ω→ΩX→ΩXX→ΩXXX，共18个候选；独立ZIP比较仅允许目标文字变化，实际候选重投影与预期文档相等。**这些scalar候选没有原生重开，不将其计入native编辑轮次。**

[5组跨包回归](../../scripts/declared-slot-font-import-parity.test.mjs)新增一组不需要安装字体的source-only二run拓扑控制：15目标、两种集合组、每run map/provenance、三轮scalar、完整非字符字节精确，以及混改map整体blocked后restore/retry。该source patch模拟观测到的拓扑，不伪称应用执行。

[112项实际反例](../../scripts/research/oracles/native-table-font-edits-checks.py)全部拒绝（reference73、filled39）：逐目标错误prepend/new-language/local字体/旧文/style绑定、可提取但遮白的Ω、PDF字体资源冲突；另含source规则变化、PDF换页、未改邻居栅格变化及真实RTL插入失败副本。

## 复查与剩余任务

本机artifact根目录：`.tmp/native-table-font-edits/run-lrdewp2z`，含reference与filled两组baseline、expected.json、三轮PPTX/PDF、append-inspection.json与append-contact-sheet.png。字体、PDF、截图不提交；[证据manifest](native-table-font-edit-evidence.json)记录hash、实际run/selector/map、字体/ink/raster、候选结果和反例。

```sh
rtk proxy node scripts/inspect-native-table-font-edits.mjs .tmp/native-table-font-edits/run-lrdewp2z
rtk proxy node --test scripts/declared-slot-font-import-parity.test.mjs
rtk proxy node --test packages/pptx-parser/test/declared-slot-font-import.test.mjs
```

接续[执行卡](../planning/advanced-capability-todos.md#表格与备注原生字体的可领取任务)：.6.1.3粗体/纯斜体/恢复regular，.6.1.4主题引用响应，之后.6.1.5汇总完整编辑回归；独立notes基线.6.4.1可另行推进。本轮不完成任何完整字体、表格、导入或round-trip父项，也不扩大source样式写权限。

后续[十五目标native样式](native-table-font-styles.md)已完成.6.1.3，记录真实face、合成stroke/shear及恢复regular；本页追加oracle和原证据保持原范围，主题与notes仍待完成。
