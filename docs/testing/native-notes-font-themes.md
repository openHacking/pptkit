# 独立 notes 主题与逐备注 override 的原生响应

2026-10-26：完成具名 Office slide 主题验证。A/B 各15页使用独立 notes master/theme；第三组30页交错 A/B，共享单个 notes master，B 的15个 notes 各自关联一个独立 fontScheme override part。三组换主题后各三轮保存、关闭、准确路径重开并打印实际纵向 Notes PDF，共180页。第三组另有三轮无编辑基线90页和首次打开30页；普通幻灯片另导出主题前60页与主题后三轮180页，不与 Notes 页混算。

任务 .6.4.2.4/.4.2 的这个具名切片完成。完整 notes 页、slide-image、字体等价、bidi/mark 正确性、跨平台、source 样式 writer 与完整 round-trip 仍未完成。

## 输入与真实操作

[准备脚本](../../scripts/prepare-native-notes-font-themes.mjs)先核上一批原生 regular 第三轮 PPTX/Notes PDF 的 manifest SHA，以及隔离 source 的 SHA，再创建新目录。A/B 为精确复制，原文和11处 Ω、局部/空/坏/major 槽保留；不是重新生成的鲜作者往返。最终目录为 .tmp/native-notes-font-themes/run-r7ySY9，准备脚本不应在恢复该目录时重跑。首次准备时分类字符串不匹配导致停止；脚本已改为验证所有输入后才创建 run，未将早期部分目录用于验收。

在 PowerPoint 16.104 中通过 Design 的 Office 主题按钮施加给所有幻灯片，并在保存前核实际选中状态。A/B 的 baseline 及第三组 native-baseline-round3 各为主题操作前输入；三轮命名 office-round1/2/3。Notes 通过打印对话框的“注释”/纵向生成；普通幻灯片通过 Save As PDF 另存为 slides.pdf。后补普通 PDF 时重开的是同一批精确 SHA 的原生 PPTX，不额外计作保存轮次。

第三组采用[source 隔离控制](notes-master-isolation.md)，原始15个 B notes 共用一个 override part。该源打开时实际出现“内容有问题，可尝试修复”提示，已取消，没有执行修复。三个定位对照的实际结果如下；它们的文件与 SHA 在[主题 manifest](native-notes-themes-evidence.json)保留：

| 对照 | 原生打开结果 |
| --- | --- |
| 同一30页合并源，移除 override | 正常30页窗口 |
| 已验收的原生15页 A 输入，只有一页加 B override | 正常15页窗口 |
| 同一30页源，B 的每个 notes 各自使用内容相同的独立 override part | 正常30页窗口，无修复提示 |
| 原始30页源，15个 notes 共用一个 override part | 内容错误/修复提示，取消 |

正向原生基线和 Office 操作使用第三行。不能据此断言共享方式违反 ISO，也不能把 reader 的 source 投影/原字节保存当成 PowerPoint 原生准入。官方 NotesSlidePart 支持 ThemeOverride 关系，见[SDK 生成源](https://github.com/dotnet/Open-XML-SDK/blob/main/generated/DocumentFormat.OpenXml/DocumentFormat.OpenXml.Generator/DocumentFormat.OpenXml.Generator.OpenXmlGenerator/Part_NotesSlidePart.g.cs)及[ThemeOverride 元素说明](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.themeoverride?view=openxml-3.0.1)。这里仅证明具名应用中的共享/独立 part 差异，完整 schema 认证及其他应用行为未推断。

## 部件关系、源运输与候选

[独立 ZIP/XML oracle](../../scripts/research/oracles/native-notes-themes.py)从 presentation 的实际页序跟踪 slide→notes→master→theme/override，同时跟踪 slide→layout→master→theme；校验唯一列出/实际 notes master、双向关系、标签、body type/idx7/99、完整 notes/master/font trees、局部声明和所有源关系角色。

第三组原生将物理 slide/notes 文件重排为展示顺序，独立主题迁移到 theme2，15个独立 override 重命名为 themeOverride1..15。oracle 按实际标签/关系建立源到输出映射，不按文件号猜测。首次到第三轮只允许具名 locks、-0、Latin 末段及冗余黑色规范化；原文、有效方案/局部槽、全部关系目标语义保持。初始 shared→distinct 转换也另核所有非 override payload 精确不变，逐 notes 关系仅更换 override 目标。

换 Office 后三组的 notes、notes master、独立 theme2、各 override、所有 notes 关系和 slide 内容 payload 逐字节不变。slide 关系只将旧 layout1 换为 Office blank layout7。新增10个 layout/关系，Office master/layout 的完整规范化树与既有已验收原生 Office 参考相同；63个 GUID creation identity、12个数字 creation identity 及共享字段身份另核。theme1 的精确 SHA 是已验收参考的 1f98605413696c4fe4029a7b7df696261233d73d8fd06413d3990abfe51ae216。

presentation 仅允许具名 slide-master ID 更新和完整参考 Office defaultTextStyle 替换；content types 的新增条目逐项核。core/app/viewProps 完整前后 XML 单列，不声称整个原生 ZIP 字节相同，也不将它们的记录视为无限通用规范化。

[runner](../../scripts/inspect-native-notes-themes.mjs)对15条实际输入验证路径执行 captured v3/notes profile 导入、精确 no-op、source 样式混改整体阻断后 restore，以及各三轮字符保存，共45个 headless 候选。候选完整重投影相等，目标字符之外所有 ZIP payload 精确相同；没有 native 重开，不计入原生页数。DOM 独立核每个 actual source/target/resolvedFrom XPath、完整脚本 map、local/空/坏槽阻断和40pt/type来源；完整 snapshot/report 写入每组 inspections 并记录 SHA。各组第三轮 browser global 的完整 document/report/bytes 与 Node 一致，精确 no-op 且零网络，共3条；不冒称所有 native 输入都运行了三消费者。

## 实际 Notes 外观

三组的 notes 正文字体链按各自独立主题/逐 notes override 保持，原文字形、Ω、font/glyph/origin、局部槽和旁路 label 保持。B 组15页 Notes 全页栅格/trace 与基线精确相同；A 组15页和第三组30页的 master header 都是 Independent Notes A，最后一个 A 的横向原点向左0.917999267578125pt，其他 glyph/位置不变。presentation 同时换成带 kern 的 Office 默认文字样式；本批未进行单变量反事实实验，不能把该相关性当成已证明的根因。

该最后字形的实际可见墨迹单独检查，变化区域以外全页栅格精确保持。三个组各自后两轮全部 trace 与栅格均和首轮相同。header 的轻微差异没有删掉或合并进“全部一致”。实际打印中源码16/40pt仍为17/41pt。合并30页中 A/local-ar 的 ArialUnicodeMS ToUnicode 子集值与单独 A 文稿不同，已按实际值记录；同一合并文稿从首次到三轮基线及换主题后保持，不将提取文本的标记差异解释为完整 bidi/mark 正确。

字体资源、非零 glyph、每段墨迹、页序和实际页面尺寸均核。Notes 中 slide-image 仍为空白，不能通过这份 PDF 证明普通 slide 外观或完整 notes artwork。

## 普通幻灯片旁路

因此另补[普通 slide oracle](../../scripts/research/oracles/native-notes-theme-slides.py)和[独立 manifest](native-notes-theme-slides-evidence.json)：三个组分别主题前和三轮主题后的720×260pt原生 PDF，共240页。它们不是“注释”PDF，也没有替代 Notes 验收。

Office 下16pt标签上移1.20001220703125pt，48pt原文上移0.959991455078125pt；字体、glyph ID、文字、横向原点保持，两个嵌入 TrueType 程序的实际字节 SHA 逐页相等。每轮实际 native raster 的 SHA 单列。在内存诊断副本中仅撤销两项具名 cm 矩阵 Y 平移后，整页栅格与基线逐字节一致；原始 PDF 不修改，不能称为原始栅格完全一致。第二/三轮的原始 trace/栅格均与首轮精确相同。

Arabic 有零 advance 的 mark trace，其字体度量 bbox 宽度为零，不能按一列像素认定不可见或忽略。oracle 从实际嵌入 TrueType 的 head/loca/glyf 按真实 GID 读取有界轮廓 bbox，用原生 glyph origin/size 定位实际墨迹；它不把 ToUnicode 的标记值当普通连字符，也不新增生产依赖。整页反向平移后的栅格仍须精确匹配。

## 实际反例与检查

[Notes/部件反例](../../scripts/research/oracles/native-notes-themes-checks.py)保存55个实际 ZIP/PDF 破坏文件，全部拒绝：文字/lang/style/local槽、master type/size、notes 借 slide theme、override 丢失/错字典、slide 邻居/layout、Office 字典/default kern、真实页序交换、字体资源错名、可提取遮白正文/被允许位移的 header A 和后轮额外墨迹。shared override 的原生修复提示不混入这些 oracle 反例计数。

普通 slide 再保存18个实际 PDF 反例，文字遮白、错字体资源、嵌入字体程序字节变化、额外几何位移、页序与后轮墨迹均拒绝。合计73个真实文件，SHA/拒绝原因保留；原基线/原生文件不覆盖。原无编辑/追加/局部样式 oracle 的规则未放宽。

source 回归增加每 notes 独立 override 的控制，现三类30页 source × Node/ESM/global × A→B→A三轮，共27个 scalar/no-op路径；各域实际 source/map 和字符外 bytes、完整候选及 browser 零网络继续核，诊断额外 master/两个 listed 拒绝仍保留。

    rtk proxy node --test scripts/notes-master-isolation-parity.test.mjs
    rtk proxy node scripts/prepare-native-notes-font-themes.mjs .tmp/native-notes-font-edits/run-t0kxZH .tmp/notes-master-isolation/run-lh85iw
    rtk proxy node scripts/inspect-native-notes-themes.mjs .tmp/native-notes-font-themes/run-r7ySY9
    rtk proxy python3 scripts/research/oracles/native-notes-theme-slides.py .tmp/native-notes-font-themes/run-r7ySY9

第二条创建新目录；恢复已捕获文件不重跑它。后三者分别准备/验证已有捕获，不会代替 GUI 操作生成原生轮次。运行主主题 runner 会更新它的 manifest；随后应重跑普通 slide oracle，使两份证据的 hash 关联一致。

build/typecheck、三组 source 回归、两套独立 native/反例 runner 与语法检查通过；lint/边界/文档及 diff 见执行记录。日志 .tmp/native-notes-themes-{build,typecheck,evidence,lint}.log、.tmp/native-notes-theme-slides-evidence.log 与 .tmp/notes-distinct-override-source-tests.log。没有声称完整 workspace runtime 套件重跑。

下一项是 .6.4.2.5 各原生编辑阶段候选/资源反例汇总。新增Ω的16项声明/原生字体差异、共享 override 准入限制、完整 notes 页/slide-image 和 source style writer 不因主题子卡完成而关闭。
