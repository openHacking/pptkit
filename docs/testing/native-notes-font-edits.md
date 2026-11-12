# 备注编辑：source 回归与原生验收准备

2026-10-26：已完成 source-only 二run/局部样式回归及A/B备注原生追加、局部三样式三轮验收；后续[具名主题响应](native-notes-font-themes.md)也已完成。本文不授予任何新的source样式写回权限，也不完成.6.4.2。

## source-only 交付

[fixture](../../tests/fixtures/native-notes-font-edits.mjs)从[独立notes主题控制](native-notes-fonts.md)生成source，A/B两组各15页，在第3/5/6/7/8/10/11/12/13/14/15页末追加独立Ω run；原run文字/language不改，新run显式en-US并保留具名局部字体声明。局部bold/纯italic/恢复regular只作用于Ω，原文保持regular；它是明确的source合同控制，**不预设PowerPoint实际生成的run拓扑、字体槽或语言与fixture相同**。

[跨消费者测试](../../scripts/notes-native-font-edits-parity.test.mjs)覆盖A/B × append/bold/italic/regular共8输入、120源页面，Node/browser ESM/global各三轮，共72个scalar保存路径和72次精确no-op。每轮验证两run/原文/ΩX增长、language、40pt/type继承、完整map（local覆盖不带map、空/坏EA阻断EA脚本键）、独立来源XPath及style值；三消费者完整document/report/输出bytes相等，浏览器零网络。source bold混改整体blocked，restore后只改Ω字符；实际候选重导入snapshot完整相等，目标字符外所有XML/ZIP payload字节精确保留，原文、标签、slide、master/theme及其他备注不变。此证据不需要安装专有字体，不计入native页面。

该测试已接入`test:import`；它补充[原生基线source回归](../../scripts/notes-native-font-parity.test.mjs)，不替代其缺theme/缺master/错误type及scope隔离负例。字体/style写回仍需IMP-09独立准入。

## 原生操作状态与恢复点

[准备脚本](../../scripts/prepare-native-notes-font-edits.mjs)从上批manifest核实A/B第三轮PPTX与Notes PDF的hash后复制新目录；原基线保留。当前目录`.tmp/native-notes-font-edits/run-t0kxZH`，A/B各有baseline.pptx、baseline-notes.pdf、expected.json。baseline在磁盘上的hash仍等于原生基线manifest。

用户解锁后先核实A准确窗口与锁屏前内容，补完其余目标；两组各11目标追加Ω，均已完成三轮另存、关闭、准确路径重开与实际纵向“注释”PDF，共6份原生PPTX、6份PDF/90页。baseline未覆盖且SHA仍与上批manifest相同。Mac锁屏恢复点现已解除，旧执行记录保留历史状态，不代表当前进度。

[独立append oracle](../../scripts/research/oracles/native-notes-append.py)按实际观察定义完整notes树变换：原run只新增dirty=0；新Ω使用en-US，CJK/韩文保留altLang，Arabic/Hebrew不新增altLang；局部字体、空EA、坏EA、major槽原样复制；major-TC新run不带其余页的dirty/err；尾段新增dirty=0。其余ZIP payload逐字节不变，只有viewProps/app/core单列源/输出XML与hash，不声称完整native ZIP字节相同。三轮完整notes语义一致，原文/标签/slide/theme/master/关系不变。

每页原glyph/face/位置完整trace与baseline相同（仅忽略绘制序号）；新增Ω具名actual face为A全11页ArialMT、B除三local外8页Georgia/三local为ArialMT。逐trace字体资源与非零glyph、实际41pt打印尺寸、可见墨迹均核实；Ω区域以外全页栅格精确匹配baseline，未改四页全页精确相同；第二/三轮全页trace和栅格与首轮精确相同。baseline中空白slide-image保持为已知限制，未隐藏。

[runner](../../scripts/inspect-native-notes-append.mjs)与[manifest](native-notes-append-evidence.json)记录6次Node精确no-op、6次source样式混改blocked并恢复、18个字符候选完整重投影/字符外payload精确。候选未native重开，不计入90页。22个新增Ω的声明选字仅6匹配actual face、16差异：两组各8个非local目标按script规则选择的Greek字体与PowerPoint用Latin槽实际绘制不同；此处是具名差异证据，不擅改通用Unicode分段合同，另拆language/字体域策略卡排查。原来的source-only fixture合同不因此冒充原生选字。

[负例oracle](../../scripts/research/oracles/native-notes-append-checks.py)生成28份实际重打包ZIP/PDF破坏文件，text/lang/bold/原run语言/local槽/master/theme/关系/slide邻页/旁路notes、可提取但遮白的新旧正文、页序和后轮额外墨迹全部拒绝并记录SHA。资源删除/错字体更深入反例与跨消费者native输入汇总仍归.6.4.2.5，严格无编辑baseline oracle未放宽。

## 原生局部样式验收

A/B均从各自append-round3另存新名字，依次只选择新增Ω施加bold、关闭bold并开启italic、关闭italic恢复regular。每阶段均另存、关闭、准确路径重开、打印实际纵向Notes PDF三轮，共18实际输入/270备注页；与追加90页分别计数。已目视复核A恢复常规后的备注内容与准确窗口URL。

[独立样式oracle](../../scripts/research/oracles/native-notes-styles.py)先以严格append oracle认证不可变基线，再逐个payload核完整notes树：新Ω及相应endParaRPr的b/i变更，major-TC也新增dirty=0/err=1；原文run、其他段落/notes、标签、slide/theme/master和所有关系payload精确不变。viewProps/app/core完整XML/hash单列。恢复regular保留显式b=0/i=0，不冒充恢复原始XML。

实际Ω字体为Arial-BoldMT/Arial-ItalicMT/ArialMT或Georgia-Bold/Georgia-Italic/Georgia。字体资源、非零glyph和可见墨迹逐页核；只有单一fill trace，无合成stroke，实际文字矩阵无shear或非等比缩放。原文和旁路trace/位置不变，Ω原点不变，变化只在新旧Ω包围区域；regular全部90页完整栅格恢复append基线，未改页面各阶段完整栅格不变；每阶段后两轮全trace与栅格均等于该阶段首轮。验证矩阵按实际170.8333文字尺寸区分41pt正文与同Arial字体17pt标签，不把标签矩阵误算为额外Ω。

[样式runner](../../scripts/inspect-native-notes-styles.mjs)和[manifest](native-notes-styles-evidence.json)记录18份实际输入的captured profile重导入、18次精确no-op和原run regular/Ω-only flags投影。[破坏验证](../../scripts/research/oracles/native-notes-styles-checks.py)生成60份实际ZIP/PDF反例：Ω错误flags（含纯italic误留bold）、原文误设italic、language/local槽/旁路notes变化、可提取遮白新旧正文、换页、人工矩阵shear和后轮额外墨迹均拒绝并记录SHA。原始字体fallback与追加16差异仍保留，未开放source样式writer。三消费者native输入候选与更完整资源/字体破坏仍归.6.4.2.5，不以source-only三消费者证据代替。

运行node scripts/inspect-native-notes-append.mjs或node scripts/inspect-native-notes-styles.mjs，参数均为.tmp/native-notes-font-edits/run-t0kxZH（shell通过rtk proxy执行）。两者只检查已有捕获，不触发GUI或制造native轮次。独立正/负例验收通过，日志.tmp/native-notes-{append,styles}-evidence.log；lint/边界/文档及diff检查另记执行记录。原生fixture依赖本机PowerPoint和具名字体，不接入缺该环境的CI；可移植source合同回归已在test:import内。

## 后续验收

待办细卡见[表格与备注原生字体任务](../planning/advanced-capability-todos.md#表格与备注原生字体的可领取任务)。原生追加与局部bold/pure-italic/regular各三轮已完成，后续[三组具名Office主题](native-notes-font-themes.md)亦已交付；独立核实际notes→master→theme、run/局部字体、PDF actual face/墨迹/glyph/位置、未改备注与slide邻居。A/B是不同文稿的notes域。规范每包最多一个Notes Master part；合法包内隔离由共享notes master下的逐notes themeOverride另验，额外notes-master只作非合规source诊断控制，见[基数与隔离证据](notes-master-isolation.md)。原shared override part在具名PowerPoint提示修复，实际正向主题证据采用每notes独立part，不把source投影当native准入。不把非合规输入的容错读取作为标准原生准入。错type/缺theme/坏slot的source阻断与actual native响应分别报告。

保留严格无编辑baseline oracle，新增编辑oracle只能按观测列明允许变化；破坏文字/language/样式/槽/关系、字体资源、仍可提取的遮白正文、页序与旁路栅格必须拒绝。Notes打印不能用普通slide导出替代。slide-image空白继续由IMP-03.3.2.1定界修复，完整notes页/fields/artwork及source writer不随字体局部操作放行。

```sh
rtk proxy node --test scripts/notes-native-font-edits-parity.test.mjs
rtk proxy node scripts/prepare-native-notes-font-edits.mjs .tmp/native-notes-fonts/run-haqFDP
```

第二条会创建新run并更新latest指针；恢复当前未保存原生工作应继续使用明确的run-t0kxZH，不重新生成或另开baseline。build/typecheck、22组Parser和新source回归已通过；lint/文档检查记录随执行记录维护，未声称完整workspace runtime套件重跑。日志`.tmp/native-notes-font-edits-{build,typecheck,parser-tests,source-tests,lint}.log`。
