# 原生表格字体：粗体、纯斜体与恢复 regular

2026-10-26：完成 IMP-05.1.3.3.2.6.1.3 的十五目标样式操作。粗体、纯斜体、恢复regular三个阶段，两组各三轮PowerPoint保存/关闭/准确路径重开/PDF，合计324个页面记录。主题、独立notes、Windows与完整字体/表格父项继续开放。

## 输入与操作

新副本来自[追加验证](native-table-font-edits.md)的两组第三轮包。reference27页的目标为minor-tc/ko/ko-KR/ar、major-tc/ko/ko-KR、collection-tc/ar/he；filled9页目标为collection-generic-tc/ko/ko-KR/ar/he，共15格。其他21页作为旁路控制。每格已有原文与Ω两个run。

PowerPoint for Mac16.104（16.104.25121423）通过CUA逐格进入文本、全选单元格文字，核对48pt和工具栏粗/斜体状态后操作。首格与RTL格选区已实际观察；不能只双击选中某个字就认作全格完成。阶段按粗体→关闭粗体并开启斜体→关闭斜体推进，每次操作后重新读取状态。每阶段第一轮保存后关闭并准确路径重开，再导出PDF；后续两轮也真实另存/关闭/重开/导出。没有重写源包或调用无界GUI脚本代替原生操作。

## 源声明与实际字体

三个阶段均保持两个run、原文字/lang、Ω的el-GR/原有altLang、48pt以及全部source字体集合；没有local Latin/EA/CS字体覆盖。源变化如下，尾段插入样式同步：

| 阶段 | 两run及endParaRPr | 其他变化 |
| --- | --- | --- |
| bold | `b=1`，没有i属性 | 对象modId更新，docProps编辑时间/修改时间/revision更新。 |
| italic | `b=0, i=1` | 同上；不是bold-italic。 |
| regular | `b=0, i=0` | 同上；显式false与原先省略不同，不能说源XML恢复原样。 |

每阶段的modId、run/source选择在后两轮保持。独立oracle比较完整slide树，只允许上述具名样式/modId变化；其他part按byte比较，docProps仅允许具名时间/revision变化，Words及其他字段保持。tableStyles、theme、关系、未改slide、notes和资源均保留原字节。

| 原正文实际regular face | bold观察 | italic观察 |
| --- | --- | --- |
| STSongti-TC-Regular | STSongti-TC-Bold | 原face合成倾斜 |
| MalgunGothic | MalgunGothicBold | 原face合成倾斜 |
| AppleSDGothicNeo-Regular | AppleSDGothicNeo-Bold | 原face合成倾斜 |
| GeezaPro | GeezaPro-Bold | 原face合成倾斜 |
| HiraginoSansGB-W3 | HiraginoSansGB-W6 | 原face合成倾斜 |
| Batang / SimSun / ArialUnicodeMS | 原face填充+描边 | 原face合成倾斜 |
| ArialMT，collection-ar | Arial-BoldMT | 原face合成倾斜 |
| ArialMT，collection-he | Arial-BoldMT | Arial-ItalicMT |
| 新增Ω，Georgia | Georgia-Bold | Georgia-Italic |

合成粗体不是重复源文字：PDF同一组glyph/坐标有type0 fill与type1 stroke两条trace，stroke linewidth=4.5；原正文文字矩阵的纵向尺度196.0784、基线1085.206，另有200的横向尺度。合成斜体使用原face，文字矩阵shear=67.96875、横纵尺度200；真实italic face及Ω不叠加shear。oracle独立检查这些具名观察值，不把同名face或工具栏标记当成实际粗斜体证据。

每页PDF校验实际font resource/descriptor一致、无glyph0、每个非空trace有可见dark ink、仅fill trace的ToUnicode顺序与样式前一致；不会删除stroke trace来隐藏合成绘制。未改页逐页栅格保持，各阶段后两轮与该阶段第一轮精确一致。**恢复regular后三轮全部36页的栅格均与样式前基线相同**；源b/i显式false仍保留。已实际查看两组bold/italic四张contact sheet。

[无编辑基线的16个选择差异](native-table-fonts.md#实际字体结果)仍是差异；本批没有把字体替代写回source family/lang，也没有新增Office resolver。RTL源组合mark完整保留，但PDF ToUnicode与source不同的事实不变；glyph/墨迹/栅格稳定不认证完整bidi、mark轮廓或PDF嵌入subset字节身份。规则/外观分离遵循[ADR0053](../architecture/decisions/0053-declared-fonts-and-native-appearance.md)。

## 交叉验证与反例

[专用oracle](../../scripts/research/oracles/native-table-font-styles.py)先验证输入hash及原追加实验的严格reader，再验证三个阶段完整slide/source树与PDF。原无编辑/追加oracle均保持各自的严格范围，不通过放宽它们接收样式编辑。

[runner](../../scripts/inspect-native-table-font-styles.mjs)核两组baseline加18份native样式输入，共20次Node精确no-op；每个run的完整map/lang/style、来源规则与Core声明选择保持。六个阶段第三轮输入另做browser global document/report一致性和精确no-op，网络请求为空。十八份native样式输入逐一验证PPTKit source bold写回仍blocked，restore后再做三轮Ω→ΩX→ΩXX→ΩXXX，共54个scalar候选；每份候选重投影相等，除目标字符外所有源字节精确不变。**这些scalar候选没有native重开，不能计入324页或原生样式验收。**

[380个实际ZIP/PDF破坏反例](../../scripts/research/oracles/native-table-font-styles-checks.py)全部拒绝：reference250、filled130。逐阶段/目标分别破坏原run bold、新run italic、尾段插入样式、local字体与文字，或遮白仍可提取的Ω、篡改字体资源、修改文字矩阵shear；每阶段另测source规则、未改页栅格与页序，另实际移除合成粗体stroke。正反例使用不同临时文件，原生证据不覆盖。

[6组跨包回归](../../scripts/declared-slot-font-import-parity.test.mjs)新增source-only三阶段×两组控制，使用[具名source assembler](../../tests/fixtures/native-table-font-styles.mjs)模拟观测到的两run/style声明，不加载真实字体。逐stage检查两个run的声明map/lang/flags、精确no-op、style混改blocked及restore后三轮scalar字符外byte精确。此fixture是source回归，不冒充PowerPoint执行。

## 复查与接续

artifacts根目录：`.tmp/native-table-font-styles/run-7o6r7z8c`，reference/filled各含baseline、expected.json、三个阶段各三轮PPTX/PDF、style-inspection.json与各阶段contact sheet。[manifest](native-table-font-style-evidence.json)记录输入输出hash、应用版本、source/实际face/合成绘制/矩阵、栅格、反例及候选结果。字体/PDF/PNG仅留本机。

```sh
rtk proxy node scripts/inspect-native-table-font-styles.mjs .tmp/native-table-font-styles/run-7o6r7z8c
rtk proxy node --test scripts/declared-slot-font-import-parity.test.mjs
rtk proxy node --test packages/pptx-parser/test/declared-slot-font-import.test.mjs
```

后续[主题引用响应与.6.1.5汇总](native-table-font-themes.md)已完成；下一步[.6.4.1独立notes基线](../planning/advanced-capability-todos.md#表格与备注原生字体的可领取任务)。本批只完成.6.1.3具名样式切片，不关闭.6.1、字体/表格/导入/round-trip父项，不扩大source style writer权限。
