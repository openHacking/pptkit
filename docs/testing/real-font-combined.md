# 完整同段脚本组合的原生局部编辑

2026-10-26：完成IMP-05.1.3.3.2.6.2.3的具名原生验收。同一个段落含Latin、SC、TC、Japanese、Korean、Arabic、Hebrew，另含Greek，共8个作者run和8个物化family。这里的SC/TC/ja是不同语言选择场景，不是三种不同的Unicode Script。此记录补齐[六组配对混合段](real-font-mixed.md)未覆盖的同段组合；完整字体父项、source表格/notes依赖、Unicode bidi/mark轮廓及跨应用能力仍开放。

## 作者输入与导出

[fixture](../../tests/fixtures/real-font-combined.mjs)通过作者API直接生成一个720×260pt页面，Probe为单段48pt、margin0、wrap false：`Latin 骨 骨 骨 한 لا שָׁ Ω`。各run分别用en-US/zh-CN/zh-TW/ja/ko/ar/he/el与完整scriptFontFamilies。fresh输出的8run三个槽分别为Arial、Hiragino Sans GB、Songti TC、Hiragino Mincho ProN、Apple SD Gothic Neo、Geeza Pro、Arial Hebrew、Georgia。三个相同Han字符必须因语言使用不同family，不能凭字符相同合并。

[生成器](../../scripts/prepare-real-font-combined.mjs)核已记录字体文件SHA、导出前后作者canonical不变、8个作者map去重为Probe上的1个located物化警告。没有XML patch或历史PPTX合并。[CI跨包测试](../../scripts/script-font-map-parity.test.mjs)独立DOM核单段/8run、lang和全部三槽，再以Parser v3重导入验证literal families、map不重建和精确no-op；CI无需安装本机字体。

## 原生操作

应用为PowerPoint for Mac16.104 (16.104.25121423)，全部UI通过CUA。fresh源先导出PDF，然后独立副本在Latin内追加X；使用明确前六ASCII字符范围切换局部bold、纯italic，最后恢复regular并应用Office主题。每阶段初始保存/PDF后，各做三轮另存、关闭、准确路径重开、导出PDF。

| 前缀 | 原生source与实际PDF |
| --- | --- |
| combined-fonts | 单段8run/8family，三种Han语言选择实际PS分别匹配。 |
| combined-edit | 插入X与原en-US Latin run合并，初始仍8run；首次重开后拼写检查将LatinX/尾空格拆成2run，总9run。 |
| combined-bold | 仅LatinX带b=1，尾空格及其余7family regular；实际Arial-BoldMT，尾空格ArialMT。 |
| combined-italic | 仅LatinX带i=1，实际Arial-ItalicMT，无额外矩阵shear；其余内容regular。 |
| combined-theme | 恢复regular，Office实际主题链与方案改变；原8family literal槽、语言与实际face保持。 |

插字首次重开的LatinX run新增dirty=0/err=1，空格单独dirty=0。oracle对这处精确拓扑与属性变化建期望；没有全局忽略proofing字段或合并run来制造不变。其他三个阶段每轮raw records与该阶段初始严格相同。原显式pPr被native编辑移除，frame/box/边距保持；不同阶段字宽/行距/位置可变，不声称整页相对原始不变。

初次native保存规范化theme XML，raw hash不同但fontScheme相同，用此前独立native PPTKit theme字节hash精确约束。Office主题必须通过真实slide→layout→master→theme链、name、raw hash与fontScheme变化联合证明。

所有文件是新的临时副本，未覆盖历史SC/五脚本/配对混合段/Cycle证据。无删除文字、字体安装或系统设置修改。

## 源顺序、实际字体与可见性

[oracle](../../scripts/research/oracles/real-font-combined.py)独立ZIP/XML核单页/Label/Probe/完整文字/8或9run、全部lang、三槽、48pt和精确局部样式；源PPTX SHA固定。PDF核实际PS与按font/text合并的完整trace序列，不能用family集合掩盖三个相同Han字符之间的字体绑定互换。ArialMT同时服务标签与尾空格时，按48pt矩阵排除16pt标签。

源run顺序为Latin→SC→TC→ja→ko→ar→he→Greek；实际非空白glyph的横向位置顺序为Latin→SC→TC→ja→ko→he→ar→Greek。两者和PDF content trace顺序分别记录并固定核对具名native结果，不以数组顺序宣称屏幕顺序，也不据此认证通用Unicode bidi正确性。

每个非空白trace的实际bbox必须有超过5个dark pixels，整段总ink也检查。八类font span被白色绘图覆盖但文字仍可提取的真实负例全部拒绝，避免“其余文字有墨迹”掩盖一段不可见。尾空格仅advance的情况不要求有墨迹，完整source空格在XML/Parser精确核验。

相邻七family的PDF glyph序列相对原始source保持；四阶段各三轮PNG与该阶段初始逐像素相同。已实际查看原始/插字/局部粗体/斜体/主题五行contact sheet，八类内容可见。本样本Arabic为无mark的lam-alef、Hebrew包含mark与presentation-form提取差异；较长带mark Arabic见[前批证据](real-font-multi.md)。邻居glyph序列稳定只证明保留native基线，不证明该基线的所有轮廓和bidi顺序符合规范。

## 投影、负例与复查

[runner](../../scripts/inspect-real-font-combined.mjs)以Parser v3重投影source+16原生PPTX：完整单段文字/8或9run、语言、literal槽、局部b/i、不恢复scriptFontFamilies均核验；17份no-op与输入bytes完全一致，canonical快照不变。没有扩大source font/style writer权限。

[55项实际破坏控制](../../scripts/research/oracles/real-font-combined-checks.py)全部拒绝：逐run family/lang、三组Han run顺序、局部样式扩散/缺失、X/Greek/Hebrew marks丢失、八类可提取但白色遮盖的span、字体资源命名冲突、三组真实PDF CJK font-selector互换、proofing字段/拓扑异常、旧主题、italic实际face冒充bold。资源元数据改名与真正内容流font引用互换分开检查，不将改名冒充字体二进制替换。

共用的PDF资源一致性读取检查BaseFont、FontDescriptor/FontName与reported face，已将此前RTL空格/正文分资源的防漏逻辑提取供两个oracle复用；既有六页配对正例和72反例重新通过。所有负例作用于临时文件，正式native artifacts不变。

本地artifacts：`.tmp/real-font-combined/run-H4I1Ud`，含combined-fonts、四阶段及-round1至-round3的PPTX/PDF，combined-inspection.json与combined-contact-sheet.png。[manifest](real-font-combined-evidence.json)记录17份文件SHA/投影/no-op、source/raw/主题、glyph/实际视觉顺序/matrix/逐span可见墨迹/raster、1物化警告和55负例。字体二进制与PDF仅留本机。

```sh
rtk proxy node scripts/prepare-real-font-combined.mjs
# 对新run另行执行native操作及每阶段三轮导出，不复用旧PDF
rtk proxy node scripts/inspect-real-font-combined.mjs .tmp/real-font-combined/run-H4I1Ud
rtk proxy pnpm test:script-fonts
```

.6.2.3具名同段组合与局部编辑退出证据已齐。父.6.2的操作测试门得到补齐，但其显式依赖.5.2/.6.1尚有同face完整外观/表格fontRef/notes等开放范围，父项不勾选；下一步回到这些依赖以及独立source主题字体编辑任务。
