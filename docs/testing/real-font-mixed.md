# 原始混合段物化与局部原生编辑

2026-10-26：六页分别使用 Latin + SC/TC/Japanese/Korean/Arabic/Hebrew + Greek 的原始混合作者run。完成fresh三family物化、局部插入X、Latin范围bold/italic、恢复regular后换Office主题；四种操作各三轮原生另存、关闭、准确路径重开及PDF导出。此处验证六组配对混合段；[后续完整同段组合](real-font-combined.md)已补齐.6.2.3，父项依赖仍开放。

## 作者输入与fresh输出

[fixture](../../tests/fixtures/real-font-mixed.mjs) 通过作者API直接生成720×260pt、48pt正文，每页一个作者run：`Latin 本脚本文字 Ω`。run具有完整scriptFontFamilies和该页language，作者canonical快照保持。没有XML patch、历史PPTX合并或native对象伪装。

[生成器](../../scripts/prepare-real-font-mixed.mjs) 写出6个具名Label/Probe控制。每个Probe原始1run在fresh XML物化为3run：`Latin ` 的三个槽均Arial、本脚本文字及尾空格的三个槽为对应family、`Ω`的三个槽均Georgia。恰有6个located script-font-map-materialized警告。此处Georgia是原作者现有字符的物化结果，与[上一批单family物化后追加Ω时的native fallback](real-font-multi.md)是不同情形，不能据此宣称可重建unused map。

[CI跨包测试](../../scripts/script-font-map-parity.test.mjs)独立DOM验证六页的三段文字与全部字体槽，随后v3重导入验证有效literal runs、无map重建、精确no-op；无需CI安装这些字体。

## 原生局部操作与反例

应用为PowerPoint for Mac16.104 (16.104.25121423)，全部操作通过CUA。打开fresh源文件先导出source PDF，然后独立副本中双击Latin、Right折叠选择为插入点、追加X。源XML验证完整逻辑文字，无既有文字替换或删除。

编辑后每页为5run：`Latin`、`X`、原空格、本脚本文字及尾空格、`Ω`。X使用en，SC/TC/ja/ko页另带原语言altLang，ar/he页不带altLang。其余run保持原语言；三个family均保持literal槽。局部样式仅作用于前6个ASCII字符LatinX；原空格和其余段保持regular。

首次用双击选词时，ko/ar/he未包含X，粗体只覆盖Latin；`mixed-bold.pptx/pdf`保留此原生反例。进一步在混合选择上按快捷键时先清除了粗体，另存的`mixed-bold-range.pptx/pdf`也保留。最终使用明确六字符范围，从统一regular状态设置bold，独立source检查每页前2run为bold、其余3run不带bold后，才认定`mixed-bold-final`为正式粗体样本。没有把toolbar状态或视觉近似当作验收。

| 阶段前缀 | 源样式与实际字体 |
| --- | --- |
| mixed-fonts | 3个fresh物化run；ArialMT / 本脚本regular PS / Georgia。 |
| mixed-edit | 5run，LatinX其余内容完整；仍为3个family的regular实际face。 |
| mixed-bold-final | 只有Latin、X带b=1；实际Arial-BoldMT；相邻空格、本脚本和Georgia保持regular。 |
| mixed-italic | 只有Latin、X带i=1；实际Arial-ItalicMT，不是regular face加shear；相邻内容保持regular。 |
| mixed-theme | 恢复regular，真实Office主题链/字体方案改变；原literal三family和实际face保持。 |

每个编辑阶段初始保存/PDF后再做三轮另存、关闭、重开、导出，文件后缀为-round1至-round3。历史SC/五脚本基线与Cycle文档保持；无字体安装、系统设置修改或删除操作。

首次native保存将theme XML从作者序列化改写为原生序列化，raw主题hash变化而fontScheme hash与内容不变。oracle用已有独立native基线的PPTKit theme hash精确约束规范化结果，不任意忽略主题字节变化；换Office主题则要求真实链、name、raw hash与fontScheme均变化。原显式pPr在native编辑后被移除，frame/box/边距保持。不同阶段的行距/位置与Latin字宽可能变化，不能把相邻family保持当作位置不变。

## 独立核验

[oracle](../../scripts/research/oracles/real-font-mixed.py)沿presentation关系定位真实六页顺序和Label/Probe，沿layout/master/theme解析主题；源PPTX SHA固定核验。逐页检查完整text、run、lang/altLang、三槽、48pt、精确局部b/i、frame/paragraph/theme和raw run attributes。四阶段的每一轮raw records与阶段初始严格一致，无需要忽略的额外dirty变动。

PDF逐页核尺寸、标签、实际PS、逐字符font归属、可见ink/无glyph0和48pt body matrices；只有16pt标签与正文共用ArialMT时按矩阵字号排除标签。实际Arial-ItalicMT的matrix无合成shear。相邻本脚本与Georgia的PDF glyph序列相对原始source PDF保持；每阶段3轮的6页PNG与该阶段初始逐页逐像素一致，共102页source/阶段/轮次记录。已实际查看6行×5列contact sheet，原始/插入/局部粗体/局部斜体/主题结果可见。

PDF空格可能仅有advance而不出现在ToUnicode；Arabic仍有一个源fatha未映射到提取文字，Hebrew使用presentation forms。原XML与Parser完整源文字另核，不全局删除空格/mark或规范化来制造相同。相邻glyph序列保持只证明本轮操作没有改变native基线的这些glyph，不认证该基线已满足完整bidi或所有mark轮廓正确。

[runner](../../scripts/inspect-real-font-mixed.mjs)另将source+16原生PPTX分别通过Parser v3重投影，逐页验证3/5run、source language、三literal槽、48pt与局部b/i、不恢复map；17次no-op与输入bytes精确一致，规范化文档快照不变。没有把native实际fallback、语言推断或主题改写烘焙进新的source writer。

[72个实际破坏/原生反例](../../scripts/research/oracles/real-font-mixed-checks.py)全部拒绝：六页各删除X、改新语言、改Latin family、删除Greek、破坏本脚本文字、丢失局部bold、bold扩散到本脚本、italic扩散到Greek、字体资源命名冲突、全部本脚本资源改名、移除本脚本glyph，共66项；再加两份真实失败选区的source/PDF四项、旧主题、实际italic冒充bold。

负例先发现RTL正文和空格可能使用不同的同名font资源：仅改首个空格资源，非空白trace仍命中原font，而改名又与现有Georgia重合。检查器因此同时核对BaseFont、FontDescriptor/FontName和trace，且新增修改全部本脚本资源的独立反例。资源重命名是元数据一致性破坏，不冒充已替换嵌入font outlines的实验。

## 复查与剩余范围

本地artifacts：`.tmp/real-font-mixed/run-7fvzgJ`，含mixed-fonts、四阶段及各三轮PPTX/PDF、两份失败原生选择样本、mixed-inspection.json、mixed-contact-sheet.png。完整[manifest](real-font-mixed-evidence.json)记录SHA、raw source/主题、PS/glyph/matrix/raster、警告、72反例及17次Parser精确no-op。字体二进制和PDF仅留本机。

```sh
rtk proxy node scripts/prepare-real-font-mixed.mjs
# 新run须另行通过GUI操作与导出，不能复用旧PDF冒充新轮次
rtk proxy node scripts/inspect-real-font-mixed.mjs .tmp/real-font-mixed/run-7fvzgJ
rtk proxy pnpm test:script-fonts
```

完整同段组合与局部编辑见[后续证据](real-font-combined.md)；接续native表格fontRef、独立notes主题、mixed-bidi/marks/emoji。此六组配对证据不替代上述范围，也不将PowerPoint native结果扩展为跨应用字体保证。
