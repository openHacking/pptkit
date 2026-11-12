# 五脚本原生编辑、样式与主题

2026-10-26：继 [SC 切片](real-font-edits.md)，完成 fresh 作者 PPTX 中 TC / Japanese / Korean / Arabic / Hebrew 五页的追加文字、全文粗体、全文斜体、恢复regular后换Office主题。每阶段初始保存/PDF后各三轮另存、关闭、准确路径重开、导出PDF。完整 author-map 编辑父任务仍开放：本轮未覆盖作者原始混合脚本段在fresh导出时的跨family物化及后续局部样式编辑，也不认证完整bidi/mark-shaping。

## 原生操作

应用为 Microsoft PowerPoint for Mac 16.104 (16.104.25121423)，所有GUI操作通过CUA。源为原fresh deck的未编辑 `fresh-round3.pptx`，新副本前缀 `fresh-edit-multi` / `fresh-bold-multi` / `fresh-italic-multi` / `fresh-theme-multi`，轮次加 `-round1` 至 `-round3`。没有修改历史SC证据、基线或Cycle文档，没有删除文字、安装字体或更改系统设置。

在48pt正文的逻辑末尾追加本脚本文字与Ω；从独立副本选择该页全部正文切换bold、关闭bold再切换italic，最后关闭italic并在设计图库应用Office主题。原始direct控制页、SC页、Arabic ligature页及Latin页不编辑，共11页旁路控制。所有内容源自作者API直接生成的fresh deck，未patch XML制造native结果。

| 页 / 追加 | native保存的run与语言 | regular / bold实际PS |
| --- | --- | --- |
| map-tc / `返Ω` | 原文zh-TW；返zh-TW/altLang en-US；Ω el-GR/altLang zh-TW | STSongti-TC-Regular / STSongti-TC-Bold |
| map-ja / `返Ω` | 原文ja；返zh-CN/altLang en-US；Ω el-GR/altLang ja | HiraMinProN-W3 / HiraMinProN-W6 |
| map-ko / `글Ω` | 原文ko；글 ko-KR/altLang en-US；Ω el-GR/altLang ko | AppleSDGothicNeo-Regular / AppleSDGothicNeo-Bold |
| map-ar / `مΩ` | 本脚本文字合为一个ar run；Ω为el-GR | GeezaPro / GeezaPro-Bold；Ω实际Cambria / Cambria-Bold |
| map-he / `םΩ` | 本脚本文字合为一个he run；Ω为el-GR | ArialHebrew / ArialHebrew-Bold；Ω实际Cambria / Cambria-Bold |

所有run的三个literal槽仍分别保持该页物化family。日文新汉字被native标为zh-CN、韩文新run为ko-KR，均按实际source记录，不重写为作者原语言。CJK/Hangul三页的Ω继续使用该页family，RTL两页则发生实际font fallback；没有一页恢复作者unused `Grek: Georgia` 规则。

五种正文family的italic均保留regular PS，以PDF text matrix `c/a=67.96875/200` 斜切；RTL页的Ω实际使用Cambria-Italic，其matrix没有额外斜切。粗体实际face、合成斜切和实际italic fallback分别验证，不以OOXML b/i标志或family声明冒充实际字型。

## 字体来源与回退证据

[runner](../../scripts/inspect-real-font-multi.mjs) 从原安装字体文件枚举face，以实际PS名称找到regular/bold候选，记录file/index/version/SHA和原face的Ω HarfBuzz结果。GeezaPro与ArialHebrew的regular/bold均返回glyph0；TC/ja/ko及Cambria候选覆盖Ω。

PowerPoint自带的Cambria候选位于其 `Contents/Resources/DFonts`：Cambria.ttc、Cambriab.ttf、Cambriai.ttf。系统 `fc-match Cambria` 实际返回Times New Roman，日文Bold查询也可能返回W3，而native实际PDF是W6。因此系统查询仅作为hint记录，不能替代PDF actual face。候选源文件与PDF同PS名字、coverage吻合；没有比较嵌入subset的轮廓指纹，不声称证明PDF字体与候选源文件字节完全同一。字体二进制保留在本机，不提交仓库。

## 保存、渲染与投影核验

[独立oracle](../../scripts/research/oracles/real-font-multi.py)复用具名ZIP关系读取，按presentation真实顺序定位Label/Probe，沿slide→layout→master→theme读取主题。逐页验证正文完整逻辑文字、run分裂/合并、lang/altLang、48pt、三个槽、b/i、frame和paragraph记录。五个编辑段的显式pPr被native移除；frame/box/边距保持。首次重开为map-he的Ω run新增dirty=0，只有这一处具名变化被显式接受并保留raw records，其余字段按阶段严格核验。

Office主题确实改变theme与fontScheme hash，literal槽与实际正文family保持。不同主题可能改变行距/位置，不能推导整页外观不变。三轮PDF分别与该阶段初始PDF逐页逐像素一致，256页记录全部保留；非换主题三个阶段的11旁路页还与原fresh基线核精确栅格。已实际查看5脚本×4阶段contact sheet，bold与italic可见，源文字与Greek均有painted ink。

PDF检查逐页核尺寸、标签顺序、实际PS、每字符所属font、非空glyph/无glyph0、可见ink和按resource区分的named body matrices。正文与fallback各自检查，不将一个face的shear套用到另一个face。

**PDF ToUnicode不等于完整源文字：** Arabic trace为 `السَّلامُمΩ`，少映射了source `السَّلَامُمΩ` 中的一个fatha；Hebrew trace为 `שָׁלוֹםםΩ`，使用presentation forms。oracle分别记录完整XML文字和具名PDF trace，固定核对这些观察结果，未全局删除mark/折叠字符来制造“相同”。因此不把PDF提取通过认作所有mark正确显示或mixed-bidi正确；mark轮廓与bidi顺序仍需独立任务。XML/Parser保留的完整文字、实际glyph/可见ink、阶段内栅格稳定是本轮各自独立的结论。

Parser `script-families-v3` 导入16份native PPTX，逐页核文本/run/lang/槽/48pt/b/i与不存在的scriptFontFamilies；16次no-op保存与输入bytes完全一致，规范化投影快照不变。实际Cambria fallback不烘焙进source family，不把字体或语言writer权限扩大。

## 负例与复查

[43项实际破坏控制](../../scripts/research/oracles/real-font-multi-checks.py)全部拒绝。五页分别修改正文、语言、强行恢复Georgia、合并run、删除bold、删除italic、清除PDF正文、移除PDF斜切，共40项；再检查旧主题冒充换主题、粗体实际face冒充regular、真实PDF资源的Cambria改名Georgia。不是只改expected，正式native artifacts保持不变。

本地artifacts为 `.tmp/real-font-fresh/run-81jnGm`，含16份PPTX/PDF、`multi-inspection.json`、`multi-contact-sheet.png`。完整[manifest](real-font-multi-evidence.json)含source/file hash、每页run/frame/paragraph/theme、glyph/font/matrix/raster、字体候选与系统hint、43项负例和16次Parser精确no-op。

```sh
rtk proxy node scripts/inspect-real-font-multi.mjs .tmp/real-font-fresh/run-81jnGm
```

该命令复核已由GUI产生的文件，不自动执行native编辑，不复制旧PDF冒充新轮次。下一步仍需作者原始混合脚本段的fresh物化与局部编辑、混合bidi/marks的视觉参照，以及native表格fontRef与独立notes主题验收。

后续：[原始配对混合段局部编辑](real-font-mixed.md)已补六组原始三family物化与Latin局部操作各三轮；完整同段七脚本组合仍开放。
