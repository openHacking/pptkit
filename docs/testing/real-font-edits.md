# 原生 SC 字体编辑、样式与主题证据

2026-10-26：在 [fresh 作者 PPTX](real-font-fresh.md) 的独立临时副本中，完成 SC map 页的追加文字、粗体、斜体、换主题四项操作。每项都有初始保存和三轮另存、关闭、重开、PDF 导出。此处只完成具名 SC 切片，其他脚本和 mixed-run 编辑仍未完成，IMP-05.1.3.3.2.6.2 保持开放。

## 操作与观察

应用：Microsoft PowerPoint for Mac 16.104 (16.104.25121423)。从 `fresh-round3.pptx` 另存编辑副本，通过 CUA 在第2页 `map-sc` 的48pt正文内操作，没有替换或删除既有文字，没有修改系统字体设置。历史 Cycle 窗口与基线文件保留。

| 阶段 / 文件前缀 | 原生操作 | 实际结果 |
| --- | --- | --- |
| `fresh-edit-sc` | 在正文末尾追加 `Ω` | 正文为 `骨直令返Ω`；2个run分别为 `zh-CN` 和 `el-GR`，后者带 `altLang=zh-CN`；所有槽仍为 Hiragino Sans GB，PDF实际为W3。 |
| `fresh-bold-sc` | 选择正文全部文字，启用粗体 | 两个run都有 `b=1`，PDF实际为 HiraginoSansGB-W6，视觉与regular不同。 |
| `fresh-italic-sc` | 关闭粗体，启用倾斜 | 两个run都有 `i=1`；PDF仍为W3，正文text matrix带 `c/a=67.96875/200` 的斜切。证明synthetic shear，不认证独立italic face。 |
| `fresh-theme-sc` | 关闭倾斜，应用 Office 主题 | 实际slide→layout→master→theme链指向Office主题；主题与fontScheme hash均变化，major/minor Latin从Arial变为等线系列。正文literal槽和实际W3保持。 |

源作者map含 `Grek: Georgia`；materialization后只剩当前SC family的literal槽。追加的 `Ω` 实际使用W3，而非Georgia；Parser v3重投影也没有恢复 `scriptFontFamilies`。语言由PowerPoint赋为el-GR是本轮native编辑结果，未由PPTKit写入或改写。

原生操作还移除了该段原有的显式 `pPr`（100%行距/0段前后距等）。text frame、边距、位置保持，但段落改为依赖默认值。换主题后的页面栅格与原regular页不同，不能把literal字体不变描述为整页外观不变。每个阶段的三轮PDF分别与该阶段初始PDF比较，全部逐页逐像素相同；不要求不同操作阶段的页面相同。

追加run的第一次重开保存新增了 `dirty="0"`。独立检查明确验证这一处proofing标记变化，逐轮保留raw attributes；不把它包装成字节完全不变，也不全局忽略未知属性。其余run/lang/altLang/槽/b/i、frame、段落与theme记录按各阶段核对。其余15页的正文/lang/槽/样式保持；非换主题的三个阶段还与基线逐页核验了这15页的精确栅格。换主题是全deck操作，仅保证其余页的文本/字体声明和实际PS，并记录各阶段栅格，不声称其余页几何或外观不变。

## 验证与负例

[runner](../../scripts/inspect-real-font-edits.mjs) 调用独立 [ZIP/XML/PDF oracle](../../scripts/research/oracles/real-font-edit.py)，再由Parser `script-families-v3`逐一导入16份原生PPTX。每份检查16页实际标签顺序、正文run、语言、三个family槽、48pt、bold/italic与未恢复的map，随后no-op保存必须与输入bytes完全相同，作者投影快照不变。

ZIP按presentation关系读取真实slide顺序，theme沿实际layout/master关系解析。PDF逐页检查尺寸、实际标签、正文PS face、非空glyph/无glyph0、可见ink、编辑页完整Unicode序列和named body text matrices。三轮记录全部页面栅格hash并与阶段初始PNG bytes比较；已实际查看4阶段contact sheet，regular/粗体/斜切/换主题结果可见。此PDF矩阵读取仅适用于具名PowerPoint输出，不是通用PDF解释器。

[10个实际破坏控制](../../scripts/research/oracles/real-font-edit-checks.py) 均被拒绝：改正文、改语言、强行恢复Georgia、合并run、去掉bold、去掉italic、用旧主题替换新主题、将粗体PDF冒充regular、从真实PDF内容流移除斜切、从PDF移除追加glyph。没有只改expected数据构造负例，正式原生文件保持不变。

完整本地artifacts在 `.tmp/real-font-fresh/run-81jnGm`，含各前缀的 `.pptx` / `.pdf` 与 `-round1` 至 `-round3`。检查输出为 `edit-inspection.json`、`edit-contact-sheet.png`；[manifest](real-font-edit-evidence.json) 记录4×4份文件SHA、16页每份的run/frame/paragraph/theme/实际PS/raster，以及16次Parser精确no-op。源PPTX hash与既有字体文件SHA重新验证；未提交字体或PDF嵌入字体二进制。

```sh
rtk proxy node scripts/inspect-real-font-edits.mjs .tmp/real-font-fresh/run-81jnGm
```

此命令复核已由GUI产生的artifact，不自动执行原生编辑。没有以CLI生成副本或复制PDF代替原生保存/重开。结果不扩展到TC/Japanese/Korean/Arabic/Hebrew编辑、跨run bidi/emoji、表格/notes，未扩展任何source字体或样式writer权限。

后续：[五脚本原生全文操作](real-font-multi.md)已补TC/ja/ko/ar/he的追加、bold/italic和Office主题各三轮；原始混合段跨family与局部编辑仍开放。
