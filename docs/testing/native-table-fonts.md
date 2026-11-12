# 原生表格 fontRef 与显式集合字体对照

2026-10-26：完成36个具名source表格控制及各组三轮PowerPoint保存重开。Parser v3的声明选择全部保持；原生实际字体有20匹配、16差异，明确保留。此批没有原生改字/样式/主题编辑，也没有独立notes主题验收，IMP-05.1.3.3.2.6.1及父项仍开放。

## 源控制与范围

[source assembler](../../tests/fixtures/native-table-fonts.mjs)先通过作者API/Exporter生成完整有效包，再替换具名tableStyles、theme字体集合和cell字体引用，供source导入实验。它是有意修改source声明的控制生成器，不是fresh作者map round-trip，更不是通用PPTX合并器。每页Label为Arial16pt，ProbeTable为1行1列、48pt正文，run只声明lang/size、没有local Latin/EA/CS字体，以免局部属性掩盖table selector。

27页组含minor fontRef、major fontRef、显式font collection各9场景：Latin、SC、TC、ja、ko、ko-KR、Arabic、Hebrew、Greek。theme major/minor的EA/CS泛用槽为空，但各有8条具体脚本声明；显式集合也有8条脚本声明，且SC/TC family故意与主题不同。补充9页只把显式集合的EA/CS泛用槽改为Arial Unicode MS，其他脚本声明相同，用来隔离空槽回退。

Microsoft的[Fonts文档](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.fonts?view=openxml-3.0.1)将table tcTxStyle中的a:font列为FontCollection，并列出latin/ea/cs/补充font子元素；[SupplementalFont.Script文档](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.supplementalfont.script?view=openxml-3.0.1)描述其script/language属性。这支持解析声明结构，但不能替代具体PowerPoint版本的选择行为证明。本批没有用SDK说明推定实际字体，也没有通过删除声明迎合应用表现。

## 实际字体结果

应用为Microsoft PowerPoint for Mac16.104 (16.104.25121423)，从当前应用Info.plist读取并记录。source文件由CUA打开、原生导出PDF，然后各组三轮另存新副本、关闭、准确路径重开和导出PDF。历史font/native/Cycle文档保持，未删除文字、安装字体或修改系统设置。

| 组 | 声明选择与native一致 | 观察到的差异 |
| --- | --- | --- |
| minor fontRef | 8/9 | ko实际MalgunGothic；ko-KR实际AppleSDGothicNeo-Regular，后者与声明一致。 |
| major fontRef | 8/9 | ko实际Batang；ko-KR实际AppleSDGothicNeo-Regular。major/minor的TC/ja差异按各自主题集合生效。 |
| 显式集合，EA/CS空 | 2/9 | SC/TC/ja实际SimSun，ko及ko-KR实际Batang，Arabic/Hebrew实际ArialMT。 |
| 显式集合，EA/CS为Arial Unicode MS | 2/9 | 七个EA/CS场景均实际ArialUnicodeMS，未使用集合内对应脚本声明的family。 |

两类显式集合的Latin/Greek恰好与其泛用Latin face同为Georgia，匹配不证明native使用了其补充脚本规则。非空泛用槽对照也显示七个差异，因此不能把全部差异归为缺省槽、字体未安装或缺glyph。

脚本声明在三轮原生保存后仍留在XML：两组的显式集合各保留8条、theme保留16条。Parser v3仍按原source合同运输这些声明并产生相同选择；原生应用在本组数据中走另一条实际字体路径。该差异是外观/interop证据，不是本轮改变v3/default/v1/v2兼容合同的授权。尚需明确source规则运输与外观准入策略，再决定是否及如何新增诊断或独立profile；不把native实际替代烘焙成源family/lang。

## 独立检查

[oracle](../../scripts/research/oracles/native-table-fonts.py)按presentation真实关系读取页序、Label与ProbeTable，沿实际layout/master/theme链和presentation→tableStyles关系找到selector。逐页核styleId、fontRef的major/minor或显式集合、全部8条脚本声明、泛用槽、完整文字/lang/48pt和缺少local fonts。不能只按文件名或manifest顺序匹配。

原生保存有三处具名source变化：bodyPr默认margin/anchor属性被移除、pPr indent -0变0且属性重排、四条显式noFill边线声明被移除。oracle只接受这三个明确变换，其余selected source records与source严格比较。没有将它们泛化为任意格式归一化；没有凭栅格相同宣称源字节或边框语义相同。

PDF按实际PS分别核对具名应用选择与source选择：20/36匹配、16差异；检查标签顺序、页尺寸、真实glyph/无glyph0、每个非空白字体span可见dark ink，比较三轮逐页PNG与该组初始PDF，144页source/轮次记录均稳定。已实际查看27页和9页contact sheet：字体替代可见，尤其filled collection的RTL字形/mark布局与参考family不同。可见ink、无glyph0和保存稳定不等于完整shaping/mark正确性；Arabic/Hebrew ToUnicode与source文字分别记录，不全局删mark求相同。

Office应用内MalgunGothic、Batang、SimSun候选记录file/index/version/SHA及HarfBuzz cmap样本，均有非零glyph；系统fc-match结果只作为hint，不能替代实际PDF。与PDF同PS名的本地候选不认证嵌入subset与源字节完全一致。字体二进制、PDF与PNG均仅留本机。

## Parser与负例

[runner](../../scripts/inspect-native-table-fonts.mjs)导入两组source+三轮共8份PPTX，逐页复核style/scriptFontFamilies/声明选择和每cell8条mapped provenance；8次no-op与各自输入bytes完全一致，canonical快照不变。因native选择不同而把source map替换成fallback的做法不会通过此检查。

新增[跨包回归](../../scripts/declared-slot-font-import-parity.test.mjs)使用独立DOM读取36个source的真实style/theme/collection链，核本地run无font覆盖、声明map/规则source定位、v3运输与精确no-op。无需CI安装字体，现有三个v3跨包组继续通过。

[79项实际ZIP/PDF负例](../../scripts/research/oracles/native-table-fonts-checks.py)全部拒绝：27页组64项、filled组15项，覆盖文本、有效ko/ko-KR互换等语言变化、styleId绑定、source脚本声明、异常native归一化、仍可提取但被白色遮盖的文字、实际字体资源冲突和PDF换页。正式native文件不变。

## 复查与接续

artifacts：`.tmp/native-table-fonts/run-IQl6QY`（27页）、`.tmp/native-table-fonts/run-sF42dC`（9页），均含table-fonts及table-fonts-round1..3的PPTX/PDF、expected.json、table-fonts-inspection.json与contact sheet。[manifest](native-table-font-evidence.json)记录8份source/native包及对应PDF hash、selector/声明/实际字体/glyph/墨迹/raster、16差异、79负例、候选字体与8次Parser精确no-op。

```sh
rtk proxy node scripts/prepare-native-table-fonts.mjs
rtk proxy node scripts/prepare-native-table-fonts.mjs --filled-collection
# 新run必须另行通过PowerPoint执行三轮真实重开/导出
rtk proxy node scripts/inspect-native-table-fonts.mjs .tmp/native-table-fonts/run-IQl6QY .tmp/native-table-fonts/run-sF42dC
rtk proxy node --test scripts/declared-slot-font-import-parity.test.mjs
```

本批只完成无编辑native基线及选择差异定界的一部分。继续.6.1：用这些对照明确reference/explicit collection source与appearance边界，补表格实际改字/样式/主题；再推进独立notes master/theme的原生字体与备注编辑。父项不勾选，source字体/style writer限制不变。

后续[15目标真实追加验证](native-table-font-edits.md)已有独立编辑oracle与三轮原生结果；本页基线证据与严格oracle保留原范围。样式/主题/notes及完整父项仍开放。
