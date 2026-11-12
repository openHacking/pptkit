# 独立 notes 字体：原生无编辑基线

2026-10-26：完成 IMP-05.1.3.3.2.6.4.1 的具名基线。两套独立备注主题各15控制，PowerPoint各三轮另存、关闭、准确路径重开并打印真正备注页PDF，共90个重开页面记录，另30页首次打开基线。24个声明选择与实际字体匹配，6个差异保留；本批没有原生改字、样式或换主题，.6.4.2及完整notes/字体/round-trip父项继续开放。

## 输入、关系和操作

[具名source assembler](../../tests/fixtures/native-notes-fonts.mjs)先用作者API生成有效基础PPTX，再patch备注master、body几何、字体声明、标签及独立theme关系。保留author-base.pptx及hash，不能称为fresh作者往返。A/B是两份独立演示文稿，每份各一个notes master；不是在一个PowerPoint包中验证多notes-master编辑。

两组均覆盖Latin、SC、TC、Japanese、短ko、ko-KR、Arabic、Hebrew、Greek，另有TC/Arabic/Hebrew三槽local覆盖、TC局部空EA、TC局部坏EA `+mn-unknown`、TC显式major引用。notes master的body idx=99，local body idx=7，但type均为body；全局notesStyle为18pt，type匹配的placeholder list为40pt。master提示文字`MASTER PROMPT NEVER INHERIT`不进入Core备注或PDF正文。装饰标签与备注正文分别核验。

slide使用theme1，slide正文固定Arial Unicode MS；notes master实际关系指向`independent-notes-theme.xml`。保留unused theme2作为控制，避免reader靠文件名猜主题。A的补充字典依次为Arial / Hiragino Sans GB / Songti TC / Hiragino Mincho ProN / Apple SD Gothic Neo / Geeza Pro / Arial Hebrew / Georgia；B改为Georgia / Songti TC / Hiragino Sans GB / Hiragino Sans GB / Arial Unicode MS（Hang/Arab/Hebr）/ Arial（Greek）。两组major/minor generic Latin取对应字典Latin，EA/CS为空；local三个槽明确设置Arial / Arial Unicode MS / Arial Unicode MS。

PowerPoint for Mac16.104（16.104.25121423）实际打开后，通过打印对话框选择“注释”、纵向、A4、颜色、“根据纸张调整大小”，PDF菜单保存本机。每轮另存PPTX、关闭、按准确文件路径重开，再重复备注页打印。这里使用真实Notes打印输出，普通slide导出不代替它。已实际查看B组原生备注页视图和两组PDF contact sheet。

## 原生保存的具名变化

首次保存删除`independent-notes-theme.xml`，将其完整主题语义迁入theme2，重绑notes master；原来的unused theme2内容被替换。slide theme1、layout/master主题链与notes域仍分离，字典、空槽、局部覆盖和坏EA字符串保留。原备注/幻灯片文字及语言不变。

独立[Python oracle](../../scripts/research/oracles/native-notes-fonts.py)追踪实际presentation→slide→notes→notesMaster→theme与notes回链、slide→layout→master→theme，并核对全部关系边、唯一ID与包part库存。完整notes/slide/theme/master/layout/tableStyles树只允许以下已观察变化：notes placeholder补入具名locks、段落indent的-0→0、Latin备注的冗余en-US endPara删除、notesStyle中显式黑色paint删除。没有将整个body、master或theme忽略。三个重开轮次的完整所检域树相同。

native还重写了docProps、viewProps，并添加presentation/presProps扩展。manifest保存这些辅助part的源/原生XML和hash，**不把这项字体基线称为完整ZIP字节保真**。Core在显式黑色被删后可显示默认主题色表示；native视觉保持不等于源语义原样保持。精确no-op和scalar候选字节约束是另外的PPTKit保存检查。

## 实际字体与差异

| 控制 | A实际PDF face | B实际PDF face |
| --- | --- | --- |
| Latin | ArialMT | Georgia |
| SC | HiraginoSansGB-W3 | STSongti-TC-Regular |
| TC / major-TC | STSongti-TC-Regular | HiraginoSansGB-W3 |
| Japanese | HiraMinProN-W3 | HiraginoSansGB-W3 |
| 短ko | MalgunGothic | Batang |
| ko-KR | AppleSDGothicNeo-Regular | ArialUnicodeMS |
| Arabic | GeezaPro | ArialUnicodeMS |
| Hebrew | ArialHebrew | ArialUnicodeMS |
| Greek | Georgia | ArialMT |
| 三组local覆盖 | ArialUnicodeMS | ArialUnicodeMS |
| 空局部EA | MicrosoftYaHei | SimSun |
| 坏局部EA | MicrosoftYaHei | MicrosoftYaHei |

两组各12匹配、3差异：短ko实际字体不等于声明Hang；空/坏EA的v3声明选择缺少EA规则，Core的预览generic fallback与native不一致。它们仍是差异，不能把actual face写回source map/lang，或把未解析槽位的预览fallback称作确定的原生选择。坏槽和空槽尤其不能合并：B组实测字体不同。规则与外观分离沿用[ADR0053](../architecture/decisions/0053-declared-fonts-and-native-appearance.md)。

source正文40pt、两类标签16pt，在本次打印PDF的texttrace分别为约41、17；纸面595×842pt。只记录该具名打印结果，不推导通用缩放公式。local页标识因连字符分成多条PDF trace，oracle核验完整分段与字符顺序，不能把trace数量当source run拓扑。

Arabic/Hebrew源mark保留；PDF ToUnicode存在presentation forms，GeezaPro提取为`السَّلامُ`，ArialHebrew为`שָׁלוֹם`。ArialUnicodeMS的Arabic在A中为`اﻟﺴ(ﻼَمُ`，B中为`اﻟﺴ&ﻼَمُ`；这些提取差异原样保留，不用Unicode归一化消除。所有非空trace核实际font resource/descriptor、glyph0和可见墨迹。三轮90页的完整glyph/font/坐标trace及整页栅格均与各组首次打开PDF精确一致；这仍不证明通用bidi、mark轮廓或嵌入subset字节身份。

**本具名source控制的slide-image placeholder在PowerPoint备注页及PDF中仍为空白。** 正文和标签可见，足以建立字体基线；完整备注页缩略图、fields/artwork与几何仍未认证。该限制没有通过修改期望图掩盖，已拆为[IMP-03.3.2.1缩略图空白定界与修复](../planning/advanced-capability-todos.md)。后续实际备注编辑必须继续区分正文和页级内容。

## 候选复验与反例

[runner](../../scripts/inspect-native-notes-fonts.mjs)核两份source和六份实际native输入：8次Node精确no-op；两个第三轮browser global完整document/report与bytes相等、零网络；全部实际输入通过独立DOM/XPath核script字体来源、属性目标和resolvedFrom，40pt的来源与body type依赖明确指向idx7/99。

六份native输入分别尝试source字体写回，全部blocked并restore；随后全15页备注各三轮追加X，合计18个scalar候选。重投影完整相等，除各notes正文字符外所有字节精确保留，slide/master/theme/标签/未知扩展不变。**这些候选未native重开，不能计入90页或冒充.6.4.2原生改字。**

[226个真实ZIP/PDF破坏反例](../../scripts/research/oracles/native-notes-fonts-checks.py)全部拒绝，两组各113：每页破坏正文、language、local字体、local type、页标识，或遮白仍可提取的正文、改font resource；另破坏实际notes主题、借用slide主题、master type/字体尺寸、实际master关系、slide邻居、页序及后轮栅格。错误关系不能回退到presentation列出的master。输入和反例使用不同临时文件，native证据不覆盖。

[新增source回归](../../scripts/notes-native-font-parity.test.mjs)在无宿主专有字体条件下构造同样两组/30源页。Node、browser ESM/global各三轮，共18个scalar路径与18次精确no-op；完整document/report/输出bytes一致，字符外字节精确，字体写回阻断后恢复。独立DOM验证实际notes主题、7/99 type-vs-index和来源路径；错误type即使idx相等也只取global18pt。缺实际theme时源引用unsupported，缺实际master时没有可继承的字体声明；二者均不借slide字体，缺master不污染其他备注。该回归接入test:import，不冒充PowerPoint执行。

## 复查与接续

artifacts：`.tmp/native-notes-fonts/run-haqFDP`，A/B各有author-base/source/expected、source-notes.pdf、native-round1..3 PPTX与-notes.pdf、notes-inspection.json及contact sheet；早期探索保留于run-THJF8c，不计入本次页数。源字体环境文件/版本/hash沿既有真实字体记录复核；native fallback仍按PDF实际face记录，不声称嵌入subset与本地字体字节相同。[manifest](native-notes-font-evidence.json)含应用版本、全部输入/输出hash、关系/源树、实际字体/glyph/ToUnicode/墨迹/栅格、辅助规范化、反例与候选结果。专有字体、PDF和PNG仅留本机。

```sh
rtk proxy node scripts/inspect-native-notes-fonts.mjs .tmp/native-notes-fonts/run-haqFDP
rtk proxy node --test scripts/notes-native-font-parity.test.mjs
rtk proxy node --test scripts/notes-style-import-parity.test.mjs
rtk proxy node --test packages/pptx-parser/test/notes-style-import.test.mjs packages/pptx-parser/test/declared-slot-font-import.test.mjs
```

build/typecheck/lint、22组Parser、3组既有notes跨包及新source回归通过；未声称重跑整个workspace runtime套件。下一步[.6.4.2备注原生改字、样式与主题操作](../planning/advanced-capability-todos.md#表格与备注原生字体的可领取任务)，仍需每阶段三轮重开和独立证据。source字体/style writer、完整notes页、字体和导入/round-trip父项均不因本批放行。
