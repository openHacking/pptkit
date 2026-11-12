# 原生表格字体：Office主题与引用响应

2026-10-26：完成 IMP-05.1.3.3.2.6.1.4 的具名主题切片，并汇总 .6.1.5 的追加、样式、主题候选复验。四组各三轮实际PowerPoint保存/关闭/准确路径重开/PDF，共216个主题后页面记录，另有72页主题前基线。独立notes、Windows、完整字体/表格/round-trip父项与source字体/style writer仍开放。

## 输入和原生操作

输入来自[样式实验](native-table-font-styles.md)的regular第三轮，保留原始证据。reference27页覆盖major/minor引用和generic空槽的显式集合；filled9页覆盖generic非空的显式集合。另各复制一份local组，用[准备脚本](../../scripts/prepare-native-table-font-themes.mjs)在原有15目标的两个run中明确patch Latin=Arial、EA/CS=Arial Unicode MS。独立reader逐part核验只有这些局部字体声明变化；这是source构造的局部覆盖控制，不是PowerPoint原生设置局部字体，也不是fresh作者往返。

四组均在PowerPoint for Mac16.104（16.104.25121423）实际打开，通过Design主题库选择Office，观察“所有幻灯片使用Office”状态，再另存、关闭、准确路径重开并导出PDF。第二、第三轮各再次另存/关闭/重开/PDF。local两组的主题前PDF也实际从PowerPoint导出；另两组主题前PDF复用已验证的regular第三轮。Mac锁屏时暂停界面操作，用户解锁后继续。未将headless修改计为原生主题操作。

## 源结构和选择规则

四组Office主题字节一致，SHA256为`1f98605413696c4fe4029a7b7df696261233d73d8fd06413d3990abfe51ae216`。名字为`Office 主题`后跟两个U+200B；major Latin为等线 Light，minor Latin为等线，generic EA/CS均为空，保留Office各script补充字典。

所有slide XML、tableStyles、notes及资源原字节保持，包括原文、Ω、language、run拓扑、regular flags及局部字体。Office新增layout2–11及关系共20part，未删除part；全部slide改连layout7，再连master1和theme1。master/layout、presentation的master ID/defaultTextStyle及具名docProps字段被应用重写，不能只检查theme1或因为字体相同就宣称引用未变。

[专用oracle](../../scripts/research/oracles/native-table-font-themes.py)以此前[SC主题实验](real-font-edits.md)的独立Office文件为资产参照；核验其完整文件SHA与既有manifest。theme字节严格相等；master/layout及关系的完整树只归一化通过格式、唯一性和共享关系检查的63个a16 creation GUID、12个p14 creation uint、两类共享field GUID，其余全部比较。三个当前native轮次的身份值还必须保持。没有把新layout整体加入忽略名单；重复ContentType、语义变化、坏/重复身份和错误共享field均会拒绝。

声明读取跟随实际关系：fontRef从新Office theme取完整可准入map；显式集合继续使用tableStyles自身原八条规则；local三槽覆盖让run不再携带继承script map。docProps的used-font inventory为等线、等线 Light、Arial，**连local组也是如此，不能把此清单当作实际使用字体证据**。

## PDF实际观察

| 控制 | minor引用 | major引用 |
| --- | --- | --- |
| Latin / SC / Greek | DengXian-Regular | DengXian-Light |
| TC | PMingLiU | PMingLiU |
| Japanese | YuGothic-Regular | YuGothic-Light |
| 短ko | Batang | Batang |
| ko-KR | MalgunGothic | MalgunGothic |
| Arabic / Hebrew | ArialMT | TimesNewRomanPSMT |
| 追加Ω | DengXian-Regular | DengXian-Light |

空generic显式集合的原正文仍为Georgia / SimSun / Batang / ArialMT，Ω仍为Georgia；filled显式集合的EA/CS正文仍为ArialUnicodeMS，Latin/Greek与Ω仍为Georgia。local目标原文均为ArialUnicodeMS、Ω为ArialMT，主题前后保持。非local控制分别按对应组行为响应。

每页label仍为ArialMT、Unicode/GID/x保持，但基线整体上移1.2pt，因此不能宣称整页主题前后栅格一致。local引用目标的正文也上移1.2pt；local显式集合及其他显式集合正文位置和y=60..260区域栅格精确保持。背景/边框path保持。位移原因尚未独立证明，不推断为某个endPara/default段落属性导致。三轮主题后的全部页面栅格精确一致；已查看四组contact sheet。

PDF subset GID不是固定字符身份：reference的collection-ar/he ArialMT正文、reference-local所有local目标的ArialMT Ω出现具名-3重编号，其他保留正文GID不变。oracle核对这些精确差值、原font/Unicode/x和具名y响应；manifest同时保留before/after原始GID。该观察不证明PDF嵌入subset字节或mark轮廓相同。

source文字不变，但Office引用的SC ToUnicode以U+2EE3开头、日本语以U+2FBB开头；Arabic按minor/major有不同presentation form，Hebrew为`שָׁלוֹם`；local ArialUnicodeMS Arabic提取为`اﻟﺴ*ﻼَمُ`。这些原样记录，不能用Unicode归一化掩盖。PDF font resource/descriptor、glyph0、逐非空trace墨迹、页序和栅格独立校验；不以可提取文字代替完整bidi/组合mark正确性。选择/外观边界沿用[ADR0053](../architecture/decisions/0053-declared-fonts-and-native-appearance.md)，原无编辑基线的16差异仍保留。

## 候选复验与回归

[runner](../../scripts/inspect-native-table-font-themes.mjs)验证4份baseline及12份Office输入：16次Node精确no-op，四份第三轮browser global完整document/report与byte精确no-op、零网络；全部实际输入通过独立DOM/XPath核源字体值、目标和resolvedFrom，关系/集合由Python独立追踪。十二份Office输入的source字体写回均blocked并restore，再各做三轮Ω→ΩX→ΩXX→ΩXXX，36个scalar候选重投影相等、字符外字节精确保留。**这些候选未native重开，不计入216页原生主题证据。**

[484个实际破坏反例](../../scripts/research/oracles/native-table-font-themes-checks.py)分别针对72页的text/lang/local/fontStyle绑定、仍可提取的遮白正文或Ω、font resource，以及四组theme槽/关系、Office layout语义、GUID格式/重复/共享、uint范围、重复ContentType、无关notes、页序、基线和后轮栅格；原生baseline不覆盖。

[source-only回归](../../scripts/table-font-theme-parity.test.mjs)独立构造filled×真实theme关系重绑定×local override八种输入，共144源页面，Node/browser ESM/global各三轮，72个scalar路径和72次no-op。rebound theme使用新part并保留旧unused theme，能抓住硬编码theme1；独立DOM跟踪实际slide/layout/master/theme/tableStyles及来源XPath，坏来源值被拒绝，source字体混改阻断后restore。此回归不需要安装专有字体，不冒充原生主题应用；已接入test:import。

## 汇总与复查

.6.1.5按本组具名范围收口：追加、样式、主题分别为8/20/16条输入验证路径，共44条Node精确no-op和12条browser对照；scalar候选18+54+36=108，实际反例112+380+484=976。这里按实验路径计数，含复用baseline，不是44份互不重复的文件。648页原生操作后记录（108追加+324样式+216主题）与72页本批主题前baseline分开计数。不关闭.6.1、独立notes或完整父任务，也不授权PPTKit源样式写回。

本机artifacts：`.tmp/native-table-font-themes/run-X74HnC`，四个子目录各含baseline、expected.json、office-round1..3 PPTX/PDF、theme-inspection.json和theme-contact-sheet.png。[manifest](native-table-font-theme-evidence.json)记录应用版本、hash、实际声明/PS/glyph/位移/栅格与反例结果；专有字体/PDF/PNG不提交。

```sh
rtk proxy node scripts/inspect-native-table-font-themes.mjs .tmp/native-table-font-themes/run-X74HnC
rtk proxy node --test scripts/table-font-theme-parity.test.mjs
rtk proxy node --test scripts/declared-slot-font-import-parity.test.mjs
rtk proxy node --test packages/pptx-parser/test/declared-slot-font-import.test.mjs
```

build/typecheck/lint及6组跨包、9组Parser回归通过，八输入主题source回归单独通过；未声称重跑完整workspace runtime套件。接续[.6.4.1独立notes字体基线](../planning/advanced-capability-todos.md#表格与备注原生字体的可领取任务)，再做notes实际操作，source主题编辑、mixed bidi、emoji和SmartArt/动画等任务仍按原队列。
