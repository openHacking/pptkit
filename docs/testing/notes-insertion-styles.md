# 空备注插入样式修复与真实输入验证

2026-10-26：[默认fresh备注实验](authored-notes-pages.md)发现PowerPoint会删除空run，连同仅保存在该run上的12pt粗体一起丢失。原失败目录run-rM4cWd保留；本批在新目录 `.tmp/notes-insertion-styles/run-VdgfBJ` 验证修正后的共享[空段落序列化](../architecture/decisions/0055-empty-paragraph-insertion-properties.md)，没有覆盖旧文件或改写旧失败结论。

## 实现与自动化

全空段落将最后一个run的字符样式同时写入endParaRPr，原run仍保留。所有当前字符属性复用同一serializer；actions留在原run，不复制到段尾，不产生新关系。非空段落的结束序列化不变。多个空run仅有一个段落插入位置，最后run的样式为最终插入样式；原生不保留各空run身份的事实仍记录。

Exporter[聚焦测试](../../packages/pptx-exporter/test/notes.test.mjs)新增回归先失败后通过，覆盖字号、字体三个槽、bold/italic/underline/strike、语言、颜色/alpha、多个空run、动作只生成一次、非空段落与text/shape/table共用路径。[跨消费者回归](../../scripts/empty-insertion-parity.test.mjs)以明确的独立XML变体删除空run，检查Node、ESM和global各三轮fresh投影、完整notes/生成bytes相等及原字节no-op；已接test:import。该变体不冒称原生操作。

## 原生控制与操作

[公共作者fixture](../../tests/fixtures/notes-insertion-styles.mjs)和[准备器](../../scripts/prepare-notes-insertion-styles.mjs)不做XML/ZIP补丁，四个控制为：

| 控制 | 最终空插入样式 |
| --- | --- |
| regular | Georgia14pt、fr-FR、直接颜色123456 |
| bold-italic | Arial12pt、bold+italic、en-US |
| script-map | Arial18pt、en-US；空字符串使用既有default-family规则，未把Latn映射Georgia猜成实际字符选择 |
| two-empty | 前一空run为Arial24pt bold，最后为Georgia16pt italic；最终插入使用后者 |

PowerPoint for Mac16.104中，首次Notes-print输出后，三轮另存为empty-round1/2/3、关闭、准确路径重开并打印纵向“注释”PDF。实际仍会删除所有空run，但完整endParaRPr字符属性树逐项与作者源相等；不是只检查size/bold，也不是只看空白栅格。

随后在第三轮原生文稿中逐页将caret置于空备注，先实际读取字体/字号、B/I和语言控件：Georgia14/法语；Arial12/B+I/英语；Arial18/英语；Georgia16/I/英语，均与源一致。没有操作字体控件，只输入同一具名单词Proof。另存为typed-round1后关闭/准确重开/打印，再做typed-round2/3。首轮重开导航曾被外部窗口变化打断，重新读取实际GoTo窗口并指定typed-round1准确路径，最终窗口和输出均核实；没有重复保存、覆盖空基线或把导航尝试计作新轮次。第三轮打印菜单需重新展开后完成PDF，文件存在和所属文稿都重新核对。

共四次实际输入、12页空阶段重开、12页输入后重开和4页首次打开输出。源fr-FR、颜色123456、全部字体槽、字号及bold/italic均由实际新run继承。输入后PowerPoint把属性放在新真实run上，省略冗余endParaRPr；oracle按这个实际具名改写分别检查，不要求非空阶段继续保留空段尾结构。

## 独立证据与反例

[ZIP/XML/PDF oracle](../../scripts/research/oracles/notes-insertion-styles.py)跟踪实际presentation顺序及slide↔notes关系，空阶段逐树核end属性，输入阶段逐树核新增run属性（只单列dirty属性的应用写入），文字严格等于Proof。实际PDF字体分别为Georgia、Arial-BoldItalicMT、ArialMT、Georgia-Italic；打印字号分别约14、12、19、17pt，源码14、12、18、16pt单列，不声称源字号与打印值全部相等。

Georgia-Italic的实际墨迹超出PDF trace的advance bbox。oracle按实际BT/ET字体资源绑定到嵌入TrueType，使用实际glyph ID读取轮廓边界，不用增加任意大遮罩通过；同一页两个ArialMT子集也按资源分开，不能只按PostScript名称选字体。在仅遮去实际新增墨迹轮廓的一像素抗锯齿边缘后，其余整页与空阶段精确相同；后两轮完整原始栅格、glyph trace和嵌入字体程序SHA与首轮精确一致。四页实际墨迹可见，已渲染并查看第三轮联系表。

[manifest](notes-insertion-styles-evidence.json)记录源和六轮实际文件的SHA、属性树、字体程序、字形墨迹边界与每页栅格。八个另存的实际PPTX破坏控制全部拒绝：空mark的size/lang/family/bold，输入run的family/color/italic/text。原run-rM4cWd的实际失败文件仍缺1200/bold，hash另保留。

[实际输入Parser复验](../../scripts/inspect-notes-insertion-styles.mjs)及[manifest](notes-insertion-parser-evidence.json)验证七份源/原生输入：7条精确no-op及混改阻断后恢复、完整snapshot/report保存。三个native空阶段仅有end属性，没有a:t标量owner，因此字符写入明确阻止；原作者源及三个typed输入各有一份实际headless字符候选，字符外所有payload精确保持，共4份。这四份未原生重开，不计入上述24页，也不授权source样式写回。

复验：

    rtk proxy node --test packages/pptx-exporter/test/notes.test.mjs scripts/empty-insertion-parity.test.mjs
    rtk proxy python3 scripts/research/oracles/notes-insertion-styles.py .tmp/notes-insertion-styles/run-VdgfBJ
    rtk proxy node scripts/inspect-notes-insertion-styles.mjs .tmp/notes-insertion-styles/run-VdgfBJ

完整build/typecheck/pnpm test已通过，import集成116项；最后增加的text/shape/table聚焦路径单独通过，无生产代码后续改动。日志 `.tmp/empty-insertion-{build,full-test,focused,parity,lint}.log` 和 `.tmp/notes-insertion-parser.log`。独立oracle的PPTX反例生成只修改negative-controls，不覆盖原生捕获。

## 反向审查

1. 不能只保留源空run：旧真实失败仍失败，新原生空mark和实际输入都验证。
2. 不能让段尾复制动作或覆盖真实正文：source动作只有一次，非空段落默认结束行为不变；单个实际插入样式不冒充保留多个零宽身份。
3. 不能用同字体名或advance bbox猜墨迹：资源、实际GID/轮廓、嵌入程序bytes及整页栅格分别核。
4. 不能扩源writer权限：原生只有插入mark时字符写入仍阻止，font/style混改仍整体拒绝。
5. 不关闭完整notes/继承/字体父项：本批只证明具名全空段落的fresh输出与原生输入继承；尾随空run、未知脚本、source样式写回、完整fields/artwork和跨应用仍开放。
