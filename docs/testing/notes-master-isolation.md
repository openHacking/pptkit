# Notes 主题隔离、基数与非合规源诊断

2026-10-26：补齐合法 notes-master 基数控制与非合规源容错的区别。新增两个30页source包，A/B实际页序交错，分别三轮A→B→A字符保存；Node/browser ESM/global共18条保存/no-op路径通过。独立runner保存8个源状态、核8次Node精确no-op，并实际重新打包80个破坏ZIP，独立oracle全部拒绝。本批source数据不计native。后续[具名原生主题实验](native-notes-font-themes.md)观察到shared override part触发修复提示，并以每notes独立part的变体完成三轮基线和Office主题；source/原生证据分别记录。

## 规范与规划纠正

Microsoft对PresentationML的说明规定，每个包最多包含一个Notes Master part。[原文与上下文](https://learn.microsoft.com/en-us/office/open-xml/presentation/structure-of-a-presentationml-document#notes-master-part)。官方[Open XML SDK schema](https://github.com/dotnet/Open-XML-SDK/blob/main/data/schemas/schemas_openxmlformats_org_presentationml_2006_main.json)的NotesMasterIdList也将notesMasterId上限设为1。

因此，“一个包内两个notes master作为标准创作/原生编辑正向准入”不是正确退出条件。合法隔离包括slide master与notes master各自的theme，以及同一notes master下逐notes的themeOverride。额外notes master只作非合规source对抗/诊断控制；不得以诊断式读取或原字节保存证明标准合规或PowerPoint可编辑。

现有Parser保持兼容：列表登记两个notesMaster直接抛`import-invalid-package`，位置为`/ppt/presentation.xml`；实际notes关系指向未列出的master时，保留`unlisted-master` finding并有界读取实际source。后者是容错投影，不是多notes-master作者能力，也不是完整schema validator。本批没有放宽基数、修改默认profile、增加writer权限或新增生产API。

此前[notes-style source证据](notes-style-import.md)使用了第二个未列出的notes master。其runner与manifest已补入非合规分类、规范链接及实际结构finding，并重新核验原有四状态/每状态22个属性与8个双重拒绝控制；历史artifacts保留。不能将这批source证据改述成合规原生文稿。

## 两组具名source

[assembler](../../tests/fixtures/notes-master-isolation.mjs)复用两份作者基础包及已说明的source patch，再构造30页。物理slide1..15为A、16..30为B，presentation顺序为1/16/2/17/…/15/30；源标签、notes回链、部件关系和主题来源必须按实际身份解析，不能按ZIP顺序或数组下标猜测。两种控制保留unused theme2。

| 控制 | 实际notes master | 字体来源 | 结构结果 |
| --- | --- | --- | --- |
| `override` | 单一master1，全部notes实际关联它 | A取master的独立theme；B的15个notes各有themeOverride关系，使用B fontScheme | notes-master基数为1，没有unlisted-master finding；完整schema仍未认证。后续具名PowerPoint对这个shared-part源提示repair，独立part变体另验。 |
| `unlisted` | A取master1，B取额外master2，列表只登记master1 | 各自实际master→theme链 | 明确非合规；每轮15个B目标都有unlisted-master finding，source字体不能借A或slide主题。 |
| `listed-two-rejected` | 两个master均登记到列表 | 不进入可用会话 | Node、browser ESM/global均拒绝；随后可独立重试前两种source。 |

每域各15控制沿用Latin/SC/TC/ja/ko/ko-KR/Arabic/Hebrew/Greek、三local覆盖、空/坏EA及major-TC。body type相同但local/master idx=7/99；placeholder字号40pt与global18pt区分匹配来源。A/B字体字典不同，local不带继承map，空/坏EA的EA脚本规则不借近层之前的值；不通过宿主字体加载或PDF外观推断source。

## 三轮保存、替换和失败行为

[跨消费者回归](../../scripts/notes-master-isolation-parity.test.mjs)在两种可诊断读取的source上先改A、再改B、再改A；原文分别增长X/X/XX。每轮检查全部30页的原文、language、40pt、generic Latin/EA/CS、完整map与实际source part，逐备注的未列出诊断保持。source字体混改整体blocked，恢复后才保存字符；实际候选重导入snapshot和结构findings相等。

Node每轮直接检查ZIP part库存及字符外精确bytes，ESM/global的完整document/report/输出bytes与Node相等，网络请求为空。未改域的所有备注、slide/标签、master/theme/override、metadata和关系均精确保留。两种source×三消费者×三轮=18条保存路径与18次no-op；此消费者矩阵不与独立runner的8次no-op相加冒充独立文件数。

另测隔离边界：

- 非合规B的实际theme或master关系缺失时，不借列表中的A；B缺theme的继承字体不可解析，B local字面槽保留，A不受影响。
- B的master body改为错误type，即使idx改成7也只能采用global18pt；A仍为40pt。
- 单master控制中，某个B notes缺失override关系时，合法回到它实际master的A字典；其他B notes继续用B，不污染共享master缓存。
- B的override fontScheme存在但缺Hant时，不从较旧的master方案补入A的Hant。local覆盖及A源状态保持。

“缺override沿实际链回退”与“存在近层组件却缺键阻断旧方案”是不同情况。错误type/缺引用也不能靠presentation列表或slide主题补出字体。这些是source行为，不能推断PowerPoint将怎样规范化坏输入。

## 独立OPC/XML与实际反例

[runner](../../scripts/inspect-notes-master-isolation.mjs)输出每组原始加三轮保存的四个PPTX、controls.json和snapshot/report，输入输出hash记录于[manifest](notes-master-isolation-evidence.json)。[独立Python oracle](../../scripts/research/oracles/notes-master-isolation.py)从ZIP重新读取content types、实际presentation页序、notes/master/theme/override关系和回链；检查物理notes master数量、列表基数、每个B的unlisted来源、type/idx、局部槽、两个font collection及源文字。每状态248个mapped字体/字号属性逐一核实际source/target/resolvedFrom及value；不把Parser输出本身当源真值。

oracle还逐字节比较A/B/A转换的所有part：只有当前域各notes的特定`a:t`字符可变，不能修改另一个域或标签。两组各四状态×十种真实破坏，共80个临时PPTX均重新打包、重新读取并记录破坏文件hash后拒绝：错误actual master、借用A主题、坏B规则、master/local错误type、language、文字、local字体、backlink和真正的页序交换。原始样本不覆盖，反例不靠伪造report替代修改真实payload。

artifacts根目录：`.tmp/notes-master-isolation/run-lh85iw`，以manifest的artifactsDirectory为准。`override`与`unlisted`分别存储，额外notes-master控制的`standardConformance=false`明确保留；override组仅验证本任务的基数/已支持fontScheme语义，没有完整schema证书。全部都是headless source证据，没有native页面或字体外观计数。

```sh
rtk proxy node --test scripts/notes-master-isolation-parity.test.mjs
rtk proxy node scripts/inspect-notes-master-isolation.mjs
rtk proxy node scripts/inspect-notes-style-import.mjs
```

新增回归接入test:import。build/typecheck、22组Parser、3组既有notes跨包及新增三组回归通过；lint/文档检查见执行记录。后续[追加/局部样式](native-notes-font-edits.md)及[Office/独立part override原生响应](native-notes-font-themes.md)的具名子卡已完成；blank slide-image及完整notes/字体/round-trip父项仍开放。

接续source回归现增加distinctParts控制：[fixture](../../tests/fixtures/notes-master-isolation.mjs)为B每个notes生成独立override part。三类source×三消费者×三轮共27个scalar/no-op路径，完整map、actual source、未改域bytes及candidate相等；原有shared和非合规额外master控制保留。旧manifest的八物化状态/80破坏仍只指原来两类source，不改写为12状态；distinct的实际包/关系/原生基线与主题证据另见新manifest。
