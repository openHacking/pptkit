# B 与逐备注 override 的 slide-image 原生验收

2026-10-26：接续[A组rect源修复](native-notes-pages.md)，补齐B独立notes主题及单master/逐notes独立fontScheme override的图像验收。本批没有生产代码或公共API修改；原字体选择差异、主题位移与shared override repair观察仍保留。

## 输入和单变量

[准备器](../../scripts/prepare-notes-image-domains.mjs)先核既有manifest和实际原生输入SHA，再创建 `.tmp/notes-image-domains/run-k4KHeU`：

- B组来自已验证的regular-round3，15页，保留原文/Ω、局部/空/坏/major字体槽及独立B notes主题。
- override组来自已验证的native-baseline-round3，30页交错A/B，单notes master，15个B notes各有独立override part；不是原来会弹修复提示的shared-part源。

各组baseline.pptx/PDF精确复制旧捕获。rect.pptx仅向实际notes master的sldImg占位新增rect几何；独立ZIP检查所有其他payload精确不变，移除新增rect后完整master XML树与baseline相同。不是fresh作者输入，也没有重新制作原生追加/样式/主题操作。

## 真实操作与结果

PowerPoint for Mac16.104中，两组rect输入均正常进入15/30页文稿窗口，无修复提示；分别输出首次纵向“注释”PDF。然后各三轮另存round1/2/3、关闭、准确路径重开和Notes-print，共45首次页和135重开页。PDF保存目录导航曾未落到目标文件夹，均在保存前停止，逐步核路径/文件名后继续；实际PPTX重开轮次与PDF数量逐份核，不把失败导航当成新轮次。

[独立oracle](../../scripts/research/oracles/notes-image-domains.py)按presentation顺序和实际关系检查：slide↔notes、唯一列出/实际notes master、独立notes与slide主题、B的15个独立override与A的无override、完整占位type/idx7/99、实际页标签。三轮全部域身份与rect输入一致，notes/slide/master/theme/override/relationships等所有payload精确保持，只有app/core元数据有变化；[manifest](notes-image-domains-evidence.json)保存这些元数据的完整前后XML与SHA，未给它们无限规范化许可。

rect前后Y≥300pt的正文/控制标签/master装饰栅格精确不变，glyph/font/origin/bbox等trace字段相同，仅忽略新增绘制对象导致的seqno编号变化。图像区域有实际墨迹并包含每页对应的独特slide标签。三轮原始整页栅格、完整trace及每个实际嵌入字体程序SHA与rect首次PDF精确一致，没有位移回正或容差。已查看override交错域的具名12页联系表；不能由此推断原有Arabic/Hebrew shaping完全正确。

每组原来图像空白的实际PDF仍拒绝，错页映射控制也拒绝，共四个控制；错页映射是检查输入的内存变换，不冒称新PDF破坏文件。另存六份实际PPTX破坏文件全部拒绝：两组各去master rect/破坏notes回链，override额外去B override/改其字体字典。原生和旧失败文件不修改。

## Parser和现有回归

[实际输入runner](../../scripts/inspect-notes-image-domains.mjs)及[证据](notes-image-domains-parser-evidence.json)检查每组rect和三轮共8输入。归一化Core文档必须与其旧空图像baseline完全相等，证明这次几何补丁没有改写声明字体或正文语义；完整snapshot/report和输入SHA保留。8条精确no-op、混合字体/字符修改整体阻断后还原，8份字符候选在实际notes owner之外所有payload精确不变；它们未原生重开，不计入135页。两组第三轮browser global完整document/report/bytes与Node相等且零网络，未声称全部输入运行了三个浏览器入口。

两套既有source回归（notes-native-font-parity和notes-master-isolation-parity）4组通过，覆盖Node/ESM/global、A→B→A/三轮保存、独立override、错误基数/域隔离。该轮无生产变化，未重复宣称完整workspace套件重跑；上一批完整build/typecheck/test证据仍独立保留。当前独立native/parser runner、source回归及lint/文档/边界/diff检查见 `.tmp/notes-image-domains-{native,parser,source-tests,lint}.log`。

复跑已有捕获：

    rtk proxy python3 scripts/research/oracles/notes-image-domains.py .tmp/notes-image-domains/run-k4KHeU
    rtk proxy node scripts/inspect-notes-image-domains.mjs

先运行独立native oracle再运行Parser runner，以保持两份manifest的hash绑定；准备器只用于新目录，不应覆盖已捕获实验。

2026-10-29 使用[隔离复检器](../../scripts/research/oracles/recheck-notes-image-domains.py)重新执行检查：临时复制捕获，在副本生成负例，比较每页语义结果、原生轮次和负控拒绝原因，不覆盖原捕获或 manifest。45 个首次页、135 个重开页、4 个空白/错页控制及 6 个实际 PPTX 破坏控制均与留存记录一致；重新打包的负控 ZIP 时间戳可不同，因此不比较其文件摘要。另将临时 manifest 的重开页数加一，确认复检器退出失败。这是已有 PowerPoint 捕获的复查，不是新一轮原生操作。

    rtk proxy python3 scripts/research/oracles/recheck-notes-image-domains.py .tmp/notes-image-domains/run-k4KHeU docs/testing/notes-image-domains-evidence.json

## 反向审查与后续

1. **改错源**：实际PPTX/PDF与旧manifest先核SHA，独立证明仅新增master rect；不把新字节套到旧字体验收。
2. **借错域或复用override**：实际图和15个独立part身份逐页核，改/去override实际文件必须拒绝。
3. **图像有字但正文/装饰被损坏**：图像墨迹/唯一标签、下部精确栅格/trace和三轮整页/字体程序分别核，旧空白实际PDF必须失败。
4. **native重写被掩盖**：仅app/core变化且完整XML保留；未放宽其他payload，也未把headless候选或保存导航尝试算作原生轮次。
5. **过度关闭父项**：A、B及distinct-override具名图像问题已有证据；共享override原生准入、完整notes fields/artwork/source-only模型、样式writer、通用字体等价与跨平台仍开放。IMP-03.3.2.1记录自身交付，而父依赖未完成不提前勾选。
