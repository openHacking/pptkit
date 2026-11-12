# Notes slide-image rect 的恢复验证

2026-10-26：从明文恢复记录继续核验 IMP-03.3.2.1。源样本构造确实遗漏了 notes master 的 slide-image `a:prstGeom prst="rect"`；现已在[构造器](../../tests/fixtures/native-notes-fonts.mjs)补齐。没有改变 Exporter 的公共默认输出，也没有把源补丁样本称为 fresh 作者往返。

## 原因与保留的对照

捕获目录 `.tmp/native-notes-pages/run-VAnIFP` 保留旧失败文件、PowerPoint 自建的两页对照和六种修改控制。使用[独立 oracle](../../scripts/research/oracles/native-notes-pages.py)核验实际 ZIP payload：`patched-original` 与 `patched-master-rect` 的部件集合相同，只有 notesMaster1.xml 不同；从修复 XML 移除新增 rect 后，完整 ElementTree 序列化与原 XML 相等。其余所有部件逐字节相同。

只移除 image 文本体、只恢复默认 idx、两者组合仍然空白；只补 master rect 就出现对应 slide 标题及图像内容。它定位的是本组源构造的必要差异，不是所有 Office slide-image 问题的通用根因。历史字体、追加、样式和主题输出都保留原样，不能把后补 rect 的结果追溯算成旧文件已支持完整备注图像。

## 当前捕获与验证

PowerPoint for Mac 16.104 中从单变量 `patched-master-rect.pptx` 起，三轮分别另存为 `rect-round1/2/3.pptx`，关闭、按准确路径重开，然后用打印对话框“注释”/纵向输出 `*-notes.pdf`。每轮均进入对应文稿窗口，未出现修复对话框。共45页重开记录，另有15页首次打开输出。

再从第三轮在缩略图窗格复制第一页，生成16页副本。另存、关闭、准确重开、打印也执行三轮，文件为 `rect-copy-round1/2/3`，共48页重开记录。重复的首两页具有相同内容，但必须使用不同 slide part 和 notes part；按实际 presentation 顺序检查每条 slide→notes→slide 双向关系、notes→master 和 master rect。其余14页严格按原顺序对应，不靠物理文件编号猜测。

[证据 manifest](native-notes-pages-evidence.json)记录所有实际 PPTX/PDF 的 SHA256、每页标签、关系、栅格 SHA 和图像区域墨迹。独立检查包括：

- 图像区域确有墨迹且包含对应 slide 的唯一标签；页尺寸为595×842pt。
- rect 前后正文和装饰区域（Y≥300pt）像素逐字节相同；字形、字体、origin/bbox 等 trace 字段相同，仅忽略插入绘制对象造成的 PDF `seqno` 编号变化。
- 三轮原始 PDF 整页栅格、完整 trace 和各字体资源的实际嵌入字体程序 SHA 均与首次 rect 输出精确一致，没有做位置校正。
- 复制页每轮全部16页整页栅格、完整 trace 与实际嵌入字体程序 SHA，与对应原页精确一致；首两页对应同一原页，其余页顺序保持。
- 四个实际空白 PDF、一个错页映射控制和缺 rect 的实际 PPTX 共六个反例均拒绝。错页映射是对检查输入的内存变换，不冒称另存的 PDF 破坏文件。

当前[跨消费者回归](../../scripts/notes-native-font-parity.test.mjs)独立使用浏览器 DOM 检查 master rect、local 占位继承及全部15页双向关系；A/B 两组继续验证 Node/ESM/global 三轮字符保存、精确 no-op、完整 document/report/bytes 和 source 样式阻断。新源构造不覆盖旧基线。

复跑已有捕获：

    rtk proxy python3 scripts/research/oracles/recheck-native-notes-pages.py .tmp/native-notes-pages/run-VAnIFP docs/testing/native-notes-pages-evidence.json
    rtk proxy node --test scripts/notes-native-font-parity.test.mjs

2026-10-28 只读复核：独立 oracle 重新检查现存捕获，并将输出按 JSON 表示归一化后与既有 manifest 比较；输入文件 SHA、首次页面、原始与复制文稿三轮、六个反例和总数均一致。该复核没有重做原生 GUI 操作，也不改写原 manifest。

2026-10-29 再次运行同一只读复核，结果为 15 个源页、45 个原文稿重开页、48 个复制文稿重开页及 6 个被拒绝的负控，与原 manifest 一致；26 份清单文件的 SHA256 也独立重新计算并全部匹配。此结果复核留存捕获，并非新的 PowerPoint 交互记录。

上述命令不代替原生 GUI 保存/重开。恢复时不要重跑旧基线准备器来覆盖捕获文件。已用 Poppler 渲染第三轮首页，并逐页查看第三轮复制文稿的16页联系表；图像、notes 正文、local 标签及 master 装饰都出现。Hebrew/Arabic 的正确 shaping/bidi 和原生选字体差异仍沿既有字体任务验收，不能从图像不再空白推出排版完全正确。

## 对抗审查与剩余边界

1. **修错层**：只修改源构造器。这次源修复时默认 Exporter 的 notes master 尚没有完整占位结构；后续[默认fresh版式](authored-notes-pages.md)已另行实现和验证，不混入本组源实验。
2. **正文或装饰被改坏**：15页首次输出及三轮均有逐像素/trace 对照；不只检查可提取文字。
3. **缩略图来自邻页或复制页回链错误**：唯一标签、实际页序、双向关系、不同 part 身份及复制后三轮全部页面对照；错页映射必须失败。
4. **重开并未真的发生**：本次新捕获执行了三轮关闭/准确重开，复制组也独立三轮；manifest 不将 PDF 重导出算成额外原生保存轮次。
5. **过度关闭任务**：只有 A 组具名源修复与原生复制得到本批完整证据。B/override 的图像原生准入、完整 notes fields/artwork 作者模型、跨平台、源样式写回与父项依赖仍开放；IMP-03.3.2.1 保持未勾选并记录局部进展。


## B/override 后续状态

[B与逐notes独立override的图像验收](notes-image-domains.md)现已补齐45首次页和135页三轮重开证据，正文/装饰保持，所有非app/core payload精确不变。以上本组A观察及旧失败文件仍原样保留；完整notes fields/artwork模型、共享override原生准入、字体等价及父依赖继续开放。
