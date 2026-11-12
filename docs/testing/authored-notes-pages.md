# 默认 fresh 作者备注页

2026-10-26：[ADR0054](../architecture/decisions/0054-default-notes-page-output.md)补齐默认 Exporter 的 notes master/local slide-image 与正文占位。保持既有540×720pt notesSize；按真实幻灯片比例将缩略图居中放入上部468×288pt区域，正文在下部独立区域。没有新增公共选项、截图资产、Core备注几何或普通slide SVG语义。

## 公共作者输入

[具名fixture](../../tests/fixtures/authored-notes-pages.mjs)只使用公共createPresentation、duplicateSlide、setNotes、replaceElement与generatePptx；[准备器](../../scripts/prepare-authored-notes-pages.mjs)生成横向720×260、纵向360×720和正方形500×500三个输入，位于 `.tmp/authored-notes-pages/run-rM4cWd`。没有对ZIP/XML打补丁。

每份四页：ONE正文与缩略图、实际Core复制后改名并有独立COPY备注、无备注的EMPTY、显式空白12pt粗体的STYLED。它们的源SHA、实际页序和文字保存在各自expected.json；旧source-patched字体控制不充当本批fresh作者证据。

## 原生操作和独立验证

PowerPoint for Mac16.104中，三个源输入分别正常进入四页文稿窗口，首次实际打印纵向“注释”PDF。每组三轮另存、关闭、准确路径重开后打印Notes PDF，命名 `wide/portrait/square-round1/2/3`，共九轮、36重开页及12首次页。没有出现修复提示。

[独立ZIP/XML/PDF oracle](../../scripts/research/oracles/authored-notes-pages.py)按presentation页序追踪实际slide→notes→slide、notes→master→theme；检查所有notes唯一身份、复制页标签/正文、母版rect和几何、notesSize、空正文占位的有无与字形。每轮整页栅格、完整font/glyph/origin trace和实际嵌入字体程序SHA与首次打开精确一致。没有做位置校正、重排或栅格容差。三个错COPY页映射反例均拒绝。

[manifest](authored-notes-pages-evidence.json)记录源文件与expected的hash、每轮PPTX/PDF hash、实际关系/几何、每页字体/位置与栅格SHA。已查看三种尺寸第三轮共12页的联系表，确认缩略图和对应正文、空页及复制页的实际页面出现；这不替代下面明确失败的空样式验证。

源正文为18pt，实际Notes-print trace为约19pt；源slide标签24pt在三种缩略图中实际为约16/10/14pt。源notesSize为540×720pt，实际打印纸张595×842pt。manifest保留各自实际几何和字体数值，不把原生打印后的坐标/字号直接称作源值相等，也不从这些数值猜测通用打印变换。正文和图像标签区域各自核真实可见墨迹，避免仅有可提取文字却实际遮白。

## 实际失败：原生空白样式丢失

源STYLED包含空run，`a:rPr sz="1200" b="1"`；三轮fresh投影仍保留。但PowerPoint第一次保存就删除空run，只留下默认 `a:endParaRPr lang="en-US"`。三种尺寸各三轮都记录 `emptyStyledRunRetained: false`，没有将其改成通过，也没有凭空白PDF声称样式一致。

本次修复交付的是可见图像/正文版式和对应关系；空插入样式的共享文字序列化缺陷另列待办，完整notes输出验收仍开放。后续需保留本组失败文件，在新目录验证修正后的end-paragraph插入属性与真实追加文字，不能覆盖本批证据后宣称此前已通过。

## 自动化和对抗检查

[跨包测试](../../scripts/notes-page-parity.test.mjs)已接入test:import：四种合法slide尺寸（另含4032×72极端比例），独立浏览器DOM检查notesSize、母版rect、正数/页内边界、上下不重叠、比例误差≤1EMU舍入、居中、无假image文本、唯一绘制身份与双向回链。Node与browser global生成bytes相等；三轮fresh内容投影和原字节no-op保持空/无正文区别；去掉实际rect的XML变体被拒绝。既有空备注样式/action、notes跳转关系及全部source字体控制继续测试。

该自动检查还分别约束图像在上部目标区域内垂直居中并至少贴合一条区域边；只按比例缩小、或只向下移动而仍在页面内的变体都被拒绝。这是针对fresh输出的几何检查，不是新的原生打印或编辑记录。

导出器计算缩略图比例时使用 `p:sldSz` 实际写出的整数EMU，而不是未舍入的输入点数。新的边界测试覆盖72与4032点、四舍五入后仍处于边界的值、越界/极端输入，以及 Node 和浏览器输出一致性；该测试不代替原生应用验证。

补充的 [LibreOfficeDev 26.8.0.0.alpha0 无界面证据](notes-size-libreoffice-evidence.json)覆盖极端横向、纵向及四舍五入后纵向三份两页样本，各三轮重存。12份实际文件经[独立ZIP/XML检查](../../scripts/research/verify-notes-size-libreoffice.py)核验：slide尺寸、每页正文和双向关系保持；应用第一次重写母版图像几何，同时给每页添加局部图像框。从第一轮到第三轮，局部框的宽和高每轮均缩小360 EMU（约0.0283 pt），坐标不变。图像实际绘制或打印效果未检查，不能根据母版变化或局部框变化推断视觉结果。该实验未验证 PowerPoint 或 GUI 操作。

可用 `python3 scripts/research/verify-notes-size-libreoffice.py` 复查留存文件，`node --test scripts/notes-size-libreoffice-evidence.test.mjs` 检查伪造记录以及重新计算哈希的正文、关系和几何破坏；运行器创建新目录，不覆盖已有捕获。

[解析器对这些实际文件的回归](../../scripts/notes-size-libreoffice-parser.test.mjs)检查全部12份源及重存文件：备注页源库存准确报告母版和每页局部图像框，Node、浏览器ESM/global报告一致，结构和内容导入的 no-op 均保留原字节。另从三份第三轮文件分别进行三轮既有备注run字符编辑；[独立ZIP检查](../../scripts/research/verify-notes-size-edit.py)要求只有目标备注部件的目标文本变化，局部框、邻页和其他payload保持。人为更改局部几何及邻页正文的反例均被拒绝。该测试依赖本机留存捕获，因此是可选证据测试；它不表示几何已进入Core或获得写回权限。

build/typecheck与完整pnpm test退出0，导入集成115项通过；日志 `.tmp/authored-notes-build.log`、`.tmp/authored-notes-full-test.log` 和 `.tmp/authored-notes-page-parity.log`。lint/文档与diff在收尾检查。当前自动化不读取native宿主字体，不把浏览器图像等同于PowerPoint图像。

五项反查：修复在Exporter而非Core；正文首shape和ID保持，source patch控制先去默认占位避免重复；复制页沿关系而非文件号映射；真实重开/首次PDF与源码投影分别计数；空样式失败、导入artwork/fields、完整notes模型、溢出续页与跨平台仍明确未完成。

复跑已有捕获：

    rtk proxy python3 scripts/research/oracles/authored-notes-pages.py .tmp/authored-notes-pages/run-rM4cWd
    rtk proxy node --test scripts/notes-page-parity.test.mjs

准备器只用于新实验，不在恢复捕获时重跑。


## 后续修复状态

[空段落插入属性修复](notes-insertion-styles.md)已在新目录验证：四组完整end属性三轮保持，四次实际输入继承样式，输入后三轮原生重开与PDF稳定。原run-rM4cWd的9条失败记录保留原样，不能把新代码的成功追溯为旧文件通过。原生仍删除零字符run身份；修复保留的是最终插入样式。完整notes模型/字段/源样式写回继续开放。
