# Fresh author PPTX native font controls

2026-10-26：本记录推进 IMP-05.1.3.3.2.5.2 / .6.2 的 regular 无编辑原生基线。16/16 个控制的实际 PDF PostScript 字体匹配，三轮 PowerPoint 保存、关闭、重开后的正文、语言、三个字体槽、run 数和完整页面栅格保持。文本编辑、粗斜体、主题切换和完整字体父任务仍未完成。

## 样本与生成边界

[fixture](../../tests/fixtures/real-font-fresh.mjs) 直接调用作者 API；[生成器](../../scripts/prepare-real-font-fresh.mjs) 调用 `generatePptx`，不修改生成 XML、不合并历史 native 控制。8 字符串各有 direct / map 两页，共 16 页，720×260 pt，正文 48 pt、无自动换行。Label 和 Probe 独立命名，标签为 Arial 16 pt，不计入正文实际字体检查。

| 样本 | 语言 | 目标 family | PDF 实际 PostScript face |
| --- | --- | --- | --- |
| 骨直令返 / SC | zh-CN | Hiragino Sans GB | HiraginoSansGB-W3 |
| 骨直令返 / TC | zh-TW | Songti TC | STSongti-TC-Regular |
| 骨直令返 / Japanese | ja | Hiragino Mincho ProN | HiraMinProN-W3 |
| 한글 | ko | Apple SD Gothic Neo | AppleSDGothicNeo-Regular |
| السَّلَامُ | ar | Geeza Pro | GeezaPro |
| لا | ar | Geeza Pro | GeezaPro |
| שָׁלוֹם | he | Arial Hebrew | ArialHebrew |
| a + combining acute | en | Arial | ArialMT |

Direct 模式保留独立 Latin / EA / CS literal 槽；map 模式将当前选中 family 写入三个槽。生成时恰有 8 个位于 map Probe 的 `script-font-map-materialized` 警告，作者规范化快照不变。CI [跨包测试](../../scripts/script-font-map-parity.test.mjs) 用独立 DOM 核对16页文本/lang/槽/单run与warning位置，无需安装这些字体。源map包含未使用的 Grek Georgia；本轮不证明原生编辑可以恢复该规则。

## 独立证据与复查

应用为 Microsoft PowerPoint for Mac 16.104 (16.104.25121423)。通过 CUA 打开 fresh 源文件、导出初始 PDF，然后依次另存 `fresh-round1` / `fresh-round2` / `fresh-round3`，每次关闭副本并从准确路径重开，再导出 PDF。保留历史 Cycle 文档，未执行删除、字体安装或系统设置修改。

本地 artifacts：`.tmp/real-font-fresh/run-81jnGm`。源 PPTX SHA-256 为 `ee318c757a9bcec08c156bf469efc0ea7b38c8b2717d75f28a1727bca2c62b03`。[manifest](real-font-fresh-evidence.json) 记录四份PPTX/PDF hash、每页字体/正文/拓扑/栅格hash；完整 `expected.json` 另存作者快照与source map。字体文件 SHA 与已有安装字体证据一致，未提交字体二进制。

[Python oracle](../../scripts/research/oracles/real-font-fresh.py) 独立读取 ZIP，按 presentation relationship 的真实顺序找到幻灯片，再以具名shape读取正文；不把文件名顺序或manifest顺序当作实际顺序。检查16页Label/Probe、文本、lang、48pt、wrap none、单run和三个槽。PDF另核尺寸、实际标签顺序、正文trace的实际PS face、非空glyph/无glyph0、可见dark ink；三份重开PDF都与源PDF逐页逐像素相同。已实际查看 contact sheet，CJK、Arabic marks/ligature、Hebrew marks、Latin accent可见。PDF trace glyph数不等同Unicode字符数，也不作为完整shaping算法认证。

[破坏测试](../../scripts/research/oracles/real-font-fresh-checks.py) 修改实际 ZIP 正文、EA 槽、标签，以及真实 PDF 页顺序、移除正文，5项均被拒绝。临时破坏文件独立保存，不改原生证据。

```sh
rtk proxy node scripts/prepare-real-font-fresh.mjs
# 对新run通过PowerPoint GUI另存、关闭重开和导出；不可复制旧PDF冒充新轮次
rtk proxy python3 scripts/research/oracles/real-font-fresh.py .tmp/real-font-fresh/run-81jnGm fresh-round1 fresh-round2 fresh-round3
rtk proxy python3 scripts/research/oracles/real-font-fresh-checks.py .tmp/real-font-fresh/run-81jnGm
rtk proxy node --test scripts/script-font-map-parity.test.mjs
```

该结论限于具名环境和16个regular、孤立脚本样本。精确栅格稳定不等于与浏览器/SVG跨引擎像素一致；未完成跨run mixed bidi、emoji、表格/notes实际字体、文本/样式/主题编辑。fontconfig 对 Hiragino Sans GB 的 Bold 查询找到 W6，但 Italic 查询返回 W3/slant0，不能把请求 Italic 或 OOXML `i` 标记当作独立斜体face存在的证据。

后续：[SC原生编辑记录](real-font-edits.md)已验证追加Ω、粗体、合成斜体和Office主题各三轮；其他脚本与mixed编辑仍开放。
