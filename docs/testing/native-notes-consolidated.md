# 备注原生编辑输入的三消费者汇总复验

2026-10-26从明文恢复后，用当前build重新执行[汇总runner](../../scripts/inspect-native-notes-consolidated.mjs)，全部37份不同字节的正向输入完成；39条逻辑路径中两条是与styles/regular第三轮完全相同的theme baseline，不重复计为新的输入字节。

[新manifest](native-notes-consolidated-evidence.json)的run为 `.tmp/native-notes-consolidated/run-MJJrkg`。每份输入在Node、browser ESM和global三个入口都执行精确no-op、font/bold/language三种混合字符+样式修改的整体阻断和还原、三轮字符候选；完整document/report/阻断报告/候选bytes与Node相同，浏览器零网络请求。独立DOM核实际source、target、resolvedFrom XPath、实际notes→master→theme/override、type匹配、完整script maps及local/空/坏槽；字符之外的所有payload保持精确字节。

共111个初始精确no-op路径、333个还原后no-op、333个混改阻断、333条字符候选路径；Node实际保存111份候选及完整快照，完成后重新逐份核SHA。所有候选均为headless，没有把它们算作原生重开或实际字体/版式操作。实际append/style/theme源文件的原生操作仍分别引用[编辑记录](native-notes-font-edits.md)和[主题记录](native-notes-font-themes.md)。

另用[只读负例验证器](../../scripts/research/oracles/verify-captured-notes-negatives.py)重新检查143份既有实际ZIP/PDF反例的SHA，并调用当前独立oracle；各阶段正向控制先通过，随后反例全部拒绝。结果在[当前负例证据](native-notes-captured-negatives-evidence.json)：追加28、样式60、主题55，涵盖text/lang/style/local槽、master/theme/override/邻页/关系、字体资源、遮白、真实页序、后轮墨迹及matrix。没有重写旧破坏文件，也没有放宽原baseline规则。

旧汇总manifest已经精确归档至旧run-MaKV81的manifest-before-recovery.json，并由[恢复hash表](native-notes-recovery-hashes.json)绑定，和本次新run分开。当前完整workspace测试通过，其中Parser678、import集成115、浏览器视觉33项。该具名候选汇总卡可收口；字体选择差异、shared override原生准入、完整notes作者/字段/预览、空插入样式、source style writer和父项仍未完成。

    rtk proxy node scripts/inspect-native-notes-consolidated.mjs
    rtk proxy python3 scripts/research/oracles/verify-captured-notes-negatives.py

第一条建立新run并在全部完成后更新manifest；恢复已有run时先核进程状态，不能重复启动。第二条只读已有反例与正向捕获，仅更新自己的JSON验收报告。
