# -*- coding: utf-8 -*-
"""逐文件讲解文案：全部来自对 src/*.cs 的实地阅读，函数清单由 ppt_inventory.json 提供。"""

# 章节：编号 / 名称 / 一句导语 / 该章文件顺序
STAGES = [
    ("01", "读之前：先建立地图", "先知道这套程序解决什么问题、由哪几层组成，再打开源码",
     []),
    ("02", "第 1 站：入口与地基", "从 Main 出发，搞懂单实例、配置、日志、Win32 与 COM 接入",
     ["Program.cs", "Settings.cs", "Log.cs", "Native.cs", "ComCtx.cs", "ComCall.cs",
      "AutoStart.cs", "VbaTrust.cs", "TrayPin.cs"]),
    ("03", "第 2 站：纯逻辑判据", "最好啃的一层：不碰 Windows、不碰 Office，全是字符串判断",
     ["Rules.cs", "Fuzzy.cs", "Editing.cs", "Parser.cs", "Scope.cs",
      "VbaBuiltins.cs", "BuiltinTables.g.cs"]),
    ("04", "第 3 站：数据与词库", "候选词从哪来：SQLite 两库、代码片段、最近使用",
     ["Sqlite.cs", "WordLibrary.cs", "WordBank.cs", "SnippetStore.cs", "Recent.cs"]),
    ("05", "第 4 站：宿主、引擎与界面", "真正的运行时：读 VBE、算候选、装钩子、画窗口",
     ["VbeBackend.cs", "Engine.cs", "App.cs", "Commands.cs", "PopupForm.cs",
      "SettingsForm.cs", "WordEditForm.cs", "SnippetForm.cs", "InfoBox.cs",
      "UiTheme.cs", "IconAssets.cs"]),
    ("06", "第 5 站：验证体系与上手", "用 587 条自检 + 15 个探针把前面每一层钉住",
     ["SelfTest.cs", "Probes.Support.cs", "Probes.Pure.cs", "Probes.Host.cs",
      "Probes.Key.cs", "Probes.Decl.cs", "Probes.Caret.cs", "Probes.Dump.cs",
      "Probes.Env.cs", "Probes.Ui.cs"]),
]

# 每个文件：职责 / 为什么排在这个位置 / 先读哪几个函数 / 带走的一句话
FILES = {
"Program.cs": dict(
    role="程序入口：只做三件事——保证只有一个实例、把开关分派给探针、启动托盘常驻。",
    why="全项目最短的入口（185 行），先读它能一眼看清「带什么参数会走哪条路」，不至于一上来陷进 1472 行的 App.cs。",
    first=[("Main", "看单实例互斥与命令行开关的分派"),
           ("RunMain", "看真正启动应用前做了哪些准备"),
           ("SignalExisting", "看第二次双击 exe 时怎么唤醒已有实例")],
    takeaway="--selftest 之类的开关会直接 return，不装钩子；只有不带参数才进托盘常驻。"),
"Settings.cs": dict(
    role="配置中心：与 Python 版共用同一个 settings.ini，负责读、改、恢复默认。",
    why="几乎每个文件都读配置。先把「环境变量 > ini > 默认值」这三层优先级弄清，后面看到 Settings.Flag 就不会迷路。",
    first=[("Flag", "看 A/B 开关怎么从环境变量读出来"),
           ("Load", "看 ini 的解析与只认 [settings] 段"),
           ("Set", "看保存时只写真正改过的项")],
    takeaway="改回默认值就把该行注释掉，而不是写死一个值；写完立刻重新 Load。"),
"Log.cs": dict(
    role="日志：默认不落盘，只有 VBECOMPLETE_LOG=1 才写队列式日志。",
    why="读懂「没日志 ≠ 没执行」这条规矩，后面用探针排查问题时才不会误判。",
    first=[("Write", "看日志怎么进队列而不是直接写文件"),
           ("DrainLoop", "看后台线程怎么秒退、为什么会丢尾部"),
           ("Rotate", "看日志文件的滚动上限")],
    takeaway="任何写日志失败都必须被吞掉——日志永远不许把主流程弄崩。"),
"Native.cs": dict(
    role="Win32 层：全局键盘钩子、按键注入、窗口与进程探测、输入法状态。",
    why="这是「钩子线程」被允许碰的全部东西。先记住它不含任何 COM 调用，才能理解项目里那些看起来绕的设计。",
    first=[("SendVk", "看补发一个真实按键用的是 INPUT 结构"),
           ("IsSlashKey", "看 / 只认 VK_OEM_2 与 VK_DIVIDE"),
           ("ImeComposing", "看输入法组字状态怎么读")],
    takeaway="钩子回调里绝不 Sleep、绝不碰 COM；一次卡住，用户的键盘就全卡住。"),
"ComCtx.cs": dict(
    role="COM 接入层：从运行中的 Office / WPS 拿到 VBE 对象并缓存，全程后期绑定。",
    why="「一个 exe 通吃 Excel、Word、WPS」靠的就是这层。先看清 Begin/End 的生命周期，才明白为什么不能常驻 COM 引用。",
    first=[("Begin", "看怎么从 ROT 拿 Application 并判断宿主类型"),
           ("End", "看用完立刻释放引用"),
           ("HostCandidates", "看 Excel / WPS 的 ProgId 候选表")],
    takeaway="打字中最多留 2 秒 COM 引用；WPS 要用窗口 PID 比对前台，不能只信 ProgId。"),
"ComCall.cs": dict(
    role="为了几个带 ByRef 出参的 VBA 方法，直接调 IDispatch::Invoke 的最小封装。",
    why="文件里没有「方法」只有结构体与字段，正好借它认识 VARIANT / DISPPARAMS，理解 C# 的 dynamic 在底层做了什么。",
    first=[("CallFourOut", "看四个 ByRef 出参怎么从 VARIANT 取回"),
           ("VARIANT / DISPPARAMS", "两个结构体：出参类型与参数数组")],
    takeaway="CodePane.GetSelection 这类方法有多个返回值，普通 dynamic 调用拿不到，才需要这一层。"),
"AutoStart.cs": dict(
    role="开机自启：在 HKCU\\...\\Run 登记一个值，免 UAC、可逆、不误删别人的项。",
    why="典型的「读写系统状态」模块：全量快照 → 只改自己的项 → 可回滚。这个套路在别的项目里也常遇到。",
    first=[("SetEnabled", "看写注册表前的意图判断"),
           ("IsOurCommand", "看删之前只认自己的命令"),
           ("Capture", "看全量快照为什么必须做")],
    takeaway="注册表是结果，autostart.txt 记的是用户意图；两者不一致时以意图为准。"),
"VbaTrust.cs": dict(
    role="读写「信任对 VBA 工程对象模型的访问」(AccessVBOM) 这个双向开关。",
    why="WPS 下 Application.VBE 会直接 E_FAIL，先知道这个前提，才不会把「读不到代码」当成程序 bug。",
    first=[("IsBlocked", "看读不到代码时的前置判断"),
           ("EnableAll", "看打开信任开关"),
           ("Read", "看注册表值的读取")],
    takeaway="这是环境前置条件，不是逻辑判据；不满足时探针要报 SKIP 而不是 FAIL。"),
"TrayPin.cs": dict(
    role="让托盘图标常驻，不被 Win11 折叠进 ^ 溢出区。",
    why="只有 64 行、1 个方法，适合作为第一个完整读懂的小文件，建立信心。",
    first=[("EnsureVisible", "看注册表里托盘图标状态位怎么改")],
    takeaway="Win11 把每个图标的显示状态记在注册表，代码只是替用户点了一下「显示」。"),

"Rules.cs": dict(
    role="纯判据合集：词边界、声明位置、关键字场合、显示列换算、宿主重排后的光标搬家。",
    why="「补全准不准」全部由这 42 个静态函数决定，且不碰 COM/UI/全局状态，是最好写测试、最适合新手逐条读的一层。",
    first=[("WordBefore", "21 行，看「光标前的半截词」怎么取"),
           ("DispCol", "看 Tab 展开与中文全角占 2 格的显示列"),
           ("DeclSlotOf", "看三态：不许名字 / 只给关键字 / 都给")],
    takeaway="列是显示列不是下标；中文行用字符列摆光标，等于偷偷把 bug 填平。"),
"Fuzzy.cs": dict(
    role="模糊匹配：判断哪几个候选命中、命中在哪几位、排第几档。",
    why="只有 180 行、5 个方法，是理解「为什么打 ar 会出 dataArr」的最短路径。",
    first=[("GreedyPositions", "看整块连续优先、否则逐字符贪心"),
           ("LongestRun", "看连续命中长度怎么影响档位"),
           ("Match", "看大小写不敏感与前缀加权")],
    takeaway="只做字符串、不懂 VBA：这一层可以被单独测到极致。"),
"Editing.cs": dict(
    role="编辑手感纯判据：块结构识别、缩进推算、自动配对、退格与回车接管。",
    why="一行 COM 都不许有，因为它跑在键盘钩子线程里。读懂「纯判据 + 真数据复核」的两档设计，就懂了项目最核心的套路。",
    first=[("EnterIndentWanted", "看严格判据：完整块头才缩进"),
           ("LooksLikeBlockHeaderPrefix", "看宽松兜底：少一个字母的 Then 也先接管"),
           ("CloserNeeded", "看 End If / Next 是怎么被补出来的")],
    takeaway="判不过就把回车还回 VBE，绝不吞键；ElseIf / Case 这类分支词必须在宽松词表里。"),
"Parser.cs": dict(
    role="VBA 源码解析：从模块全文里收集工程自己定义的名字。",
    why="完全不碰 COM、只吃字符串。搞清楚「候选池里到底有什么」，后面看 Engine 才不会凭空想象。",
    first=[("ExtractProcedures", "看过程头与 StartLine 怎么收集"),
           ("ExtractIdentifiers", "看声明里的新名字怎么收"),
           ("OptionExplicit", "看为什么它会把「收不到」放大成「不提示」")],
    takeaway="注释与字符串里的名字要跳过；点号后面的名字也不算新标识符。"),
"Scope.cs": dict(
    role="作用域分析：算出每个名字属于谁、在别处看不看得见，并给出 1~4 的档位。",
    why="档位决定候选排序。先理解「本过程局部 > 本模块 > 工程公有 > 内建」，再看 EnsurePool 会顺很多。",
    first=[("Analyze", "看模块文本怎么变成 ModuleScope"),
           ("IsAtStatementStart", "看语句起点含 : / Then / Else"),
           ("IsMemberAccess", "看点号后面为什么不收名字")],
    takeaway="过程头行算进过程体是形参作用域需要，但关键字场合判定必须把它单独排除。"),
"VbaBuiltins.cs": dict(
    role="内建名字的内存镜像：语言关键字、内建函数、内建类型。",
    why="事实来源其实是两个 SQLite 词库，本类只是加载后的缓存；读懂这一点才知道改词库该改哪儿。",
    first=[("Pool", "看把词库读进内存镜像"),
           ("IsBuiltin", "看关键字过滤用在哪儿"),
           ("ApplySeed", "看表空了怎么灌内置种子")],
    takeaway="声明位置里关键字要用 IsBuiltin 过滤，避免把用户自己的名字当关键字挡掉。"),
"BuiltinTables.g.cs": dict(
    role="自动生成的内置表：由 Python 版 vba_builtins.py 导出，请勿手工编辑。",
    why="0 个方法、85 行常量。它提醒你：这个项目有「唯一事实来源」的概念，C# 与 PY 两边必须逐字一致。",
    first=[("Functions（数组）", "171 条内建函数"),
           ("TypeNames（数组）", "内建类型全集"),
           ("Keywords（数组）", "语言关键字全集")],
    takeaway="改词库不要改这里，改 SQLite；这个文件只是生成物。"),

"Sqlite.cs": dict(
    role="极薄 SQLite 封装：P/Invoke 系统自带 winsqlite3.dll，不引任何 NuGet 包。",
    why="「一个 exe 双击就能跑」的承诺靠它兑现。顺便可以学 P/Invoke 与 using 释放非托管句柄。",
    first=[("Open", "看 sqlite3_open 的调用"),
           ("Exec", "看 SQL 执行与错误码"),
           ("Query", "看逐行读列的循环")],
    takeaway="零第三方依赖是有代价的：查询、取列、释放都得自己写。"),
"WordLibrary.cs": dict(
    role="词库装载：两个用户自维护的 SQLite 库（关键字 / 函数与类型）读成内存镜像。",
    why="候选池里除了代码里定义的名字，还有用户自己加的词；先看清 WantsSpace 这类补全后处理。",
    first=[("Load", "看表空了自动灌种子"),
           ("LoadKeywords", "看关键字库怎么读"),
           ("WantsSpace", "看补完关键字后要不要跟空格")],
    takeaway="两个库绝不与 Python 版共用；C# 版有自己的 keywords-cs.db / functions-cs.db。"),
"WordBank.cs": dict(
    role="词库读写层：给设置界面做增删改与启用/停用。",
    why="v182 的惨案就在这里：ListView 的程序性重建会触发 ItemChecked，曾把 182 条全写成停用。读它学会「改完要核数据本体」。",
    first=[("Save", "看新增/改一条词条"),
           ("SetEnabled", "看单条启用停用写库"),
           ("All", "看读出全部词条")],
    takeaway="界面勾选必须能真的写库；反向断言比「看起来勾上了」重要。"),
"SnippetStore.cs": dict(
    role="代码片段库：SQLite 两列表（title 主键 + code），外加标题模糊匹配。",
    why="与词库同样的存储层，但多了一层「标题命中高亮」的匹配逻辑，适合对照着读。",
    first=[("Filter", "看标题模糊匹配的边界判定"),
           ("Score", "看命中位置怎么排序"),
           ("Add", "看新增一条片段")],
    takeaway="代码片段是 C# 版自己的库，和 Python 版不共用文件。"),
"Recent.cs": dict(
    role="最近使用：与 Python 版共用 recent.json，扁平字典 {名字: 递增序号}。",
    why="78 行、4 个方法，是理解「排序权重怎么持久化」的最小例子。",
    first=[("Note", "看选中一次就递增序号"),
           ("Rank", "看序号越大越靠前"),
           ("Flush", "看什么时候真正落盘")],
    takeaway="读坏了就当空、绝不抛异常——辅助功能不能拖垮主流程。"),

"VbeBackend.cs": dict(
    role="VBE 后端：读光标快照、读模块全文、把补全写回去、执行改一行的操作。",
    why="全项目唯一大面积碰 COM 的地方，42 个方法都在 UI 线程调用。看懂 Snapshot() 只管光标那行，就看懂了一半。",
    first=[("Snapshot", "看快照只取光标那行的文本与显示列"),
           ("TryGetSelection", "看 ByRef 出参怎么拿真实选区"),
           ("FormControls", "看窗体控件名为什么只能从 Designer 取")],
    takeaway="任何 COM 异常都被吞掉并返回 false；打字中不要常驻 COM 引用。"),
"Engine.cs": dict(
    role="候选引擎：攒池子、模糊匹配、排序、决定弹不弹、选完之后补什么。",
    why="不引用 WinForms，窗口靠回调接进来。它是「纯逻辑」与「真界面」之间的桥，读它之前先读 Rules 与 Parser。",
    first=[("EnsurePool", "看半截词被挖成空格后再解析的原因"),
           ("Trigger", "看该不该弹的全部判断"),
           ("AcceptName", "看选中之后写回什么")],
    takeaway="EnsurePool 解析前把半截词挖成空格，这个细节直接造成过「函数头前补 Private 不提示」。"),
"App.cs": dict(
    role="主流程：全局钩子 + 自适应轮询 + 按键路由 + 改一行类操作。",
    why="唯一有权决定「吃不吃一个键」的地方。放到第 4 站最后读，因为要同时懂 Win32、COM、判据三层。",
    first=[("HookProcCore", "看钩子线程只读内存状态"),
           ("PollOnce", "看 UI 线程读快照、判重排、搬光标"),
           ("NoteSave", "看 Ctrl+S 那一刻把保存前的行存成锚")],
    takeaway="快照不是原子的（先文本后光标）；Ctrl+S 重排的锚必须用 VBE 保留的原选区列。"),
"Commands.cs": dict(
    role="行操作命令：Ctrl+D / Ctrl+K / Ctrl+/ / Ctrl+Enter / Ctrl+↑↓。",
    why="会改用户代码，所以每条都必须「先读真区间、再整段替换」。对照 VbeBackend 看它怎么保持一字不差。",
    first=[("ToggleComment", "看用 VBE 自己的命令按 ID 找控件"),
           ("DuplicateLine", "看复制当前行"),
           ("LeadingWs", "看缩进怎么被保留")],
    takeaway="Ctrl+/ 走 FindControl(1, 192|2552)，按 ID 不按 Caption；失败绝不补发裸键。"),
"PopupForm.cs": dict(
    role="候选窗：无边框、置顶、永不抢焦点、全部自绘。",
    why="四条不能动的规矩都在这里。读懂 OnPaint 与命中高亮，才知道为什么不用 ListView 现成控件。",
    first=[("OnPaint", "看一个绘制循环画完整张候选表"),
           ("WndProc", "看怎么做到永不抢焦点"),
           ("SetData", "看候选数据进来后怎么分页与选中")],
    takeaway="抢一下焦点，用户在 VBE 的输入就断了；字体要去 Light/Thin，字号比 VBE 小一号。"),
"SettingsForm.cs": dict(
    role="设置界面：常用 / 快捷键 / 候选 / 外观 / 词库 五页。",
    why="界面代码里藏着最典型的坑：页签只刷第一个、ItemChecked 延迟通知。读它学会「凡见 _xxx[0] 都要问其余元素谁管」。",
    first=[("SelectPage", "看切页时刷全部词库页签"),
           ("ReloadBank", "看 ListView 重建"),
           ("SaveAll", "看只写改过的项")],
    takeaway="勾上 = 启用标题所述行为；NO_XXX 这类键走反相映射，没实现的开关不要摆上去。"),
"WordEditForm.cs": dict(
    role="添加 / 编辑词条小窗：一行一个批量粘贴，用户不用再开 DB Browser。",
    why="312 行、7 个方法，是「一个小窗体从搭建到写库」的完整闭环，适合作为读大窗体前的热身。",
    first=[("BuildUi", "看控件怎么摆"),
           ("Fill", "看编辑已有词条时回填"),
           ("Commit", "看批量写回词库")],
    takeaway="关键字库与函数库问的问题不同：后者还要问「补括号」和「跟空格」。"),
"SnippetForm.cs": dict(
    role="代码片段窗（Ctrl+Q）：左侧自绘列表 + 右侧标题与代码编辑。",
    why="和 PopupForm 一样是自绘列表，但它是独立窗口、允许拿焦点，正好对照着看两种窗体的差别。",
    first=[("Refilter", "看搜索命中怎么逐字符高亮"),
           ("OnPaint", "看列表绘制与首行预览"),
           ("DoInsert", "看把片段插回 VBE")],
    takeaway="自绘是为了命中高亮；Confirm 这种危险操作要先问用户。"),
"InfoBox.cs": dict(
    role="托盘菜单的反馈框（模态）。",
    why="70 行、2 个方法。为什么不用 ShowBalloonTip：Win11 下气泡可能一个像素都不显示。",
    first=[("Show", "看模态反馈框怎么弹出")],
    takeaway="系统通知不可靠时，自己画一个确定能看见的框。"),
"UiTheme.cs": dict(
    role="设置界面与词条编辑窗共用的观感：配色、按钮、复选框、分组标题。",
    why="102 行、5 个方法，看「把样式集中管理」的最小实践。",
    first=[("Primary", "看主按钮统一样式"),
           ("Check", "看复选框统一样式"),
           ("Hex", "看颜色与配置字符串互转")],
    takeaway="与 PopupForm / SnippetForm 的自绘体系分开，互不污染。"),
"IconAssets.cs": dict(
    role="程序图标：从 exe 内嵌的 app.ico 里按尺寸挑最合适的一帧。",
    why="67 行、3 个方法。单文件发布时 exe 旁边没有散落资源，这解释了为什么要内嵌。",
    first=[("Get", "看从程序集资源流读图标"),
           ("Tray", "看按尺寸挑最合适的一帧"),
           ("Window", "看窗口图标用哪一帧")],
    takeaway="发布成单文件后，「旁边那个文件」都不存在了，资源必须内嵌。"),

"SelfTest.cs": dict(
    role="自检：587 条纯逻辑断言，结果写进 selftest.log。",
    why="这是全项目最好的「标准答案」：每条断言都写明了某个判据应该是什么行为。改代码前先跑它。",
    first=[("Run", "看断言怎么分组与计数"),
           ("Ok", "看一条断言怎么写"),
           ("TestParser", "看解析相关的断言组")],
    takeaway="用文件而不是 Console：程序是 WinExe，没有控制台可看。"),
"Probes.Support.cs": dict(
    role="探针脚手架：打字、发组合键、重建临时代码块、读行、清场。",
    why="所有真机探针都靠它摆场景。先读它，后面 9 个探针文件就只是在写「断言」而已。",
    first=[("TypeTyping", "看默认 Unicode 注入 + 补记打字"),
           ("SendCombo", "看真按 Ctrl+S 这类组合键"),
           ("RebuildScaffold", "看临时代码块怎么搭与还原")],
    takeaway="输入法模式位不可信；默认走 Unicode 注入，验真 VK 才开 PROBE_VK。"),
"Probes.Pure.cs": dict(
    role="合成后端探针：不碰真 VBE，用假模块跑完整 Engine（--enginetest / --sprobe）。",
    why="最便宜的第二层验证：把候选池、作用域、关键字场合全在内存里跑一遍。",
    first=[("EngineTest", "看一条用例怎么断言候选名单"),
           ("AddForm", "看假后端里造一个带控件的窗体"),
           ("SnippetTest", "看片段库的离屏断言")],
    takeaway="A/B 口子（如 NO_PROC_HEADER_FIX）在这里能确定性复现用户现象。"),
"Probes.Host.cs": dict(
    role="真后端只读探针：--formprobe / --hdrprobe / --pooldump，读真 VBE 但不改代码。",
    why="介于合成与真机之间：用真 COM 数据、内存副本判据，不动用户模块、不抢前台。",
    first=[("HdrProbe", "看函数头前补 Private 的真机只读断言"),
           ("FormProbe", "看控件名候选的真机断言"),
           ("BareProcHead", "看反向对照怎么造")],
    takeaway="只读探针也要有反向对照，否则「全 PASS」可能只是没验到。"),
"Probes.Key.cs": dict(
    role="键盘真机探针：--enterprobe / --hktest / --paretest，驱动真实低级钩子。",
    why="回车补结构、自动配对这类功能只能在真钩子上验。它展示了「强制造出落后快照」的诚实断言写法。",
    first=[("EnterProbe", "看 C3 用例怎么造出 want=False"),
           ("HkTest", "看钩子收键断言"),
           ("ParenTest", "看自动配对断言")],
    takeaway="造不出前置条件就报 SKIP，绝不带着假前提跑到底刷绿。"),
"Probes.Decl.cs": dict(
    role="声明与格式化真机探针：--declprobe / --fmtprobe（后者会真按 Ctrl+S）。",
    why="最复杂的一支：40 条声明用例 + 7 条格式化用例，覆盖中文全角与保存重排。",
    first=[("DeclProbe", "看 40 条用例表与两档采样窗口"),
           ("FmtProbe", "看真按 Ctrl+S 后的轨迹采样")],
    takeaway="采样只读 TryGetSelection / LineAt；碰快照缓存等于替产品抹掉 bug。"),
"Probes.Caret.cs": dict(
    role="光标真机探针：--carettest，验证显示列与字符列换算。",
    why="165 行、3 个方法，是「中文全角占 2 格」这条规矩最直接的真机证据。",
    first=[("CaretTest", "看显示列断言"),
           ("WholeWordIndex", "看整词边界的辅助计算")],
    takeaway="中文用例必须用显示列摆光标。"),
"Probes.Dump.cs": dict(
    role="现场快照探针：--dump / --probe / --pooldump，把模块全文与解析结果写成文本。",
    why="排 bug 时第一件事就是「看看现在到底是什么样」；--dump 还负责收尾逐字节比对。",
    first=[("DumpModule", "看模块全文怎么落盘（首行含光标位置）"),
           ("PoolDump", "看候选池与 curProc 等中间量"),
           ("WriteTest", "看写回操作的真机验证")],
    takeaway="按 VK 往真机打过字，收尾必须 --dump 逐字节比对；比对前先还原光标。"),
"Probes.Env.cs": dict(
    role="环境探针：--envtest，检查 VBE 是否打开、AccessVBOM、前台宿主等前置条件。",
    why="「红了先怀疑环境」这条铁律的落地：前置不满足一律 SKIP + rc=3。",
    first=[("StartupTest", "看开机自启前置与还原"),
           ("TrustTest", "看 AccessVBOM 前置检查")],
    takeaway="SKIP 不是失败，但它也绝不能算通过。"),
"Probes.Ui.cs": dict(
    role="界面探针：--uitest，离屏生成设置页与候选窗预览 PNG。",
    why="改外观不用盯屏幕：跑一次出图，页面与词库页签都能核。",
    first=[("UiTest", "看离屏渲染与出图"),
           ("DumpControls", "看控件树与词库页签行数反向断言")],
    takeaway="界面回归要有反向断言：模拟勾选必须能真的写库。"),
}

# 非文件页文案
INTRO = [
    ("这个项目到底在补什么", "VBE 自带列表只认「已经存在的成员」；你自己刚声明的变量、函数、窗体控件，它不提示。"
     "这个工具在后台读代码、解析出名字，在你打字时弹出自己的候选窗。"),
    ("技术栈：为什么只有 1 个 exe", "C# / .NET 10 WinExe + WinForms，零 NuGet 依赖；SQLite 直接 P/Invoke 系统 winsqlite3.dll；"
     "Office 全走 dynamic 后期绑定，因此不绑 Office 版本；发布成单文件 dist\\VbeComplete.exe。"),
    ("分层与两条线程", "纯判据（Rules/Editing/Parser/Scope）→ 引擎（Engine）→ 宿主（VbeBackend）→ 主流程（App）→ 界面（PopupForm）。"
     "钩子线程快但绝不碰 COM，UI 线程慢但能读真代码，两者靠内存快照接力。"),
]

OUTRO = [
    ("三层验证金字塔", "587 条纯逻辑自检（不依赖桌面）→ 合成后端探针（假模块跑真 Engine）→ 真机探针（真 VBE、真钩子、真按键）。"
     "每一层都有 A/B 开关：关掉修复必须能确定性复现用户现象。"),
    ("命令速查", "dotnet build src/VbeComplete.csproj -c Release；"
     "dist\\VbeComplete.exe --selftest（纯逻辑）；--enginetest（合成）；--declprobe / --fmtprobe / --hdrprobe（真机）；"
     "dotnet publish ... -o dist；check-cs.bat 一键跑全套 12 步。"),
    ("7 步上手路线", "① 跑 --selftest 看 587/0 ② 读 Rules.WordBefore（21 行）③ 在 SelfTest.cs 加一条断言跑通 "
     "④ 读 Fuzzy.GreedyPositions ⑤ 跑 --enginetest ⑥ 用 A/B 开关亲手复现一个 bug ⑦ 再进 App.cs。"),
]
