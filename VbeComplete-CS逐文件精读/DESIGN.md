# 设计说明（逐文件精读课件）

## 全局视觉约定

- 画布 1280×720；padding 上下左右各 60px；内容区 1160×600。
- 配色：主色 #2563EB / 深主色 #1E3A8A / 辅色 #06B6D4 / 强调 #F59E0B / 底色 #FFFFFF / 正文 #1E293B。
- 字体：Microsoft YaHei（中文）+ Arial（西文）；代码命令用 Consolas。
- 字号：封面 76 / 章节 52 / 页面主标 34 / 卡片小标 17-22 / 正文与表格 16 / 页脚 14。
- 视觉母题：标题下 1px 贯通细线；固定页脚（章节页/封面/目录/封底无页脚）。
- 素材：全部为结构化图形（Box / SVG 式布局），无外部图片；教学查读型，不使用装饰性配图。
- 版式循环：左浅色信息卡 + 右函数表 / 右深色信息卡 + 左函数表 / 双色卡 + 左函数表，逐页交替。

## 逐页字段表

| 页 | 类型 | 角色 | 节奏 | 版式 | 内容 | 视觉 | 禁止项 |
|---|---|---|---|---|---|---|---|
| P01 | cover | hero | peak | 大字报式 | 封面 | 超大主标题 | 禁止数据指标 / 禁止页脚 |
| P02 | catalog | supporting | valley | 左右不对称杂志式 | 目录 | 大字目录 + 六章清单 | 禁止统计块 / 禁止阅读顺序复述条 |
| P03 | section | hero | peak | 章节大字（深底） | 读之前：先建立地图 | 巨型章节编号 | 禁止加页脚 / 禁止章节内看点列表 |
| P04 | roadmap | hero | peak | 流程链 | 5 站阅读路线 | 五节点流程 + 结论条 | 禁止等宽卡片横排 |
| P05 | intro | supporting | valley | 对比式 | 这个项目到底在补什么 | 结构化图形 | 禁止抽象标题 / 禁止空容器 |
| P06 | intro | supporting | valley | 巨型数字+清单 | 技术栈：为什么只有 1 个 exe | 结构化图形 | 禁止抽象标题 / 禁止空容器 |
| P07 | intro | supporting | valley | 分层图+双线程卡 | 分层与两条线程 | 结构化图形 | 禁止抽象标题 / 禁止空容器 |
| P08 | section | hero | peak | 章节大字（深底） | 第 1 站：入口与地基 | 巨型章节编号 | 禁止加页脚 / 禁止章节内看点列表 |
| P09 | file | supporting | valley | 双色卡+左函数表 | Program.cs（1/1） | L3 母版眉标 + 函数表（12 行） | 禁止裸表、禁止缩字号塞内容、禁止函数说明留空 |
| P10 | file | supporting | valley | 左信息卡+右函数表 | Settings.cs（1/1） | L3 母版眉标 + 函数表（12 行） | 禁止裸表、禁止缩字号塞内容、禁止函数说明留空 |
| P11 | file | supporting | valley | 右深卡+左函数表 | Log.cs（1/1） | L3 母版眉标 + 函数表（12 行） | 禁止裸表、禁止缩字号塞内容、禁止函数说明留空 |
| P12 | file | supporting | valley | 双色卡+左函数表 | Native.cs（1/3） | L3 母版眉标 + 函数表（12 行） | 禁止裸表、禁止缩字号塞内容、禁止函数说明留空 |
| P13 | file | supporting | valley | 左信息卡+右函数表 | Native.cs（2/3） | L3 母版眉标 + 函数表（12 行） | 禁止裸表、禁止缩字号塞内容、禁止函数说明留空 |
| P14 | file | supporting | valley | 右深卡+左函数表 | Native.cs（3/3） | L3 母版眉标 + 函数表（12 行） | 禁止裸表、禁止缩字号塞内容、禁止函数说明留空 |
| P15 | file | supporting | valley | 双色卡+左函数表 | ComCtx.cs（1/2） | L3 母版眉标 + 函数表（12 行） | 禁止裸表、禁止缩字号塞内容、禁止函数说明留空 |
| P16 | file | supporting | valley | 左信息卡+右函数表 | ComCtx.cs（2/2） | L3 母版眉标 + 函数表（12 行） | 禁止裸表、禁止缩字号塞内容、禁止函数说明留空 |
| P17 | file | supporting | valley | 右深卡+左函数表 | ComCall.cs（1/1） | L3 母版眉标 + 函数表（12 行） | 禁止裸表、禁止缩字号塞内容、禁止函数说明留空 |
| P18 | file | supporting | valley | 双色卡+左函数表 | AutoStart.cs（1/2） | L3 母版眉标 + 函数表（12 行） | 禁止裸表、禁止缩字号塞内容、禁止函数说明留空 |
| P19 | file | supporting | valley | 左信息卡+右函数表 | AutoStart.cs（2/2） | L3 母版眉标 + 函数表（12 行） | 禁止裸表、禁止缩字号塞内容、禁止函数说明留空 |
| P20 | file | supporting | valley | 右深卡+左函数表 | VbaTrust.cs（1/1） | L3 母版眉标 + 函数表（12 行） | 禁止裸表、禁止缩字号塞内容、禁止函数说明留空 |
| P21 | file | supporting | valley | 双色卡+左函数表 | TrayPin.cs（1/1） | L3 母版眉标 + 函数表（12 行） | 禁止裸表、禁止缩字号塞内容、禁止函数说明留空 |
| P22 | section | hero | peak | 章节大字（深底） | 第 2 站：纯逻辑判据 | 巨型章节编号 | 禁止加页脚 / 禁止章节内看点列表 |
| P23 | file | supporting | valley | 右深卡+左函数表 | Rules.cs（1/4） | L3 母版眉标 + 函数表（12 行） | 禁止裸表、禁止缩字号塞内容、禁止函数说明留空 |
| P24 | file | supporting | valley | 双色卡+左函数表 | Rules.cs（2/4） | L3 母版眉标 + 函数表（12 行） | 禁止裸表、禁止缩字号塞内容、禁止函数说明留空 |
| P25 | file | supporting | valley | 左信息卡+右函数表 | Rules.cs（3/4） | L3 母版眉标 + 函数表（12 行） | 禁止裸表、禁止缩字号塞内容、禁止函数说明留空 |
| P26 | file | supporting | valley | 右深卡+左函数表 | Rules.cs（4/4） | L3 母版眉标 + 函数表（12 行） | 禁止裸表、禁止缩字号塞内容、禁止函数说明留空 |
| P27 | file | supporting | valley | 双色卡+左函数表 | Fuzzy.cs（1/1） | L3 母版眉标 + 函数表（12 行） | 禁止裸表、禁止缩字号塞内容、禁止函数说明留空 |
| P28 | file | supporting | valley | 左信息卡+右函数表 | Editing.cs（1/4） | L3 母版眉标 + 函数表（12 行） | 禁止裸表、禁止缩字号塞内容、禁止函数说明留空 |
| P29 | file | supporting | valley | 右深卡+左函数表 | Editing.cs（2/4） | L3 母版眉标 + 函数表（12 行） | 禁止裸表、禁止缩字号塞内容、禁止函数说明留空 |
| P30 | file | supporting | valley | 双色卡+左函数表 | Editing.cs（3/4） | L3 母版眉标 + 函数表（12 行） | 禁止裸表、禁止缩字号塞内容、禁止函数说明留空 |
| P31 | file | supporting | valley | 左信息卡+右函数表 | Editing.cs（4/4） | L3 母版眉标 + 函数表（12 行） | 禁止裸表、禁止缩字号塞内容、禁止函数说明留空 |
| P32 | file | supporting | valley | 右深卡+左函数表 | Parser.cs（1/2） | L3 母版眉标 + 函数表（12 行） | 禁止裸表、禁止缩字号塞内容、禁止函数说明留空 |
| P33 | file | supporting | valley | 双色卡+左函数表 | Parser.cs（2/2） | L3 母版眉标 + 函数表（12 行） | 禁止裸表、禁止缩字号塞内容、禁止函数说明留空 |
| P34 | file | supporting | valley | 左信息卡+右函数表 | Scope.cs（1/2） | L3 母版眉标 + 函数表（12 行） | 禁止裸表、禁止缩字号塞内容、禁止函数说明留空 |
| P35 | file | supporting | valley | 右深卡+左函数表 | Scope.cs（2/2） | L3 母版眉标 + 函数表（12 行） | 禁止裸表、禁止缩字号塞内容、禁止函数说明留空 |
| P36 | file | supporting | valley | 双色卡+左函数表 | VbaBuiltins.cs（1/1） | L3 母版眉标 + 函数表（12 行） | 禁止裸表、禁止缩字号塞内容、禁止函数说明留空 |
| P37 | file | supporting | valley | 左信息卡+右函数表 | BuiltinTables.g.cs（1/1） | L3 母版眉标 + 函数表（12 行） | 禁止裸表、禁止缩字号塞内容、禁止函数说明留空 |
| P38 | section | hero | peak | 章节大字（深底） | 第 3 站：数据与词库 | 巨型章节编号 | 禁止加页脚 / 禁止章节内看点列表 |
| P39 | file | supporting | valley | 双色卡+左函数表 | Sqlite.cs（1/1） | L3 母版眉标 + 函数表（12 行） | 禁止裸表、禁止缩字号塞内容、禁止函数说明留空 |
| P40 | file | supporting | valley | 左信息卡+右函数表 | WordLibrary.cs（1/1） | L3 母版眉标 + 函数表（12 行） | 禁止裸表、禁止缩字号塞内容、禁止函数说明留空 |
| P41 | file | supporting | valley | 右深卡+左函数表 | WordBank.cs（1/1） | L3 母版眉标 + 函数表（12 行） | 禁止裸表、禁止缩字号塞内容、禁止函数说明留空 |
| P42 | file | supporting | valley | 双色卡+左函数表 | SnippetStore.cs（1/2） | L3 母版眉标 + 函数表（12 行） | 禁止裸表、禁止缩字号塞内容、禁止函数说明留空 |
| P43 | file | supporting | valley | 左信息卡+右函数表 | SnippetStore.cs（2/2） | L3 母版眉标 + 函数表（12 行） | 禁止裸表、禁止缩字号塞内容、禁止函数说明留空 |
| P44 | file | supporting | valley | 右深卡+左函数表 | Recent.cs（1/1） | L3 母版眉标 + 函数表（12 行） | 禁止裸表、禁止缩字号塞内容、禁止函数说明留空 |
| P45 | section | hero | peak | 章节大字（深底） | 第 4 站：宿主、引擎与界面 | 巨型章节编号 | 禁止加页脚 / 禁止章节内看点列表 |
| P46 | file | supporting | valley | 左信息卡+右函数表 | VbeBackend.cs（1/4） | L3 母版眉标 + 函数表（12 行） | 禁止裸表、禁止缩字号塞内容、禁止函数说明留空 |
| P47 | file | supporting | valley | 右深卡+左函数表 | VbeBackend.cs（2/4） | L3 母版眉标 + 函数表（12 行） | 禁止裸表、禁止缩字号塞内容、禁止函数说明留空 |
| P48 | file | supporting | valley | 双色卡+左函数表 | VbeBackend.cs（3/4） | L3 母版眉标 + 函数表（12 行） | 禁止裸表、禁止缩字号塞内容、禁止函数说明留空 |
| P49 | file | supporting | valley | 左信息卡+右函数表 | VbeBackend.cs（4/4） | L3 母版眉标 + 函数表（12 行） | 禁止裸表、禁止缩字号塞内容、禁止函数说明留空 |
| P50 | file | supporting | valley | 右深卡+左函数表 | Engine.cs（1/2） | L3 母版眉标 + 函数表（12 行） | 禁止裸表、禁止缩字号塞内容、禁止函数说明留空 |
| P51 | file | supporting | valley | 双色卡+左函数表 | Engine.cs（2/2） | L3 母版眉标 + 函数表（12 行） | 禁止裸表、禁止缩字号塞内容、禁止函数说明留空 |
| P52 | file | supporting | valley | 左信息卡+右函数表 | App.cs（1/3） | L3 母版眉标 + 函数表（12 行） | 禁止裸表、禁止缩字号塞内容、禁止函数说明留空 |
| P53 | file | supporting | valley | 右深卡+左函数表 | App.cs（2/3） | L3 母版眉标 + 函数表（12 行） | 禁止裸表、禁止缩字号塞内容、禁止函数说明留空 |
| P54 | file | supporting | valley | 双色卡+左函数表 | App.cs（3/3） | L3 母版眉标 + 函数表（12 行） | 禁止裸表、禁止缩字号塞内容、禁止函数说明留空 |
| P55 | file | supporting | valley | 左信息卡+右函数表 | Commands.cs（1/1） | L3 母版眉标 + 函数表（12 行） | 禁止裸表、禁止缩字号塞内容、禁止函数说明留空 |
| P56 | file | supporting | valley | 右深卡+左函数表 | PopupForm.cs（1/3） | L3 母版眉标 + 函数表（12 行） | 禁止裸表、禁止缩字号塞内容、禁止函数说明留空 |
| P57 | file | supporting | valley | 双色卡+左函数表 | PopupForm.cs（2/3） | L3 母版眉标 + 函数表（12 行） | 禁止裸表、禁止缩字号塞内容、禁止函数说明留空 |
| P58 | file | supporting | valley | 左信息卡+右函数表 | PopupForm.cs（3/3） | L3 母版眉标 + 函数表（12 行） | 禁止裸表、禁止缩字号塞内容、禁止函数说明留空 |
| P59 | file | supporting | valley | 右深卡+左函数表 | SettingsForm.cs（1/4） | L3 母版眉标 + 函数表（12 行） | 禁止裸表、禁止缩字号塞内容、禁止函数说明留空 |
| P60 | file | supporting | valley | 双色卡+左函数表 | SettingsForm.cs（2/4） | L3 母版眉标 + 函数表（12 行） | 禁止裸表、禁止缩字号塞内容、禁止函数说明留空 |
| P61 | file | supporting | valley | 左信息卡+右函数表 | SettingsForm.cs（3/4） | L3 母版眉标 + 函数表（12 行） | 禁止裸表、禁止缩字号塞内容、禁止函数说明留空 |
| P62 | file | supporting | valley | 右深卡+左函数表 | SettingsForm.cs（4/4） | L3 母版眉标 + 函数表（12 行） | 禁止裸表、禁止缩字号塞内容、禁止函数说明留空 |
| P63 | file | supporting | valley | 双色卡+左函数表 | WordEditForm.cs（1/1） | L3 母版眉标 + 函数表（12 行） | 禁止裸表、禁止缩字号塞内容、禁止函数说明留空 |
| P64 | file | supporting | valley | 左信息卡+右函数表 | SnippetForm.cs（1/4） | L3 母版眉标 + 函数表（12 行） | 禁止裸表、禁止缩字号塞内容、禁止函数说明留空 |
| P65 | file | supporting | valley | 右深卡+左函数表 | SnippetForm.cs（2/4） | L3 母版眉标 + 函数表（12 行） | 禁止裸表、禁止缩字号塞内容、禁止函数说明留空 |
| P66 | file | supporting | valley | 双色卡+左函数表 | SnippetForm.cs（3/4） | L3 母版眉标 + 函数表（12 行） | 禁止裸表、禁止缩字号塞内容、禁止函数说明留空 |
| P67 | file | supporting | valley | 左信息卡+右函数表 | SnippetForm.cs（4/4） | L3 母版眉标 + 函数表（12 行） | 禁止裸表、禁止缩字号塞内容、禁止函数说明留空 |
| P68 | file | supporting | valley | 右深卡+左函数表 | InfoBox.cs（1/1） | L3 母版眉标 + 函数表（12 行） | 禁止裸表、禁止缩字号塞内容、禁止函数说明留空 |
| P69 | file | supporting | valley | 双色卡+左函数表 | UiTheme.cs（1/1） | L3 母版眉标 + 函数表（12 行） | 禁止裸表、禁止缩字号塞内容、禁止函数说明留空 |
| P70 | file | supporting | valley | 左信息卡+右函数表 | IconAssets.cs（1/1） | L3 母版眉标 + 函数表（12 行） | 禁止裸表、禁止缩字号塞内容、禁止函数说明留空 |
| P71 | section | hero | peak | 章节大字（深底） | 第 5 站：验证体系与上手 | 巨型章节编号 | 禁止加页脚 / 禁止章节内看点列表 |
| P72 | file | supporting | valley | 双色卡+左函数表 | SelfTest.cs（1/2） | L3 母版眉标 + 函数表（12 行） | 禁止裸表、禁止缩字号塞内容、禁止函数说明留空 |
| P73 | file | supporting | valley | 左信息卡+右函数表 | SelfTest.cs（2/2） | L3 母版眉标 + 函数表（12 行） | 禁止裸表、禁止缩字号塞内容、禁止函数说明留空 |
| P74 | file | supporting | valley | 右深卡+左函数表 | Probes.Support.cs（1/1） | L3 母版眉标 + 函数表（12 行） | 禁止裸表、禁止缩字号塞内容、禁止函数说明留空 |
| P75 | file | supporting | valley | 双色卡+左函数表 | Probes.Pure.cs（1/1） | L3 母版眉标 + 函数表（12 行） | 禁止裸表、禁止缩字号塞内容、禁止函数说明留空 |
| P76 | file | supporting | valley | 左信息卡+右函数表 | Probes.Host.cs（1/2） | L3 母版眉标 + 函数表（12 行） | 禁止裸表、禁止缩字号塞内容、禁止函数说明留空 |
| P77 | file | supporting | valley | 右深卡+左函数表 | Probes.Host.cs（2/2） | L3 母版眉标 + 函数表（12 行） | 禁止裸表、禁止缩字号塞内容、禁止函数说明留空 |
| P78 | file | supporting | valley | 双色卡+左函数表 | Probes.Key.cs（1/1） | L3 母版眉标 + 函数表（12 行） | 禁止裸表、禁止缩字号塞内容、禁止函数说明留空 |
| P79 | file | supporting | valley | 左信息卡+右函数表 | Probes.Decl.cs（1/1） | L3 母版眉标 + 函数表（12 行） | 禁止裸表、禁止缩字号塞内容、禁止函数说明留空 |
| P80 | file | supporting | valley | 右深卡+左函数表 | Probes.Caret.cs（1/1） | L3 母版眉标 + 函数表（12 行） | 禁止裸表、禁止缩字号塞内容、禁止函数说明留空 |
| P81 | file | supporting | valley | 双色卡+左函数表 | Probes.Dump.cs（1/1） | L3 母版眉标 + 函数表（12 行） | 禁止裸表、禁止缩字号塞内容、禁止函数说明留空 |
| P82 | file | supporting | valley | 左信息卡+右函数表 | Probes.Env.cs（1/1） | L3 母版眉标 + 函数表（12 行） | 禁止裸表、禁止缩字号塞内容、禁止函数说明留空 |
| P83 | file | supporting | valley | 右深卡+左函数表 | Probes.Ui.cs（1/1） | L3 母版眉标 + 函数表（12 行） | 禁止裸表、禁止缩字号塞内容、禁止函数说明留空 |
| P84 | outro | supporting | valley | 金字塔 | 三层验证金字塔 | 结构化图形 | 禁止把结尾内容塞进内容页 |
| P85 | outro | supporting | valley | 命令清单 | 命令速查 | 结构化图形 | 禁止把结尾内容塞进内容页 |
| P86 | outro | supporting | valley | 七步流程 | 7 步上手路线 | 结构化图形 | 禁止把结尾内容塞进内容页 |
| P87 | ending | hero | peak | 深底金句 + 落款 | 收束金句 | 大字两行 | 禁止只写「谢谢」 |
