# AddInLoader v1.0.0 升级日志

## 版本信息
- 版本号：v1.0.0
- 版本描述：初始版本
- 发布日期：2026-11-23
- Python 版本：3.12.10
- PyInstaller 版本：6.22.3
- 依赖版本：comtypes 1.5.0

## 改动内容
- 自动发现已登记加载项（Options 的 OPEN/OPENn）与默认 AddIns 目录下的自定义 xlam/xla，排除 Office 内置项与临时文件；勾选默认全选、可取消，支持刷新。
- 常驻扫描已打开的 Excel 窗口，经官方 Native OM 接口附着后按勾选装载，幂等不重复；新增项自动纳入发现列表。
- 临时启动：不写注册表、不写长期登记；该项留在列表中由装载器持续装载，Excel 关了重开仍会装上，装载器退出或取消勾选后失效。
- 长期启动：把 xlam 复制到统一管理目录（%APPDATA%\AddInLoader），按选项决定是否注册该目录为 Excel 受信任位置，并写入 OPEN 登记；下次 Excel 启动自动加载。同名本工具副本覆盖升级，非本工具同名文件拒绝覆盖。
- 卸载：撤销 OPEN 登记、删除目录副本、清加载项列表记录，并同步关闭当前 Excel 中的该加载项；目录内无其它副本时撤销受信任位置登记。
- 生命周期：关闭或点 X 隐藏到系统托盘继续驻留，托盘右键退出；支持开机启动。
- 主界面复用安装器式布局：安装信息与选项左右并排，日志区固定可见。

## 交付前修复（长期启动加载失败）
- 修复 OPEN 登记格式：由 `/A "路径"` 改为 Excel 原生格式 `"路径"`。`AddIns.Add` 前缀导致 Excel 不认该登记，加载项出现在列表但 `Installed=False`、VBE 中不加载。
- 修复当前会话装载方式：`provision` 长期分支弃用在本机必失败的 `AddIns.Add()`（该调用抛异常时会令 Excel 退出反向清除 OPEN 登记），改用与临时启动一致的 `Workbooks.Open`；已由 OPEN 自动加载的项识别为 already，不重复装载。
- 转长期新增前置检查：检测到 EXCEL.EXE 正在运行时弹窗提示先保存并关闭所有 Excel 窗口再继续。运行中的 Excel 退出会用内存状态覆盖 OPEN 登记，导致本次长期启动失效；关闭后再登记可一次成功（已实测重启后 `Installed=True`）。
- 关闭 Excel 改为装载器代劳：确认后由装载器强制结束所有 EXCEL.EXE 进程（等同任务管理器“结束任务”），无需用户手动去任务管理器；关干净后再继续转长期。未保存内容会丢失，弹窗已提示。
- 卸载长期启动项同样先关 Excel：确认后强制结束所有 Excel 再撤销登记、删副本，避免运行中 Excel 退出把 OPEN 登记写回（实测曾出现“复查：登记和加载状态被重新写回”）。

## 缺陷修复（托盘图标）
- 问题：explorer 重启后托盘图标消失且不恢复。原实现仅在程序启动时注册一次图标，explorer 重启会清空所有托盘图标。
- 修复：监听 `TaskbarCreated` 系统广播（explorer 重启后必发），收到后重新注册图标。已实测注册→删除→广播→重建全链路通过。

## 补发修订（原 v1.0.0 tag 重写并强推远端）
- 关于窗在「发布页」右侧新增「检查更新」按钮：后台请求 GitHub releases/latest，对比本地版本弹窗提示并引导下载；网络或解析失败时询问后降级打开发布页。仅此一项 UI 改动，版本号仍为 v1.0.0，原有装载逻辑不变。
- 原 v1.0.0 已发布 exe 的版本识别依赖文件名 `AddInLoader_v1.0.0.exe`，本次重打包保持该命名，供 v1.0.1 的高版本接管逻辑识别。
- 许可文件归位：LICENSE / Python-LICENSE.txt / THIRD_PARTY_NOTICES.md / third_party_licenses/ / third_party_licenses.zip 移回项目根目录；仓库保留 MIT LICENSE。
- 检查更新补 tags 降级：远端只推了 Git tag、未在 GitHub 创建 Release 时，`releases/latest` 返回 404，导致点击「检查更新」弹出「HTTP Error 404: Not Found」。现改为 404 时降级读取 `/tags` 并取**版本号最大**的 tag（不依赖 API 返回顺序，兼容跳号如 v1.3.2 / v2.0.0），本地版本不低于远端统一提示「当前已是最新版 vX」。v1.0.0 与 v1.0.1 均已同步。
- 启动静默检查更新（v1.0.1 同步回移）：每次运行延迟 0.6 秒后台检查一次；无新版或网络失败完全不弹窗；发现新版弹「检测到新版本 vX，当前版本 vY。是否前往发布页下载？」，用户确认才打开发布页，取消则无动作。手动检查更新入口保持原样。

## 补发修订（启动崩溃修复）
- 修复启动阶段崩溃：Excel 加载项列表对 `%APPDATA%\Microsoft\AddIns` 库目录内的加载项只登记裸文件名（如 `OPEN2="AttachLink.xlam"`），`registered_at()` 用该裸文件名与完整路径比较，永远匹配不上而返回 `None`；`refresh_install_state()` 随后对 `None` 解包，抛出 `TypeError: cannot unpack non-iterable NoneType object`，程序在绘制主界面时退出。
- 修复方式：集合推导先过滤 `None` 再计数，不再对返回值解包；不改动其它装载逻辑。
- 复现与验证：用 AttachLink 的裸文件名登记格式对注册表只读复现，修复前必现 TypeError、修复后正常；`py_compile` 通过，回归脚本全部通过，`--selfcheck` 退出 0。
- 后续版本同步同一处修复，修复记录只保留在本日志。

## 四版本统一调整：移除一键升级模块
同日稍晚，因使用方反馈只需打开发布页手动下载、不再需要程序内自动替换，三个版本一并重写（不改版本号，覆盖原 tag/Release/附件）。此项为各版本共同改动，不是某一版独有。

- 删除 `download_file()`、`write_upgrade_bat()`、`start_upgrade_bat()`、`_bat_quote()`、`run_upgrade()` 及 `_UPDATE_ASSET`。
- `fetch_latest_release()` 只读清单的 `version` 与 `page`，不再解析 `asset`/`sha256`。
- 手动「检查更新」与启动静默检查的弹窗统一为「是否前往发布页下载？」，确认后打开发布页；删去代理提示。
- 回归脚本相应调整，83 项全部 PASS。

## 文件清单
- AddInLoader_v1.0.0.exe（SHA-256 `0e2edbf47a7be176a33ffe92d46c6b093fe4525eb14ca6c440c127377cdc2ed3`，12,236,579 字节）
- AddInLoader_v1.0.0.py
- upgrade_log_v1.0.0.md
