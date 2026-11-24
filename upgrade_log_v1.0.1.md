# AddInLoader v1.0.1 升级日志

版本描述：安全加固 + 可靠性补全 + 计划任务开机启动 + 版本仲裁 + 检查更新。后续统一调整见文末。

## 修复清单（fix-list）

| # | 级别 | 改动 | 函数/结构锚点 | 说明 |
|---|---|---|---|---|
| 1 | P0 | 宿主进程校验 | `window_process_name()`、`find_excel_doc_windows()` | 枚举 `XLMAIN/EXCEL7` 后复核所属进程必须是 `EXCEL.EXE`；WPS 同类名宿主跳过，识别不了也跳过。 |
| 2 | P0 | 路径可信四态，临时/长期共用 | `TRUST_*`、`trust_state()`、`xlam_allowed()`、`provision()`、`build_list()` | 统一为「受信任位置 / 会话临时信任 / 同名异路径 / 未受信任」四态；显式选择或 `--xlam` 为会话临时信任，退出即失效；同名异路径副本返回 `foreign`，不打开第二份、不执行宏。列表显示四态。 |
| 3 | P1 | `_TEMP_WBS` 用完释放 + 判重加固 | `provision()`、`unload_from_excel()`、`_verify_unloaded()` | 删除 `_TEMP_WBS`，不再长期持有加载项工作簿 COM 引用；按文件名取回工作簿并核对 `FullName` 后关闭。判重改为「在位 + 路径一致」，同名异路径直接拒绝。 |
| 4 | P1 | 去掉秒杀 Excel：等待 → 超时询问 | `_wait_excel_close()`、`ensure_excel_closed()`、`close_excel_processes()`、`on_register()`、`on_unregister()` | 卸载/长期启动前提示保存并轮询最多 45 秒；Excel 自行退出则继续。超时后默认明确询问「检测到 Excel 可能回写加载项状态，是否强制结束？」，确认后才 `TerminateProcess`；用户拒绝则取消操作。高级选项「卸载时允许强制结束进程」改为超时自动强制（默认关）。 |
| 5 | P2 | 落盘日志 / selfcheck 哈希 / 回归脚本 | `write_run_log()`、`append_log()`、`selfcheck()`、`tests/test_v101.py` | 日志追加到 `%APPDATA%\AddInLoader\addinloader.log`（1 MB 滚动）；自检打印源文件/exe 的 SHA-256 与超时强制选项；新增不依赖 Excel 的回归脚本。 |
| 6 | P1 | 计划任务开机启动 + 自动取消 | `TASK_NAME`、`_schtasks()`、`task_info()`、`set_autostart()`、`migrate_autostart()`、`autostart_health_fix()`、`sync_autostart_with_targets()` | 由 HKCU Run 改为计划任务 `AddInLoader Background`：`ONLOGON` + `/IT`（仅用户登录时运行）+ `/RL HIGHEST`（最高权限），登录不再弹 UAC。旧 Run 键启动时自动迁移并删除；任务路径漂移/权限降级/登录类型错误自动重建；程序文件被删除自动删除任务；加载项被取消勾选或文件删除导致**可装载目标归零**时自动取消开机启动（启动后 30 秒保护期）。计划任务不可用时回退 Run 键并在日志明确提示。 |
| 7 | P1 | 多版本仲裁：高版本接管低版本 | `version_tuple()`、`exe_version()`、`iter_loader_processes()`、`running_loader_versions()`、`take_over_older_instances()`、`clear_lower_autostart()` | 启动时枚举运行中的 `AddInLoader*.exe` 并从文件名解析版本。仅当确认存在**更低版本**实例时结束其进程，并清除指向该低版本 exe 的开机启动（旧 Run 键 + 计划任务），然后自己接管 mutex。存在同版本/更高版本/版本无法识别的实例时一律让位（沿用原单实例逻辑），避免误杀。`running_loader_versions()` 排除自身 PID 与自身父进程 PID：PyInstaller onefile 的 bootloader 父进程与子进程同名同版本，不排除会被误判为「同版本已在运行」而让位，继而低版本持有的 mutex 挡住高版本，表现为双击高版本无反应。mutex 名保持不变，因此低版本先启动、高版本后启动的接管由本项负责；高版本在跑时低版本自然退出。 |
| 8 | P2 | 检查更新按钮 | `fetch_latest_release()`、`compare_versions()`、`LoaderApp.check_for_update()` | 关于窗「发布页」右侧新增「检查更新」：后台线程请求 GitHub releases/latest，对比本地版本后弹窗提示并引导打开发布页；网络或解析失败时询问后降级打开发布页。同步回移到 v1.0.0 exe。 |
| 9 | P1 | 检查更新 404 降级读 tags | `TAGS_API`、`fetch_latest_release()`、`LoaderApp.check_for_update()` | 远端只推了 Git tag、未创建 GitHub Release 时，`releases/latest` 返回 404，弹「HTTP Error 404」。现 404 时降级读取 `/tags` 并取**版本号最大**的 tag（不依赖 API 顺序，兼容跳号 v1.3.2 / v2.0.0）；有新版引导前往发布页下载，本地不低于远端提示「当前已是最新版 vX」。v1.0.0 同步回移。原一键升级模块已在四版本统一移除。 |
| 10 | P1 | ~~一键升级~~（已移除） | `TREE_API`、`RAW_BASE`、`remote_exe_path()`、`download_file()`、`write_upgrade_bat()`、`start_upgrade_bat()`、`LoaderApp.run_upgrade()` | 此条为版本沿革记录：当时实现过 git tree 定位 + raw 下载 + bat 替换脚本的一键升级。**该模块已在后续统一调整中从全部版本移除**，现改为引导发布页手动下载。 |
| 11 | P1 | 启动静默检查更新 | `silent_update_target()`、`LoaderApp.silent_check_update()` | 每次运行延迟 0.6 秒后台静默检查一次版本；无新版或网络失败完全不弹窗、不打开发布页；发现新版弹「检测到新版本 vX，当前版本 vY。是否前往发布页下载？」，用户确认才打开，取消则无动作。手动「检查更新」流程不受影响。v1.0.0 同步回移。 |
| 12 | P2 | 管理目录标签不换行 | `LoaderApp.build_ui()` 中 `dir_link` 的 `wraplength` | 安装信息区「管理目录」蓝色路径的换行宽度由 260 调为 340，路径仍在右侧留白时可单行显示，不再把末尾几个字符挤到第二行。 |

## 明确未做（不在本版范围）

- **未启动 Excel/WPS 做装载联动实测**：本轮为源码级改动与不依赖 Excel 的回归验证，真实宿主联动未在本机执行。
- **任务计划程序联动未做本机实测**：未在本机真实注册/删除任务。命令按 Windows `schtasks` 文档语法实现，回归脚本用模拟 XML 覆盖了解析、自修复、删除取消、旧键迁移与目标归零分支；首次在界面勾选开机启动时会实际注册任务。

## 行为变化与兼容提示

- 管理目录未加入 Excel 受信任位置时，长期启动项不会自动装载；在选项区勾选受信任位置，或重新执行一次长期启动。
- 卸载或转长期启动不再秒杀 Excel：先等 45 秒，超时才询问；取消后操作整体中止，不会留下半卸载状态。
- 开机启动从注册表迁移到计划任务；取消勾选/托盘关闭会立即删除任务。若列表中不再有任何可装载目标（全部取消勾选或文件删除），任务会自动删除，界面勾选同步取消。
- `prefs.json` 新增 `allow_force_close` 键；旧版本升级后为默认值，无迁移负担。
- 运行日志落盘，无需管理员权限写入（目录即管理目录）。

## 验证记录

- `venv_build\Scripts\python.exe -m py_compile AddInLoader.py`：通过。
- `venv_build\Scripts\python.exe tests\test_v101.py`：**83 项全部 PASS**（四态、判重与引用释放、功能移除断言、等待/拒绝/确认/高级强制四分支、计划任务 XML 解析/自修复/文件删除取消/旧键迁移/目标归零自动取消、日志落盘、版本仲裁四分支与低版本开机启动清除、仲裁排除自身与 bootloader 父进程、检查更新版本比较/API 解析/失败降级/404 降级读 tags/乱序与跳号取最大版本、无一键升级模块断言、启动静默检查的无更新/失败不弹窗与有新版询问后确认/取消分支）；回归脚本不调用真实 `schtasks.exe`、不联网。
- 版本号解析验证：跳号 `v1.3.2`（无 1.2.x）与大版本跳跃 `v2.0.0` 均能解析并比较；文件名必须带 `v` 与三段数字（`AddInLoader_v2.0.0.exe`），`version_tuple("v2.0.0") == (2,0,0)`。
- 真实 API 验证（联网）：升级前 v1.0.1 与 v1.0.0 的 `fetch_latest_release()` 均返回 `('v1.0.0', releases 页, None)`，无错误；raw 下载远端 v1.0.0 exe 返回 HTTP 200 / 12,514,363 字节。真实端到端升级（新版下载→替换→重启）未在本机执行，仅逻辑回归。
- PyInstaller 6.22.3 `--onefile --windowed --uac-admin` 打包成功；exe 内嵌 `requireAdministrator` 清单。
- 真实仲裁实测（本机）：先启动 v1.0.0（父+子两个进程），再启动 v1.0.1；1.2 秒时两个 v1.0.0 进程被结束，v1.0.1 父+子进程持续存活，不再静默退出。修复前同一场景 v1.0.1 子进程约 1.6 秒即消失、退出码 0。
- exe 大小 12,490,906 字节；SHA-256 `9099799fbd0081b6077119d9d6004c8611424c02aadde5d05aed5cd9e3f4a8b6`。
- 源码快照 SHA-256 `1f5a8ec298d2618016d96e221dfcbadd22bdedfb15be8934a052948cfc2e1595`（与根目录源码一致）。
- v1.0.0 重打包 exe SHA-256 `b47ae10e704028744d8cfbacbf3e50eaa1f2ff63ef7b8902e0af7dc8e02f8d93`（12,514,363 字节，含检查更新 + tags 取最大 + 启动静默检查，供高版本接管识别）。
- 本轮未执行改动后的 `--selfcheck`（它会查询任务）；源码自检在上一轮改动时已通过，本轮仅编译与回归验证。

## 四版本统一调整：移除一键升级模块
同日稍晚，因使用方反馈只需打开发布页手动下载、不再需要程序内自动替换，三个版本一并重写（不改版本号，覆盖原 tag/Release/附件）。此项为各版本共同改动，不是某一版独有。

- 删除 `download_file()`、`write_upgrade_bat()`、`start_upgrade_bat()`、`_bat_quote()`、`run_upgrade()` 及 `_UPDATE_ASSET`。
- `fetch_latest_release()` 只读清单的 `version` 与 `page`，不再解析 `asset`/`sha256`。
- 手动「检查更新」与启动静默检查的弹窗统一为「是否前往发布页下载？」，确认后打开发布页；删去代理提示。
- 回归脚本相应调整，83 项全部 PASS。

## 最终文件清单
- AddInLoader_v1.0.1.exe（SHA-256 `9099799fbd0081b6077119d9d6004c8611424c02aadde5d05aed5cd9e3f4a8b6`，12,490,906 字节）
- AddInLoader_v1.0.1.py
- upgrade_log_v1.0.1.md
