# AddInLoader v1.0.2 升级日志

版本描述：受信任位置状态即时刷新的补丁版本。基于 v1.0.1，仅修复信任缓存滞后问题，不改动装载、仲裁、更新、开机启动等既有功能。后续统一调整见文末。

## 修复清单（fix-list）

| # | 级别 | 改动 | 函数/结构锚点 | 说明 |
|---|---|---|---|---|
| 1 | P1 | 受信任位置状态即时刷新 | `ensure_trusted_location()`、`remove_trusted_location()`、`_TRUST_CACHE`、`trusted_locations()` | 登记或撤销 Excel 受信任位置后立即把 `_TRUST_CACHE["ts"]` 置零，使下次 `trusted_locations()` 重读注册表。修复此前写入成功后缓存仍保留旧值、界面最多滞后 30 秒才显示「受信任位置」的问题；同时消除日志里「已加入受信任位置」与「未受信任」自相矛盾的两行输出。覆盖三条路径：新建登记、已存在登记、撤销删除。 |
| 2 | P2 | 源码包瘦身 | `.gitattributes` | 新增 `export-ignore` 规则，GitHub 自动生成的 Source code (zip/tar.gz) 排除本地发布目录与全部 exe，仅保留纯源码、许可、README 与依赖清单。仅影响归档导出，不影响仓库文件。 |
| 3 | P2 | 回归脚本版本解耦 | `tests/test_v101.py` | 静默检查断言不再硬编码 `v1.0.1`，改用 `m.VERSION` 动态取值，避免版本号升级后误报。 |
| 4 | P2 | 交付文档审计修正 | `README.md`、`upgrade_log_v1.0.1.md` | README 自检覆盖清单删去「404 降级读 tags」「tree 路径解析」等内部实现描述；v1.0.1 日志删除对不存在的 `_docs\audit_report.md` 的引用，并把 raw 下载字节数由 12,512,285 更正为 12,514,363。 |

## 明确未做（不在本版范围）

- **未启动 Excel/WPS 做装载联动实测**：本版仅修复信任缓存，未连接真实 Excel 宿主。
- **任务计划程序联动未做本机实测**：未执行 `schtasks.exe`，回归脚本继续用模拟 XML 覆盖各分支。

## 行为变化与兼容提示

- 点「长期启动」后，受信任位置状态立即生效，不再需要手动刷新或等待 30 秒。
- 卸载清理受信任位置后，状态同样立即更新，避免「已撤销但仍显示受信任」的假象。
- 检查更新一律引导前往发布页手动下载；程序内不再自动下载或替换 exe。

## 验证记录

- `venv_build\Scripts\python.exe -m py_compile AddInLoader.py`：通过。
- `venv_build\Scripts\python.exe tests\test_v101.py`：**83 项全部 PASS**；回归脚本不调用真实 `schtasks.exe`、不联网。
- 源码快照 SHA-256 `c7345981f013dbc3662754cb5c716db6153463dc280711ad13bb370bea0bbce3`（根目录源码与发布快照一致）。
- `.gitattributes` 导出验证：`git archive --worktree-attributes HEAD` 输出 7 项，排除本地发布目录与 exe；GitHub Source code zip 由 16 项 / 24,834,290 字节降为纯源码。
- 真实注册表旁证（本机 v1.0.1 现场）：`HKCU\Software\Microsoft\Office\16.0\Excel\Security\Trusted Locations\Location6` = `%APPDATA%\AddInLoader`，写入成功但界面显示「未受信任」，确认是缓存未失效而非写入失败。
- PyInstaller 6.22.3 `--onefile --windowed --uac-admin` 打包成功；exe 内嵌 `requireAdministrator` 清单，机器码 0x8664（x64）。
- `--selfcheck` 实际执行：输出 `AddInLoader v1.0.2`、comtypes OK、文件哈希 `62b02a7039a90fe9e0bb1a309ad9a26ee7c0634e31b9754d8606941310dd1983`、进程结构 OK，退出码 0（计划任务查询未执行，程序按既有降级路径继续）。
- exe 大小 12,487,908 字节；SHA-256 `62b02a7039a90fe9e0bb1a309ad9a26ee7c0634e31b9754d8606941310dd1983`。

## 四版本统一调整：移除一键升级模块
同日稍晚，因使用方反馈只需打开发布页手动下载、不再需要程序内自动替换，三个版本一并重写（不改版本号，覆盖原 tag/Release/附件）。此项为各版本共同改动，不是某一版独有。

- 删除 `download_file()`、`write_upgrade_bat()`、`start_upgrade_bat()`、`_bat_quote()`、`run_upgrade()` 及 `_UPDATE_ASSET`。
- `fetch_latest_release()` 只读清单的 `version` 与 `page`，不再解析 `asset`/`sha256`。
- 手动「检查更新」与启动静默检查的弹窗统一为「是否前往发布页下载？」，确认后打开发布页；删去代理提示。
- 回归脚本相应调整，83 项全部 PASS。

## 最终文件清单
- AddInLoader_v1.0.2.exe（SHA-256 `62b02a7039a90fe9e0bb1a309ad9a26ee7c0634e31b9754d8606941310dd1983`，12,487,908 字节）
- AddInLoader_v1.0.2.py
- upgrade_log_v1.0.2.md
