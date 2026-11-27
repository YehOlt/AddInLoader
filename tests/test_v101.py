# -*- coding: utf-8 -*-
"""AddInLoader 回归脚本：不依赖真实 Excel/WPS，直接验证关键分支。"""
import base64
import contextlib
import os
import sys
import tempfile

SRC = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))),
                   "AddInLoader.py")
sys.path.insert(0, os.path.dirname(SRC))
import AddInLoader as m  # noqa: E402

PASS = []


def check(name, cond):
    if not cond:
        raise AssertionError(name)
    PASS.append(name)
    print("PASS  %s" % name)


@contextlib.contextmanager
def patched(module, name, value):
    old = getattr(module, name)
    setattr(module, name, value)
    try:
        yield
    finally:
        setattr(module, name, old)


class FakeWB:
    def __init__(self, fullname):
        self.FullName = fullname


class FakeWorkbooks:
    def __init__(self, result=None):
        self.result = result
        self.opened = []

    def Open(self, path):
        self.opened.append(path)
        return self.result


class FakeApp:
    def __init__(self, existing=None, opened=None):
        self.existing = existing
        self.Workbooks = FakeWorkbooks(opened)


def fake_workbook_by_name(app, name):
    return app.existing


class DummyForce:
    def __init__(self, value):
        self.value = value

    def get(self):
        return self.value


def make_stub(wait_state, force=False, ask=True):
    stub = object.__new__(m.LoaderApp)
    stub.root = None
    stub.var_force = DummyForce(force)
    stub.logs = []
    stub.append_log = stub.logs.append
    stub._wait_excel_close = lambda reason: wait_state
    stub.ask = ask
    return stub


def test_trust_states():
    with tempfile.TemporaryDirectory() as td:
        trusted = os.path.join(td, "trusted.xlam")
        session = os.path.join(td, "session.xlam")
        foreign = os.path.join(td, "foreign.xlam")
        denied = os.path.join(td, "denied.xlam")
        for p in (trusted, session, foreign, denied):
            with open(p, "wb"):
                pass
        with patched(m, "is_auto_trusted", lambda p: m.norm_path(p) == m.norm_path(trusted)):
            m._SESSION_TRUSTED.add(m.norm_path(session))
            check("路径四态-受信任位置",
                  m.trust_state(trusted)[0] == m.TRUST_TRUSTED)
            check("路径四态-会话临时信任",
                  m.trust_state(session)[0] == m.TRUST_SESSION)
            check("路径四态-未受信任",
                  m.trust_state(denied)[0] == m.TRUST_DENIED)
            with patched(m, "workbook_by_name", lambda app, name: FakeWB(
                    os.path.join(td, "other.xlam"))):
                state, why = m.trust_state(foreign, app=object())
                check("路径四态-同名异路径",
                      state == m.TRUST_FOREIGN and "异路径" in why)
            check("xlam_allowed 与四态一致",
                  m.xlam_allowed(trusted)[0] and m.xlam_allowed(denied)[0] is False)
            m._SESSION_TRUSTED.discard(m.norm_path(session))


def test_provision_reuses_and_releases():
    with tempfile.TemporaryDirectory() as td:
        target = os.path.join(td, "AIL_probe.xlam")
        with open(target, "wb"):
            pass
        m._LOADED_APPS.clear()
        app = FakeApp(existing=FakeWB(target))
        with patched(m, "workbook_by_name", fake_workbook_by_name):
            check("已在工作簿判重不重复打开",
                  m.provision(app, 7, target) == "already" and
                  not app.Workbooks.opened)
        app = FakeApp(existing=FakeWB(os.path.join(td, "copy.xlam")))
        with patched(m, "workbook_by_name", fake_workbook_by_name):
            check("同名异路径不打开第二份",
                  m.provision(app, 7, target) == "foreign" and
                  not app.Workbooks.opened)
        app = FakeApp(existing=None, opened=FakeWB(target))
        with patched(m, "workbook_by_name", fake_workbook_by_name):
            check("首次装载只打开一次且不留引用",
                  m.provision(app, 7, target) == "loaded" and
                  len(app.Workbooks.opened) == 1 and
                  not hasattr(m, "_TEMP_WBS"))


def test_excel_close_flow():
    calls = {"kill": 0, "ask": 0}

    def is_running():
        return calls["kill"] == 0

    def kill():
        calls["kill"] += 1
        return calls["kill"]

    def ask(*_a, **_k):
        calls["ask"] += 1
        return True

    with patched(m, "is_excel_running", is_running), \
            patched(m, "close_excel_processes", kill), \
            patched(m.messagebox, "askyesno", ask):
        stub = make_stub(False)
        check("用户取消等待不强制结束",
              m.LoaderApp.ensure_excel_closed(stub, "卸载") is False and calls["kill"] == 0)
        calls.update(kill=0, ask=0)
        stub = make_stub("timeout", ask=False)
        with patched(m.messagebox, "askyesno", lambda *_a, **_k: False):
            check("超时且用户拒绝不强制结束",
                  m.LoaderApp.ensure_excel_closed(stub, "卸载") is False and calls["kill"] == 0)
        calls.update(kill=0, ask=0)
        stub = make_stub("timeout")
        check("超时确认后才强制结束",
              m.LoaderApp.ensure_excel_closed(stub, "卸载") is True and
              calls["kill"] == 1 and calls["ask"] == 1)
        calls.update(kill=0, ask=0)
        stub = make_stub("timeout", force=True)
        check("高级选项超时自动强制结束",
              m.LoaderApp.ensure_excel_closed(stub, "卸载") is True and
              calls["kill"] == 1 and calls["ask"] == 0)


TASK_XML = """<?xml version="1.0"?>
<Task xmlns="http://schemas.microsoft.com/windows/2004/02/mit/task">
  <Principals><Principal>
    <RunLevel>HighestAvailable</RunLevel>
    <LogonType>InteractiveToken</LogonType>
  </Principal></Principals>
  <Settings><Enabled>true</Enabled></Settings>
  <Actions><Exec>
    <Command>C:\\Tools\\AddInLoader_v1.0.1.exe</Command>
    <Arguments>--background</Arguments>
  </Exec></Actions>
</Task>
"""


def test_autostart_tasks():
    with patched(m, "_schtasks", lambda args: (0, TASK_XML, "")):
        info = m.task_info()
        check("任务 XML 解析-Highest", info["run_level"] == "HighestAvailable")
        check("任务 XML 解析-交互式登录", info["logon_type"] == "InteractiveToken")
        check("任务动作还原",
              m.task_command(info) == "C:\\Tools\\AddInLoader_v1.0.1.exe --background")
    check("任务命令归一化忽略引号大小写",
          m.norm_task_command('"C:\\A.EXE" --background')
          == m.norm_task_command("c:\\a.exe --background"))

    with patched(m, "task_info", lambda: {
            "enabled": True, "command": "old.exe", "arguments": "",
            "run_level": "Limited", "logon_type": "Password"}), \
            patched(m, "autostart_targets_exist", lambda: True), \
            patched(m, "_legacy_run_value", lambda: None), \
            patched(m, "set_autostart", lambda want: acts.append(want)):
        acts = []
        note = m.autostart_health_fix()
        check("权限降级/路径漂移自动重建任务",
              acts == [True] and "修复" in note)

    with patched(m, "task_info", lambda: {
            "enabled": True, "command": "x.exe", "arguments": "",
            "run_level": "HighestAvailable", "logon_type": "InteractiveToken"}), \
            patched(m, "autostart_targets_exist", lambda: False), \
            patched(m, "_legacy_run_value", lambda: None), \
            patched(m, "set_autostart", lambda want: acts.append(want)):
        acts = []
        note = m.autostart_health_fix()
        check("程序文件删除自动取消开机启动",
              acts == [False] and "删除" in note)

    with patched(m, "_legacy_run_value", lambda: '"old.exe" --background'), \
            patched(m, "task_info", lambda: None), \
            patched(m, "_remove_legacy_run", lambda: None), \
            patched(m, "set_autostart", lambda want: acts.append(want)):
        acts = []
        note = m.migrate_autostart()
        check("旧 Run 键迁移到计划任务", acts == [True] and "迁移" in note)

    with patched(m, "task_info", lambda: {
            "enabled": True, "command": "x.exe", "arguments": "",
            "run_level": "HighestAvailable", "logon_type": "InteractiveToken"}), \
            patched(m, "_legacy_run_value", lambda: None):
        check("计划任务存在即视为开机启动已启用", m.is_autostart() is True)

    def make_autostart_stub(targets):
        stub = object.__new__(m.LoaderApp)
        stub.root = None
        stub.started_at = 0
        stub.logs = []
        stub.append_log = stub.logs.append
        stub.get_targets = lambda: targets
        stub.applied = []
        stub._apply_autostart = lambda want: stub.applied.append(want)
        return stub

    with patched(m, "is_autostart", lambda: True):
        stub = make_autostart_stub([])
        m.LoaderApp.sync_autostart_with_targets(stub)
        check("目标为空自动取消开机启动", stub.applied == [False])
        stub = make_autostart_stub(["xlam"])
        m.LoaderApp.sync_autostart_with_targets(stub)
        check("仍有目标时保留开机启动", stub.applied == [])


def test_clear_unchecked_temp():
    """v1.0.4 清除未勾选的临时启动项：只清临时且未勾选、未登记者。"""
    with tempfile.TemporaryDirectory() as td:
        temp_unchecked = os.path.join(td, "temp_unchecked.xlam")
        temp_checked = os.path.join(td, "temp_checked.xlam")
        discovered_unchecked = os.path.join(td, "discovered_unchecked.xlam")
        registered = os.path.join(td, "registered.xlam")
        for p in (temp_unchecked, temp_checked, discovered_unchecked, registered):
            with open(p, "wb"):
                pass
        deselected = {m.norm_path(p) for p in (temp_unchecked, discovered_unchecked,
                                               registered)}

        check("未勾选临时项被选中",
              m.unchecked_temp_paths([temp_unchecked, temp_checked], deselected)
              == [temp_unchecked])
        with patched(m, "registered_at",
                     lambda p: "reg" if m.norm_path(p) == m.norm_path(registered) else None):
            check("已登记项即使未勾选也不清除",
                  m.unchecked_temp_paths([registered], deselected) == [])
        check("未勾选的自动发现项不在临时列表则不受影响",
              m.unchecked_temp_paths([], deselected) == [])
        check("路径规范化后匹配",
              m.unchecked_temp_paths([temp_unchecked.upper()], deselected)
              == [temp_unchecked.upper()])

        # 清除方法：同步卸载 + 队列兜底 + 撤销会话信任，不动勾选与登记项
        m._SESSION_TRUSTED.add(m.norm_path(temp_unchecked))
        m._SESSION_TRUSTED.add(m.norm_path(temp_checked))
        unloaded = []
        removed_loadtime = []
        removed_manager = []
        with patched(m, "find_excel_doc_windows", lambda: [101, 102]), \
             patched(m, "get_excel_app", lambda hwnd: "app%d" % hwnd), \
             patched(m, "unload_from_excel",
                     lambda app, hwnd, p: unloaded.append((app, hwnd, p))), \
             patched(m, "clear_addin_load_time",
                     lambda p: removed_loadtime.append(p) or [p]), \
             patched(m, "clear_addin_manager_entry",
                     lambda p: removed_manager.append(p) or [p]):
            stub = object.__new__(m.LoaderApp)
            stub.lock = __import__("threading").Lock()
            stub.extra = [temp_unchecked, temp_checked]
            stub.deselected = deselected
            stub.unloadq = __import__("queue").Queue()
            stub.kick = __import__("threading").Event()
            stub.logs = []
            stub.append_log = stub.logs.append
            stub.build_list = lambda: None
            stub.set_initial_status = lambda: None
            stub.sync_autostart_with_targets = lambda: None
            m.LoaderApp.clear_unchecked_temp(stub)

        check("同步卸载覆盖所有 Excel 窗口",
              len(unloaded) == 2
              and all(p == temp_unchecked for _, _, p in unloaded)
              and {app for app, _, _ in unloaded} == {"app101", "app102"})
        check("卸载队列已入队兜底",
              stub.unloadq.qsize() == 1 and stub.unloadq.get() == temp_unchecked)
        check("kick 已置位", stub.kick.is_set())
        check("临时列表只移除未勾选者", stub.extra == [temp_checked])
        check("deselected 保留原路径供自动发现项继续使用",
              m.norm_path(temp_unchecked) in stub.deselected)
        check("会话信任只撤销被清除项",
              m.norm_path(temp_unchecked) not in m._SESSION_TRUSTED
              and m.norm_path(temp_checked) in m._SESSION_TRUSTED)
        check("清除 Excel 加载记录",
              removed_loadtime == [temp_unchecked]
              and removed_manager == [temp_unchecked])
        check("源文件未删除", os.path.isfile(temp_unchecked))
        check("清除日志正确",
              any("已清除 1 项未勾选的临时启动" in s for s in stub.logs))
        m._SESSION_TRUSTED.discard(m.norm_path(temp_checked))

        # 没有任何未勾选临时项：不卸载、不写日志以外的动作
        stub2 = object.__new__(m.LoaderApp)
        stub2.lock = __import__("threading").Lock()
        stub2.extra = [temp_checked]
        stub2.deselected = set()
        stub2.logs = []
        stub2.append_log = stub2.logs.append
        with patched(m, "find_excel_doc_windows",
                     lambda: (_ for _ in ()).throw(AssertionError("不应卸载"))):
            m.LoaderApp.clear_unchecked_temp(stub2)
        check("无可清除项时仅写提示日志",
              stub2.extra == [temp_checked]
              and any("没有已取消勾选" in s for s in stub2.logs))


def test_run_log():
    with tempfile.TemporaryDirectory() as td:
        with patched(m, "managed_dir", lambda: td):
            m.write_run_log("regression line")
            text = open(os.path.join(td, "addinloader.log"), encoding="utf-8").read()
            check("日志落盘含时间戳", "regression line" in text and "20" in text[:4])


def test_version_arbitration():
    check("版本元组解析", m.version_tuple("v1.0.1") == (1, 0, 1)
          and m.version_tuple("1.2.10") == (1, 2, 10))
    check("exe 文件名版本解析", m.exe_version(
        r"C:\x\AddInLoader_v1.0.0.exe") == (1, 0, 0))
    check("无版本 exe 识别不了", m.exe_version(
        r"C:\x\AddInLoader.exe") is None)

    OLDER = r"C:\v100\AddInLoader_v1.0.0.exe"
    HIGHER = r"C:\v110\AddInLoader_v1.1.0.exe"
    MINE = r"C:\v101\AddInLoader_v1.0.1.exe"

    cases = []
    with patched(m, "VERSION", "1.0.1"):
        with patched(m, "running_loader_versions",
                     lambda: [(11, OLDER, (1, 0, 0))]):
            with patched(m, "terminate_pids", lambda p: cases.append(("kill", list(p))) or len(p)):
                with patched(m, "clear_lower_autostart",
                             lambda p: cases.append(("clear", list(p))) or ["旧计划任务"]):
                    msg = m.take_over_older_instances()
    check("接管低版本：返回说明含结束与清除", "结束 1 个进程" in (msg or "")
          and "旧计划任务" in (msg or ""))
    check("接管低版本：只结束低版本 PID",
          ("kill", [11]) in cases and ("clear", [OLDER]) in cases)

    with patched(m, "VERSION", "1.0.1"):
        with patched(m, "running_loader_versions",
                     lambda: [(12, HIGHER, (1, 1, 0))]):
            with patched(m, "terminate_pids", lambda p: (_ for _ in ()).throw(
                    AssertionError("不应结束更高版本"))):
                check("更高版本在跑则让位", m.take_over_older_instances() is None)

    with patched(m, "VERSION", "1.0.1"):
        with patched(m, "running_loader_versions",
                     lambda: [(13, MINE.replace("1.0.1", "X"), None)]):
            with patched(m, "terminate_pids", lambda p: (_ for _ in ()).throw(
                    AssertionError("不应结束未知版本"))):
                check("版本识别不了则让位", m.take_over_older_instances() is None)

    with patched(m, "VERSION", "1.0.1"):
        with patched(m, "running_loader_versions",
                     lambda: [(14, MINE, (1, 0, 1))]):
            with patched(m, "terminate_pids", lambda p: (_ for _ in ()).throw(
                    AssertionError("不应结束同版本"))):
                check("同版本在跑则让位", m.take_over_older_instances() is None)

    fired = []
    with patched(m, "_legacy_run_value", lambda: '"%s" --background' % OLDER):
        with patched(m, "_remove_legacy_run", lambda: fired.append("run")):
            with patched(m, "task_info", lambda: None):
                with patched(m, "set_autostart", lambda v: fired.append("set:%s" % v)):
                    notes = m.clear_lower_autostart([OLDER])
    check("清除指向低版本的旧 Run 键并重建",
          "run" in fired and "set:True" in fired
          and "旧启动项已改指向当前版本" in notes)

    fired = []
    with patched(m, "_legacy_run_value", lambda: None):
        with patched(m, "task_info", lambda: {"command": MINE, "arguments": "",
                                             "run_level": "Highest",
                                             "logon_type": "InteractiveToken"}):
            with patched(m, "set_autostart", lambda v: fired.append(v)):
                notes = m.clear_lower_autostart([OLDER])
    check("不误删指向本版本的任务", fired == [] and notes == [])

    # PyInstaller onefile：bootloader 父进程与子进程同名同版本，必须排除，
    # 否则仲裁误判同版本已运行而让位，导致高版本被低版本 mutex 挡住静默退出
    all_procs = [(100, MINE),        # 本进程（子进程）
                 (99, MINE),         # 本进程的 bootloader 父进程
                 (11, OLDER)]        # 真实低版本
    with patched(m.os, "getpid", lambda: 100), \
            patched(m.os, "getppid", lambda: 99), \
            patched(m, "iter_loader_processes", lambda: all_procs):
        seen = m.running_loader_versions()
    check("仲裁排除自身与 bootloader 父进程",
          all(v != (1, 0, 1) for _p, _path, v in seen)
          and len(seen) == 1 and seen[0][0] == 11)


def test_check_update_logic():
    check("版本比较：更新/更旧/相同",
          m.compare_versions("v1.0.2", "1.0.1") == 1
          and m.compare_versions("1.0.0", "v1.0.1") == -1
          and m.compare_versions("v1.0.1", "1.0.1") == 0)
    check("版本比较：多位数段",
          m.compare_versions("1.2.10", "1.2.9") == 1)

    class Resp:
        def __enter__(self):
            return self

        def __exit__(self, *_a):
            return False

        def read(self):
            return b'{"tag_name":"v1.0.2","html_url":"https://x/releases/tag/v1.0.2"}'

    with patched(m.urllib.request, "urlopen", lambda *_a, **_k: Resp()):
        tag, url, err = m.fetch_latest_release()
    check("API 解析 tag 与链接", tag == "v1.0.2"
          and url.endswith("v1.0.2") and err is None)

    def boom(*_a, **_k):
        raise OSError("offline")

    with patched(m.urllib.request, "urlopen", boom):
        tag, url, err = m.fetch_latest_release()
    check("API 失败返回错误不抛异常", tag == "" and url == "" and "offline" in err)

    # 远端只有 Git tag、未创建 Release：404 降级读 tags
    class TagsResp:
        def __enter__(self):
            return self

        def __exit__(self, *_a):
            return False

        def read(self):
            return b'[{"name":"v1.0.0"},{"name":"v0.9.0"}]'

    calls = []

    def fallback(req, *_a, **_k):
        calls.append(req.full_url)
        if req.full_url in (m.MANIFEST_API, m.UPDATE_MANIFEST):
            raise m.urllib.error.HTTPError(
                req.full_url, 404, "Not Found", {}, None)
        if req.full_url == m.UPDATE_API:
            raise m.urllib.error.HTTPError(
                req.full_url, 404, "Not Found", {}, None)
        return TagsResp()

    with patched(m.urllib.request, "urlopen", fallback):
        tag, url, err = m.fetch_latest_release()
    check("404 降级读 tags 取首个版本", tag == "v1.0.0"
          and url == m.RELEASE_PAGE_URL and err is None and len(calls) == 4)

    # 跳号 + 乱序 tags：取版本号最大者，不依赖 API 返回顺序
    class JumpedResp:
        def __enter__(self):
            return self

        def __exit__(self, *_a):
            return False

        def read(self):
            return b'[{"name":"v1.0.1"},{"name":"v1.3.2"},{"name":"v0.9.0"}]'

    def jump_fallback(req, *_a, **_k):
        if req.full_url in (m.MANIFEST_API, m.UPDATE_MANIFEST):
            raise m.urllib.error.HTTPError(
                req.full_url, 404, "Not Found", {}, None)
        if req.full_url == m.UPDATE_API:
            raise m.urllib.error.HTTPError(
                req.full_url, 404, "Not Found", {}, None)
        return JumpedResp()

    with patched(m.urllib.request, "urlopen", jump_fallback):
        tag2, _url2, err2 = m.fetch_latest_release()
    check("乱序 tags 取最高版本 v1.3.2", tag2 == "v1.3.2" and err2 is None)

    class BigJumpResp(JumpedResp):
        def read(self):
            return b'[{"name":"v1.0.1"},{"name":"v2.0.0"}]'

    def big_fallback(req, *_a, **_k):
        if req.full_url in (m.MANIFEST_API, m.UPDATE_MANIFEST):
            raise m.urllib.error.HTTPError(
                req.full_url, 404, "Not Found", {}, None)
        if req.full_url == m.UPDATE_API:
            raise m.urllib.error.HTTPError(
                req.full_url, 404, "Not Found", {}, None)
        return BigJumpResp()

    with patched(m.urllib.request, "urlopen", big_fallback):
        tag3, _url3, err3 = m.fetch_latest_release()
    check("大版本跳跃取 v2.0.0", tag3 == "v2.0.0" and err3 is None)


def test_upgrade_flow():
    # 清单只需给出版本号与发布页；下载地址/校验值已不再解析
    payload = (b'{"version":"v2.0.0",'
               b'"page":"https://x/releases/tag/v2.0.0",'
               b'"asset":"https://x/AddInLoader_v2.0.0.exe",'
               b'"sha256":"' + b"a" * 64 + b'"}')

    class ApiManifestResp:
        def __enter__(self):
            return self

        def __exit__(self, *_a):
            return False

        def read(self):
            import json as _json
            return _json.dumps({"content": base64.b64encode(payload).decode()}).encode()

    seen = []

    def only_manifest(req, *_a, **_k):
        seen.append(req.full_url)
        if req.full_url == m.MANIFEST_API:
            return ApiManifestResp()
        raise OSError("unexpected url")

    with patched(m.urllib.request, "urlopen", only_manifest):
        tag, page, err = m.fetch_latest_release()
    check("API 清单解析版本与发布页", tag == "v2.0.0"
          and page.endswith("v2.0.0") and err is None)
    check("API 清单优先命中不再请求 raw", seen == [m.MANIFEST_API])

    # API 清单失败时降级 raw 清单仍可用
    class RawResp:
        def __enter__(self):
            return self

        def __exit__(self, *_a):
            return False

        def read(self):
            return payload

    def api_down_raw_ok(req, *_a, **_k):
        if req.full_url == m.MANIFEST_API:
            raise OSError("api offline")
        if req.full_url == m.UPDATE_MANIFEST:
            return RawResp()
        raise OSError("unexpected url")

    with patched(m.urllib.request, "urlopen", api_down_raw_ok):
        tag_raw, page_raw, err_raw = m.fetch_latest_release()
    check("API 清单失败降级 raw 清单", tag_raw == "v2.0.0" and err_raw is None
          and page_raw.endswith("v2.0.0"))


def test_silent_update():
    target = m.silent_update_target
    check("静默检查-无更新返回空",
          target("v1.0.1", "https://x", None, current="1.0.1") == "")
    check("静默检查-旧版返回空",
          target("v1.0.0", "https://x", None, current="1.0.1") == "")
    check("静默检查-错误返回空",
          target("v2.0.0", "https://x", "网络失败", current="1.0.1") == "")
    check("静默检查-空tag返回空",
          target("", "https://x", None, current="1.0.1") == "")
    check("静默检查-有新版返回发布页",
          target("v2.0.0", "https://release", None,
                 current="1.0.1") == "https://release")
    check("静默检查-新版无URL回退默认发布页",
          target("v1.3.2", "", None, current="1.0.1") == m.RELEASE_PAGE_URL)

    # 不弹窗：无更新/失败路径不得打开浏览器；有新版直接打开
    class SyncThread:
        def __init__(self, target, args=(), daemon=False):
            self.target, self.args = target, args

        def start(self):
            self.target(*self.args)

    class JobRoot:
        def __init__(self):
            self.jobs = []

        def after(self, _ms, fn):
            self.jobs.append(fn)

    def run_check(silent, manifest, err, choice="skip"):
        calls, choices, logs, infos, downloaded = [], [], [], [], []
        root = JobRoot()
        stub = object.__new__(m.LoaderApp)
        stub.root = root
        stub.append_log = logs.append
        stub.open_release_page = lambda u: calls.append(u)
        stub._handle_update_found = lambda parent, man, latest: choices.append(latest)
        with patched(m, "fetch_update_manifest", lambda: (manifest, err)), \
                patched(m.threading, "Thread", SyncThread), \
                patched(m.messagebox, "askyesno", lambda *a, **k: False), \
                patched(m.messagebox, "showinfo", lambda *a, **k: infos.append(a)):
            if silent:
                m.LoaderApp.silent_check_update(stub)
            else:
                m.LoaderApp.check_for_update(stub, root)
            root.jobs.pop(0)()
        return calls, choices, logs, infos

    latest = {"version": "v2.0.0", "exe": "AddInLoader_v2.0.0.exe",
              "sha256": "0" * 64, "url": "https://x/a.exe"}

    check("静默检查-无更新不打扰但记日志",
          run_check(True, {"version": "v" + m.VERSION}, None)
          == ([], [], ["当前已是最新版 v%s。" % m.VERSION], []))
    check("静默检查-失败不打扰但记日志",
          run_check(True, None, "网络失败")
          == ([], [], ["检查更新失败：网络失败"], []))
    check("静默检查-有新版走三选一分派",
          run_check(True, latest, None)[1] == ["v2.0.0"])
    manual_same = run_check(False, {"version": "v" + m.VERSION}, None)
    check("手动检查-无新版只记日志不弹窗",
          manual_same[2] == ["当前已是最新版 v%s。" % m.VERSION]
          and manual_same[3] == [])
    check("手动检查-失败弹确认窗",
          "检查更新失败" in "".join(run_check(False, None, "网络失败")[2]))
    check("手动检查-有新版同样走三选一",
          run_check(False, latest, None)[1] == ["v2.0.0"])

    # 关于窗没有日志面板：结果走 report 回调写状态行
    seen = []
    stub = object.__new__(m.LoaderApp)
    stub.root = JobRoot()
    stub.append_log = lambda _t: None
    stub.open_release_page = lambda _u: None
    with patched(m, "fetch_update_manifest",
                 lambda: ({"version": "v" + m.VERSION}, None)), \
            patched(m.threading, "Thread", SyncThread):
        m.LoaderApp.check_for_update(stub, None, report=seen.append)
        stub.root.jobs.pop(0)()
    check("关于窗报告-先进度后结果",
          seen == ["正在检查更新…", "当前已是最新版 v%s。" % m.VERSION])

    # open_release_page：URL 为空回退默认发布页，且始终写日志
    opened = []
    logs = []
    stub = object.__new__(m.LoaderApp)
    stub.append_log = logs.append
    with patched(m.webbrowser, "open", lambda u: opened.append(u)):
        m.LoaderApp.open_release_page(stub, "")
    check("打开发布页-空URL回退默认页并记日志",
          opened == [m.RELEASE_PAGE_URL]
          and logs == ["打开发布页 %s" % m.RELEASE_PAGE_URL])




def test_trust_cache_invalidation():
    # 注册表写入/删除成功后必须立即清信任缓存，避免界面滞后 30 秒
    calls = []

    class FakeKey:
        def __init__(self, values):
            self.values = values

    def fake_open(hive, path, *a):
        calls.append(("open", path))
        if "Trusted Locations\\" in path and path.endswith("\\Location0"):
            return FakeKey({"Path": m.managed_dir()})
        if path.endswith("\\Trusted Locations"):
            return FakeKey({})
        raise OSError("missing")

    def fake_enum(key, idx):
        if idx == 0:
            return "Location0"
        raise OSError("end")

    def fake_query(key, name):
        return key.values[name], 0

    m._TRUST_CACHE["ts"] = 9999999.0
    m._TRUST_CACHE["locs"] = [("stale", True)]
    with patched(m.winreg, "OpenKey", fake_open), \
         patched(m.winreg, "EnumKey", fake_enum), \
         patched(m.winreg, "QueryValueEx", fake_query), \
         patched(m.winreg, "CloseKey", lambda k: None), \
         patched(m.winreg, "CreateKeyEx", lambda h, p: FakeKey({})), \
         patched(m.winreg, "SetValueEx", lambda k, n, i, t, v: None), \
         patched(m.winreg, "DeleteKey", lambda h, p: None):
        ok, msg = m.ensure_trusted_location()
        check("ensure_trusted 已有位置返回已受信任", (ok, msg) == (False, "已受信任"))
        check("ensure_trusted 已有位置清缓存", m._TRUST_CACHE["ts"] == 0.0)
        m._TRUST_CACHE["ts"] = 9999999.0
        m._TRUST_CACHE["locs"] = [("stale", True)]
        m.remove_trusted_location()
        check("remove_trusted 删除后清缓存", m._TRUST_CACHE["ts"] == 0.0)


if __name__ == "__main__":
    test_trust_states()
    test_trust_cache_invalidation()
    test_provision_reuses_and_releases()
    test_excel_close_flow()
    test_autostart_tasks()
    test_run_log()
    test_clear_unchecked_temp()
    test_version_arbitration()
    test_check_update_logic()
    test_upgrade_flow()
    test_silent_update()
    print("\nALL PASS: %d checks" % len(PASS))
