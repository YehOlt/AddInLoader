# -*- coding: utf-8 -*-
"""AddInLoader v1.0.1 回归脚本：不依赖真实 Excel/WPS，直接验证关键分支。"""
import base64
import contextlib
import hashlib
import importlib.util
import os
import subprocess
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
    check("_TEMP_WBS 已移除", not hasattr(m, "_TEMP_WBS"))
    check("装载后验证功能已移除",
          not any(hasattr(m, n) for n in
                  ("VERIFY_PREF_KEY", "verify_loaded", "read_verify_cfg",
                   "menu_tag_exists", "_VERIFY_REPORTED")))
    check("父进程白名单功能已移除",
          not any(hasattr(m, n) for n in
                  ("PARENT_PREF_KEY", "parent_whitelist",
                   "window_parent_process_name")))
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
    check("本地 v1.0.1 比远端 v1.0.0 新时判为本地更新",
          m.compare_versions(tag, m.VERSION) == -1)

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
    check("跳号版本比较 1.0.1 vs 1.3.2",
          m.compare_versions(tag2, m.VERSION) == 1)

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
    check("清单不含下载地址字段解析",
          not hasattr(m, "_UPDATE_ASSET") and not hasattr(m, "download_file"))
    check("无资源解析仍识别新版",
          m.compare_versions(tag, m.VERSION) > 0)

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

    v100 = os.path.join(os.path.dirname(SRC),
                        "AddInLoader_v1.0.0_InitialRelease",
                        "AddInLoader_v1.0.0.py")
    spec = importlib.util.spec_from_file_location("ail_v100_silent", v100)
    m100 = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(m100)
    check("v1.0.0 静态检查含静默更新入口",
          "self.silent_check_update" in open(v100, encoding="utf-8").read())
    check("v1.0.0 静默检查判定一致",
          m100.silent_update_target("v1.0.1", "https://r", None) == "https://r"
          and m100.silent_update_target("v1.0.0", "https://r", None) == "")

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
    check("手动检查-无新版弹提示",
          run_check(False, {"version": "v" + m.VERSION}, None)[3] != [])
    check("手动检查-失败弹确认窗",
          "检查更新失败" in "".join(run_check(False, None, "网络失败")[2]))
    check("手动检查-有新版同样走三选一",
          run_check(False, latest, None)[1] == ["v2.0.0"])
    with open(SRC, encoding="utf-8") as _f:
        _src_text = _f.read()
    check("静默与手动共用同一实现",
          "self.check_for_update(self.root, silent=True)" in _src_text
          and "def check_for_update(self, parent, silent=False):" in _src_text)

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

    # v1.0.0 同样具备日志能力
    v100 = os.path.join(os.path.dirname(SRC),
                        "AddInLoader_v1.0.0_InitialRelease",
                        "AddInLoader_v1.0.0.py")
    spec = importlib.util.spec_from_file_location("ail_v100_log", v100)
    m100 = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(m100)
    opened100, logs100 = [], []
    stub100 = object.__new__(m100.LoaderApp)
    stub100.append_log = logs100.append
    with patched(m100.webbrowser, "open", lambda u: opened100.append(u)):
        m100.LoaderApp.open_release_page(stub100, "https://v100")
    check("v1.0.0 打开发布页同样记日志",
          opened100 == ["https://v100"]
          and logs100 == ["打开发布页 https://v100"])



def test_api_upgrade_all_versions():
    base = os.path.dirname(SRC)
    versions = [
        "AddInLoader.py",
        "AddInLoader_v1.0.0_InitialRelease/AddInLoader_v1.0.0.py",
        "AddInLoader_v1.0.1_P0SecurityFix/AddInLoader_v1.0.1.py",
        "AddInLoader_v1.0.2_TrustCacheFix/AddInLoader_v1.0.2.py",
    ]
    for rel in versions:
        src = open(os.path.join(base, rel), encoding="utf-8").read()
        check("%s 含 API 清单入口" % rel,
              'MANIFEST_API = "https://api.github.com/repos/' in src
              and "base64.b64decode(envelope" in src)
        check("%s 无旧式一键升级模块" % rel,
              "def run_upgrade" not in src
              and "def write_upgrade_bat" not in src
              and "def start_upgrade_bat" not in src)
        check("%s 保留静默检查与三选一更新" % rel,
              "def silent_check_update(self):" in src
              and "def open_release_page(self, url):" in src
              and "def ask_update_choice(self, parent, latest):" in src
              and 'text="直接下载"' in src
              and 'text="发布页"' in src
              and 'text="暂不更新"' in src
              and "是否前往发布页下载？" not in src)


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


def test_trust_option_removed_in_v103():
    """v1.0.3 去掉受信任位置勾选框，长期启动无条件登记；历史版本保留原样。"""
    base = os.path.dirname(SRC)
    v103 = open(SRC, encoding="utf-8").read()
    check("v1.0.3 无受信任位置勾选框",
          "var_trust" not in v103 and "将管理目录加入" not in v103)
    check("v1.0.3 长期启动无条件登记受信任位置",
          "trusted_ok, trusted_msg = ensure_trusted_location()" in v103
          and v103.count("if self.var_trust.get():") == 0)
    check("v1.0.3 登记失败不中止长期启动",
          "受信任位置登记失败（%s）；" in v103
          and "加载项仍会登记" in v103
          and 'trust_note = "未加入受信任位置"' in v103)
    check("v1.0.3 版本号为 1.0.3", 'VERSION = "1.0.3"' in v103)
    for rel in ("AddInLoader_v1.0.0_InitialRelease/AddInLoader_v1.0.0.py",
                "AddInLoader_v1.0.1_P0SecurityFix/AddInLoader_v1.0.1.py",
                "AddInLoader_v1.0.2_TrustCacheFix/AddInLoader_v1.0.2.py"):
        src = open(os.path.join(base, rel), encoding="utf-8").read()
        check("%s 仍保留勾选框（历史版本不改）" % rel,
              "var_trust" in src and "将管理目录加入" in src)


if __name__ == "__main__":
    test_trust_states()
    test_trust_cache_invalidation()
    test_provision_reuses_and_releases()
    test_excel_close_flow()
    test_autostart_tasks()
    test_run_log()
    test_version_arbitration()
    test_check_update_logic()
    test_upgrade_flow()
    test_silent_update()
    test_api_upgrade_all_versions()
    test_trust_option_removed_in_v103()
    print("\nALL PASS: %d checks" % len(PASS))
