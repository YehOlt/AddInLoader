# -*- coding: utf-8 -*-
r"""
AddInLoader —— 通用 Excel 自定义加载项装载器

常驻扫描已打开的 Excel 窗口，把勾选的自定义加载项装载进由第三方软件
以 COM 自动化拉起的 Excel 实例，让对应功能区选项卡出现。

自动发现来源：
  1. 注册表 HKCU\Software\Microsoft\Office\<ver>\Excel\Options 的 OPEN/OPENn
  2. %APPDATA%\Microsoft\AddIns\ 下的 .xlam / .xla
排除 Office 自带 Library，不碰 XLL。

临时启动 / 长期启动：
  临时启动：文件对话框选文件后立即装载进所有已打开的 Excel，不写注册表、
  不复制文件；该项留在装载列表中，后台扫描持续装载，Excel 关了重开仍会装上，
  直到装载器退出或取消勾选才失效。
  长期启动：把 xlam 复制到 %APPDATA%\AddInLoader（装载器统一管理目录），副本文件名
  统一加 AIL_ 前缀以便在 Excel 加载项列表中识别来源；按选项勾选把该目录注册为 Excel
  受信任位置，并写入 OPEN 登记；下次 Excel 启动自动加载。同名副本覆盖升级，不留旧版。
  卸载：撤销 OPEN 登记、删除目录副本、移出装载器清单、关闭当前 Excel 中的该加载项；
  目录内已无其它长期副本时，一并撤销该受信任位置登记。

生命周期：点「关闭」或窗口 X 隐藏到系统托盘、后台继续扫描；托盘左键
显隐窗口，右键菜单含显示窗口／开机启动开关／退出；只有「退出」才真正
结束进程。开机启动用计划任务（ONLOGON/交互式/RunLevel=Highest，--background
静默驻留托盘），旧 Run 键自动迁移；没有可装载目标时自动取消开机启动。

v1.0.1 安全加固：
  1. 宿主校验：枚举到 XLMAIN/EXCEL7 后复核所属进程必须是 EXCEL.EXE，
     WPS 表格复用同类窗口名一律跳过。
  2. 路径信任四态：受信任位置 / 会话临时信任 / 同名异路径 / 未受信任；
     临时启动与长期启动共用判据，同名异路径副本不装载、不执行宏。
  3. 不持有加载项工作簿的 COM 引用，按名字取回并核对 FullName 后关闭；
     判重逻辑同步改用工作簿在位检查。
  4. 关闭 Excel 前先提示保存并轮询等待（最多 45 秒），超时经确认才强制
     结束；高级选项可改为超时自动强制结束。
  5. 日志落盘，--selfcheck 输出源文件/exe 的 SHA-256，tests/test_v101.py
     为不依赖 Excel 的回归脚本。
  6. README 醒目标注 %APPDATA% 管理目录作为用户可写受信任位置的残余风险。
  7. 开机启动改为计划任务 ONLOGON + /IT + /RL HIGHEST；旧 Run 键自动迁移，
     任务路径/权限异常自修复，程序文件或加载目标消失时自动取消。

v1.0.2 修复：
  受信任位置登记/撤销后即时清除信任缓存，界面状态与日志不再滞后 30 秒。

通道：oleacc!AccessibleObjectFromWindow(OBJID_NATIVEOM) 取 Excel Application，
微软文档化的 Office Native OM 附着方式；不挂钩子、不注入、不写 HKLM。

用法：
  AddInLoader.exe                 图形界面
  AddInLoader.exe --background    开机启动用：直接驻留托盘、不显示窗口
  AddInLoader.exe --selfcheck     环境与依赖自检，打印发现结果与开机启动状态
  AddInLoader.exe --xlam <路径>   仅装载指定 xlam（排障用，跳过自动发现）
"""
# 版本信息（统一在此维护，随版本更新）
VERSION = "1.0.2"

import ctypes
import base64
import hashlib
import json
import os
import queue
import re
import shutil
import subprocess
import sys
import threading
import time
import tkinter as tk
import urllib.error
import urllib.request
import webbrowser
import winreg
import xml.etree.ElementTree as ET
from collections import deque
from ctypes import wintypes
from tkinter import font as tkfont
from tkinter import scrolledtext, ttk, filedialog, messagebox

import comtypes
import comtypes.automation as auto

TOOL_NAME = "AddInLoader"
AIL_PREFIX = "AIL_"
EXCEL_CLOSE_TIMEOUT = 45      # 卸载/长期启动前等待 Excel 自行退出的秒数
EXCEL_CLOSE_POLL_MS = 500
FORCE_PREF_KEY = "allow_force_close"
TRUST_TRUSTED = "trusted"
TRUST_SESSION = "session"
TRUST_FOREIGN = "foreign"
TRUST_DENIED = "denied"
TRUST_LABELS = {
    TRUST_TRUSTED: "受信任位置",
    TRUST_SESSION: "会话临时信任",
    TRUST_FOREIGN: "同名异路径副本",
    TRUST_DENIED: "未受信任",
}
RELEASE_PAGE_URL = "https://github.com/YehOlt/AddInLoader/releases"
UPDATE_API = "https://api.github.com/repos/YehOlt/AddInLoader/releases/latest"
TAGS_API = "https://api.github.com/repos/YehOlt/AddInLoader/tags"
UPDATE_MANIFEST = "https://raw.githubusercontent.com/YehOlt/AddInLoader/refs/heads/main/update.json"
MANIFEST_API = "https://api.github.com/repos/YehOlt/AddInLoader/contents/update.json?ref=main"
SCAN_INTERVAL = 2.0
DISCOVER_INTERVAL = 5.0
LOG_LIMIT = 200

OBJID_NATIVEOM = 0xFFFFFFF0
IID_IDispatch = comtypes.GUID("{00020400-0000-0000-C000-000000000046}")
OFFICE_VERS = ("16.0", "15.0", "14.0")

# ---- 开机启动 ----
RUN_KEY = r"Software\Microsoft\Windows\CurrentVersion\Run"
RUN_NAME = "AddInLoader"
TASK_NAME = "AddInLoader Background"
AUTOSTART_GRACE = 30       # 启动后留 30 秒给自动发现，避免误判“没有目标”
AUTOSTART_CHECK_SEC = 60
STARTUP_NOTICE = ""

# ---- 单实例 ----
MUTEX_NAME = "Local\\AddInLoader_SingleInstance"
WM_SHOW_APP = 0x0400 + 77

# ---- 系统托盘 ----
WM_USER = 0x0400
TRAY_MSG = WM_USER + 20
NIM_ADD, NIM_MODIFY, NIM_DELETE = 0, 1, 2
NIF_MESSAGE, NIF_ICON, NIF_TIP = 0x1, 0x2, 0x4
GWLP_WNDPROC = -4
IDI_APPLICATION = 32512
WM_LBUTTONUP, WM_RBUTTONUP, WM_LBUTTONDBLCLK = 0x202, 0x205, 0x203
MF_STRING, MF_CHECKED, MF_SEPARATOR = 0x0, 0x8, 0x800
TPM_RIGHTBUTTON, TPM_RETURNCMD = 0x2, 0x100
ID_SHOW, ID_AUTOSTART, ID_EXIT = 1001, 1002, 1003


class NOTIFYICONDATAW(ctypes.Structure):
    _fields_ = [
        ("cbSize", wintypes.DWORD),
        ("hWnd", wintypes.HWND),
        ("uID", wintypes.UINT),
        ("uFlags", wintypes.UINT),
        ("uCallbackMessage", wintypes.UINT),
        ("hIcon", wintypes.HICON),
        ("szTip", wintypes.WCHAR * 128),
        ("dwState", wintypes.DWORD),
        ("dwStateMask", wintypes.DWORD),
        ("szInfo", wintypes.WCHAR * 256),
        ("uTimeoutOrVersion", wintypes.UINT),
        ("szInfoTitle", wintypes.WCHAR * 64),
        ("dwInfoFlags", wintypes.DWORD),
        ("guidItem", ctypes.c_byte * 16),
        ("hBalloonIcon", wintypes.HICON),
    ]


_oleacc = ctypes.windll.oleacc
_oleacc.AccessibleObjectFromWindow.argtypes = [
    wintypes.HWND, wintypes.DWORD,
    ctypes.POINTER(comtypes.GUID), ctypes.POINTER(ctypes.c_void_p),
]
_oleacc.AccessibleObjectFromWindow.restype = ctypes.HRESULT

_user32 = ctypes.windll.user32
_shell32 = ctypes.windll.shell32
_kernel32 = ctypes.windll.kernel32
PROCESS_QUERY_LIMITED_INFORMATION = 0x1000
_user32.GetWindowThreadProcessId.argtypes = [
    wintypes.HWND, ctypes.POINTER(wintypes.DWORD)]
_user32.GetWindowThreadProcessId.restype = wintypes.DWORD
_kernel32.OpenProcess.argtypes = [wintypes.DWORD, wintypes.BOOL, wintypes.DWORD]
_kernel32.OpenProcess.restype = wintypes.HANDLE
_kernel32.CloseHandle.argtypes = [wintypes.HANDLE]
_kernel32.CloseHandle.restype = wintypes.BOOL
_kernel32.QueryFullProcessImageNameW.argtypes = [
    wintypes.HANDLE, wintypes.DWORD, wintypes.LPWSTR,
    ctypes.POINTER(wintypes.DWORD)]
_kernel32.QueryFullProcessImageNameW.restype = wintypes.BOOL
_user32.RegisterWindowMessageW.argtypes = [wintypes.LPCWSTR]
_user32.RegisterWindowMessageW.restype = wintypes.UINT
TASKBAR_CREATED = _user32.RegisterWindowMessageW("TaskbarCreated")
_user32.GetClassNameW.argtypes = [wintypes.HWND, wintypes.LPWSTR, ctypes.c_int]
_user32.GetClassNameW.restype = ctypes.c_int
_WNDENUMPROC = ctypes.WINFUNCTYPE(wintypes.BOOL, wintypes.HWND, wintypes.LPARAM)

LRESULT = ctypes.c_ssize_t
WNDPROC = ctypes.WINFUNCTYPE(LRESULT, wintypes.HWND, wintypes.UINT,
                             wintypes.WPARAM, wintypes.LPARAM)
_user32.GetWindowLongPtrW.argtypes = [wintypes.HWND, ctypes.c_int]
_user32.GetWindowLongPtrW.restype = ctypes.c_ssize_t
_user32.SetWindowLongPtrW.argtypes = [wintypes.HWND, ctypes.c_int, ctypes.c_ssize_t]
_user32.SetWindowLongPtrW.restype = ctypes.c_ssize_t
_user32.CallWindowProcW.argtypes = [ctypes.c_ssize_t, wintypes.HWND, wintypes.UINT,
                                    wintypes.WPARAM, wintypes.LPARAM]
_user32.CallWindowProcW.restype = LRESULT
_shell32.Shell_NotifyIconW.argtypes = [wintypes.DWORD, ctypes.POINTER(NOTIFYICONDATAW)]
_shell32.Shell_NotifyIconW.restype = wintypes.BOOL
_user32.LoadIconW.argtypes = [wintypes.HINSTANCE, wintypes.LPVOID]
_user32.LoadIconW.restype = wintypes.HICON
_user32.CreatePopupMenu.restype = wintypes.HMENU
_user32.AppendMenuW.argtypes = [wintypes.HMENU, wintypes.UINT,
                                wintypes.UINT, wintypes.LPCWSTR]
_user32.TrackPopupMenu.argtypes = [wintypes.HMENU, wintypes.UINT, ctypes.c_int,
                                   ctypes.c_int, ctypes.c_int, wintypes.HWND,
                                   wintypes.LPVOID]
_user32.PostMessageW.argtypes = [wintypes.HWND, wintypes.UINT,
                                 wintypes.WPARAM, wintypes.LPARAM]


def norm_path(path):
    return os.path.normcase(os.path.normpath(str(path)))


def is_office_builtin(path):
    """排除 Office 安装目录下的内置加载项（Library 等）。"""
    low = norm_path(path)
    return "\\microsoft office\\" in low and "\\library\\" in low


def parse_open_value(val):
    """从注册表 OPEN 值提取 xlam/xla 路径，兼容引号与 /A 前缀。"""
    s = str(val).strip()
    m = re.search(r'"([^"]+\.xla(?:m)?)"', s, re.IGNORECASE)
    if m:
        return m.group(1)
    m = re.match(r'(?:^|/A\s+)([^\s"]+\.xla(?:m)?)', s, re.IGNORECASE)
    return m.group(1) if m else ""


def options_key_path(ver):
    return r"Software\Microsoft\Office\%s\Excel\Options" % ver


def open_entries(ver):
    """返回 [(值名, 原值, 路径), ...]，按 OPEN, OPEN1, OPEN2 顺序。"""
    out = []
    try:
        key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, options_key_path(ver))
    except OSError:
        return out
    try:
        idx = 0
        while True:
            try:
                name, val, _ = winreg.EnumValue(key, idx)
            except OSError:
                break
            if name == "OPEN" or re.fullmatch(r"OPEN\d+", name):
                out.append((name, val, parse_open_value(val)))
            idx += 1
    finally:
        winreg.CloseKey(key)

    def slot(t):
        return 0 if t[0] == "OPEN" else int(t[0][4:])

    return sorted(out, key=slot)


def registered_at(path):
    """返回登记所在的 (版本, 值名)，未登记返回 None。"""
    n = norm_path(path)
    for ver in OFFICE_VERS:
        for name, _val, p in open_entries(ver):
            if p and norm_path(p) == n:
                return ver, name
    return None


def register_addin(path):
    """把 xlam 写入 Excel 加载项列表。返回 (成功与否, 说明)。"""
    if not os.path.isfile(path):
        return False, "文件不存在"
    if registered_at(path):
        return True, "已登记"
    # 优先写已有 OPEN 登记的版本（Excel 实际在用的），其次已有 Options 键，最后新建
    target = None
    for ver in OFFICE_VERS:
        if open_entries(ver):
            target = ver
            break
    if target is None:
        for ver in OFFICE_VERS:
            try:
                k = winreg.OpenKey(winreg.HKEY_CURRENT_USER, options_key_path(ver))
                winreg.CloseKey(k)
                target = ver
                break
            except OSError:
                continue
    if target is None:
        target = OFFICE_VERS[0]
    key_path = options_key_path(target)
    try:
        key = winreg.CreateKeyEx(winreg.HKEY_CURRENT_USER, key_path, 0,
                                 winreg.KEY_READ | winreg.KEY_WRITE)
    except OSError as exc:
        return False, str(exc)
    try:
        used = {0 if name == "OPEN" else int(name[4:])
                for name, _v, _p in open_entries(target)}
        slot = 0
        while slot in used:
            slot += 1
        name = "OPEN" if slot == 0 else "OPEN%d" % slot
        winreg.SetValueEx(key, name, 0, winreg.REG_SZ, '"%s"' % path)
        return True, "%s\\%s" % (target, name)
    except OSError as exc:
        return False, str(exc)
    finally:
        winreg.CloseKey(key)


def unregister_addin(path):
    """撤销 OPEN 登记并把剩余 OPENn 紧凑化。返回 (成功与否, 说明)。"""
    n = norm_path(path)
    hit = registered_at(path)
    if not hit:
        return True, "本来未登记"
    ver, _slot_name = hit
    key_path = options_key_path(ver)
    try:
        key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, key_path, 0,
                             winreg.KEY_READ | winreg.KEY_WRITE)
    except OSError as exc:
        return False, str(exc)
    try:
        entries = open_entries(ver)
        keep = [(name, val) for name, val, p in entries
                if not (p and norm_path(p) == n)]
        for name, _val, _p in entries:
            try:
                winreg.DeleteValue(key, name)
            except FileNotFoundError:
                pass
        for i, (_old, val) in enumerate(keep):
            name = "OPEN" if i == 0 else "OPEN%d" % i
            winreg.SetValueEx(key, name, 0, winreg.REG_SZ, val)
        return True, ver
    except OSError as exc:
        return False, str(exc)
    finally:
        winreg.CloseKey(key)


def clear_addin_load_time(path):
    """删除 AddInLoadTimes 中该路径的记录，使 Excel 加载项列表不再显示。"""
    target = norm_path(path)
    removed = []
    for ver in OFFICE_VERS:
        key_path = r"Software\Microsoft\Office\%s\Excel\AddInLoadTimes" % ver
        try:
            key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, key_path, 0,
                                 winreg.KEY_READ | winreg.KEY_WRITE)
        except OSError:
            continue
        try:
            names = []
            idx = 0
            while True:
                try:
                    name = winreg.EnumValue(key, idx)[0]
                except OSError:
                    break
                if norm_path(name) == target:
                    names.append(name)
                idx += 1
            for name in names:
                try:
                    winreg.DeleteValue(key, name)
                    removed.append("%s\\%s" % (ver, name))
                except FileNotFoundError:
                    pass
        finally:
            winreg.CloseKey(key)
    return removed


def clear_addin_manager_entry(path):
    """删除 Excel“可用加载宏”对话框里该路径的历史浏览记录（Add-in Manager 键）。

    只清参数指定的路径，不动 OPEN 登记（长期启动本身）。
    """
    target = norm_path(path)
    removed = []
    for ver in OFFICE_VERS:
        key_path = r"Software\Microsoft\Office\%s\Excel\Add-in Manager" % ver
        try:
            key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, key_path, 0,
                                 winreg.KEY_READ | winreg.KEY_WRITE)
        except OSError:
            continue
        try:
            names = []
            idx = 0
            while True:
                try:
                    name = winreg.EnumValue(key, idx)[0]
                except OSError:
                    break
                if norm_path(name) == target:
                    names.append(name)
                idx += 1
            for name in names:
                try:
                    winreg.DeleteValue(key, name)
                    removed.append("%s\\%s" % (ver, name))
                except FileNotFoundError:
                    pass
        finally:
            winreg.CloseKey(key)
    return removed


def detect_office_versions():
    """返回 [(版本, 位数)]，如 [('16.0','64 位')]，去重。"""
    found = []
    for ver in OFFICE_VERS:
        for base, bits in ((r"SOFTWARE\Microsoft\Office", "64 位"),
                           (r"SOFTWARE\WOW6432Node\Microsoft\Office", "32 位")):
            try:
                winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE,
                               base + "\\" + ver + r"\Excel")
                found.append((ver, bits))
                break
            except OSError:
                continue
    return found


def managed_dir():
    return os.path.join(
        os.environ.get("APPDATA", os.path.expanduser("~")), "AddInLoader"
    )


def _managed_path():
    return os.path.join(managed_dir(), "managed_copies.json")


def _trusted_key_path(ver):
    return r"Software\Microsoft\Office\%s\Excel\Security\Trusted Locations" % ver


def ensure_trusted_location():
    """把统一管理目录注册为 Excel 受信任位置。返回 (是否新增, 说明)。"""
    target = managed_dir()
    for ver in OFFICE_VERS:
        base = _trusted_key_path(ver)
        try:
            key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, base)
        except OSError:
            try:
                key = winreg.CreateKeyEx(winreg.HKEY_CURRENT_USER, base)
            except OSError:
                continue
        try:
            subs = []
            idx = 0
            while True:
                try:
                    subs.append(winreg.EnumKey(key, idx))
                except OSError:
                    break
                idx += 1
            for sub in subs:
                try:
                    sk = winreg.OpenKey(winreg.HKEY_CURRENT_USER, base + "\\" + sub)
                    val, _ = winreg.QueryValueEx(sk, "Path")
                    winreg.CloseKey(sk)
                except OSError:
                    continue
                if val and norm_path(os.path.expandvars(val)) == norm_path(target):
                    _TRUST_CACHE["ts"] = 0.0
                    return False, "已受信任"
            max_n = -1
            for sub in subs:
                m = re.fullmatch(r"Location(\d+)", sub)
                if m:
                    max_n = max(max_n, int(m.group(1)))
            name = "Location%d" % (max_n + 1)
            sk = winreg.CreateKeyEx(winreg.HKEY_CURRENT_USER, base + "\\" + name)
            try:
                winreg.SetValueEx(sk, "Path", 0, winreg.REG_SZ, target)
                winreg.SetValueEx(sk, "Description", 0, winreg.REG_SZ,
                                  "AddInLoader 加载项目录")
            finally:
                winreg.CloseKey(sk)
            _TRUST_CACHE["ts"] = 0.0
            return True, "%s\\%s" % (ver, name)
        except OSError as exc:
            return False, str(exc)
        finally:
            winreg.CloseKey(key)
    return False, "未找到 Excel 版本键"


def remove_trusted_location():
    """撤销统一管理目录的受信任位置登记。返回是否删除。"""
    target = norm_path(managed_dir())
    removed = False
    for ver in OFFICE_VERS:
        base = _trusted_key_path(ver)
        try:
            key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, base, 0,
                                 winreg.KEY_READ | winreg.KEY_WRITE)
        except OSError:
            continue
        try:
            subs = []
            idx = 0
            while True:
                try:
                    subs.append(winreg.EnumKey(key, idx))
                except OSError:
                    break
                idx += 1
            for sub in subs:
                try:
                    sk = winreg.OpenKey(winreg.HKEY_CURRENT_USER, base + "\\" + sub)
                    val, _ = winreg.QueryValueEx(sk, "Path")
                    winreg.CloseKey(sk)
                except OSError:
                    continue
                if val and norm_path(os.path.expandvars(val)) == target:
                    try:
                        winreg.DeleteKey(winreg.HKEY_CURRENT_USER, base + "\\" + sub)
                        removed = True
                    except OSError:
                        pass
        finally:
            winreg.CloseKey(key)
    if removed:
        _TRUST_CACHE["ts"] = 0.0
    return removed


# ---- 路径信任（v1.0.1）----
# 自动装载只认 Excel 受信任位置；文件对话框显式选择 / --xlam 指定的文件
# 记入会话信任集，退出装载器即失效，不写注册表。
_SESSION_TRUSTED = set()
_TRUST_CACHE = {"ts": 0.0, "locs": []}
_TRUST_TTL = 30.0


def _norm_dir(path):
    s = norm_path(os.path.abspath(os.path.expandvars(str(path))))
    return s.rstrip("\\") or s


def trusted_locations(force=False):
    """读取 Excel 受信任位置 [(目录, 是否含子目录)]，30 秒缓存。"""
    now = time.time()
    if not force and _TRUST_CACHE["locs"] and now - _TRUST_CACHE["ts"] < _TRUST_TTL:
        return _TRUST_CACHE["locs"]
    locs = []
    bases = []
    for ver in OFFICE_VERS:
        bases.append(r"Software\Microsoft\Office\%s\Excel\Security\Trusted Locations" % ver)
        bases.append(r"Software\WOW6432Node\Microsoft\Office\%s\Excel\Security\Trusted Locations" % ver)
    for hive in (winreg.HKEY_CURRENT_USER, winreg.HKEY_LOCAL_MACHINE):
        for base in bases:
            try:
                key = winreg.OpenKey(hive, base)
            except OSError:
                continue
            try:
                idx = 0
                while True:
                    try:
                        sub = winreg.EnumKey(key, idx)
                    except OSError:
                        break
                    idx += 1
                    try:
                        sk = winreg.OpenKey(hive, base + "\\" + sub)
                        try:
                            val, _ = winreg.QueryValueEx(sk, "Path")
                            allow = 1
                            try:
                                a, _ = winreg.QueryValueEx(sk, "AllowSubFolders")
                                allow = int(a)
                            except OSError:
                                pass
                        finally:
                            winreg.CloseKey(sk)
                        if val:
                            locs.append((_norm_dir(val), bool(allow)))
                    except OSError:
                        continue
            finally:
                winreg.CloseKey(key)
    _TRUST_CACHE["ts"] = now
    _TRUST_CACHE["locs"] = locs
    return locs


def is_auto_trusted(path):
    """自动装载策略：文件所在目录属于任一受信任位置（含子目录按 AllowSubFolders）。"""
    d = _norm_dir(os.path.dirname(os.path.abspath(path)))
    for base, allow in trusted_locations():
        if d == base or (allow and d.startswith(base + "\\")):
            return True
    return False


def mark_session_trusted(path):
    _SESSION_TRUSTED.add(norm_path(path))
    _TRUST_CACHE["ts"] = 0.0


def trust_state(path, app=None):
    """统一四态：受信任位置 / 会话临时信任 / 同名异路径 / 未受信任。

    临时启动与长期启动共用此判据；传入 app 时额外检测同名异路径副本。
    """
    if not os.path.isfile(path):
        return TRUST_DENIED, "文件不存在"
    n = norm_path(path)
    if app is not None:
        existing = workbook_by_name(app, os.path.basename(path))
        if existing is not None:
            try:
                ep = existing.FullName
            except Exception:
                ep = ""
            if ep and norm_path(ep) != n:
                return TRUST_FOREIGN, "存在同名异路径副本"
    if n in _SESSION_TRUSTED:
        return TRUST_SESSION, "会话临时信任"
    if is_auto_trusted(path):
        return TRUST_TRUSTED, "受信任位置"
    return TRUST_DENIED, "不在 Excel 受信任位置"


def xlam_allowed(path):
    """路径层判据（不含运行时的同名异路径态），返回 (是否允许, 原因)。"""
    state, why = trust_state(path)
    return state in (TRUST_SESSION, TRUST_TRUSTED), why


def _read_managed():
    try:
        with open(_managed_path(), "r", encoding="utf-8") as fh:
            data = json.load(fh)
        return data if isinstance(data, dict) else {}
    except Exception:
        return {}


def _write_managed(data):
    path = _managed_path()
    os.makedirs(os.path.dirname(path), exist_ok=True)
    tmp = path + ".tmp"
    with open(tmp, "w", encoding="utf-8") as fh:
        json.dump(data, fh, ensure_ascii=False, indent=2)
    os.replace(tmp, path)


def managed_copy_name(path):
    """管理目录副本名：统一加 AIL_ 前缀，已带则不重复。"""
    name = os.path.basename(path)
    return name if name.lower().startswith(AIL_PREFIX.lower()) else AIL_PREFIX + name


def copy_to_managed(path):
    """复制 xlam 到管理目录（文件名加 AIL_ 前缀）。同名本工具副本覆盖升级；非本工具同名文件拒绝覆盖。"""
    if not os.path.isfile(path):
        return None, "源文件不存在"
    dst = os.path.join(managed_dir(), managed_copy_name(path))
    n = norm_path(dst)
    managed = _read_managed()
    ours = n in managed
    if norm_path(path) == n:
        return dst, "已在管理目录"
    if os.path.exists(dst) and not ours:
        return None, "管理目录已存在同名文件，且非本工具副本"
    try:
        os.makedirs(managed_dir(), exist_ok=True)
        shutil.copy2(path, dst)
    except OSError as exc:
        return None, str(exc)
    managed[n] = norm_path(path)
    try:
        _write_managed(managed)
    except OSError as exc:
        if not ours:
            try:
                os.remove(dst)
            except OSError:
                pass
        return None, str(exc)
    return dst, "已覆盖升级" if ours else "已复制到管理目录"


def remove_from_managed(path):
    """仅删除由本工具复制到管理目录的副本。目录空后撤销受信任位置。返回是否删除。"""
    n = norm_path(path)
    managed = _read_managed()
    if n not in managed:
        return False
    try:
        os.remove(path)
    except FileNotFoundError:
        pass
    except OSError:
        return False
    managed.pop(n, None)
    _write_managed(managed)
    if not managed:
        remove_trusted_location()
    return True


def discover_addins():
    """返回 [(路径, 来源), ...]，按路径去重，仅保留存在的自定义 xlam/xla。"""
    found = {}

    for ver in OFFICE_VERS:
        key_path = options_key_path(ver)
        try:
            key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, key_path)
        except OSError:
            continue
        idx = 0
        try:
            while True:
                try:
                    name, val, _ = winreg.EnumValue(key, idx)
                except OSError:
                    break
                if name == "OPEN" or re.fullmatch(r"OPEN\d+", name):
                    p = parse_open_value(val)
                    if p and os.path.isfile(p) and not is_office_builtin(p):
                        found.setdefault(norm_path(p), (p, "Excel 加载项列表"))
                idx += 1
        finally:
            winreg.CloseKey(key)

    addins_dir = os.path.join(
        os.environ.get("APPDATA", os.path.expanduser("~")), "Microsoft", "AddIns"
    )
    if os.path.isdir(addins_dir):
        for fn in os.listdir(addins_dir):
            if fn.startswith("~$"):
                continue
            if fn.lower().endswith((".xlam", ".xla")):
                p = os.path.join(addins_dir, fn)
                if os.path.isfile(p) and not is_office_builtin(p):
                    found.setdefault(norm_path(p), (p, "AddIns 目录"))

    return sorted(found.values(), key=lambda t: os.path.basename(t[0]).lower())


def is_excel_running():
    """检测 EXCEL.EXE 进程是否存在（含隐藏实例）。用 Toolhelp32 枚举，不弹窗。"""
    TH32CS_SNAPPROCESS = 0x00000002
    INVALID = ctypes.c_void_p(-1).value

    class PROCESSENTRY32W(ctypes.Structure):
        _fields_ = [
            ("dwSize", wintypes.DWORD),
            ("cntUsage", wintypes.DWORD),
            ("th32ProcessID", wintypes.DWORD),
            ("th32DefaultHeapID", ctypes.POINTER(ctypes.c_ulong)),
            ("th32ModuleID", wintypes.DWORD),
            ("cntThreads", wintypes.DWORD),
            ("th32ParentProcessID", wintypes.DWORD),
            ("pcPriClassBase", ctypes.c_long),
            ("dwFlags", wintypes.DWORD),
            ("szExeFile", wintypes.WCHAR * 260),
        ]

    kernel32 = ctypes.windll.kernel32
    snap = kernel32.CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, 0)
    if not snap or snap == INVALID:
        return False
    try:
        entry = PROCESSENTRY32W()
        entry.dwSize = ctypes.sizeof(PROCESSENTRY32W)
        ok = kernel32.Process32FirstW(snap, ctypes.byref(entry))
        while ok:
            if entry.szExeFile.lower() == "excel.exe":
                return True
            ok = kernel32.Process32NextW(snap, ctypes.byref(entry))
        return False
    finally:
        kernel32.CloseHandle(snap)


def close_excel_processes():
    """强制结束所有 EXCEL.EXE 进程，返回结束数量。不依赖 COM，不卡在保存对话框。"""
    TH32CS_SNAPPROCESS = 0x00000002
    INVALID = ctypes.c_void_p(-1).value
    PROCESS_TERMINATE = 0x0001

    class PROCENTRY(ctypes.Structure):
        _fields_ = [
            ("dwSize", wintypes.DWORD),
            ("cntUsage", wintypes.DWORD),
            ("th32ProcessID", wintypes.DWORD),
            ("th32DefaultHeapID", ctypes.POINTER(ctypes.c_ulong)),
            ("th32ModuleID", wintypes.DWORD),
            ("cntThreads", wintypes.DWORD),
            ("th32ParentProcessID", wintypes.DWORD),
            ("pcPriClassBase", ctypes.c_long),
            ("dwFlags", wintypes.DWORD),
            ("szExeFile", wintypes.WCHAR * 260),
        ]

    kernel32 = ctypes.windll.kernel32
    snap = kernel32.CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, 0)
    if not snap or snap == INVALID:
        return 0
    pids = []
    try:
        entry = PROCENTRY()
        entry.dwSize = ctypes.sizeof(PROCENTRY)
        ok = kernel32.Process32FirstW(snap, ctypes.byref(entry))
        while ok:
            if entry.szExeFile.lower() == "excel.exe":
                pids.append(entry.th32ProcessID)
            ok = kernel32.Process32NextW(snap, ctypes.byref(entry))
    finally:
        kernel32.CloseHandle(snap)
    killed = 0
    for pid in pids:
        h = kernel32.OpenProcess(PROCESS_TERMINATE, False, pid)
        if h:
            try:
                if kernel32.TerminateProcess(h, 0):
                    killed += 1
            finally:
                kernel32.CloseHandle(h)
    if killed:
        time.sleep(1.0)
    return killed


def window_process_name(hwnd):
    """取窗口所属进程名；识别不了返回空串（调用方必须 fail-closed）。"""
    pid = wintypes.DWORD()
    _user32.GetWindowThreadProcessId(hwnd, ctypes.byref(pid))
    if not pid.value:
        return ""
    h = _kernel32.OpenProcess(PROCESS_QUERY_LIMITED_INFORMATION, False, pid.value)
    if not h:
        return ""
    try:
        size = wintypes.DWORD(1024)
        buf = ctypes.create_unicode_buffer(1024)
        if not _kernel32.QueryFullProcessImageNameW(h, 0, buf, ctypes.byref(size)):
            return ""
        return os.path.basename(buf.value).lower()
    finally:
        _kernel32.CloseHandle(h)


def find_excel_doc_windows():
    """枚举 EXCEL.EXE 的 XLMAIN/EXCEL7 文档窗口。"""
    tops = []

    def top_cb(hwnd, _lp):
        buf = ctypes.create_unicode_buffer(64)
        _user32.GetClassNameW(hwnd, buf, 64)
        if buf.value != "XLMAIN" or window_process_name(hwnd) != "excel.exe":
            return True
        tops.append(hwnd)
        return True

    _user32.EnumWindows(_WNDENUMPROC(top_cb), 0)
    docs = []
    for top in tops:
        def child_cb(hwnd, _lp):
            buf = ctypes.create_unicode_buffer(64)
            _user32.GetClassNameW(hwnd, buf, 64)
            if buf.value == "EXCEL7":
                docs.append(hwnd)
            return True

        _user32.EnumChildWindows(top, _WNDENUMPROC(child_cb), 0)
    return docs


def get_excel_app(hwnd):
    """从 EXCEL7 窗口取 Excel Application 动态对象，失败返回 None。"""
    out = ctypes.c_void_p()
    hr = _oleacc.AccessibleObjectFromWindow(
        hwnd, wintypes.DWORD(OBJID_NATIVEOM),
        ctypes.byref(IID_IDispatch), ctypes.byref(out),
    )
    if hr != 0 or not out.value:
        return None
    try:
        unk = ctypes.cast(out, ctypes.POINTER(comtypes.IUnknown))
        disp = unk.QueryInterface(auto.IDispatch)
        dispid = disp.GetIDsOfNames("Application")[0]
        return disp.Invoke(dispid, _invkind=auto.DISPATCH_PROPERTYGET)
    except Exception:
        return None


_LOADED_APPS = set()
_SKIP_LOGGED = set()


def _prefs_path():
    return os.path.join(managed_dir(), "prefs.json")


def read_prefs():
    try:
        with open(_prefs_path(), "r", encoding="utf-8") as fh:
            data = json.load(fh)
        return data if isinstance(data, dict) else {}
    except Exception:
        return {}


def write_pref(key, value):
    try:
        data = read_prefs()
        data[key] = value
        path = _prefs_path()
        os.makedirs(os.path.dirname(path), exist_ok=True)
        tmp = path + ".tmp"
        with open(tmp, "w", encoding="utf-8") as fh:
            json.dump(data, fh, ensure_ascii=False, indent=2)
        os.replace(tmp, path)
        return True
    except OSError:
        return False


def _run_log_path():
    return os.path.join(managed_dir(), "addinloader.log")


def write_run_log(text):
    """追加落盘日志，超过 1 MB 滚动一次；失败不影响主流程。"""
    path = _run_log_path()
    try:
        if os.path.isfile(path) and os.path.getsize(path) >= 1024 * 1024:
            os.replace(path, path + ".1")
        with open(path, "a", encoding="utf-8") as fh:
            fh.write("%s  %s\n" % (time.strftime("%Y-%m-%d %H:%M:%S"), text))
    except OSError:
        pass


def workbook_by_name(app, name):
    try:
        return app.Workbooks.Item(name)
    except Exception:
        return None


def workbook_in_place(app, xlam):
    """按名字取工作簿并核对完整路径：加载项工作簿不在 Workbooks 计数里，但 Item(name) 可取。"""
    wb = workbook_by_name(app, os.path.basename(xlam))
    if wb is None:
        return False, None
    try:
        p = wb.FullName
    except Exception:
        p = ""
    return (bool(p) and norm_path(p) == norm_path(xlam)), wb


def provision(app, hwnd, xlam):
    """幂等装载到当前会话；用完即释放 COM 引用，按工作簿在位判重。"""
    target = norm_path(xlam)
    key = (int(hwnd), target)
    state, _why = trust_state(xlam, app)
    if state == TRUST_FOREIGN:
        return "foreign"
    existing = workbook_by_name(app, os.path.basename(xlam))
    if existing is not None:
        try:
            ep = existing.FullName
        except Exception:
            ep = ""
        if ep and norm_path(ep) != target:
            return "foreign"
        if not ep:
            return "failed"
        _LOADED_APPS.add(key)
        return "already"
    # 用 add-in 工作簿打开：IsAddin=True，功能区生效，不进 Workbooks 集合。
    # 不持有返回对象；卸载时按名字取回工作簿并核对 FullName 后关闭。
    try:
        wb = app.Workbooks.Open(xlam)
        if wb is None:
            return "failed"
    except Exception:
        return "failed"
    _LOADED_APPS.add(key)
    return "loaded"


def unload_from_excel(app, hwnd, xlam):
    """关闭指定加载项：临时工作簿与长期 AddIns 项都处理。"""
    closed = 0
    target = norm_path(xlam)
    wb = workbook_by_name(app, os.path.basename(xlam))
    if wb is not None:
        try:
            if norm_path(wb.FullName) == target:
                wb.Close(SaveChanges=False)
                closed += 1
        except Exception:
            pass
    try:
        for i in range(1, app.AddIns.Count + 1):
            try:
                ad = app.AddIns.Item(i)
                if ad.FullName and norm_path(ad.FullName) == target and ad.Installed:
                    ad.Installed = False
                    closed += 1
            except Exception:
                pass
    except Exception:
        pass
    _LOADED_APPS.discard((int(hwnd), target))
    for vkey in [k for k in _VERIFIED if k[0] == int(hwnd) and k[1] == target]:
        _VERIFIED.discard(vkey)
    return closed


def worker(stop, outq, unloadq, kick, get_targets, interval):
    comtypes.CoInitialize()
    try:
        cycle = 0
        while not stop.is_set():
            cycle += 1
            kick.clear()
            try:
                wins = find_excel_doc_windows()
                while True:
                    try:
                        up = unloadq.get_nowait()
                    except queue.Empty:
                        break
                    for hwnd in wins:
                        app = get_excel_app(hwnd)
                        if app is None:
                            continue
                        closed = unload_from_excel(app, hwnd, up)
                        if closed:
                            outq.put(("log", "已从 Excel 卸载 %s"
                                      % os.path.basename(up)))
                targets = get_targets()
                seen_apps = set()
                loaded_paths = set()
                for hwnd in wins:
                    if stop.is_set():
                        break
                    app = get_excel_app(hwnd)
                    if app is None:
                        continue
                    key = int(hwnd)
                    if key in seen_apps:
                        continue
                    seen_apps.add(key)
                    for xlam in targets:
                        if stop.is_set():
                            break
                        allowed, why = xlam_allowed(xlam)
                        if not allowed:
                            if norm_path(xlam) not in _SKIP_LOGGED:
                                _SKIP_LOGGED.add(norm_path(xlam))
                                outq.put(("log", "跳过 %s：%s"
                                          % (os.path.basename(xlam), why)))
                            continue
                        try:
                            result = provision(app, hwnd, xlam)
                            if result == "foreign":
                                if norm_path(xlam) not in _SKIP_LOGGED:
                                    _SKIP_LOGGED.add(norm_path(xlam))
                                    outq.put(("log", "跳过 %s：存在同名异路径副本"
                                              % os.path.basename(xlam)))
                                continue
                            if result in ("loaded", "already"):
                                loaded_paths.add(norm_path(xlam))
                                if result == "loaded":
                                    outq.put(("log", "已装载 %s"
                                              % os.path.basename(xlam)))
                            elif result == "failed":
                                outq.put(("log", "装载 %s 失败"
                                          % os.path.basename(xlam)))
                        except Exception as exc:
                            outq.put(("log", "装载 %s 失败：%s"
                                      % (os.path.basename(xlam), exc)))
                outq.put(("status", cycle, len(wins), len(targets),
                          len(loaded_paths), len(loaded_paths)))
            except Exception as exc:
                outq.put(("log", "扫描异常：%s" % exc))
            deadline = interval
            while deadline > 0 and not stop.is_set() and not kick.is_set():
                stop.wait(min(0.05, deadline))
                deadline -= 0.05
    finally:
        comtypes.CoUninitialize()


# ---- 开机启动：计划任务 Highest ----
def autostart_command():
    exe = sys.executable
    if getattr(sys, "frozen", False):
        return '"%s" --background' % exe
    return '"%s" "%s" --background' % (exe, os.path.abspath(__file__))


def _schtasks(args):
    try:
        p = subprocess.run(["schtasks"] + args, capture_output=True, text=True,
                           creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
                           timeout=20)
        return p.returncode, p.stdout, p.stderr
    except (OSError, subprocess.TimeoutExpired) as exc:
        return -1, "", str(exc)


def _local_tag(tag):
    return tag.rsplit("}", 1)[-1]


def task_info():
    """读取计划任务的动作/权限/登录类型；XML 标签不受系统语言影响。"""
    rc, out, _err = _schtasks(["/Query", "/TN", TASK_NAME, "/XML"])
    if rc != 0:
        return None
    try:
        root = ET.fromstring(out.lstrip("\ufeff \r\n\t"))
    except ET.ParseError:
        return None
    info = {"enabled": None, "command": "", "arguments": "",
            "run_level": "", "logon_type": ""}
    for el in root.iter():
        name = _local_tag(el.tag)
        text = (el.text or "").strip()
        if name == "Enabled" and info["enabled"] is None:
            info["enabled"] = text.lower() == "true"
        elif name == "Command" and not info["command"]:
            info["command"] = text
        elif name == "Arguments" and not info["arguments"]:
            info["arguments"] = text
        elif name == "RunLevel" and not info["run_level"]:
            info["run_level"] = text
        elif name == "LogonType" and not info["logon_type"]:
            info["logon_type"] = text
    return info


def task_command(info):
    args = info.get("arguments", "")
    return (info.get("command", "") + (" " + args if args else "")).strip()


def norm_task_command(text):
    return " ".join(text.replace('"', "").lower().split())


def _legacy_run_value():
    try:
        key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, RUN_KEY)
        val, _ = winreg.QueryValueEx(key, RUN_NAME)
        winreg.CloseKey(key)
        return val
    except OSError:
        return None


def _remove_legacy_run():
    try:
        key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, RUN_KEY, 0,
                             winreg.KEY_SET_VALUE)
        try:
            winreg.DeleteValue(key, RUN_NAME)
        finally:
            winreg.CloseKey(key)
    except FileNotFoundError:
        pass
    except OSError:
        pass


def is_autostart():
    info = task_info()
    if info and info.get("enabled") is not False:
        return True
    return bool(_legacy_run_value())


def set_autostart(enabled):
    if not enabled:
        _schtasks(["/Delete", "/TN", TASK_NAME, "/F"])
        _remove_legacy_run()
        return
    rc, _out, err = _schtasks([
        "/Create", "/TN", TASK_NAME, "/SC", "ONLOGON",
        "/TR", autostart_command(), "/RL", "HIGHEST", "/IT", "/F"])
    if rc != 0:
        # 计划任务不可用时保留旧 Run 键作为回退，并明确报告
        key = winreg.CreateKeyEx(winreg.HKEY_CURRENT_USER, RUN_KEY, 0,
                                 winreg.KEY_SET_VALUE)
        try:
            winreg.SetValueEx(key, RUN_NAME, 0, winreg.REG_SZ,
                              autostart_command())
        finally:
            winreg.CloseKey(key)
        raise RuntimeError(
            "计划任务注册失败，已回退注册表 Run（登录会弹 UAC）：%s"
            % err.strip()[:200])
    _remove_legacy_run()


def migrate_autostart():
    """把 v1.0.0 的 HKCU Run 一次性迁移到计划任务，避免两处都启动。"""
    if not _legacy_run_value():
        return None
    if task_info() is None:
        set_autostart(True)
        return "旧注册表开机启动已迁移为计划任务（Highest）"
    _remove_legacy_run()
    return "旧注册表开机启动已清理，计划任务保持启用"


def autostart_targets_exist():
    if not os.path.isfile(sys.executable):
        return False
    return (getattr(sys, "frozen", False)
            or os.path.isfile(os.path.abspath(__file__)))


def autostart_health_fix():
    """修正路径漂移/权限降级/旧键残留；程序文件删除则取消任务。返回提示或 None。"""
    notes = []
    if _legacy_run_value():
        _remove_legacy_run()
        notes.append("已清理残留的注册表开机启动项")
    info = task_info()
    if info is None:
        return "；".join(notes) or None
    if not autostart_targets_exist():
        set_autostart(False)
        notes.append("程序文件已删除，已自动取消开机启动")
        return "；".join(notes)
    expected = norm_task_command(autostart_command())
    actual = norm_task_command(task_command(info))
    wrong_level = info.get("run_level") not in ("HighestAvailable", "Highest")
    wrong_logon = info.get("logon_type") not in ("InteractiveToken", "Interactive")
    if expected != actual or wrong_level or wrong_logon:
        set_autostart(True)
        notes.append("开机启动任务已修复（路径/权限/登录类型）")
    return "；".join(notes) or None


_LOADER_EXE_RE = re.compile(r"^AddInLoader.*\.exe$", re.I)
_EXE_VERSION_RE = re.compile(r"[_-]v?(\d+)\.(\d+)\.(\d+)", re.I)


def version_tuple(text):
    m = re.search(r"(\d+)\.(\d+)\.(\d+)", str(text))
    return tuple(int(x) for x in m.groups()) if m else (0, 0, 0)


def exe_version(path):
    """从 exe 文件名解析版本；识别不了返回 None（保守让位，不误杀）。"""
    m = _EXE_VERSION_RE.search(os.path.basename(str(path)))
    return tuple(int(x) for x in m.groups()) if m else None


def iter_loader_processes():
    """枚举运行中的 AddInLoader*.exe 进程，yield (pid, exe_path)。"""
    TH32CS_SNAPPROCESS = 0x00000002
    INVALID = ctypes.c_void_p(-1).value

    class PROCENTRY(ctypes.Structure):
        _fields_ = [
            ("dwSize", wintypes.DWORD),
            ("cntUsage", wintypes.DWORD),
            ("th32ProcessID", wintypes.DWORD),
            ("th32DefaultHeapID", ctypes.POINTER(ctypes.c_ulong)),
            ("th32ModuleID", wintypes.DWORD),
            ("cntThreads", wintypes.DWORD),
            ("th32ParentProcessID", wintypes.DWORD),
            ("pcPriClassBase", ctypes.c_long),
            ("dwFlags", wintypes.DWORD),
            ("szExeFile", wintypes.WCHAR * 260),
        ]

    snap = _kernel32.CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, 0)
    if not snap or snap == INVALID:
        return
    try:
        entry = PROCENTRY()
        entry.dwSize = ctypes.sizeof(PROCENTRY)
        ok = _kernel32.Process32FirstW(snap, ctypes.byref(entry))
        while ok:
            name = str(entry.szExeFile).split("\x00", 1)[0]
            pid = int(entry.th32ProcessID)
            if name and _LOADER_EXE_RE.match(name):
                h = _kernel32.OpenProcess(
                    PROCESS_QUERY_LIMITED_INFORMATION, False, pid)
                if h:
                    try:
                        size = wintypes.DWORD(1024)
                        buf = ctypes.create_unicode_buffer(1024)
                        if _kernel32.QueryFullProcessImageNameW(
                                h, 0, buf, ctypes.byref(size)):
                            yield pid, buf.value
                    finally:
                        _kernel32.CloseHandle(h)
            ok = _kernel32.Process32NextW(snap, ctypes.byref(entry))
    finally:
        _kernel32.CloseHandle(snap)


def running_loader_versions():
    """返回 [(pid, exe_path, 版本元组或None)]，排除自身与自身的父进程。

    PyInstaller onefile 的 bootloader 父进程与实际子进程同名（都是
    AddInLoader_vX.Y.Z.exe）；若不排除父进程，仲裁会误判"同版本实例已在运行"
    而让位，随后单实例 mutex 被低版本占用，高版本静默退出。
    """
    me = os.getpid()
    parent = os.getppid()
    return [(pid, path, exe_version(path))
            for pid, path in iter_loader_processes()
            if pid != me and pid != parent]


def terminate_pids(pids):
    killed = 0
    for pid in pids:
        h = _kernel32.OpenProcess(0x0001, False, pid)  # PROCESS_TERMINATE
        if not h:
            continue
        try:
            if _kernel32.TerminateProcess(h, 0):
                killed += 1
        finally:
            _kernel32.CloseHandle(h)
    if killed:
        time.sleep(0.5)
    return killed


def clear_lower_autostart(paths):
    """清除指向低版本 exe 的开机启动，并重注册指向当前版本。"""
    cleared = False
    run_val = _legacy_run_value()
    if run_val and any(norm_path(p) in norm_path(run_val) for p in paths):
        _remove_legacy_run()
        cleared = True
    info = task_info()
    if info:
        cmd = task_command(info)
        if any(norm_path(p) in norm_path(cmd) for p in paths):
            set_autostart(False)
            cleared = True
    if not cleared:
        return []
    try:
        set_autostart(True)
        return ["旧启动项已改指向当前版本"]
    except Exception as exc:
        return ["旧启动项已清除但重建失败：%s" % exc]


def take_over_older_instances():
    """版本仲裁：更高/未知版本在跑则让位；仅低版本在跑则结束并接管。"""
    my = version_tuple(VERSION)
    others = running_loader_versions()
    if any(ver is None or ver >= my for _pid, _path, ver in others):
        return None
    lower = [(pid, path) for pid, path, ver in others
             if ver is not None and ver < my]
    if not lower:
        return None
    killed = terminate_pids([pid for pid, _ in lower])
    notes = clear_lower_autostart([path for _, path in lower])
    msg = "已接管低版本实例：结束 %d 个进程" % killed
    if notes:
        msg += "，清除%s" % "、".join(notes)
    return msg


def fetch_latest_release(timeout=8):
    """取最新版本号与链接，返回 (tag, html_url, err)。

    优先读更新清单；失败则降级 releases/latest，再读 tags，
    这样只有 Git tag 的仓库也能检查更新。
    """
    for manifest in (MANIFEST_API, UPDATE_MANIFEST):
        try:
            req = urllib.request.Request(
                manifest, headers={"User-Agent": "AddInLoader/%s" % VERSION})
            with urllib.request.urlopen(req, timeout=timeout) as resp:
                payload = resp.read().decode("utf-8")
            if manifest == MANIFEST_API:
                envelope = json.loads(payload)
                info = json.loads(base64.b64decode(envelope["content"]).decode("utf-8"))
            else:
                info = json.loads(payload)
            tag = str((info or {}).get("version") or "").strip()
            page = str((info or {}).get("page") or RELEASE_PAGE_URL).strip()
            if tag and version_tuple(tag) != (0, 0, 0):
                return tag, page, None
        except (urllib.error.URLError, OSError, ValueError, KeyError, AttributeError):
            continue
    err = ""
    try:
        req = urllib.request.Request(
            UPDATE_API, headers={"User-Agent": "AddInLoader/%s" % VERSION})
        with urllib.request.urlopen(req, timeout=timeout) as resp:
            data = json.loads(resp.read().decode("utf-8"))
        tag = str(data.get("tag_name") or "").strip()
        if tag:
            url = str(data.get("html_url") or RELEASE_PAGE_URL).strip()
            return tag, url, None
        err = "发布页未返回版本信息"
    except urllib.error.HTTPError as exc:
        if exc.code == 404:
            err = "发布页未创建 Release"
        else:
            return "", "", str(exc)
    except (urllib.error.URLError, OSError, ValueError, KeyError) as exc:
        return "", "", str(exc)

    try:
        req = urllib.request.Request(
            TAGS_API, headers={"User-Agent": "AddInLoader/%s" % VERSION})
        with urllib.request.urlopen(req, timeout=timeout) as resp:
            data = json.loads(resp.read().decode("utf-8"))
        tags = [str((item or {}).get("name") or "").strip()
                for item in data if isinstance(data, list)]
        tags = [t for t in tags if t and version_tuple(t) != (0, 0, 0)]
        if not tags:
            return "", "", err or "未找到有效版本标签"
        best = max(tags, key=version_tuple)
        return best, RELEASE_PAGE_URL, None
    except (urllib.error.URLError, OSError, ValueError, KeyError) as exc:
        return "", "", "%s；tags 读取失败：%s" % (err, exc)
    return "", "", err or "未找到版本标签"


def compare_versions(a, b):
    """返回 1 表示 a 更新，-1 更旧，0 相同。"""
    ta, tb = version_tuple(a), version_tuple(b)
    return (ta > tb) - (ta < tb)


# ---- 一键更新：清单 + 下载校验 + 校验即启动 ----
# 清单与 exe 同源（自己的 GitHub Releases），HTTPS 保障传输。SHA-256 用于发现下载损坏/截断，
# 不承担防恶意替换职责（哈希与文件同源）。若以后接入第三方加速镜像，必须补回发布签名。
UPDATE_MANIFEST_URL = "https://raw.githubusercontent.com/YehOlt/AddInLoader/refs/heads/main/update.json"
UPDATE_MAX_MANIFEST = 64 * 1024
UPDATE_MAX_EXE = 200 * 1024 * 1024
_LOOPBACK_HOSTS = {"127.0.0.1", "localhost"}


def validate_update_manifest(m):
    """字段校验；非法抛 ValueError。"""
    if not isinstance(m, dict):
        raise ValueError("清单必须是 JSON 对象")
    if version_tuple(m.get("version")) == (0, 0, 0):
        raise ValueError("清单 version 非法")
    exe = m.get("exe")
    if not isinstance(exe, str) or not exe.lower().endswith(".exe"):
        raise ValueError("清单 exe 必须以 .exe 结尾")
    if re.search(r"[\\/]", exe) or ".." in exe:
        raise ValueError("清单 exe 不得含路径分隔或 ..")
    if not re.search(r"[_-]v?(\d+\.\d+\.\d+)", exe, re.I):
        raise ValueError("清单 exe 文件名必须含可解析版本号")
    sha = m.get("sha256")
    if not isinstance(sha, str) or not re.fullmatch(r"[0-9a-fA-F]{64}", sha):
        raise ValueError("清单 sha256 必须是 64 位十六进制")
    url = m.get("url")
    if not isinstance(url, str) or not url:
        raise ValueError("清单缺少 url")
    size = m.get("size")
    if size is not None and (isinstance(size, bool) or not isinstance(size, int) or size < 0):
        raise ValueError("清单 size 必须是非负整数")
    return m


def fetch_update_manifest(url=None, timeout=8):
    """取清单 → 字段校验。返回 (清单, 错误说明)。"""
    target = url or UPDATE_MANIFEST_URL
    try:
        req = urllib.request.Request(target, headers={"User-Agent": "AddInLoader/%s" % VERSION})
        with urllib.request.urlopen(req, timeout=timeout) as resp:
            raw = resp.read(UPDATE_MAX_MANIFEST + 1)
        if len(raw) > UPDATE_MAX_MANIFEST:
            return None, "清单异常大"
        m = json.loads(raw.decode("utf-8"))
        validate_update_manifest(m)
        return m, None
    except Exception as exc:
        return None, str(exc)


def download_verified_exe(manifest, dest_dir, on_progress=None, timeout=30):
    """下载到 dest_dir 的 .part 文件，边下边验魔数/大小，收尾验 SHA-256；
    通过后 os.replace 到清单指定的文件名。返回 (最终路径, 错误说明)。"""
    name = manifest["exe"]
    expected_sha = manifest["sha256"].lower()
    expected_size = manifest.get("size")
    part = os.path.join(dest_dir, name + ".part")
    final = os.path.join(dest_dir, name)
    url = manifest.get("url") or ""
    if not url:
        return None, "清单缺少下载地址"
    host = (urllib.parse.urlsplit(url).hostname or "").lower()
    if host not in _LOOPBACK_HOSTS and not url.startswith("https://"):
        return None, "下载地址必须是 https"
    hasher = hashlib.sha256()
    state = {"n": 0, "head": b"", "err": None}

    def fail(reason):
        if state["err"] is None:
            state["err"] = reason

    try:
        os.makedirs(dest_dir, exist_ok=True)
        req = urllib.request.Request(url, headers={"User-Agent": "AddInLoader/%s" % VERSION})
        with urllib.request.urlopen(req, timeout=timeout) as resp, open(part, "wb") as out:
            while True:
                block = resp.read(256 * 1024)
                if not block:
                    break
                if state["err"]:
                    break
                if len(state["head"]) < 2:
                    state["head"] = (state["head"] + block[:2])[:2]
                    if state["head"] != b"MZ":
                        fail("下载内容不是 exe（缺少 MZ 魔数）")
                        break
                state["n"] += len(block)
                if state["n"] > UPDATE_MAX_EXE:
                    fail("文件超过大小上限")
                    break
                out.write(block)
                hasher.update(block)
                if on_progress:
                    on_progress(state["n"], expected_size or -1)
        if state["err"]:
            return None, state["err"]
        if expected_size is not None and state["n"] != expected_size:
            return None, "大小与清单不符（收到 %d，清单 %s）" % (state["n"], expected_size)
        if hasher.hexdigest() != expected_sha:
            return None, "SHA-256 校验不符"
        os.replace(part, final)
        return final, None
    except Exception as exc:
        return None, str(exc)
    finally:
        if os.path.exists(part):
            try:
                os.remove(part)
            except OSError:
                pass


def launch_verified_exe(path):
    """启动已校验的新版 exe（不走 shell）。返回 (是否成功, 说明)。"""
    try:
        subprocess.Popen([path], close_fds=True)
        return True, ""
    except Exception as exc:
        return False, str(exc)


def run_one_click_update(manifest_url=None, log=print, dry_run=False, current=VERSION):
    """一键更新：清单验签 → 版本比较 → 下载校验 → （非 dry_run）启动新版。
    返回用户可读的汇总文本。"""
    log("检查更新：正在获取版本清单…")
    m, err = fetch_update_manifest(manifest_url)
    if err:
        return "失败：%s" % err
    latest = m["version"]
    if compare_versions(latest, current) <= 0:
        return "当前已是最新版 v%s。" % current
    log("发现新版本 %s（当前 v%s），开始下载…" % (latest, current))
    exe_dir = os.path.dirname(os.path.abspath(sys.executable if getattr(sys, "frozen", False) else __file__))
    path, err = download_verified_exe(m, exe_dir)
    if err:
        return "失败：%s" % err
    log("下载完成并通过 SHA-256 校验：%s" % path)
    if dry_run:
        return "试运行完成：已下载并校验 %s（未启动）。" % path
    ok, err = launch_verified_exe(path)
    if not ok:
        return "失败：新版已下载并校验，但启动失败：%s（文件：%s）" % (err, path)
    return "已启动新版本 %s；旧版窗口即将被新版接管。" % latest


def silent_update_target(tag, url, err, current=VERSION):
    """静默检查的落点：有新版返回发布页 URL，其余返回空串（不弹窗、不打开）。"""
    if err or not tag:
        return ""
    if compare_versions(tag, current) <= 0:
        return ""
    return url or RELEASE_PAGE_URL


def acquire_single_instance():
    """返回 (mutex句柄, 是否首个实例)。非首个实例会广播唤起消息。"""
    _kernel32.GetLastError.restype = wintypes.DWORD
    mutex = _kernel32.CreateMutexW(None, False, MUTEX_NAME)
    last = _kernel32.GetLastError()
    if last == 183:  # ERROR_ALREADY_EXISTS
        _user32.PostMessageW(0xFFFF, WM_SHOW_APP, 0, 0)  # HWND_BROADCAST
        return mutex, False
    return mutex, True


def selfcheck(xlam):
    print("%s v%s" % (TOOL_NAME, VERSION))
    print("Python %s" % sys.version.split()[0])
    ok = True
    print("comtypes        : OK")
    try:
        wins = find_excel_doc_windows()
        print("窗口枚举        : OK（当前 EXCEL.EXE 的 EXCEL7 窗口 %d 个）" % len(wins))
    except Exception as exc:
        print("窗口枚举        : FAIL %s" % exc)
        ok = False
    info = task_info()
    if info:
        print("开机启动        : 计划任务 %s（%s / %s / %s）" % (
            "已启用" if info.get("enabled") is not False else "已禁用",
            TASK_NAME, info.get("run_level") or "?", info.get("logon_type") or "?"))
    else:
        print("开机启动        : %s" % (
            "已启用（旧 Run 键，待迁移）" if _legacy_run_value() else "未启用"))
    print("超时自动强制    : %s" % (
        "已启用" if read_prefs().get(FORCE_PREF_KEY) else "未启用（超时后询问）"))
    target = sys.executable if getattr(sys, "frozen", False) else __file__
    try:
        digest = hashlib.sha256()
        with open(target, "rb") as fh:
            for chunk in iter(lambda: fh.read(1 << 20), b""):
                digest.update(chunk)
        print("文件哈希        : SHA-256 %s（%s）"
              % (digest.hexdigest(), os.path.basename(target)))
    except OSError as exc:
        print("文件哈希        : FAIL %s" % exc)
    print("托盘结构        : OK（%d bytes）" % ctypes.sizeof(NOTIFYICONDATAW))
    for ver in OFFICE_VERS:
        entries = open_entries(ver)
        if entries:
            print("OPEN 登记       : %s（%d 项）" % (ver, len(entries)))
            for name, _val, p in entries:
                print("  · %s = %s" % (name, p or "?"))
            break
    else:
        print("OPEN 登记       : 未发现")
    print("受信任位置      : %d 处（自动装载仅接受这些目录）" % len(trusted_locations(True)))
    print("管理目录        : %s（%s）" % (
        managed_dir(), "已受信任" if is_auto_trusted(
            os.path.join(managed_dir(), "probe.xlam")) else "未受信任"))
    if xlam:
        if os.path.isfile(xlam):
            print("指定加载项      : OK %s" % xlam)
            print("登记状态        : %s" % (registered_at(xlam) or "未登记"))
        else:
            print("指定加载项      : 缺失 %s" % xlam)
            ok = False
    else:
        items = discover_addins()
        print("自动发现        : OK（自定义加载项 %d 个）" % len(items))
        for p, src in items:
            print("  · %s  [%s]" % (os.path.basename(p), src))
            print("    %s" % p)
        if not items:
            print("  未发现自定义加载项，请检查安装。")
    return 0 if ok else 3


class ToolTip:
    """鼠标悬停提示。"""
    MAX_WIDTH = 400

    def __init__(self, widget, text, delay=400):
        self.widget = widget
        self.text = text
        self.delay = delay
        self.tip = None
        self.timer = None
        widget.bind("<Enter>", self._enter)
        widget.bind("<Leave>", self._leave)

    def _wrapped_text(self):
        font = tkfont.Font(family="Microsoft YaHei", size=9)
        lines, cur = [], ""
        s = self.text
        i = 0
        while i < len(s):
            ch = s[i]
            if ch.isascii() and (ch.isalnum() or ch in "/.-_"):
                j = i + 1
                while j < len(s) and s[j].isascii() and (
                        s[j].isalnum() or s[j] in "/.-_"):
                    j += 1
                tok = s[i:j]
            else:
                tok, j = ch, i + 1
            if cur and font.measure(cur + tok) > self.MAX_WIDTH:
                lines.append(cur)
                cur = tok.lstrip()
            else:
                cur += tok
            i = j
        if cur:
            lines.append(cur)
        return "\n".join(lines)

    def _enter(self, _e=None):
        self.timer = self.widget.after(self.delay, self._show)

    def _leave(self, _e=None):
        if self.timer:
            self.widget.after_cancel(self.timer)
            self.timer = None
        if self.tip:
            self.tip.destroy()
            self.tip = None

    def _show(self):
        self.timer = None
        if self.tip:
            return
        self.tip = tk.Toplevel(self.widget)
        self.tip.wm_overrideredirect(True)
        lbl = tk.Label(self.tip, text=self._wrapped_text(), justify="left",
                       anchor="w", fg="#333333", bg="#ffffe0", relief="solid",
                       bd=1, font=("Microsoft YaHei", 9), padx=6, pady=4)
        lbl.pack()
        self.tip.update_idletasks()
        sw = self.tip.winfo_screenwidth()
        x = self.widget.winfo_rootx() + 20
        x = min(x, sw - self.tip.winfo_width() - 8)
        y = self.widget.winfo_rooty() + self.widget.winfo_height() + 4
        self.tip.wm_geometry("+%d+%d" % (x, y))


class LoaderApp:
    def __init__(self, root, fixed=None, background=False):
        self.root = root
        self.fixed = fixed
        self.background = background
        for p in (fixed or []):
            mark_session_trusted(p)
        self.deselected = set()
        self.addins = [(p, "手动指定") for p in fixed] if fixed else discover_addins()
        self.extra = []
        self.discover_sig = self._sig(self.addins)
        self.last_discover = time.monotonic()
        self.started_at = time.monotonic()
        self.last_autostart_check = time.monotonic()
        self.lock = threading.Lock()
        self.stop = threading.Event()
        self.kick = threading.Event()
        self.unloadq = queue.Queue()
        self.outq = queue.Queue()
        self.log_lines = deque(maxlen=LOG_LIMIT)
        self.tray_nid = None
        self.old_wndproc = None
        self.new_wndproc = None
        self.tray_hwnd = None
        self.menu_open = False
        self._close_tip_shown = False

        root.title("%s v%s" % (TOOL_NAME, VERSION))
        root.geometry("760x680")
        root.minsize(700, 560)
        try:
            root.iconbitmap(default="")
        except Exception:
            pass

        btnf = ttk.Frame(root)
        btnf.pack(fill="x", padx=12, pady=(10, 6))
        self.btn_pick = ttk.Button(btnf, text="临时启动",
                                   command=self.choose_xlam, width=12)
        self.btn_pick.pack(side="left")
        self.btn_refresh = ttk.Button(btnf, text="刷新", command=self.refresh,
                                      width=11)
        self.btn_refresh.pack(side="left", padx=6)
        self._make_about_badge(btnf)

        top = ttk.Frame(root)
        top.pack(fill="x", padx=12, pady=6)
        info_area = ttk.LabelFrame(top, text=" 安装信息 ")
        info_area.pack(side="left", fill="both", expand=True, padx=(0, 3))
        versions = detect_office_versions()
        office_text = ("、".join("%s（%s）" % (v, b) for v, b in versions)
                       if versions else "未检测到 Microsoft Office")
        tk.Label(info_area, anchor="w", justify="left", wraplength=330,
                 font=("Microsoft YaHei", 9),
                 text="检测到 Office：%s" % office_text).pack(
                     fill="x", padx=8, pady=(6, 2))
        dir_row = tk.Frame(info_area)
        dir_row.pack(fill="x", padx=8, pady=2)
        tk.Label(dir_row, anchor="w", font=("Microsoft YaHei", 9),
                 text="管理目录：").pack(side="left")
        dir_link = tk.Label(dir_row, anchor="w", justify="left", wraplength=340,
                            font=("Microsoft YaHei", 9, "underline"),
                            fg="#0563c1", cursor="hand2", text=managed_dir())
        dir_link.pack(side="left")
        dir_link.bind("<Button-1>", lambda _e: self.open_managed_dir())
        self.lbl_install_state = tk.Label(info_area, anchor="w", justify="left",
                                          wraplength=330,
                                          font=("Microsoft YaHei", 9), fg="#8a6d00")
        self.lbl_install_state.pack(fill="x", padx=8, pady=(2, 6))
        self.refresh_install_state()

        opt = ttk.LabelFrame(top, text=" 选项 ")
        opt.pack(side="right", fill="both", padx=(3, 0))
        self.var_autostart = tk.BooleanVar(value=is_autostart())
        ttk.Checkbutton(
            opt, text="开机启动",
            variable=self.var_autostart,
            command=self.on_autostart_ui).pack(anchor="w", padx=8, pady=(6, 2))
        self.var_trust = tk.BooleanVar(value=True)
        trust_cb = ttk.Checkbutton(
            opt, text="将管理目录加入 Excel 受信任位置（推荐）",
            variable=self.var_trust)
        trust_cb.pack(anchor="w", padx=8, pady=(2, 6))
        ToolTip(trust_cb,
                "勾选后打开 Excel 不再弹出“禁用宏/启用内容”安全提示；"
                "不勾选也可以长期启动，但打开 Excel 时可能需手动点一次“启用内容”。")
        self.var_force = tk.BooleanVar(
            value=bool(read_prefs().get(FORCE_PREF_KEY, False)))
        force_cb = ttk.Checkbutton(
            opt, text="卸载时允许强制结束进程",
            variable=self.var_force,
            command=lambda: write_pref(FORCE_PREF_KEY, self.var_force.get()))
        force_cb.pack(anchor="w", padx=8, pady=(2, 2))
        ToolTip(force_cb,
                "高级选项：默认关闭。等待 Excel 自行退出超时后不再二次确认，"
                "直接强制结束；长期启动前的关闭流程同样适用。")

        listf = ttk.LabelFrame(root, text=" 加载项 ")
        listf.pack(fill="x", padx=12, pady=6)
        canvas = tk.Canvas(listf, height=150, highlightthickness=0)
        bar = ttk.Scrollbar(listf, orient="vertical", command=canvas.yview)
        self.listbox = ttk.Frame(canvas)
        self.listbox.bind(
            "<Configure>", lambda _e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        canvas.create_window((0, 0), window=self.listbox, anchor="nw")
        canvas.configure(yscrollcommand=bar.set)
        canvas.pack(side="left", fill="both", expand=True, padx=(6, 0), pady=6)
        bar.pack(side="right", fill="y", pady=6)
        self.canvas = canvas
        root.bind_all("<MouseWheel>", self._on_list_wheel)
        self.checks = {}
        self.build_list()

        info = ttk.LabelFrame(root, text=" 运行状态 ")
        info.pack(fill="x", padx=12, pady=6)
        self.lbl_state = tk.Label(info, anchor="w", font=("Microsoft YaHei", 9),
                                  fg="#1a6b1a")
        self.lbl_state.pack(fill="x", padx=8, pady=(6, 2))
        self.lbl_stat = tk.Label(info, anchor="w", font=("Consolas", 9), fg="#555555")
        self.lbl_stat.pack(fill="x", padx=8, pady=(2, 6))
        self.set_initial_status()

        logf = ttk.LabelFrame(root, text=" 日志 ")
        logf.pack(fill="both", expand=True, padx=12, pady=(6, 10))
        self.log = scrolledtext.ScrolledText(
            logf, height=5, wrap="word", font=("Consolas", 9), state="disabled"
        )
        self.log.pack(fill="both", expand=True, padx=6, pady=6)

        self.append_log("装载器已启动，发现 %d 个自定义加载项。" % len(self.addins))
        if STARTUP_NOTICE:
            self.append_log(STARTUP_NOTICE)
        self.append_log("打开 Excel，等待对应选项卡出现。")

        root.update_idletasks()
        self.setup_tray()

        threading.Thread(target=worker,
                         args=(self.stop, self.outq, self.unloadq, self.kick,
                               self.get_targets, SCAN_INTERVAL),
                         daemon=True).start()
        root.protocol("WM_DELETE_WINDOW", self.on_close_request)
        root.after(120, self.poll)
        root.after(600, self.silent_check_update)
        if not self.background:
            root.after(400, lambda: self.show_about(auto_close=True))

        if self.background:
            self.hide_to_tray()
        elif is_autostart():
            self.append_log("已设为开机启动（可在“选项”区取消）。")

    def open_managed_dir(self):
        """在资源管理器中打开管理目录，不存在则先创建。"""
        d = managed_dir()
        try:
            os.makedirs(d, exist_ok=True)
            os.startfile(d)
        except Exception as exc:
            self.append_log("打开管理目录失败：%s" % exc)

    def _make_about_badge(self, parent):
        """右上角圆形感叹号按钮，点击打开关于窗。"""
        cv = tk.Canvas(parent, width=24, height=24, highlightthickness=0,
                       bd=0, cursor="hand2")
        f = tkfont.nametofont("TkDefaultFont").copy()
        f.configure(weight="bold", size=12)
        cv._about_font = f

        def repaint(hover=False):
            cv.delete("all")
            cv.create_oval(
                2, 2, 22, 22,
                fill="#dcdcdc" if hover else "#ededed",
                outline="#8a8a8a" if hover else "#b5b5b5",
            )
            cv.create_text(12, 12, text="!", fill="#555555", font=f)

        cv._repaint = repaint
        repaint()
        cv.bind("<Enter>", lambda _e: repaint(True))
        cv.bind("<Leave>", lambda _e: repaint(False))
        cv.bind("<Button-1>", lambda _e: self.show_about())
        cv.pack(side="right", padx=(4, 0))
        return cv

    def silent_check_update(self):
        """启动时自动检查一次，复用 check_for_update 的清单校验与三选一流程。"""
        self.check_for_update(self.root, silent=True)

    def ask_update_choice(self, parent, latest):
        """三选一：直接下载 / 发布页 / 暂不更新。返回 download/page/skip。"""
        chosen = {"value": "skip"}
        dlg = tk.Toplevel(parent)
        dlg.title("发现新版本")
        dlg.transient(parent)
        dlg.resizable(False, False)
        dlg.grab_set()

        def close(choice):
            chosen["value"] = choice
            dlg.destroy()

        page = ttk.Frame(dlg, padding=(20, 16))
        page.pack(fill="both", expand=True)
        ttk.Label(page, text="发现新版本 %s，当前版本 v%s。" % (latest, VERSION),
                  font=("Microsoft YaHei", 10, "bold")).pack(anchor="w")
        ttk.Label(page, text="\n请选择处理方式：", anchor="w").pack(anchor="w")
        row = ttk.Frame(page)
        row.pack(pady=(14, 0), fill="x")
        ttk.Button(row, text="直接下载", width=12,
                   command=lambda: close("download")).pack(side="left")
        ttk.Button(row, text="发布页", width=10,
                   command=lambda: close("page")).pack(side="left", padx=(8, 0))
        ttk.Button(row, text="暂不更新", width=10,
                   command=lambda: close("skip")).pack(side="left", padx=(8, 0))

        dlg.protocol("WM_DELETE_WINDOW", lambda: close("skip"))
        dlg.bind("<Escape>", lambda _e: close("skip"))
        dlg.update_idletasks()
        x = parent.winfo_rootx() + (parent.winfo_width() - dlg.winfo_width()) // 2
        y = parent.winfo_rooty() + (parent.winfo_height() - dlg.winfo_height()) // 2
        dlg.geometry("+%d+%d" % (max(x, 0), max(y, 0)))
        dlg.wait_window()
        return chosen["value"]

    def _handle_update_found(self, parent, manifest, latest):
        """发现新版：三选一后分派。"""
        choice = self.ask_update_choice(parent, latest)
        if choice == "download":
            self.append_log("选择直接下载：%s" % latest)
            self._download_and_launch(parent, manifest, latest)
        elif choice == "page":
            target = manifest.get("page") or RELEASE_PAGE_URL
            self.append_log("选择前往发布页：%s" % target)
            self.open_release_page(target)
        else:
            self.append_log("选择暂不更新：%s" % latest)

    def check_for_update(self, parent, silent=False):
        """检查更新：手动按钮与启动检查共用。清单校验 → 三选一。

        silent=True 时抑制「已是最新版」与失败弹窗，发现新版照常三选一。
        """
        box = queue.Queue()

        def worker():
            box.put(fetch_update_manifest())

        threading.Thread(target=worker, daemon=True).start()

        def poll():
            try:
                manifest, err = box.get_nowait()
            except queue.Empty:
                self.root.after(100, poll)
                return
            if err:
                self.append_log("检查更新失败：%s" % err)
                if silent:
                    return
                answer = messagebox.askyesno(
                    "检查更新",
                    "检查更新失败：%s\n\n是否打开发布页手动查看？" % err,
                    parent=parent)
                if answer:
                    self.open_release_page(RELEASE_PAGE_URL)
                return
            latest = manifest["version"]
            if compare_versions(latest, VERSION) <= 0:
                self.append_log("当前已是最新版 v%s。" % VERSION)
                if not silent:
                    messagebox.showinfo(
                        "检查更新", "当前已是最新版 v%s。" % VERSION, parent=parent)
                return
            self.append_log("发现新版本 %s，当前版本 v%s。" % (latest, VERSION))
            self._handle_update_found(parent, manifest, latest)

        self.root.after(100, poll)

    def _download_and_launch(self, parent, manifest, latest):
        """后台下载校验，完成后回主线程启动新版。"""
        box = queue.Queue()

        def worker():
            def progress(done, total):
                self.root.after(0, lambda: self.append_log(
                    "下载中：%d / %s 字节" % (done, total)))
            box.put(download_verified_exe(
                manifest,
                os.path.dirname(os.path.abspath(sys.executable
                                                if getattr(sys, "frozen", False) else __file__)),
                on_progress=progress))

        threading.Thread(target=worker, daemon=True).start()

        def poll():
            try:
                path, err = box.get_nowait()
            except queue.Empty:
                self.root.after(100, poll)
                return
            if err:
                self.append_log("下载或校验失败：%s" % err)
                messagebox.showerror("更新失败", "下载或校验失败：%s" % err, parent=parent)
                return
            self.append_log("下载完成并通过 SHA-256 校验：%s" % path)
            ok, err = launch_verified_exe(path)
            if not ok:
                messagebox.showerror("更新失败", "新版已下载并校验，但启动失败：%s\n文件：%s" % (err, path),
                                     parent=parent)
                return
            messagebox.showinfo(
                "更新完成",
                "新版 %s 已下载并启动。\n\n当前窗口即将被新版接管，旧版文件保留不变。"
                % latest, parent=parent)

        self.root.after(100, poll)


    def show_about(self, auto_close=False):
        """关于窗：含两种启动模式对比表，发布页与检查更新按钮。"""
        dlg = tk.Toplevel(self.root)
        dlg.title("关于")
        dlg.transient(self.root)
        dlg.resizable(False, False)
        dlg.grab_set()

        def close(_e=None):
            timer = getattr(dlg, "_about_timer", None)
            if timer:
                try:
                    self.root.after_cancel(timer)
                except tk.TclError:
                    pass
            dlg.destroy()

        page = ttk.Frame(dlg, padding=(18, 14))
        page.pack(fill="both", expand=True)

        title_font = tkfont.nametofont("TkDefaultFont").copy()
        title_font.configure(weight="bold", size=16)
        dlg._about_title_font = title_font
        ttk.Label(page, text="%s v%s" % (TOOL_NAME, VERSION),
                  font=title_font).pack(pady=(4, 0))

        rows = [
            ("", "临时启动", "长期启动"),
            ("作用范围", "装载器运行时持续装载", "下次手动打开 Excel 自动加载"),
            ("写注册表", "不写", "写 OPEN + 受信任位置"),
            ("文件处理", "直接用原文件", "复制副本到管理目录（加 AIL_ 前缀）"),
            ("覆盖的场景", "临时用一下，不常驻", "日常手动开 Excel 时免得每次都装"),
            ("装载器退出后", "失效", "仍然有效"),
        ]
        table = tk.Frame(page, bg="#d8e4d8", bd=1)
        table.pack(fill="x", pady=(14, 2))
        for r, row in enumerate(rows):
            for c, text in enumerate(row):
                if r == 0:
                    bg, fg, bold = "#e4f0e4", "#2f3b2f", True
                elif c == 0:
                    bg, fg, bold = "#eef2ee", "#444444", False
                else:
                    bg, fg, bold = "#ffffff", "#333333", False
                tk.Label(table, text=text, bg=bg, fg=fg,
                         font=("Microsoft YaHei", 9, "bold" if bold else "normal"),
                         anchor="w", justify="left", padx=6, pady=5,
                         wraplength=150).grid(row=r, column=c, sticky="nsew",
                                              padx=1, pady=1)

        ttk.Label(
            page,
            text="本程序不作任何担保。详情见 MIT 许可协议。",
            wraplength=380, justify="center", foreground="#777777",
        ).pack(pady=(8, 4))

        btn_row = ttk.Frame(page)
        btn_row.pack(pady=(12, 0))
        ttk.Button(btn_row, text="发布页",
                   command=lambda: self.open_release_page(RELEASE_PAGE_URL)
                   ).pack(side="left")
        ttk.Button(btn_row, text="检查更新",
                   command=lambda: self.check_for_update(dlg)
                   ).pack(side="left", padx=(8, 0))

        dlg.bind("<Escape>", close)
        dlg.protocol("WM_DELETE_WINDOW", close)
        dlg.update_idletasks()
        w, h = dlg.winfo_width(), dlg.winfo_height()
        x = self.root.winfo_rootx() + (self.root.winfo_width() - w) // 2
        y = self.root.winfo_rooty() + (self.root.winfo_height() - h) // 2
        dlg.geometry("+%d+%d" % (max(0, x), max(0, y)))
        dlg.focus_set()
        if auto_close:
            dlg._about_timer = self.root.after(5000, close)

    def refresh_install_state(self):
        """安装信息区的长期启动状态：已登记项数 + 管理目录副本数。"""
        if not hasattr(self, "lbl_install_state"):
            return
        n_reg = len({reg for p, _ in self.all_items()
                     for reg in [registered_at(p)] if reg})
        self.lbl_install_state["text"] = "当前状态：长期启动 %d 项" % n_reg

    def refresh_state_text(self):
        text = "状态：运行中"
        if self.var_autostart.get():
            text += " ｜ 开机启动已开"
        self.lbl_state["text"] = text

    def set_initial_status(self):
        self.refresh_state_text()
        self.lbl_stat["text"] = "扫描 0 次 ｜ 勾选 %d ｜ 已装载 0" % len(
            self.get_targets()
        )

    def all_items(self):
        with self.lock:
            items = list(self.addins)
            known = {norm_path(p) for p, _ in items}
            for p in self.extra:
                n = norm_path(p)
                if n not in known:
                    items.append((p, "临时启动"))
                    known.add(n)
            return sorted(items, key=lambda t: os.path.basename(t[0]).lower())

    def _on_list_wheel(self, event):
        """鼠标在加载项列表区时滚轮滚动该区，不影响日志区。"""
        x, y = self.root.winfo_pointerxy()
        bx = self.canvas.winfo_rootx()
        by = self.canvas.winfo_rooty()
        if (bx <= x < bx + self.canvas.winfo_width()
                and by <= y < by + self.canvas.winfo_height()):
            self.canvas.yview_scroll(-1 if event.delta > 0 else 1, "units")
            return "break"

    def build_list(self):
        for w in list(self.listbox.winfo_children()):
            w.destroy()
        self.checks = {}
        for path, src in self.all_items():
            n = norm_path(path)
            row = ttk.Frame(self.listbox)
            row.pack(fill="x", pady=3)

            reg = registered_at(path)
            var = tk.BooleanVar(value=True if reg else n not in self.deselected)
            cb = ttk.Checkbutton(row, variable=var,
                                 state="disabled" if reg else "normal",
                                 command=lambda p=path, v=var: self.on_toggle(p, v))
            cb.grid(row=0, column=0, sticky="w")
            tk.Label(row, anchor="w", font=("Microsoft YaHei", 9),
                     text="%s  [%s]" % (os.path.basename(path), src)).grid(
                row=0, column=1, sticky="w", padx=4)

            trust, _why = trust_state(path)
            state = ("已登记" if reg else "未登记") + "·" + TRUST_LABELS[trust]
            tk.Label(row, anchor="w", font=("Microsoft YaHei", 8),
                     fg="#1a6b1a" if trust in (TRUST_TRUSTED, TRUST_SESSION) else (
                         "#c0392b" if trust == TRUST_DENIED else "#8a6d00"),
                     text=state).grid(row=0, column=2, sticky="w", padx=6)
            ttk.Button(row, text="卸载" if reg else "长期启动", width=10,
                       command=lambda p=path, r=reg: (
                           self.on_unregister(p) if r else self.on_register(p)
                       )).grid(row=0, column=3, sticky="e", padx=4)

            tk.Label(row, anchor="w", font=("Consolas", 8), fg="#888888",
                     text=path).grid(row=1, column=1, columnspan=3, sticky="w",
                                     padx=4)
            self.checks[n] = var

    def on_toggle(self, path, var):
        n = norm_path(path)
        with self.lock:
            if var.get():
                self.deselected.discard(n)
            else:
                self.deselected.add(n)
        self.sync_autostart_with_targets()

    def get_targets(self):
        with self.lock:
            items = list(self.addins)
            known = {norm_path(p) for p, _ in items}
            for p in self.extra:
                if norm_path(p) not in known:
                    items.append((p, "临时启动"))
                    known.add(norm_path(p))
            return [p for p, _ in items
                    if (registered_at(p) or norm_path(p) not in self.deselected)
                    and xlam_allowed(p)[0]]

    @staticmethod
    def _sig(items):
        return frozenset(norm_path(p) for p, _ in items)

    def apply_discovery(self, items):
        """主线程中应用新发现列表：重建复选框，保留用户已取消的勾选。"""
        with self.lock:
            self.addins = items
            self.discover_sig = self._sig(items)
            valid = self._sig(items) | {norm_path(p) for p in self.extra}
            self.deselected.intersection_update(valid)
        self.build_list()
        self.canvas.configure(scrollregion=self.canvas.bbox("all"))
        self.set_initial_status()
        self.refresh_install_state()
        self.sync_autostart_with_targets()

    def auto_discover(self):
        """定期刷新：注册表/目录变化后自动刷新窗口列表。"""
        if self.fixed:
            return
        items = discover_addins()
        if self._sig(items) != self.discover_sig:
            self.apply_discovery(items)
            self.append_log("加载项列表已更新：%d 个自定义加载项。" % len(items))

    def refresh(self):
        if self.fixed:
            return
        items = discover_addins()
        self.apply_discovery(items)
        self.append_log("已刷新：%d 个自定义加载项。" % len(items))

    def choose_xlam(self):
        path = filedialog.askopenfilename(
            title="选择 Excel 加载项",
            filetypes=[("Excel 加载项", "*.xlam *.xla"), ("所有文件", "*.*")],
        )
        if not path:
            return
        with self.lock:
            known = {norm_path(p) for p in self.extra}
            if norm_path(path) not in known:
                self.extra.append(path)
            self.deselected.discard(norm_path(path))
        mark_session_trusted(path)
        _SKIP_LOGGED.discard(norm_path(path))
        self.kick.set()
        self.build_list()
        self.set_initial_status()
        note = ("；该文件不在 Excel 受信任位置，按会话临时信任装载"
                if not is_auto_trusted(path) else "")
        self.append_log("已选择 %s，正在临时装载（装载器退出或取消勾选后失效，未写注册表%s）"
                        % (os.path.basename(path), note))

    def register_pick(self):
        path = filedialog.askopenfilename(
            title="选择要长期启动的 Excel 加载项",
            filetypes=[("Excel 加载项", "*.xlam *.xla"), ("所有文件", "*.*")],
        )
        if path:
            self.on_register(path)

    def _wait_excel_close(self, reason):
        """提示保存并轮询等待 Excel 退出。True=已退出，False=取消，'timeout'=超时。"""
        result = {"value": None}
        dlg = tk.Toplevel(self.root)
        dlg.title("关闭 Excel")
        dlg.transient(self.root)
        dlg.resizable(False, False)

        def finish(value):
            tid = timer["id"]
            if tid is not None:
                try:
                    dlg.after_cancel(tid)
                except tk.TclError:
                    pass
            result["value"] = value
            dlg.destroy()

        frame = ttk.Frame(dlg, padding=(16, 12))
        frame.pack(fill="both", expand=True)
        ttk.Label(frame, text=reason, wraplength=390,
                  justify="left").pack(anchor="w")
        ttk.Label(frame, text="请先保存工作并关闭所有 Excel 窗口。Excel 退出时会写回加载项状态，"
                  "所以这一步不能跳过。", wraplength=390, justify="left",
                  foreground="#8a6d00").pack(anchor="w", pady=(8, 0))
        lbl = tk.Label(frame, anchor="w", font=("Consolas", 9),
                       foreground="#555555")
        lbl.pack(anchor="w", fill="x", pady=(10, 2))
        bf = ttk.Frame(frame)
        bf.pack(fill="x", pady=(8, 0))
        ttk.Button(bf, text="我已关闭", width=10,
                   command=lambda: check()).pack(side="right")
        ttk.Button(bf, text="取消", width=10,
                   command=lambda: finish(False)).pack(side="right", padx=(8, 0))

        timer = {"id": None}
        deadline = time.monotonic() + EXCEL_CLOSE_TIMEOUT

        def check():
            if not is_excel_running():
                finish(True)
                return
            left = max(0, int(deadline - time.monotonic() + 0.999))
            lbl["text"] = "剩余 %d 秒 ｜ 仍检测到 Excel 运行" % left
            if left <= 0:
                finish("timeout")
                return
            timer["id"] = dlg.after(EXCEL_CLOSE_POLL_MS, check)

        dlg.protocol("WM_DELETE_WINDOW", lambda: finish(False))
        dlg.grab_set()
        check()
        self.root.wait_window(dlg)
        return result["value"]

    def ensure_excel_closed(self, action):
        """返回 True 可继续；默认等用户保存，超时必须确认才强制结束。"""
        if not is_excel_running():
            return True
        state = self._wait_excel_close("%s前需要关闭 Excel。" % action)
        if state is True:
            self.append_log("Excel 已自行退出，可以继续%s。" % action)
            return True
        if state is False:
            self.append_log("%s已取消：Excel 仍在运行。" % action)
            return False
        if not self.var_force.get():
            if not messagebox.askyesno(
                "Excel 未退出",
                "等待 %d 秒后仍检测到 Excel 运行。\n"
                "检测到 Excel 可能回写加载项状态，是否强制结束？\n"
                "强制结束会丢失未保存的内容。" % EXCEL_CLOSE_TIMEOUT,
                parent=self.root,
            ):
                self.append_log("%s已取消：用户未允许强制结束 Excel。" % action)
                return False
        else:
            self.append_log("等待 %d 秒超时，按高级选项直接强制结束 Excel。"
                            % EXCEL_CLOSE_TIMEOUT)
        n = close_excel_processes()
        self.append_log("已强制关闭 %d 个 Excel 进程" % n)
        if is_excel_running():
            self.append_log("%s失败：Excel 仍在运行，请手动关闭后重试。" % action)
            return False
        return True

    def on_register(self, path):
        if not self.ensure_excel_closed("长期启动"):
            return
        # 源路径若已直接登记，先撤销，避免副本与源留下两条 OPEN
        if registered_at(path):
            unregister_addin(path)
        copy, how = copy_to_managed(path)
        if not copy:
            self.append_log("长期启动 %s 失败：%s" % (os.path.basename(path), how))
            return
        if self.var_trust.get():
            trusted_ok, trusted_msg = ensure_trusted_location()
            if not trusted_ok and trusted_msg != "已受信任":
                self.append_log("长期启动 %s 失败：受信任位置登记失败（%s）"
                                % (os.path.basename(path), trusted_msg))
                return
            trust_note = "已加入受信任位置"
        else:
            trust_note = "未加入受信任位置"
        ok, msg = register_addin(copy)
        if not ok:
            self.append_log("长期启动 %s 失败：%s" % (os.path.basename(path), msg))
            return
        # 从临时切到长期：关掉源路径的临时工作簿，清除其加载时间记录与
        # “可用加载宏”对话框里的历史浏览记录，只留管理副本那一条
        if norm_path(path) != norm_path(copy):
            for hwnd in find_excel_doc_windows():
                app = get_excel_app(hwnd)
                if app is not None:
                    unload_from_excel(app, hwnd, path)
            clear_addin_load_time(path)
            clear_addin_manager_entry(path)
        with self.lock:
            self.extra = [p for p in self.extra if norm_path(p) != norm_path(path)]
            known = {norm_path(p) for p in self.extra}
            if norm_path(copy) not in known:
                self.extra.append(copy)
            self.deselected.discard(norm_path(copy))
        # 同步装载当前 Excel，不等下次启动（卸载侧已有同样的同步兜底）
        trust_ok = is_auto_trusted(copy)
        if trust_ok:
            for hwnd in find_excel_doc_windows():
                app = get_excel_app(hwnd)
                if app is not None:
                    provision(app, hwnd, copy)
        else:
            self.append_log("警告：管理目录未加入 Excel 受信任位置，%s 已登记但不会自动装载；"
                            "请在选项区勾选受信任位置" % os.path.basename(copy))
        self.kick.set()
        self.refresh()
        self.append_log("已长期启动 %s（%s，%s，登记 %s）%s"
                        % (os.path.basename(path), how, trust_note, msg,
                           "，当前 Excel 已装载，下次启动自动加载" if trust_ok else ""))

    def on_unregister(self, path):
        name = os.path.basename(path)
        if not messagebox.askyesno(
            "确认卸载",
            "卸载 %s 将撤销 Excel 加载项登记、删除加载项目录中的副本。\n"
            "若 Excel 正在运行，会先提示你保存并关闭；等待超时后才询问是否强制结束。\n"
            "确定继续？" % name,
            parent=self.root,
        ):
            return
        if not self.ensure_excel_closed("卸载"):
            return
        ok, msg = unregister_addin(path)
        if not ok:
            self.append_log("卸载 %s 失败：%s" % (name, msg))
            return
        cleared = clear_addin_load_time(path)
        with self.lock:
            self.extra = [p for p in self.extra if norm_path(p) != norm_path(path)]
            self.deselected.discard(norm_path(path))
        # 同步关一次，不依赖异步 worker（Excel 忙时 worker 可能吞异常）
        for hwnd in find_excel_doc_windows():
            app = get_excel_app(hwnd)
            if app is not None:
                unload_from_excel(app, hwnd, path)
        removed = remove_from_managed(path)
        self.unloadq.put(path)
        self.kick.set()
        self.refresh()
        self.append_log("已卸载 %s（登记已撤销%s%s）"
                        % (name,
                           "，加载项列表记录已清除" if cleared else "",
                           "，目录副本已删除" if removed else ""))
        self.sync_autostart_with_targets()
        # 延迟复查：防止登记/加载状态被写回
        self.root.after(3000, lambda: self._verify_unloaded(path))

    def _verify_unloaded(self, path):
        """卸载 3 秒后复查：登记或 Installed 若被写回，再次清除并告警。"""
        name = os.path.basename(path)
        fixed = []
        if registered_at(path):
            unregister_addin(path)
            fixed.append("登记")
        if clear_addin_load_time(path):
            fixed.append("列表记录")
        for hwnd in find_excel_doc_windows():
            app = get_excel_app(hwnd)
            if app is None:
                continue
            target = norm_path(path)
            wb = workbook_by_name(app, os.path.basename(path))
            if wb is not None:
                try:
                    if norm_path(wb.FullName) == target:
                        wb.Close(SaveChanges=False)
                        fixed.append("临时装载")
                except Exception:
                    pass
            for i in range(1, app.AddIns.Count + 1):
                try:
                    ad = app.AddIns.Item(i)
                    if ad.FullName and norm_path(ad.FullName) == target and ad.Installed:
                        ad.Installed = False
                        fixed.append("加载状态")
                except Exception:
                    pass
        if fixed:
            self.append_log("复查：%s 的%s被重新写回，已再次清除。"
                            % (name, "和".join(fixed)))

    # ---- 托盘 ----
    def setup_tray(self):
        hwnd = wintypes.HWND(int(self.root.frame(), 16))
        self.tray_hwnd = hwnd
        self.old_wndproc = _user32.GetWindowLongPtrW(hwnd, GWLP_WNDPROC)

        app = self
        def wndproc(h, msg, wp, lp):
            # 回调运行在 Windows 消息循环线程，绝不直接调用 tkinter（会破坏 GIL
            # 导致 PyEval_RestoreThread 致命崩溃）；只入队，由 poll 在主线程消费。
            if msg == TRAY_MSG:
                if app.menu_open:
                    return 0  # 菜单模态期间忽略新托盘点击，避免积压重入
                event = int(lp) & 0xFFFF
                if event in (WM_LBUTTONUP, WM_LBUTTONDBLCLK):
                    app.outq.put(("tray", "toggle"))
                elif event == WM_RBUTTONUP:
                    app.outq.put(("tray", "menu"))
                return 0
            if msg == TASKBAR_CREATED:
                app.outq.put(("tray", "rebuild"))
                return 0
            if msg == WM_SHOW_APP:
                app.outq.put(("tray", "show"))
                return 0
            return _user32.CallWindowProcW(app.old_wndproc, h, msg, wp, lp)

        self.new_wndproc = WNDPROC(wndproc)
        proc_ptr = ctypes.cast(self.new_wndproc, ctypes.c_void_p).value
        _user32.SetWindowLongPtrW(hwnd, GWLP_WNDPROC, proc_ptr)

        self.add_tray_icon()

    def add_tray_icon(self):
        """注册（或 explorer 重启后重建）托盘图标，可重复调用。"""
        nid = NOTIFYICONDATAW()
        nid.cbSize = ctypes.sizeof(NOTIFYICONDATAW)
        nid.hWnd = self.tray_hwnd
        nid.uID = 1
        nid.uFlags = NIF_MESSAGE | NIF_ICON | NIF_TIP
        nid.uCallbackMessage = TRAY_MSG
        nid.hIcon = _user32.LoadIconW(None, IDI_APPLICATION)
        nid.szTip = "%s v%s" % (TOOL_NAME, VERSION)
        self.tray_nid = nid
        if not _shell32.Shell_NotifyIconW(NIM_ADD, ctypes.byref(nid)):
            self.append_log("警告：系统托盘图标创建失败。")

    def show_tray_menu(self):
        self.menu_open = True
        menu = _user32.CreatePopupMenu()
        flags = MF_STRING | (MF_CHECKED if is_autostart() else 0)
        _user32.AppendMenuW(menu, MF_STRING, ID_SHOW, "显示窗口")
        _user32.AppendMenuW(menu, flags, ID_AUTOSTART, "开机启动")
        _user32.AppendMenuW(menu, MF_SEPARATOR, 0, None)
        _user32.AppendMenuW(menu, MF_STRING, ID_EXIT, "退出")
        pt = wintypes.POINT()
        _user32.GetCursorPos(ctypes.byref(pt))
        _user32.SetForegroundWindow(self.tray_hwnd)
        cmd = _user32.TrackPopupMenu(menu, TPM_RIGHTBUTTON | TPM_RETURNCMD,
                                     pt.x, pt.y, 0, self.tray_hwnd, None)
        _user32.DestroyMenu(menu)
        _user32.PostMessageW(self.tray_hwnd, 0x0, 0, 0)  # WM_NULL，确保菜单关闭
        self.menu_open = False
        if cmd == ID_SHOW:
            self.show_window()
        elif cmd == ID_AUTOSTART:
            self.toggle_autostart()
        elif cmd == ID_EXIT:
            self.exit_app()

    def sync_autostart_with_targets(self):
        """目标为空时自动取消开机启动：取消勾选或文件删除都会走到这里。"""
        if not is_autostart():
            return
        if time.monotonic() - self.started_at < AUTOSTART_GRACE:
            return
        if self.get_targets():
            return
        self._apply_autostart(False)
        self.append_log("没有可装载的加载项，已自动取消开机启动")

    def _apply_autostart(self, want):
        """统一写入开机启动状态，并回读同步主界面勾选框与托盘菜单。"""
        try:
            set_autostart(want)
            self.append_log("开机启动已%s。" % ("启用" if want else "取消"))
        except Exception as exc:
            self.append_log("设置开机启动失败：%s" % exc)
        self.var_autostart.set(is_autostart())
        self.refresh_state_text()

    def on_autostart_ui(self):
        self._apply_autostart(self.var_autostart.get())

    def toggle_autostart(self):
        self._apply_autostart(not is_autostart())

    def show_window(self):
        if self.stop.is_set():
            return
        self.root.deiconify()
        self.root.lift()
        try:
            self.root.attributes("-topmost", True)
            self.root.after(300, lambda: self.root.attributes("-topmost", False))
        except Exception:
            pass

    def on_close_request(self):
        """点窗口 X：提示后台运行并隐藏；后续不再重复提示。"""
        if not self._close_tip_shown:
            self._close_tip_shown = True
            messagebox.showinfo(
                "后台运行",
                "装载器将在后台继续运行。\n"
                "需要退出时，请右键点击屏幕右下角的系统托盘图标，选择「退出」。",
                parent=self.root,
            )
        self.hide_to_tray()

    def hide_to_tray(self):
        self.root.withdraw()

    def toggle_window(self):
        if self.root.state() == "withdrawn":
            self.show_window()
        else:
            self.hide_to_tray()

    def exit_app(self):
        self.stop.set()
        self.kick.set()
        self.lbl_state["text"] = "状态：已退出"
        if self.tray_nid is not None:
            _shell32.Shell_NotifyIconW(NIM_DELETE, ctypes.byref(self.tray_nid))
            self.tray_nid = None
        if self.old_wndproc is not None:
            _user32.SetWindowLongPtrW(self.tray_hwnd, GWLP_WNDPROC, self.old_wndproc)
            self.old_wndproc = None
        self.root.destroy()

    # ---- 日志/轮询 ----
    def append_log(self, text):
        write_run_log(text)
        self.log_lines.append(text)
        self.log.configure(state="normal")
        self.log.delete("1.0", "end")
        self.log.insert("end", "\n".join(self.log_lines))
        self.log.see("end")
        self.log.configure(state="disabled")

    def open_release_page(self, url):
        url = url or RELEASE_PAGE_URL
        self.append_log("打开发布页 %s" % url)
        webbrowser.open(url)

    def poll(self):
        try:
            while True:
                item = self.outq.get_nowait()
                if item[0] == "log":
                    self.append_log(item[1])
                elif item[0] == "status":
                    self.lbl_stat["text"] = (
                        "扫描 %d 次 ｜ 勾选 %d ｜ 已装载 %d"
                        % (item[1], item[3], item[4])
                    )
                elif item[0] == "tray":
                    action = item[1]
                    if action == "toggle":
                        self.toggle_window()
                    elif action == "menu":
                        self.show_tray_menu()
                    elif action == "show":
                        self.show_window()
                    elif action == "rebuild":
                        self.add_tray_icon()
        except queue.Empty:
            pass
        now = time.monotonic()
        if now - self.last_discover >= DISCOVER_INTERVAL:
            self.last_discover = now
            self.auto_discover()
        if now - self.last_autostart_check >= AUTOSTART_CHECK_SEC:
            self.last_autostart_check = now
            try:
                note = autostart_health_fix()
                if note:
                    self.append_log(note)
                    self.var_autostart.set(is_autostart())
                    self.refresh_state_text()
            except Exception as exc:
                self.append_log("开机启动自检失败：%s" % exc)
            self.sync_autostart_with_targets()
        self.root.after(120, self.poll)


def parse_fixed(argv):
    if "--xlam" in argv:
        idx = argv.index("--xlam")
        if idx + 1 < len(argv):
            return [argv[idx + 1]]
    return None


def main():
    global STARTUP_NOTICE
    argv = sys.argv[1:]
    fixed = parse_fixed(argv)
    if "--selfcheck" in argv:
        sys.exit(selfcheck(fixed[0] if fixed else None))

    # 命令行更新入口：--update-manifest <url> [--update-dry-run]
    if "--update-manifest" in argv:
        idx = argv.index("--update-manifest")
        target = argv[idx + 1] if idx + 1 < len(argv) else ""
        if not target:
            print("FAIL: --update-manifest 缺少地址")
            sys.exit(1)
        dry = "--update-dry-run" in argv
        print(run_one_click_update(target, dry_run=dry))
        sys.exit(0)

    notice_takeover = ""
    try:
        notice_takeover = take_over_older_instances() or ""
    except Exception as exc:
        notice_takeover = "版本仲裁失败：%s" % exc

    _mutex, first = acquire_single_instance()
    if not first:
        sys.exit(0)

    notices = [n for n in (notice_takeover,) if n]
    try:
        boot = migrate_autostart() or autostart_health_fix() or ""
        if boot:
            notices.append(boot)
    except Exception as exc:
        notices.append("开机启动自检失败：%s" % exc)
    STARTUP_NOTICE = "；".join(notices)

    root = tk.Tk()
    LoaderApp(root, fixed, background=("--background" in argv))
    root.mainloop()


if __name__ == "__main__":
    main()
