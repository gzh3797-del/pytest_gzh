"""Acuview 2 可执行文件自动发现 —— 让项目在不同同事 PC 上免配置即可运行。

背景: Acuview 2 是 *按用户* 安装在 `%USERPROFILE%\\Acuview2\\Acuview 2.exe`(不是
Program Files)。换 PC 仅用户名不同，所以靠 expanduser 就能覆盖绝大多数机器；
再加注册表 / Program Files / 运行中进程兜底。

两条入口:
  resolve_acuview_exe(configured) -> str|None   解析"用来启动"的 exe 路径(按序尝试)
  running_acuview_exe()           -> str|None   读取"已在运行"的 Acuview 2 进程全路径

均不引入第三方依赖(标准库 winreg + ctypes)。
"""
from __future__ import annotations

import ctypes
import glob
import os
from ctypes import wintypes

EXE_NAME = "Acuview 2.exe"
_PER_USER = os.path.join("~", "Acuview2", EXE_NAME)


def _ok(path: str | None) -> str | None:
    if path:
        p = os.path.expanduser(os.path.expandvars(path))
        if os.path.isfile(p):
            return p
    return None


def _from_program_files() -> str | None:
    roots = [os.environ.get("ProgramFiles", r"C:\Program Files"),
             os.environ.get("ProgramFiles(x86)", r"C:\Program Files (x86)")]
    for root in roots:
        if not root:
            continue
        for hit in glob.glob(os.path.join(root, "Acuview*", EXE_NAME)):
            if os.path.isfile(hit):
                return hit
    return None


def _from_registry() -> str | None:
    """扫卸载项, DisplayName 含 'acuview' 的 InstallLocation 下找 exe。"""
    try:
        import winreg
    except ImportError:
        return None
    roots = [(winreg.HKEY_CURRENT_USER, r"Software\Microsoft\Windows\CurrentVersion\Uninstall"),
             (winreg.HKEY_LOCAL_MACHINE, r"Software\Microsoft\Windows\CurrentVersion\Uninstall"),
             (winreg.HKEY_LOCAL_MACHINE,
              r"Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall")]
    for hive, sub in roots:
        try:
            base = winreg.OpenKey(hive, sub)
        except OSError:
            continue
        try:
            for i in range(winreg.QueryInfoKey(base)[0]):
                try:
                    k = winreg.OpenKey(base, winreg.EnumKey(base, i))
                    name = str(winreg.QueryValueEx(k, "DisplayName")[0])
                    if "acuview" not in name.lower():
                        continue
                    loc = winreg.QueryValueEx(k, "InstallLocation")[0]
                    hit = _ok(os.path.join(loc, EXE_NAME))
                    if hit:
                        return hit
                except OSError:
                    continue
        finally:
            base.Close()
    return None


def resolve_acuview_exe(configured: str | None) -> str | None:
    """返回可用于启动的 Acuview 2 exe 路径; 都找不到返回 None。

    顺序: config 配置 -> %USERPROFILE%\\Acuview2 -> Program Files -> 注册表 -> 运行中进程。
    """
    return (_ok(configured)
            or _ok(_PER_USER)
            or _from_program_files()
            or _from_registry()
            or running_acuview_exe())


# --------------------------------------------------------------------------
# 读取已运行的 Acuview 2 进程全路径(复用 gui_driver 同款进程枚举思路)
# --------------------------------------------------------------------------
def running_acuview_exe() -> str | None:
    TH32CS_SNAPPROCESS = 0x00000002
    PROCESS_QUERY_LIMITED_INFORMATION = 0x1000
    INVALID = ctypes.c_void_p(-1).value
    k32 = ctypes.windll.kernel32

    class PROCESSENTRY32W(ctypes.Structure):
        _fields_ = [
            ("dwSize", ctypes.c_ulong), ("cntUsage", ctypes.c_ulong),
            ("th32ProcessID", ctypes.c_ulong), ("th32DefaultHeapID", ctypes.c_void_p),
            ("th32ModuleID", ctypes.c_ulong), ("cntThreads", ctypes.c_ulong),
            ("th32ParentProcessID", ctypes.c_ulong), ("pcPriClassBase", ctypes.c_long),
            ("dwFlags", ctypes.c_ulong), ("szExeFile", ctypes.c_wchar * 260),
        ]

    snap = k32.CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, 0)
    if snap == INVALID:
        return None
    try:
        entry = PROCESSENTRY32W()
        entry.dwSize = ctypes.sizeof(PROCESSENTRY32W)
        if not k32.Process32FirstW(snap, ctypes.byref(entry)):
            return None
        while True:
            if entry.szExeFile.lower() == EXE_NAME.lower():
                pid = entry.th32ProcessID
                h = k32.OpenProcess(PROCESS_QUERY_LIMITED_INFORMATION, False, pid)
                if h:
                    try:
                        buf = ctypes.create_unicode_buffer(32768)
                        size = wintypes.DWORD(len(buf))
                        if k32.QueryFullProcessImageNameW(h, 0, buf, ctypes.byref(size)):
                            return buf.value
                    finally:
                        k32.CloseHandle(h)
            if not k32.Process32NextW(snap, ctypes.byref(entry)):
                break
    finally:
        k32.CloseHandle(snap)
    return None


if __name__ == "__main__":
    from .config import get_config
    cfg = get_config()
    configured = (cfg.app.get("exe_path") or "").strip()
    print("configured :", configured or "(空)")
    print("running    :", running_acuview_exe() or "(未运行)")
    print("resolved   :", resolve_acuview_exe(configured) or "(未找到 — 请在 config.app.exe_path 填绝对路径)")
