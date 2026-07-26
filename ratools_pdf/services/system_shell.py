"""系统 shell 集成：打开文件/目录、资源管理器定位、Acrobat 探测。

集中此前散落在 MainController 中的 4 处平台三分支（Windows/Darwin/Linux）
重复实现。全部函数失败时抛出原始异常，由调用方决定如何提示用户。
"""

import os
import platform
import re
import subprocess


def open_with_default_app(path):
    """用系统默认程序打开文件。"""
    sys_plat = platform.system()
    if sys_plat == "Windows":
        os.startfile(path)
    elif sys_plat == "Darwin":
        subprocess.Popen(["open", path])
    else:
        subprocess.Popen(["xdg-open", path])


def open_directory(dir_path):
    """打开目录本身（不选中内部条目）。"""
    open_with_default_app(dir_path)


def reveal_in_file_manager(path):
    """在系统文件管理器中定位并尽量高亮该路径。"""
    sys_plat = platform.system()
    if sys_plat == "Windows":
        if os.path.isfile(path):
            subprocess.Popen(["explorer", "/select,", os.path.normpath(path)])
        else:
            os.startfile(path)
    elif sys_plat == "Darwin":
        subprocess.Popen(["open", "-R", path])
    else:
        target_dir = os.path.dirname(path) if os.path.isfile(path) else path
        subprocess.Popen(["xdg-open", target_dir])


def extract_executable_from_open_command(command):
    """从注册表 shell\\open\\command 值中提取 exe 路径。"""
    command = str(command or "").strip()
    if not command:
        return ""
    quoted = re.match(r'^\s*"([^"]+?\.exe)"', command, flags=re.IGNORECASE)
    if quoted:
        return quoted.group(1)
    unquoted = re.match(r'^\s*(.+?\.exe)(?:\s+|$)', command, flags=re.IGNORECASE)
    if unquoted:
        return unquoted.group(1).strip()
    first = command.split()[0] if command.split() else ""
    return first if first.lower().endswith(".exe") else ""


def find_acrobat_executable():
    """探测本机 Adobe Acrobat 可执行文件路径，找不到返回空串。"""
    candidates = [
        os.environ.get("RATOOLS_ACROBAT_PATH", ""),
        r"C:\Program Files\Adobe\Acrobat DC\Acrobat\Acrobat.exe",
        r"C:\Program Files\Adobe\Acrobat\Acrobat\Acrobat.exe",
        r"C:\Program Files (x86)\Adobe\Acrobat DC\Acrobat\Acrobat.exe",
    ]

    if platform.system() == "Windows":
        try:
            import winreg

            registry_locations = [
                (winreg.HKEY_CLASSES_ROOT, r"Acrobat.Document.DC\shell\open\command", ""),
                (winreg.HKEY_CLASSES_ROOT, r"Acrobat.Document\shell\open\command", ""),
                (winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\Acrobat.exe", ""),
                (winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\App Paths\Acrobat.exe", ""),
            ]
            for hive, key_path, value_name in registry_locations:
                try:
                    with winreg.OpenKey(hive, key_path) as key:
                        value, _value_type = winreg.QueryValueEx(key, value_name)
                    exe_path = extract_executable_from_open_command(value)
                    if exe_path:
                        candidates.append(exe_path)
                except OSError:
                    continue
        except Exception:
            pass

    for candidate in candidates:
        candidate = str(candidate or "").strip().strip('"')
        if candidate and os.path.isfile(candidate):
            return candidate
    return ""


def open_pdf_in_acrobat_or_default(pdf_path, acrobat_path=None):
    """优先用 Acrobat 打开 PDF；未提供或无效时退回系统默认程序。"""
    if acrobat_path and os.path.isfile(acrobat_path):
        subprocess.Popen([acrobat_path, pdf_path])
        return
    open_with_default_app(pdf_path)
