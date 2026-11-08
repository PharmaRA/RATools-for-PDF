"""Bounded, recoverable cleanup for files owned by a processing task."""

import os
from pathlib import Path
import shutil
import subprocess


def check_unlinked_path(path):
    """Reject links/reparse points before writing, replacing or recycling a path."""
    candidate = Path(path).absolute()
    for entry in (candidate, *candidate.parents):
        if not os.path.lexists(entry):
            continue
        info = entry.lstat()
        if entry.is_symlink() or getattr(info, 'st_file_attributes', 0) & 0x400:
            raise ValueError(f'路径包含链接或重解析点，需人工确认: {entry}')
        if entry.is_file() and info.st_nlink > 1:
            raise ValueError(f'文件包含硬链接，需人工确认: {entry}')
    return candidate


def _recycle_windows(path):
    import pythoncom
    from win32com.shell import shell

    pythoncom.CoInitialize()
    try:
        operation = pythoncom.CoCreateInstance(
            shell.CLSID_FileOperation, None, pythoncom.CLSCTX_ALL, shell.IID_IFileOperation,
        )
        # SILENT | NOCONFIRMATION | NOERRORUI | RECYCLEONDELETE | EARLYFAILURE.
        # RECYCLEONDELETE (Windows 8+) explicitly requests recycling, never a delete fallback.
        operation.SetOperationFlags(0x4 | 0x10 | 0x400 | 0x80000 | 0x100000)
        item = shell.SHCreateItemFromParsingName(str(path), None, shell.IID_IShellItem)
        operation.DeleteItem(item, None)
        operation.PerformOperations()
        if operation.GetAnyOperationsAborted() or path.exists():
            raise OSError(f'未能移入回收站: {path}')
    finally:
        pythoncom.CoUninitialize()


def recycle_owned_path(path, owned_root):
    """Recycle only a verified task-owned path; failures leave the files in place."""
    candidate = check_unlinked_path(path)
    root = check_unlinked_path(owned_root).resolve()
    if not candidate.resolve().is_relative_to(root):
        raise ValueError('回收目标超出本次任务目录')
    if not candidate.exists():
        return
    if candidate.is_dir():
        for parent, dirs, files in os.walk(candidate, followlinks=False):
            for name in dirs + files:
                check_unlinked_path(Path(parent) / name)
    if os.name == 'nt':
        _recycle_windows(candidate)
    else:
        # GIO provides a desktop trash operation. Never use a permanent-delete fallback.
        gio = shutil.which('gio')
        if not gio:
            raise OSError(f'系统回收站接口不可用，临时文件已保留: {candidate}')
        subprocess.run([gio, 'trash', str(candidate)], check=True, capture_output=True)
        if candidate.exists():
            raise OSError(f'未能移入回收站: {candidate}')
