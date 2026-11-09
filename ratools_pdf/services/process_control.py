"""Owned subprocess lifetime and conservative admission for PDF processing."""

import ctypes
import math
import os
import shutil
import signal
import subprocess
import threading
import time

from ctypes import wintypes


class _SecurityAttributes(ctypes.Structure):
    _fields_ = [('length', wintypes.DWORD), ('descriptor', ctypes.c_void_p), ('inherit', wintypes.BOOL)]


class _StartupInfo(ctypes.Structure):
    _fields_ = [('cb', wintypes.DWORD), ('reserved', wintypes.LPWSTR),
                ('desktop', wintypes.LPWSTR), ('title', wintypes.LPWSTR),
                ('x', wintypes.DWORD), ('y', wintypes.DWORD), ('x_size', wintypes.DWORD),
                ('y_size', wintypes.DWORD), ('x_chars', wintypes.DWORD), ('y_chars', wintypes.DWORD),
                ('fill', wintypes.DWORD), ('flags', wintypes.DWORD),
                ('show', wintypes.WORD), ('reserved_size', wintypes.WORD),
                ('reserved_bytes', ctypes.POINTER(wintypes.BYTE)),
                ('stdin', wintypes.HANDLE), ('stdout', wintypes.HANDLE), ('stderr', wintypes.HANDLE)]


class _StartupInfoEx(ctypes.Structure):
    _fields_ = [('startup', _StartupInfo), ('attributes', ctypes.c_void_p)]


class _ProcessInformation(ctypes.Structure):
    _fields_ = [('hProcess', wintypes.HANDLE), ('hThread', wintypes.HANDLE),
                ('dwProcessId', wintypes.DWORD), ('dwThreadId', wintypes.DWORD)]


def _create_windows_process(kernel, command, startup):
    """The native creation boundary: Job association is part of this call."""
    info = _ProcessInformation()
    command_line = ctypes.create_unicode_buffer(subprocess.list2cmdline(command))
    flags = 0x80000 | 0x08000000 | 0x4  # EXTENDED_STARTUPINFO_PRESENT, NO_WINDOW, SUSPENDED.
    if not kernel.CreateProcessW(command[0], command_line, None, None, True, flags,
                                 None, None, ctypes.byref(startup), ctypes.byref(info)):
        code = ctypes.get_last_error()
        if code in (2, 3):
            raise FileNotFoundError('未找到 qpdf 工具，请检查安装或 QPDF_PATH')
        raise OSError(f'无法在受控进程树中启动 qpdf（Windows 错误 {code}）')
    return info


class _WindowsProcess:
    """Native handles and anonymous-pipe collection for one atomically owned child."""

    def __init__(self, job, info, stdout, stderr):
        self.job = job
        self.kernel = job.kernel
        self._handle = info.hProcess
        self._thread_handle = info.hThread
        self.pid = info.dwProcessId
        self.returncode = None
        self._pipes = [stdout, stderr]
        self._chunks = [[], []]
        self._readers = []
        self._read_errors = []

    def poll(self):
        state = self.kernel.WaitForSingleObject(self._handle, 0)
        if state == 0 and self.returncode is None:
            code = wintypes.DWORD()
            if not self.kernel.GetExitCodeProcess(self._handle, ctypes.byref(code)):
                raise OSError('无法读取受控 qpdf 退出状态')
            self.returncode = int(code.value)
        elif state == 0xffffffff:
            raise OSError('无法等待受控 qpdf 进程')
        return self.returncode

    def kill(self):
        if self.poll() is None and not self.kernel.TerminateProcess(self._handle, 1):
            # Job close may already have initiated termination: Windows can
            # reject a second termination while the process is not yet signaled.
            if self.kernel.WaitForSingleObject(self._handle, 5000) != 0:
                raise OSError('无法终止受控 qpdf 进程')

    def _read_pipe(self, index):
        buffer = ctypes.create_string_buffer(65536)
        count = wintypes.DWORD()
        while self.kernel.ReadFile(self._pipes[index], buffer, len(buffer), ctypes.byref(count), None):
            if not count.value:
                return
            self._chunks[index].append(buffer.raw[:count.value])
        if ctypes.get_last_error() not in (0, 109):  # ERROR_BROKEN_PIPE is normal EOF.
            self._read_errors.append(True)

    def communicate(self, timeout=None):
        if not self._readers:
            for index in range(2):
                reader = threading.Thread(target=self._read_pipe, args=(index,), daemon=True)
                self._readers.append(reader)
                reader.start()
        deadline = time.monotonic() + timeout if timeout is not None else None
        while self.poll() is None:
            remaining = deadline - time.monotonic() if deadline is not None else None
            if remaining is not None and remaining <= 0:
                raise subprocess.TimeoutExpired(['qpdf'], timeout)
            delay = min(100, max(1, math.ceil(remaining * 1000))) if remaining is not None else 100
            self.kernel.WaitForSingleObject(self._handle, delay)
        # Descendants may hold inherited output handles after the root exits.
        # Closing the Job releases them before joining the pipe reader threads.
        self.job.close()
        for reader in self._readers:
            reader.join(5)
            if reader.is_alive():
                raise OSError('受控 qpdf 输出管道未能收尾')
        if self._read_errors:
            raise OSError('无法读取受控 qpdf 输出')
        return tuple(b''.join(chunks).decode('utf-8', errors='replace') for chunks in self._chunks)

    def close(self):
        for handle in [self._thread_handle, self._handle, *self._pipes]:
            if handle:
                self.kernel.CloseHandle(handle)
        self._thread_handle = self._handle = None
        self._pipes = []


def normalize_file_timeout(value):
    try:
        seconds = float(value)
    except (TypeError, ValueError):
        raise ValueError('每文件超时必须为 0 至 86400 秒（0 为关闭）') from None
    if not math.isfinite(seconds) or not 0 <= seconds <= 86400:
        raise ValueError('每文件超时必须为 0 至 86400 秒（0 为关闭）')
    return seconds


def available_memory_bytes():
    """Available, not installed, memory. Unavailable readings stay unknown."""
    if os.name != 'nt':
        return None
    class MemoryStatus(ctypes.Structure):
        _fields_ = [('length', ctypes.c_ulong), ('load', ctypes.c_ulong),
                    ('total', ctypes.c_ulonglong), ('available', ctypes.c_ulonglong),
                    ('total_page', ctypes.c_ulonglong), ('available_page', ctypes.c_ulonglong),
                    ('total_virtual', ctypes.c_ulonglong), ('available_virtual', ctypes.c_ulonglong),
                    ('extended', ctypes.c_ulonglong)]
    status = MemoryStatus()
    status.length = ctypes.sizeof(status)
    kernel = ctypes.WinDLL('kernel32', use_last_error=True)
    kernel.GlobalMemoryStatusEx.argtypes = [ctypes.POINTER(MemoryStatus)]
    kernel.GlobalMemoryStatusEx.restype = ctypes.c_int
    return int(status.available) if kernel.GlobalMemoryStatusEx(ctypes.byref(status)) else None


def concurrency_limit(requested, input_sizes, available):
    """Reserve half available RAM; estimate at least 256 MiB or 8x input/task."""
    requested = max(1, int(requested))
    if not available or any(size is None for size in input_sizes):
        return 1
    estimate = max(256 * 1024**2, max(input_sizes, default=0) * 8)
    return max(1, min(requested, int(available * 0.5) // estimate))


def ensure_output_capacity(directory, input_size):
    try:
        free = shutil.disk_usage(directory).free
    except OSError:
        raise OSError('无法访问输出位置，请检查磁盘或网络连接后重试') from None
    required = max(64 * 1024**2, input_size * 3)
    if free < required:
        raise OSError('输出磁盘可用空间不足，请释放空间或选择其他输出位置后重试')


class WindowsJob:
    """Kill-on-close ownership is supplied atomically to CreateProcessW."""

    def __init__(self):
        class BasicLimits(ctypes.Structure):
            _fields_ = [('process_time', ctypes.c_longlong), ('job_time', ctypes.c_longlong),
                        ('flags', wintypes.DWORD), ('min_ws', ctypes.c_size_t),
                        ('max_ws', ctypes.c_size_t), ('active', wintypes.DWORD),
                        ('affinity', ctypes.c_size_t), ('priority', wintypes.DWORD),
                        ('scheduling', wintypes.DWORD)]

        class IoCounters(ctypes.Structure):
            _fields_ = [(name, ctypes.c_ulonglong) for name in
                        ('read_ops', 'write_ops', 'other_ops', 'read_bytes', 'write_bytes', 'other_bytes')]

        class ExtendedLimits(ctypes.Structure):
            _fields_ = [('basic', BasicLimits), ('io', IoCounters),
                        ('process_memory', ctypes.c_size_t), ('job_memory', ctypes.c_size_t),
                        ('peak_process', ctypes.c_size_t), ('peak_job', ctypes.c_size_t)]

        self.kernel = ctypes.WinDLL('kernel32', use_last_error=True)
        self.kernel.CreateJobObjectW.argtypes = [ctypes.c_void_p, wintypes.LPCWSTR]
        self.kernel.CreateJobObjectW.restype = wintypes.HANDLE
        self.kernel.SetInformationJobObject.argtypes = [wintypes.HANDLE, ctypes.c_int,
                                                       ctypes.c_void_p, wintypes.DWORD]
        self.kernel.SetInformationJobObject.restype = wintypes.BOOL
        self.kernel.AssignProcessToJobObject.argtypes = [wintypes.HANDLE, wintypes.HANDLE]
        self.kernel.AssignProcessToJobObject.restype = wintypes.BOOL
        declarations = {
            'IsProcessInJob': ([wintypes.HANDLE, wintypes.HANDLE, ctypes.POINTER(wintypes.BOOL)], wintypes.BOOL),
            'CreatePipe': ([ctypes.POINTER(wintypes.HANDLE), ctypes.POINTER(wintypes.HANDLE),
                            ctypes.POINTER(_SecurityAttributes), wintypes.DWORD], wintypes.BOOL),
            'SetHandleInformation': ([wintypes.HANDLE, wintypes.DWORD, wintypes.DWORD], wintypes.BOOL),
            'CreateFileW': ([wintypes.LPCWSTR, wintypes.DWORD, wintypes.DWORD,
                             ctypes.POINTER(_SecurityAttributes), wintypes.DWORD,
                             wintypes.DWORD, wintypes.HANDLE], wintypes.HANDLE),
            'InitializeProcThreadAttributeList': ([ctypes.c_void_p, wintypes.DWORD, wintypes.DWORD,
                                                   ctypes.POINTER(ctypes.c_size_t)], wintypes.BOOL),
            'UpdateProcThreadAttribute': ([ctypes.c_void_p, wintypes.DWORD, ctypes.c_size_t,
                                           ctypes.c_void_p, ctypes.c_size_t, ctypes.c_void_p,
                                           ctypes.c_void_p], wintypes.BOOL),
            'DeleteProcThreadAttributeList': ([ctypes.c_void_p], None),
            'CreateProcessW': ([wintypes.LPCWSTR, wintypes.LPWSTR, ctypes.c_void_p, ctypes.c_void_p,
                                wintypes.BOOL, wintypes.DWORD, ctypes.c_void_p, wintypes.LPCWSTR,
                                ctypes.POINTER(_StartupInfoEx), ctypes.POINTER(_ProcessInformation)], wintypes.BOOL),
            'ResumeThread': ([wintypes.HANDLE], wintypes.DWORD),
            'WaitForSingleObject': ([wintypes.HANDLE, wintypes.DWORD], wintypes.DWORD),
            'GetExitCodeProcess': ([wintypes.HANDLE, ctypes.POINTER(wintypes.DWORD)], wintypes.BOOL),
            'TerminateProcess': ([wintypes.HANDLE, wintypes.UINT], wintypes.BOOL),
            'ReadFile': ([wintypes.HANDLE, ctypes.c_void_p, wintypes.DWORD,
                          ctypes.POINTER(wintypes.DWORD), ctypes.c_void_p], wintypes.BOOL),
        }
        for name, (arguments, result) in declarations.items():
            api = getattr(self.kernel, name)
            api.argtypes, api.restype = arguments, result
        self.kernel.CloseHandle.argtypes = [wintypes.HANDLE]
        self.kernel.CloseHandle.restype = wintypes.BOOL
        self.handle = self.kernel.CreateJobObjectW(None, None)
        if not self.handle:
            raise OSError('无法建立受控进程树，已停止启动 qpdf')
        limits = ExtendedLimits()
        limits.basic.flags = 0x2000  # JOB_OBJECT_LIMIT_KILL_ON_JOB_CLOSE; no breakaway.
        if not self.kernel.SetInformationJobObject(self.handle, 9, ctypes.byref(limits), ctypes.sizeof(limits)):
            self.close()
            raise OSError('无法配置受控进程树，已停止启动 qpdf')

    def assign(self, process):
        # Confirm the creation-time postcondition; never assign after launch.
        owned = wintypes.BOOL()
        if not self.kernel.IsProcessInJob(process._handle, self.handle, ctypes.byref(owned)) or not owned.value:
            raise OSError('qpdf 创建时未进入受控进程树，已安全停止')

    def resume(self, process):
        if self.kernel.ResumeThread(process._thread_handle) == 0xffffffff:
            raise OSError('无法恢复受控 qpdf 进程，已安全停止')
        self.kernel.CloseHandle(process._thread_handle)
        process._thread_handle = None

    def create_process(self, command):
        """Windows 10+ JOB_LIST and HANDLE_LIST, both set before child creation."""
        handles = []
        attributes = None
        info = None
        transferred = False
        security = _SecurityAttributes(ctypes.sizeof(_SecurityAttributes), None, True)
        try:
            def pipe():
                read, write = wintypes.HANDLE(), wintypes.HANDLE()
                if not self.kernel.CreatePipe(ctypes.byref(read), ctypes.byref(write), ctypes.byref(security), 0):
                    raise OSError('无法建立受控 qpdf 输出管道')
                handles.extend([read.value, write.value])
                if not self.kernel.SetHandleInformation(read, 1, 0):
                    raise OSError('无法限制 qpdf 管道继承')
                return read.value, write.value
            stdout_read, stdout_write = pipe()
            stderr_read, stderr_write = pipe()
            stdin = self.kernel.CreateFileW('NUL', 0x80000000, 3, ctypes.byref(security), 3, 0, None)
            if stdin == ctypes.c_void_p(-1).value:
                raise OSError('无法建立受控 qpdf 输入句柄')
            handles.append(stdin)
            size = ctypes.c_size_t()
            self.kernel.InitializeProcThreadAttributeList(None, 2, 0, ctypes.byref(size))
            if not size.value:
                raise OSError('系统不支持创建时进程树控制，未启动 qpdf（需要 Windows 10 或更新版本）')
            storage = ctypes.create_string_buffer(size.value)
            pointer = ctypes.cast(storage, ctypes.c_void_p)
            if not self.kernel.InitializeProcThreadAttributeList(pointer, 2, 0, ctypes.byref(size)):
                raise OSError('无法初始化创建时进程树控制，未启动 qpdf')
            attributes = pointer
            inherited = (wintypes.HANDLE * 3)(stdin, stdout_write, stderr_write)
            jobs = (wintypes.HANDLE * 1)(self.handle)
            for attribute, value in ((0x20002, inherited), (0x2000d, jobs)):
                if not self.kernel.UpdateProcThreadAttribute(attributes, 0, attribute, ctypes.byref(value),
                                                             ctypes.sizeof(value), None, None):
                    raise OSError('系统或宿主不支持创建时 Job 关联，未启动 qpdf')
            startup = _StartupInfoEx()
            startup.startup.cb = ctypes.sizeof(startup)
            startup.startup.flags = 0x100 | 1  # USESTDHANDLES | USESHOWWINDOW; SW_HIDE is zero.
            startup.startup.stdin, startup.startup.stdout, startup.startup.stderr = stdin, stdout_write, stderr_write
            startup.attributes = attributes
            info = _create_windows_process(self.kernel, command, startup)
            process = _WindowsProcess(self, info, stdout_read, stderr_read)
            transferred = True
            return process
        finally:
            if attributes:
                self.kernel.DeleteProcThreadAttributeList(attributes)
            for handle in handles:
                if not transferred or handle not in (stdout_read, stderr_read):
                    self.kernel.CloseHandle(handle)
            if info is not None and not transferred:
                self.kernel.TerminateProcess(info.hProcess, 1)
                self.kernel.WaitForSingleObject(info.hProcess, 5000)
                self.kernel.CloseHandle(info.hThread)
                self.kernel.CloseHandle(info.hProcess)

    def close(self):
        if self.handle:
            self.kernel.CloseHandle(self.handle)
            self.handle = None


def run_owned_command(command, timeout=None, startupinfo=None):
    """Run one owned tree, returning CompletedProcess without changing qpdf API.

    Windows job handles are owned by the calling Python process, never inherited:
    killing/crashing that worker closes the handle and terminates descendants.
    Source/non-Windows timeout uses a process group; owner-crash qualification
    there is intentionally not claimed.
    """
    job = WindowsJob() if os.name == 'nt' else None
    process = None
    try:
        if job:
            process = job.create_process(command)
            job.assign(process)
            job.resume(process)
        else:
            process = subprocess.Popen(command, startupinfo=startupinfo, start_new_session=True,
                                       stdout=subprocess.PIPE, stderr=subprocess.PIPE,
                                       text=True, encoding='utf-8', errors='replace')
        try:
            stdout, stderr = process.communicate(timeout=timeout)
        except subprocess.TimeoutExpired:
            # Do not embed the command or captured diagnostics: callers may have
            # supplied a PDF password in argv. Public exception type is preserved.
            raise subprocess.TimeoutExpired(['qpdf'], timeout) from None
        return subprocess.CompletedProcess(command, process.returncode, stdout, stderr)
    finally:
        if job:
            job.close()
        elif process:
            try:
                os.killpg(process.pid, signal.SIGKILL)
            except ProcessLookupError:
                pass
        if process:
            try:
                if process.poll() is None:
                    process.kill()
                # Reap even when ownership verification/resume failed before
                # any child code ran. Closing handles must survive pipe errors.
                process.communicate()
            finally:
                if job:
                    process.close()
