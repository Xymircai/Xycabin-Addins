"""Noninteractive build processes and read-only foreground diagnostics on Windows."""
from pathlib import Path
import ctypes
from ctypes import wintypes
import json
import os
import subprocess
import threading
import time

BACKGROUND_ENV = 'XY_BUILD_BACKGROUND'
TRACE_ENV = 'XY_BUILD_FOCUS_DIR'


def configure_background(enabled):
    if enabled:
        os.environ[BACKGROUND_ENV] = '1'


def process_options(options=None):
    result = dict(options or {})
    if 'input' not in result and result.get('stdin') is None:
        result['stdin'] = subprocess.DEVNULL
    if os.name != 'nt' or os.environ.get(BACKGROUND_ENV) != '1':
        return result
    flags = result.get('creationflags', 0)
    if flags & (subprocess.CREATE_NEW_CONSOLE | subprocess.DETACHED_PROCESS):
        raise ValueError('后台构建不能创建或分离新的控制台')
    result['creationflags'] = flags | subprocess.CREATE_NO_WINDOW
    startup = result.get('startupinfo') or subprocess.STARTUPINFO()
    startup.dwFlags |= subprocess.STARTF_USESHOWWINDOW
    startup.wShowWindow = subprocess.SW_HIDE
    result['startupinfo'] = startup
    if result.get('env') is not None:
        result['env'] = dict(result['env'], **{BACKGROUND_ENV: '1'})
        if os.environ.get(TRACE_ENV):
            result['env'][TRACE_ENV] = os.environ[TRACE_ENV]
    return result


def _append(path, value):
    with Path(path).open('a', encoding='utf-8') as stream:
        stream.write(json.dumps(value, ensure_ascii=False) + '\n')


def popen_process(*args, **kwargs):
    options = process_options(kwargs)
    child = subprocess.Popen(*args, **options)
    directory = os.environ.get(TRACE_ENV)
    if directory:
        command = args[0] if args else kwargs.get('args', [])
        executable = str(command[0]) if isinstance(command, (list, tuple)) and command else str(command)
        _append(Path(directory) / ('launches-%s.jsonl' % os.getpid()),
                dict(event='child_start', time=time.time(), parent_pid=os.getpid(), pid=child.pid,
                     executable=Path(executable).name, gate=os.environ.get('XY_BUILD_GATE'),
                     no_window=bool(options.get('creationflags', 0) & getattr(subprocess, 'CREATE_NO_WINDOW', 0)),
                     stdin='closed' if options.get('stdin') == subprocess.DEVNULL else 'explicit'))
    return child


def run_process(*popenargs, input=None, capture_output=False, timeout=None, check=False, **kwargs):
    # Match subprocess.run, while using our Popen policy at every nesting level.
    if input is not None:
        if kwargs.get('stdin') is not None:
            raise ValueError('stdin and input arguments may not both be used')
        kwargs['stdin'] = subprocess.PIPE
    if capture_output:
        if kwargs.get('stdout') is not None or kwargs.get('stderr') is not None:
            raise ValueError('stdout and stderr arguments may not be used with capture_output')
        kwargs['stdout'] = kwargs['stderr'] = subprocess.PIPE
    with popen_process(*popenargs, **kwargs) as child:
        try:
            stdout, stderr = child.communicate(input, timeout=timeout)
        except subprocess.TimeoutExpired as error:
            child.kill()
            error.stdout, error.stderr = child.communicate()
            raise
        except BaseException:
            child.kill()
            child.wait()
            raise
        code = child.poll()
        if check and code:
            raise subprocess.CalledProcessError(code, child.args, output=stdout, stderr=stderr)
    return subprocess.CompletedProcess(child.args, code, stdout, stderr)


class FocusTrace:
    """Record observations, never activate windows or infer intent from a user click."""
    def __init__(self, directory, stage):
        self.directory, self.stage = Path(directory), stage
        self.path = self.directory / 'focus-events.jsonl'
        self.stop_event = threading.Event()
        self.thread = None
        self.count = 0
        self.errors = []
        self.previous_directory = None
        self.stopped = False

    def start(self):
        if os.name != 'nt' or os.environ.get(BACKGROUND_ENV) != '1' or os.environ.get('XY_REPORT_NESTED') == '1':
            return
        self.directory.mkdir(parents=True, exist_ok=True)
        self.previous_directory = os.environ.get(TRACE_ENV)
        os.environ[TRACE_ENV] = str(self.directory.resolve())
        self.thread = threading.Thread(target=self._observe, name='Build foreground observer', daemon=True)
        self.thread.start()

    def _observe(self):
        try:
            user = ctypes.WinDLL('user32', use_last_error=True)
            kernel = ctypes.WinDLL('kernel32', use_last_error=True)
            kernel.OpenProcess.argtypes = [wintypes.DWORD, wintypes.BOOL, wintypes.DWORD]
            kernel.OpenProcess.restype = wintypes.HANDLE
            kernel.QueryFullProcessImageNameW.argtypes = [wintypes.HANDLE, wintypes.DWORD, wintypes.LPWSTR, ctypes.POINTER(wintypes.DWORD)]
            kernel.CloseHandle.argtypes = [wintypes.HANDLE]
            user.GetForegroundWindow.restype = wintypes.HWND
            user.GetWindowThreadProcessId.argtypes = [wintypes.HWND, ctypes.POINTER(wintypes.DWORD)]
            get_style = getattr(user, 'GetWindowLongPtrW', user.GetWindowLongW)
            get_style.argtypes = [wintypes.HWND, ctypes.c_int]
            get_style.restype = ctypes.c_ssize_t
            class LastInput(ctypes.Structure):
                _fields_ = [('cbSize', wintypes.UINT), ('dwTime', wintypes.DWORD)]
            previous = None
            while not self.stop_event.is_set():
                handle = user.GetForegroundWindow()
                pid = wintypes.DWORD()
                user.GetWindowThreadProcessId(handle, ctypes.byref(pid))
                state = (int(handle or 0), pid.value)
                if state != previous:
                    image = None
                    process = kernel.OpenProcess(0x1000, False, pid.value)
                    if process:
                        try:
                            size = wintypes.DWORD(32768)
                            name = ctypes.create_unicode_buffer(size.value)
                            if kernel.QueryFullProcessImageNameW(process, 0, name, ctypes.byref(size)):
                                image = Path(name.value).name
                        finally:
                            kernel.CloseHandle(process)
                    last = LastInput(ctypes.sizeof(LastInput), 0)
                    known = bool(user.GetLastInputInfo(ctypes.byref(last)))
                    age = ((kernel.GetTickCount() & 0xffffffff) - last.dwTime) & 0xffffffff if known else None
                    _append(self.path, dict(event='initial' if previous is None else 'foreground_change',
                                            time=time.time(), gate=self.stage(), hwnd=state[0], pid=state[1],
                                            executable=image,
                                            topmost=bool(get_style(handle, -20) & 8) if handle else False,
                                            input_age_ms=age, user_input_recent=known and age <= 300))
                    self.count += 1
                    previous = state
                self.stop_event.wait(.05)
        except Exception as error:
            self.errors.append(type(error).__name__ + ': ' + str(error))

    def stop(self):
        self.stop_event.set()
        if self.thread:
            self.thread.join(timeout=2)
            if not self.stopped:
                if self.previous_directory is None:
                    os.environ.pop(TRACE_ENV, None)
                else:
                    os.environ[TRACE_ENV] = self.previous_directory
        self.stopped = True
        return dict(status='error' if self.errors else 'observed' if self.thread else 'not_collected',
                    path=str(self.path.resolve()) if self.thread else None,
                    events=self.count, errors=self.errors, attribution='observation_only')
