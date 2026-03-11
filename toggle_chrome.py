"""
TaobaoScraper - Chrome 显示/隐藏切换器
按 F9 切换 Chrome 窗口的显示和隐藏状态
"""

import ctypes
import ctypes.wintypes
import time
import sys

# ── Win32 常量 ──
SW_HIDE = 0
SW_SHOW = 5
SW_RESTORE = 9
VK_F9 = 0x78

user32 = ctypes.windll.user32

# Win32 函数
EnumWindows = user32.EnumWindows
GetWindowTextW = user32.GetWindowTextW
GetWindowTextLengthW = user32.GetWindowTextLengthW
IsWindowVisible = user32.IsWindowVisible
ShowWindow = user32.ShowWindow
SetForegroundWindow = user32.SetForegroundWindow
GetClassNameW = user32.GetClassNameW
GetAsyncKeyState = user32.GetAsyncKeyState
IsWindow = user32.IsWindow


def find_chrome_windows():
    """查找所有 Chrome 主窗口句柄"""
    chrome_hwnds = []

    @ctypes.WINFUNCTYPE(ctypes.c_bool, ctypes.wintypes.HWND, ctypes.wintypes.LPARAM)
    def enum_callback(hwnd, lparam):
        # 获取类名
        class_name = ctypes.create_unicode_buffer(256)
        GetClassNameW(hwnd, class_name, 256)
        if class_name.value != "Chrome_WidgetWin_1":
            return True

        # 获取窗口标题（只要有标题的 Chrome 窗口）
        length = GetWindowTextLengthW(hwnd)
        if length > 0:
            chrome_hwnds.append(hwnd)
        return True

    EnumWindows(enum_callback, 0)
    return chrome_hwnds


class ChromeToggler:
    def __init__(self):
        self.is_hidden = False
        self.hidden_hwnds = []

    def toggle(self):
        """切换 Chrome 显示/隐藏"""
        if self.is_hidden:
            self._show_chrome()
        else:
            self._hide_chrome()

    def _hide_chrome(self):
        """隐藏所有 Chrome 窗口"""
        hwnds = find_chrome_windows()
        if not hwnds:
            print(f"  [{time.strftime('%H:%M:%S')}] F9 按下 - 未找到 Chrome 窗口")
            return

        self.hidden_hwnds = []
        for hwnd in hwnds:
            if IsWindowVisible(hwnd):
                ShowWindow(hwnd, SW_HIDE)
                self.hidden_hwnds.append(hwnd)

        if self.hidden_hwnds:
            self.is_hidden = True
            print(f"  [{time.strftime('%H:%M:%S')}] F9 按下 - Chrome 已隐藏（{len(self.hidden_hwnds)} 个窗口）")
        else:
            # Chrome 窗口存在但都不可见（可能在屏幕外），记录它们
            self.hidden_hwnds = hwnds
            self.is_hidden = True
            print(f"  [{time.strftime('%H:%M:%S')}] F9 按下 - Chrome 已标记隐藏")

    def _show_chrome(self):
        """显示之前隐藏的 Chrome 窗口"""
        shown = 0

        # 先尝试恢复之前记录的窗口
        for hwnd in self.hidden_hwnds:
            try:
                if IsWindow(hwnd):
                    ShowWindow(hwnd, SW_SHOW)
                    ShowWindow(hwnd, SW_RESTORE)
                    shown += 1
            except:
                pass

        # 如果没恢复到，重新搜索所有 Chrome 窗口
        if shown == 0:
            hwnds = find_chrome_windows()
            for hwnd in hwnds:
                ShowWindow(hwnd, SW_SHOW)
                ShowWindow(hwnd, SW_RESTORE)
                shown += 1

        # 把窗口带到前台
        try:
            targets = [h for h in self.hidden_hwnds if IsWindow(h)]
            if targets:
                SetForegroundWindow(targets[0])
            elif find_chrome_windows():
                SetForegroundWindow(find_chrome_windows()[0])
        except:
            pass

        self.is_hidden = False
        self.hidden_hwnds = []
        print(f"  [{time.strftime('%H:%M:%S')}] F9 按下 - Chrome 已显示（{shown} 个窗口）")


def main():
    print()
    print("  ================================================")
    print("    TaobaoScraper - Chrome 显示/隐藏切换器")
    print("  ================================================")
    print()
    print("  F9     = 隐藏/显示 Chrome 窗口")
    print("  Ctrl+C = 退出（自动恢复Chrome显示）")
    print()
    print("  ✅ 正在监听 F9 按键...\n")

    toggler = ChromeToggler()
    key_was_pressed = False

    try:
        while True:
            # 检测 F9 键状态（最高位=1 表示当前被按下）
            state = GetAsyncKeyState(VK_F9)
            is_pressed = (state & 0x8000) != 0

            if is_pressed and not key_was_pressed:
                # F9 刚按下（上升沿）
                toggler.toggle()

            key_was_pressed = is_pressed
            time.sleep(0.05)  # 50ms 轮询，CPU 占用极低

    except KeyboardInterrupt:
        pass
    finally:
        # 退出前确保 Chrome 显示
        if toggler.is_hidden:
            toggler._show_chrome()
            print(f"\n  [退出] 已恢复 Chrome 显示")
        print("  已退出")


if __name__ == "__main__":
    main()
