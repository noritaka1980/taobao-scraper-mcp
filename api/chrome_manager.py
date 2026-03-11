"""Chrome WebDriver 单例管理器

从 scraper_gui.py 中提取 Chrome 初始化/连接逻辑，
提供 acquire/release 模式保证同时只有一个任务使用 driver。
"""
import threading
import logging
import socket
import subprocess
import time
import os
import tempfile
from typing import Optional

log = logging.getLogger("chrome_manager")


class ChromeManager:

    def __init__(self, settings):
        self._driver = None
        self._lock = threading.Lock()        # 保护 driver 创建/销毁
        self._use_lock = threading.Lock()    # 保证同时只有一个任务使用 driver
        self._settings = settings
        self._chromedriver_path = None
        self._is_busy = False
        self._current_task_id: Optional[str] = None

    @property
    def is_busy(self) -> bool:
        return self._is_busy

    @property
    def current_task_id(self) -> Optional[str]:
        return self._current_task_id

    # ── Chrome 查找与连接（提取自 scraper_gui.py:1206-1387）──

    def _ensure_chromedriver(self):
        """确保 ChromeDriver 路径可用"""
        if self._chromedriver_path:
            return
        from webdriver_manager.chrome import ChromeDriverManager
        log.info("正在初始化 ChromeDriver（首次需要下载）...")
        self._chromedriver_path = ChromeDriverManager().install()
        log.info(f"ChromeDriver 路径: {self._chromedriver_path}")

    def _try_connect(self, port: int):
        """尝试连接指定端口的 Chrome 调试端口"""
        try:
            s = socket.socket()
            s.settimeout(0.5)
            result = s.connect_ex(("127.0.0.1", port))
            s.close()
            if result == 0:
                from selenium import webdriver
                from selenium.webdriver.chrome.service import Service
                from selenium.webdriver.chrome.options import Options
                opts = Options()
                opts.add_experimental_option("debuggerAddress", f"127.0.0.1:{port}")
                d = webdriver.Chrome(
                    service=Service(self._chromedriver_path),
                    options=opts,
                )
                return d, port
        except Exception:
            pass
        return None, None

    def _find_chrome(self) -> Optional[str]:
        """查找 Chrome 可执行文件路径"""
        import winreg
        candidates = [
            os.path.join(os.environ.get("PROGRAMFILES", ""), "Google", "Chrome", "Application", "chrome.exe"),
            os.path.join(os.environ.get("PROGRAMFILES(X86)", ""), "Google", "Chrome", "Application", "chrome.exe"),
            os.path.join(os.environ.get("LOCALAPPDATA", ""), "Google", "Chrome", "Application", "chrome.exe"),
        ]
        for p in candidates:
            if p and os.path.isfile(p):
                return p
        try:
            key = winreg.OpenKey(
                winreg.HKEY_LOCAL_MACHINE,
                r"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\chrome.exe",
            )
            val, _ = winreg.QueryValueEx(key, "")
            winreg.CloseKey(key)
            if os.path.isfile(val):
                return val
        except Exception:
            pass
        return None

    def _launch_chrome_with_debug(self, port: int = 9222) -> bool:
        """强制关闭 Chrome 后以调试端口重新启动"""
        import urllib.request as _urlreq

        chrome_path = self._find_chrome()
        if not chrome_path:
            log.error("未找到 Chrome，请确认已安装 Google Chrome")
            return False

        log.info(f"Chrome 路径: {chrome_path}")

        # 关闭所有 Chrome 及辅助进程（CrashHandler 持有文件句柄会阻碍重启）
        log.info("正在关闭现有 Chrome 进程...")
        for proc_name in ["chromedriver.exe", "chrome.exe",
                          "GoogleCrashHandler.exe", "GoogleCrashHandler64.exe"]:
            subprocess.call(
                ["taskkill", "/f", "/t", "/im", proc_name],
                stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL,
            )
        time.sleep(5)

        # 验证 Chrome 已退出
        for _retry in range(10):
            result = subprocess.run(
                ["tasklist", "/fi", "imagename eq chrome.exe"],
                capture_output=True, text=True,
            )
            if "chrome.exe" not in result.stdout.lower():
                log.info("Chrome 进程已全部退出")
                break
            log.info(f"Chrome 仍在运行，重试关闭（{_retry + 1}/10）...")
            subprocess.call(
                ["taskkill", "/f", "/t", "/im", "chrome.exe"],
                stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL,
            )
            time.sleep(2)

        # 清除锁文件
        # 使用独立目录，避免 Chrome 136+ 默认目录静默忽略 --remote-debugging-port
        user_data = os.path.join(
            os.environ.get("LOCALAPPDATA", ""),
            "Google", "Chrome", "Debug Profile",
        )
        for lock_name in ["SingletonLock", "SingletonSocket", "SingletonCookie"]:
            lock_path = os.path.join(user_data, lock_name)
            try:
                if os.path.exists(lock_path):
                    os.remove(lock_path)
                    log.info(f"已清除锁文件: {lock_name}")
            except Exception as e:
                log.warning(f"清除锁文件失败 {lock_name}: {e}")

        time.sleep(2)

        # 通过临时 BAT 文件启动 Chrome
        bat_content = f'''@echo off
start "" "{chrome_path}" ^
  --remote-debugging-port={port} ^
  --user-data-dir="{user_data}" ^
  --disable-blink-features=AutomationControlled ^
  --no-first-run ^
  --no-default-browser-check ^
  --no-restore-last-session
'''
        bat_path = os.path.join(tempfile.gettempdir(), "_chrome_launch.bat")
        with open(bat_path, "w", encoding="utf-8") as f:
            f.write(bat_content)

        # 清除脏环境变量，避免继承已死进程的 crashpad 管道导致初始化异常
        clean_env = {k: v for k, v in os.environ.items()
                     if k != "CHROME_CRASHPAD_PIPE_NAME"}

        log.info(f"正在启动 Chrome（调试端口 {port}）...")
        subprocess.Popen(
            ["cmd", "/c", bat_path],
            stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL,
            env=clean_env,
        )

        # 等待端口就绪（最多 45 秒）
        log.info("等待调试端口就绪...")
        for i in range(90):
            time.sleep(0.5)
            if i > 0 and i % 20 == 0:
                log.info(f"仍在等待... ({i // 2}秒)")
            try:
                s = socket.socket()
                s.settimeout(0.5)
                rc = s.connect_ex(("127.0.0.1", port))
                s.close()
                if rc != 0:
                    continue
            except Exception:
                continue
            try:
                req = _urlreq.Request(f"http://127.0.0.1:{port}/json/version")
                resp = _urlreq.urlopen(req, timeout=2)
                if resp.status == 200:
                    log.info(f"Chrome 已启动，调试端口 {port} 就绪")
                    try:
                        os.remove(bat_path)
                    except Exception:
                        pass
                    return True
            except Exception:
                pass

        # 备用：用独立 user-data-dir 启动
        log.warning("使用默认 User Data 启动失败，尝试用独立目录启动...")
        subprocess.call(
            ["taskkill", "/f", "/t", "/im", "chrome.exe"],
            stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL,
        )
        time.sleep(3)

        alt_data = os.path.join(tempfile.gettempdir(), "chrome_scraper_profile")
        bat_content2 = f'''@echo off
start "" "{chrome_path}" ^
  --remote-debugging-port={port} ^
  --user-data-dir="{alt_data}" ^
  --disable-blink-features=AutomationControlled ^
  --no-first-run ^
  --no-default-browser-check
'''
        with open(bat_path, "w", encoding="utf-8") as f:
            f.write(bat_content2)

        log.info(f"使用独立目录: {alt_data}")
        subprocess.Popen(
            ["cmd", "/c", bat_path],
            stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL,
            env=clean_env,
        )

        for i in range(40):
            time.sleep(0.5)
            try:
                req = _urlreq.Request(f"http://127.0.0.1:{port}/json/version")
                resp = _urlreq.urlopen(req, timeout=2)
                if resp.status == 200:
                    log.info(f"Chrome 已用独立目录启动，端口 {port} 就绪")
                    log.warning("使用了独立目录，需要重新登录淘宝")
                    try:
                        os.remove(bat_path)
                    except Exception:
                        pass
                    return True
            except Exception:
                pass

        log.error("两种方式都无法启动带调试端口的 Chrome")
        try:
            os.remove(bat_path)
        except Exception:
            pass
        return False

    # ── 公开接口 ──

    def get_or_create_driver(self):
        """获取或创建 driver（线程安全）"""
        with self._lock:
            self._ensure_chromedriver()

            if self._driver:
                try:
                    self._driver.current_url
                    return self._driver
                except Exception:
                    log.warning("现有 driver 已失效，将重新连接")
                    self._driver = None

            # 尝试连接已有的 Chrome 调试端口
            for port in self._settings.chrome_fallback_ports:
                d, p = self._try_connect(port)
                if d:
                    log.info(f"已连接到 Chrome（端口 {p}）")
                    self._driver = d
                    return d

            # 连不上就报错，不自动杀 Chrome 重启（避免丢失登录态和标签页）
            raise RuntimeError(
                "无法连接 Chrome 调试端口，请先启动 Chrome"
                "（运行 START_GUI.bat 或以 --remote-debugging-port=9222 启动）"
            )

    def acquire(self, task_id: str, timeout: float = 300):
        """
        获取 driver 使用权（阻塞等待直到可用或超时）。
        调用方必须在 finally 中调用 release()。
        """
        acquired = self._use_lock.acquire(timeout=timeout)
        if not acquired:
            raise TimeoutError(
                f"等待 Chrome 超时（{timeout}秒），当前任务: {self._current_task_id}"
            )
        self._is_busy = True
        self._current_task_id = task_id
        try:
            return self.get_or_create_driver()
        except Exception:
            self._use_lock.release()
            self._is_busy = False
            self._current_task_id = None
            raise

    def release(self):
        """归还 driver 使用权，断开 chromedriver 让用户正常使用浏览器"""
        # 恢复 Chrome 窗口到正常位置
        if self._driver:
            try:
                from scraper_core import restore_chrome_window
                restore_chrome_window(self._driver)
            except Exception:
                pass
            self._driver = None
        # 杀残余 chromedriver 进程
        try:
            subprocess.call(
                ["taskkill", "/f", "/im", "chromedriver.exe"],
                stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL,
            )
        except Exception:
            pass
        self._is_busy = False
        self._current_task_id = None
        try:
            self._use_lock.release()
        except RuntimeError:
            pass

    def restart(self):
        """断开 chromedriver 后重新连接（不杀 Chrome 进程）"""
        with self._lock:
            self._driver = None
            # 杀残余 chromedriver 进程
            try:
                subprocess.call(
                    ["taskkill", "/f", "/im", "chromedriver.exe"],
                    stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL,
                )
            except Exception:
                pass
            time.sleep(1)
        return self.get_or_create_driver()

    def shutdown(self):
        """关闭 driver（应用退出时调用）"""
        with self._lock:
            if self._driver:
                try:
                    from scraper_core import restore_chrome_window
                    restore_chrome_window(self._driver)
                except Exception:
                    pass
                try:
                    self._driver.quit()
                except Exception:
                    pass
                self._driver = None

    def status(self) -> dict:
        """返回当前状态"""
        driver_alive = False
        if self._driver:
            try:
                self._driver.current_url
                driver_alive = True
            except Exception:
                pass
        return {
            "driver_alive": driver_alive,
            "is_busy": self._is_busy,
            "current_task_id": self._current_task_id,
        }
