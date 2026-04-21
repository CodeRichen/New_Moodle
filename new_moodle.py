# pip install selenium webdriver-manager colorama requests py7zr patool rarfile gdown
# 不要在開啟Moodle網頁的狀態執行程式

import os
import sys
import ctypes
import contextlib
import io
import re
import shutil

# Moodle 主站（需要在第一次輸入帳密前就可用）
MOODLE_BASE_URL = "https://elearningv4.nuk.edu.tw"

# --- 登入錯誤快取（讓 test_login / load_credentials / ensure_logged_in 能顯示明確原因） ---
_LAST_LOGIN_ERROR_KIND = None
_LAST_LOGIN_ERROR_TEXT = None

def clear_last_login_error():
    global _LAST_LOGIN_ERROR_KIND, _LAST_LOGIN_ERROR_TEXT
    _LAST_LOGIN_ERROR_KIND = None
    _LAST_LOGIN_ERROR_TEXT = None

def set_last_login_error(kind, text):
    global _LAST_LOGIN_ERROR_KIND, _LAST_LOGIN_ERROR_TEXT
    _LAST_LOGIN_ERROR_KIND = kind
    _LAST_LOGIN_ERROR_TEXT = text

def get_last_login_error():
    return _LAST_LOGIN_ERROR_KIND, _LAST_LOGIN_ERROR_TEXT

# ========== 使用 Windows API 設定終端機視窗為全螢幕 ==========
def maximize_console_window():
    """將終端機視窗最大化（僅在非 exe 環境下執行）"""
    try:
        if getattr(sys, 'frozen', False):
            return
        if os.name == 'nt':  # Windows only
            hwnd = ctypes.windll.kernel32.GetConsoleWindow()
            if hwnd:
                ctypes.windll.user32.ShowWindow(hwnd, 3)  # SW_MAXIMIZE = 3
        # macOS / Linux：不需要額外操作，terminal 本身管理視窗
    except Exception:
        pass

# 執行視窗最大化
maximize_console_window()

# ========== 降噪：過濾 Chrome GPU 洗版錯誤（只過濾已知無害訊息） ==========
# 例如：gpu_channel_manager.cc:919 Failed to create shared context for virtualization.
# 這些通常不影響 Selenium 功能，但會洗版終端。
_DISABLE_STDERR_FILTER = os.environ.get("DISABLE_STDERR_FILTER", "0").strip().lower() in {"1", "true", "yes", "on"}

if not _DISABLE_STDERR_FILTER:
    try:
        _NOISE_PATTERNS = [
            re.compile(r"gpu_channel_manager\.cc", re.IGNORECASE),
            re.compile(r"failed to create shared context for virtualization", re.IGNORECASE),
        ]

        class _StderrNoiseFilter(io.TextIOBase):
            def __init__(self, underlying):
                self._u = underlying
                self._buf = ""

            def write(self, s):
                try:
                    if not isinstance(s, str):
                        s = str(s)
                except Exception:
                    return 0

                self._buf += s
                written = 0
                while "\n" in self._buf:
                    line, self._buf = self._buf.split("\n", 1)
                    if any(p.search(line) for p in _NOISE_PATTERNS):
                        continue
                    try:
                        self._u.write(line + "\n")
                        written += len(line) + 1
                    except Exception:
                        pass
                return written

            def flush(self):
                # flush 時，把尚未換行的內容也嘗試輸出（但仍過濾）
                try:
                    if self._buf:
                        tail = self._buf
                        self._buf = ""
                        if not any(p.search(tail) for p in _NOISE_PATTERNS):
                            self._u.write(tail)
                    self._u.flush()
                except Exception:
                    pass

            def isatty(self):
                try:
                    return self._u.isatty()
                except Exception:
                    return False

            @property
            def encoding(self):
                try:
                    return getattr(self._u, "encoding", "utf-8")
                except Exception:
                    return "utf-8"

        sys.stderr = _StderrNoiseFilter(sys.stderr)
    except Exception:
        pass

os.environ['WDM_LOG_LEVEL'] = '0' # 針對 webdriver-manager 的日誌屏蔽
os.environ['TF_CPP_MIN_LOG_LEVEL'] = '3'  #關閉 TensorFlow 的日誌
import subprocess
import platform
import atexit
import urllib.parse
import http.server
import socketserver
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import StaleElementReferenceException, SessionNotCreatedException, WebDriverException, TimeoutException, InvalidSessionIdException
from selenium.webdriver.chrome.options import Options
import time
import tempfile
import zipfile
import py7zr
import requests  # 用於直接下載圖片
import threading
from concurrent.futures import ThreadPoolExecutor, as_completed
from collections import defaultdict
import json
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager
from urllib.parse import urlparse, parse_qs

@contextlib.contextmanager
def _suppress_stdio_ctx():
    try:
        with open(os.devnull, 'w') as devnull:
            with contextlib.redirect_stdout(devnull), contextlib.redirect_stderr(devnull):
                yield
    except Exception:
        yield

# 全域變數儲存 ChromeDriver 路徑，避免重複下載
_cached_driver_path = None

def get_chrome_driver_path():
    """獲取 ChromeDriver 路徑，包含重試機制和回退方案"""
    global _cached_driver_path

    # 若使用者或安裝流程已提供明確路徑，優先使用
    env_driver = os.environ.get("CHROMEDRIVER_PATH")
    if env_driver and os.path.exists(env_driver):
        _cached_driver_path = env_driver
        return _cached_driver_path

    # macOS：若已安裝 Chrome for Testing，優先使用其對應 chromedriver
    if platform.system() == "Darwin":
        try:
            home_dir = os.path.expanduser("~")
            arch = (platform.machine() or "").lower()
            plat = "mac-arm64" if "arm" in arch else "mac-x64"
            candidate = os.path.join(home_dir, "chrome-for-testing", f"chromedriver-{plat}", "chromedriver")
            if os.path.exists(candidate):
                _cached_driver_path = candidate
                return _cached_driver_path
        except Exception:
            pass

    # 允許在 Linux/Docker 環境使用系統內建的 chromedriver，
    # 避免 webdriver-manager 下載到不相容版本。
    if os.environ.get("USE_SYSTEM_CHROMEDRIVER", "0") == "1":
        return None
    
    # 如果已經快取，直接返回
    if _cached_driver_path:
        return _cached_driver_path
    
    # 嘗試使用 webdriver-manager（最多重試3次）
    for attempt in range(3):
        try:
            # webdriver-manager 在網路不穩時可能會自行印一堆錯誤；這裡抑制輸出避免洗版
            with _suppress_stdio_ctx():
                _cached_driver_path = ChromeDriverManager().install()
            return _cached_driver_path
        except Exception as e:
            if attempt < 2:
                time.sleep(1)  # 等待1秒後重試
                continue
    
    # 回退方案：返回 None，讓 Selenium 自動尋找 ChromeDriver
    return None

def get_chrome_binary_path():
    """跨平台尋找 Chrome/Chromium 執行檔路徑。"""
    env_binary = os.environ.get("CHROME_BINARY")
    if env_binary and os.path.exists(env_binary):
        return env_binary

    system_name = platform.system()
    candidates = []

    if system_name == "Darwin":
        home_dir = os.path.expanduser("~")
        # Big Sur (macOS 11.x) 上，使用者可能已安裝「太新」的系統 Chrome，啟動會直接 crash。
        # 因此預設優先找 Chrome for Testing（若存在），最後才用 /Applications 的 system Chrome。
        try:
            ver = platform.mac_ver()[0] or ""
            parts = [int(p) for p in ver.split('.') if p.isdigit()]
            mac_major = parts[0] if parts else 0
        except Exception:
            mac_major = 0

        cft_candidates = [
            os.path.join(home_dir, "chrome-for-testing", "chrome-mac-x64", "Google Chrome for Testing.app", "Contents", "MacOS", "Google Chrome for Testing"),
            os.path.join(home_dir, "chrome-for-testing", "chrome-mac-arm64", "Google Chrome for Testing.app", "Contents", "MacOS", "Google Chrome for Testing"),
            "/Applications/Google Chrome for Testing.app/Contents/MacOS/Google Chrome for Testing",
        ]
        system_candidates = [
            "/Applications/Chromium.app/Contents/MacOS/Chromium",
            "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome",
        ]
        if mac_major and mac_major <= 11 and os.environ.get("PREFER_SYSTEM_CHROME", "0") != "1":
            candidates = cft_candidates + system_candidates
        else:
            candidates = [
                "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome",
                "/Applications/Google Chrome for Testing.app/Contents/MacOS/Google Chrome for Testing",
                "/Applications/Chromium.app/Contents/MacOS/Chromium",
                os.path.join(home_dir, "chrome-for-testing", "chrome-mac-x64", "Google Chrome for Testing.app", "Contents", "MacOS", "Google Chrome for Testing"),
                os.path.join(home_dir, "chrome-for-testing", "chrome-mac-arm64", "Google Chrome for Testing.app", "Contents", "MacOS", "Google Chrome for Testing"),
            ]
    elif system_name == "Windows":
        local_appdata = os.environ.get("LOCALAPPDATA", "")
        program_files = os.environ.get("PROGRAMFILES", "")
        program_files_x86 = os.environ.get("PROGRAMFILES(X86)", "")
        candidates = [
            os.path.join(local_appdata, "Google", "Chrome", "Application", "chrome.exe"),
            os.path.join(program_files, "Google", "Chrome", "Application", "chrome.exe"),
            os.path.join(program_files_x86, "Google", "Chrome", "Application", "chrome.exe"),
            os.path.join(program_files, "Chromium", "Application", "chrome.exe"),
            os.path.join(program_files_x86, "Chromium", "Application", "chrome.exe"),
        ]
    else:
        candidates = [
            "/usr/bin/google-chrome",
            "/usr/bin/google-chrome-stable",
            "/usr/bin/chromium",
            "/usr/bin/chromium-browser",
            "/snap/bin/chromium",
        ]

    for candidate in candidates:
        if candidate and os.path.exists(candidate):
            return candidate

    return None

def apply_chrome_binary_option(chrome_options):
    """若找到 Chrome 可執行檔，套用到 Selenium options。"""
    binary_path = get_chrome_binary_path()
    if binary_path:
        chrome_options.binary_location = binary_path
    return binary_path

def should_use_safari_fallback():
    """在 macOS 是否回退 Safari。

    依使用者要求：預設優先使用 Chrome；只有在明確指定 `FORCE_SAFARI=1` 時才回退。
    """
    if platform.system() != "Darwin":
        return False
    force_safari = os.environ.get("FORCE_SAFARI", "0").strip().lower() in {"1", "true", "yes", "on"}
    force_chrome = os.environ.get("FORCE_CHROME", "0").strip().lower() in {"1", "true", "yes", "on"}
    return bool(force_safari and not force_chrome)

def _is_chrome_startup_issue(exc):
    """判斷是否屬於可透過更換 profile 重試的 Chrome 啟動問題。"""
    msg = str(exc).lower()
    keywords = ["devtoolsactiveport", "chrome failed to start", "session not created", "chrome not reachable", "crashed"]
    return any(k in msg for k in keywords)

def _kill_existing_chrome_processes():
    """（已改為預設不使用）殺死現有 Chrome 進程（Windows）。

    先前為了處理 DevToolsActivePort 等啟動問題而加入，但會把使用者正在使用的
    Chrome 視窗全部關掉。現在改為：預設不殺 Chrome；只在明確允許時才執行。

    允許方式：設定環境變數 `ALLOW_KILL_CHROME=1`。
    """
    if os.name != 'nt':
        return
    allow = os.environ.get("ALLOW_KILL_CHROME", "0").strip().lower() in {"1", "true", "yes", "on"}
    if not allow:
        return
    try:
        import subprocess as _sp
        _sp.run(["taskkill", "/IM", "chrome.exe", "/F"],
                stdout=_sp.DEVNULL, stderr=_sp.DEVNULL, timeout=2)
        time.sleep(0.5)
    except Exception:
        pass

def _prepare_windows_chrome_options(chrome_options):
    """Windows 啟動穩定化：補齊必要參數。"""
    if os.name != 'nt' or chrome_options is None:
        return chrome_options

    existing_args = set(chrome_options.arguments)

    def _add_arg(arg):
        if arg not in existing_args:
            chrome_options.add_argument(arg)
            existing_args.add(arg)

    # 遠端調試 port 使用動態分配（避免 DevToolsActivePort 衝突）
    # _add_arg("--remote-debugging-port=0")
    
    # 基礎穩定化參數
    _add_arg("--disable-gpu")
    # 避免 GPU 子程序反覆噴錯（特別是開啟虛擬化/遠端桌面時）
    _add_arg("--use-gl=swiftshader")
    _add_arg("--disable-dev-shm-usage")
    _add_arg("--no-sandbox")
    _add_arg("--disable-software-rasterizer")
    _add_arg("--disable-features=RendererCodeIntegrity")
    
    # 新版 Chrome 解決 DevToolsActivePort 找不到的方案
    _add_arg("--remote-debugging-pipe")

    _add_arg("--no-first-run")
    _add_arg("--no-default-browser-check")
    _add_arg("--disable-extensions")
    
    return chrome_options


def create_webdriver(chrome_options=None, *, hide_windows_console=False):
    """統一建立 WebDriver，必要時在 macOS 回退 Safari。"""
    force_chrome = os.environ.get("FORCE_CHROME", "0").strip().lower() in {"1", "true", "yes", "on"}

    if should_use_safari_fallback():
        print("[INFO] macOS 未偵測到 Chrome，改用 Safari WebDriver。")
        return webdriver.Safari()

    if os.name == 'nt' and chrome_options is not None:
        _prepare_windows_chrome_options(chrome_options)

    def _create_chrome_driver(options_to_use=None):
        active_options = options_to_use if options_to_use is not None else chrome_options
        driver_path = get_chrome_driver_path()
        if driver_path:
            # 盡量把 ChromeDriver/Chrome 的雜訊導向 devnull，避免終端洗版
            try:
                service = Service(driver_path, log_output=subprocess.DEVNULL, service_args=["--log-level=OFF"])
            except TypeError:
                service = Service(driver_path, log_path=os.devnull)
            if os.name == 'nt' and hide_windows_console:
                service.creation_flags = subprocess.CREATE_NO_WINDOW
            if active_options is not None:
                return webdriver.Chrome(options=active_options, service=service)
            return webdriver.Chrome(service=service)

        if active_options is not None:
            return webdriver.Chrome(options=active_options)
        return webdriver.Chrome()

    # macOS 上即使找到 Chrome，仍可能因安裝不完整造成啟動失敗，這時回退 Safari。
    if platform.system() == "Darwin" and not force_chrome:
        try:
            return _create_chrome_driver()
        except (SessionNotCreatedException, WebDriverException) as e:
            print(f"[WARN] Chrome 啟動失敗，改用 Safari。原因: {e}")
            return webdriver.Safari()

    # Windows：不再使用「臨時 profile 重試」邏輯（避免產生額外配置檔/行為不一致）。
    return _create_chrome_driver()

# 嘗試導入 RAR 解壓工具
try:
    import patool
    HAS_PATOOL = True
except ImportError:
    HAS_PATOOL = False

try:
    import rarfile
    HAS_RARFILE = True
except ImportError:
    HAS_RARFILE = False

# 偵測終端是否支援 ANSI 顏色，並在 Windows 上主動啟用 VT 模式
import sys as _sys
def _supports_color():
    if os.environ.get("NO_COLOR"):   # https://no-color.org/ 明確要求關閉
        return False
    if os.name == 'nt':
        # Win10 v1511+ 原生支援，但需透過 SetConsoleMode 明確啟用
        # ENABLE_VIRTUAL_TERMINAL_PROCESSING = 0x0004
        try:
            import ctypes, ctypes.wintypes
            kernel32 = ctypes.windll.kernel32
            hOut = kernel32.GetStdHandle(-11)  # STD_OUTPUT_HANDLE
            mode = ctypes.wintypes.DWORD()
            if kernel32.GetConsoleMode(hOut, ctypes.byref(mode)):
                kernel32.SetConsoleMode(hOut, mode.value | 0x0004)
                return True  # 成功啟用 VT mode
        except Exception:
            pass
        # fallback：VS Code terminal / ConPTY / Windows Terminal 等偽 tty 也支援
        if os.environ.get("TERM_PROGRAM") or os.environ.get("WT_SESSION") or os.environ.get("COLORTERM"):
            return True
        return False  # 真的不支援（舊版 cmd.exe、無色彩環境）
    # macOS / Linux：只要 stdout 是 tty 就支援
    return hasattr(_sys.stdout, 'isatty') and _sys.stdout.isatty()

_USE_COLOR = _supports_color()

if _USE_COLOR:
    YELLOW   = "\033[33m"
    # macOS 終端機改用較穩定、接近原色的紅綠色
    if platform.system() == "Darwin":
        RED      = "\033[38;2;255;99;132m"
        GREEN    = "\033[38;2;46;204;113m"
    else:
        RED      = "\033[38;2;255;105;180m"
        GREEN    = "\033[38;2;055;205;180m"
    BLUE     = "\033[34m"
    BBLUE    = "\033[94m"
    ITALIC   = "\033[3;37m"
    MIKU     = "\033[36m"
    LOWGREEN = "\033[32m"
    PURPLE   = "\033[38;5;129m"
    ORANGE   = "\033[38;5;214m"
    RESET    = "\033[0m"
    PINK     = "\033[38;2;255;220;255m"
else:
    YELLOW = RED = BLUE = BBLUE = ITALIC = MIKU = ""
    LOWGREEN = GREEN = PURPLE = ORANGE = RESET = PINK = ""

# ========== TODO 路徑設定區域（修改這裡可以改變所有檔案存放位置）==========
# 主要下載目錄 - 修改這裡就能改變所有檔案的存放位置
# 例如：r"D:\Moodle" 或 r"E:\課程資料"
BASE_DOWNLOAD_DIR = os.path.join(os.path.expanduser("~"), "Downloads", "class")
# ================================================================

# ========== 下載暫存資料夾清理（確保每次執行結束都會刪除）==========
_TEMP_DL_DIRS = set()

# ========== 模擬瀏覽器（點本機連結自動登入並開啟活動）==========
_SIM_SERVER = None
_SIM_BASE_URL = None
_SIM_DRIVER = None
_SIM_DRIVER_LOCK = threading.Lock()
_SIM_LINK_LOCK = threading.Lock()
_SIM_LINK_MAP = {}          # token -> original_url
_SIM_LINK_REVERSE = {}      # original_url -> token
_SIM_LINK_COUNTER = 0

def _base36(n: int) -> str:
    chars = "0123456789abcdefghijklmnopqrstuvwxyz"
    if n <= 0:
        return "0"
    out = []
    while n:
        n, r = divmod(n, 36)
        out.append(chars[r])
    return "".join(reversed(out))

def _register_short_link(url: str) -> str:
    """回傳短代碼 token；同一個 url 會重用相同 token。"""
    global _SIM_LINK_COUNTER
    if not url:
        return ""
    with _SIM_LINK_LOCK:
        if url in _SIM_LINK_REVERSE:
            return _SIM_LINK_REVERSE[url]
        _SIM_LINK_COUNTER += 1
        token = _base36(_SIM_LINK_COUNTER)
        _SIM_LINK_MAP[token] = url
        _SIM_LINK_REVERSE[url] = token
        return token

def _build_simulator_driver_options():
    opts = Options()
    opts.add_argument("--log-level=3")
    opts.add_experimental_option("excludeSwitches", ["enable-logging"])
    opts.add_argument("--disable-gpu")
    opts.add_argument("--disable-dev-shm-usage")
    opts.add_argument("--remote-debugging-pipe")
    opts.add_argument("--disable-software-rasterizer")
    if os.name == 'nt':
        opts.add_argument("--no-sandbox")
        opts.add_argument("--disable-features=RendererCodeIntegrity")
    apply_chrome_binary_option(opts)
    return opts

def _ensure_simulator_driver():
    """確保可見的 Selenium 瀏覽器存在且已登入。"""
    global _SIM_DRIVER
    with _SIM_DRIVER_LOCK:
        # 先檢查現有 session
        if _SIM_DRIVER is not None:
            try:
                _ = _SIM_DRIVER.current_url
            except Exception:
                try:
                    _SIM_DRIVER.quit()
                except Exception:
                    pass
                _SIM_DRIVER = None

        if _SIM_DRIVER is None:
            _SIM_DRIVER = create_webdriver(_build_simulator_driver_options(), hide_windows_console=True)

        # 確保已登入（若被導向登入頁會自動登入）
        try:
            ensure_logged_in(_SIM_DRIVER, USERNAME, PASSWORD, silent=True)
        except Exception:
            pass
        return _SIM_DRIVER

def _open_in_simulator(url: str) -> bool:
    if not url:
        return False
    try:
        drv = _ensure_simulator_driver()
        drv.get(url)
        # 若被重導向登入頁：登入後再跳回原活動
        try:
            cur = (drv.current_url or "").lower()
        except Exception:
            cur = ""
        if "login" in cur:
            try:
                ensure_logged_in(drv, USERNAME, PASSWORD, silent=True)
                drv.get(url)
            except Exception:
                pass
        return True
    except Exception:
        return False

class _SimulatorRequestHandler(http.server.BaseHTTPRequestHandler):
    def do_GET(self):
        parsed = urllib.parse.urlparse(self.path)
        if parsed.path == "/health":
            self.send_response(200)
            self.send_header("Content-Type", "text/plain; charset=utf-8")
            self.end_headers()
            self.wfile.write("ok".encode("utf-8"))
            return

        # 短連結：/o/<token>
        if parsed.path.startswith("/o/"):
            token = parsed.path.split("/", 2)[2] if len(parsed.path.split("/", 2)) >= 3 else ""
            url = ""
            with _SIM_LINK_LOCK:
                url = _SIM_LINK_MAP.get(token, "")
            if not url:
                self.send_response(404)
                self.send_header("Content-Type", "text/plain; charset=utf-8")
                self.end_headers()
                self.wfile.write("找不到短連結".encode("utf-8"))
                return
            ok = _open_in_simulator(url)
            self.send_response(200 if ok else 500)
            self.send_header("Content-Type", "text/plain; charset=utf-8")
            self.end_headers()
            msg = f"已在模擬器中開啟：{url}" if ok else f"開啟失敗：{url}"
            self.wfile.write(msg.encode("utf-8"))
            return

        if parsed.path != "/open":
            self.send_response(404)
            self.end_headers()
            return

        qs = urllib.parse.parse_qs(parsed.query)
        url = (qs.get("url") or [""])[0]
        url = urllib.parse.unquote(url)
        if not (url.startswith("http://") or url.startswith("https://")):
            self.send_response(400)
            self.send_header("Content-Type", "text/plain; charset=utf-8")
            self.end_headers()
            self.wfile.write("無效的 URL".encode("utf-8"))
            return

        ok = _open_in_simulator(url)
        self.send_response(200 if ok else 500)
        self.send_header("Content-Type", "text/plain; charset=utf-8")
        self.end_headers()
        msg = f"已在模擬器中開啟：{url}" if ok else f"開啟失敗：{url}"
        self.wfile.write(msg.encode("utf-8"))

    def log_message(self, format, *args):
        # 靜默，不污染終端輸出
        return

def start_simulator_server():
    """啟動本機 HTTP 轉發器，回傳 base url（例如 http://127.0.0.1:12345）。"""
    global _SIM_SERVER, _SIM_BASE_URL
    if os.environ.get("DISABLE_SIM_SERVER", "0").strip() in {"1", "true", "yes", "on"}:
        return None
    if _SIM_BASE_URL:
        return _SIM_BASE_URL
    try:
        class _ThreadingHTTPServer(socketserver.ThreadingMixIn, http.server.HTTPServer):
            daemon_threads = True

        _SIM_SERVER = _ThreadingHTTPServer(("127.0.0.1", 0), _SimulatorRequestHandler)
        port = _SIM_SERVER.server_address[1]
        _SIM_BASE_URL = f"http://127.0.0.1:{port}"
        threading.Thread(target=_SIM_SERVER.serve_forever, daemon=True).start()

        def _shutdown():
            try:
                if _SIM_SERVER:
                    _SIM_SERVER.shutdown()
            except Exception:
                pass
            try:
                if _SIM_SERVER:
                    _SIM_SERVER.server_close()
            except Exception:
                pass
            try:
                if _SIM_DRIVER is not None:
                    _SIM_DRIVER.quit()
            except Exception:
                pass

        atexit.register(_shutdown)
        return _SIM_BASE_URL
    except Exception:
        _SIM_SERVER = None
        _SIM_BASE_URL = None
        return None

def make_simulator_open_link(original_url: str) -> str:
    """把 Moodle/外部連結包成可點擊的本機短連結（若伺服器不可用則回傳原網址）。"""
    if not original_url or not isinstance(original_url, str):
        return original_url
    base = _SIM_BASE_URL
    if not base:
        return original_url
    try:
        token = _register_short_link(original_url)
        if token:
            return f"{base}/o/{token}"
        encoded = urllib.parse.quote(original_url, safe="")
        return f"{base}/open?url={encoded}"
    except Exception:
        return original_url

def _register_temp_dl_dir(path: str) -> None:
    try:
        if path:
            _TEMP_DL_DIRS.add(path)
    except Exception:
        pass

def _cleanup_temp_dl_dirs() -> None:
    """刪除所有課程的 `_temp_dl` 暫存資料夾。只清理已登錄的新下載資料夾，不遍歷全目錄。"""
    import shutil
    import stat

    def _on_rm_error(func, path, exc_info):
        try:
            os.chmod(path, stat.S_IWRITE)
            func(path)
        except Exception:
            pass

    def _rmtree_retry(path: str, retries: int = 6) -> bool:
        if not path or not os.path.isdir(path):
            return True
        for i in range(retries):
            try:
                shutil.rmtree(path, onerror=_on_rm_error)
                return True
            except Exception:
                # Windows 常見：Chrome 尚未完全釋放檔案鎖，稍等重試
                try:
                    time.sleep(0.3 + i * 0.1)
                except Exception:
                    pass
        return False

    # 僅清理已註冊的（不再 os.walk 整個大目錄）
    try:
        for p in sorted(_TEMP_DL_DIRS, key=lambda s: len(s or ""), reverse=True):
            _rmtree_retry(p)
        _TEMP_DL_DIRS.clear()
    except Exception:
        pass

atexit.register(_cleanup_temp_dl_dirs)

# 根據 BASE_DOWNLOAD_DIR 自動設定其他檔案路徑
OUTPUT_FILE = os.path.join(BASE_DOWNLOAD_DIR, "cless.txt")
SUBMITTED_ASSIGNMENTS_FILE = os.path.join(BASE_DOWNLOAD_DIR, "submitted_assignments.json")
PENDING_ASSIGNMENTS_FILE = os.path.join(BASE_DOWNLOAD_DIR, "pending_assignments.txt")
ASSIGNMENT_CHECK_LIMIT = 12
PASSWORD_FILE = os.path.join(BASE_DOWNLOAD_DIR, "password.txt")
BUILDERROR_MARKER = "builderror"
NONPOP_MARKER = "nonpop"
POPUP_DISABLE_MARKERS = {"nonpop", "nopop"}

# 確保主目錄存在
os.makedirs(BASE_DOWNLOAD_DIR, exist_ok=True)

# ========== 全域錯誤捕捉：出錯時輸出至 error_log.txt ==========
import traceback as _traceback
_ERROR_LOG = os.path.join(BASE_DOWNLOAD_DIR, "error_log.txt")

def _write_error_log(exc_type, exc_value, exc_tb, *, source="主執行緒"):
    import datetime
    try:
        with open(_ERROR_LOG, 'a', encoding='utf-8') as _f:
            _f.write(f"\n{'='*60}\n")
            _f.write(f"時間：{datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
            _f.write(f"來源：{source}\n")
            _f.write(_traceback.format_exc() if exc_tb is None
                     else ''.join(_traceback.format_exception(exc_type, exc_value, exc_tb)))
            _f.write(f"{'='*60}\n")
    except Exception:
        pass

def _global_excepthook(exc_type, exc_value, exc_tb):
    _write_error_log(exc_type, exc_value, exc_tb)
    # print(f"\n\033[31mX 程式發生未預期錯誤，詳情已儲存至：\n   {_ERROR_LOG}\033[0m")
    sys.__excepthook__(exc_type, exc_value, exc_tb)

sys.excepthook = _global_excepthook

def _thread_excepthook(args):
    _write_error_log(args.exc_type, args.exc_value, args.exc_traceback,
                     source=f"背景執行緒 {getattr(args.thread, 'name', '?')}")

import threading as _threading_err
_threading_err.excepthook = _thread_excepthook
# ================================================================

# 初始化第一次使用標記
IS_FIRST_TIME = False
AUTO_OPEN_NEW_ACTIVITY_FOLDERS = True
IS_BUILD_ENV = False

def should_emit_login_failure_messages() -> bool:
    """是否需要輸出『登入失敗』提示字樣。

    只有在 password.txt 仍含 builderror（環境尚未成功建置）時才輸出，
    避免在日常執行中因偶發網路/跳轉造成洗版。
    """
    try:
        if IS_BUILD_ENV:
            return True
    except Exception:
        pass

    try:
        if os.path.exists(PASSWORD_FILE):
            with open(PASSWORD_FILE, 'r', encoding='utf-8') as f:
                lines = f.read().splitlines()
            return any((line or '').strip().lower() == BUILDERROR_MARKER for line in lines)
    except Exception:
        pass

    return False

def get_password_input(prompt):
    """自定義密碼輸入函數，顯示星號（Windows）或隱藏輸入（macOS/Linux）"""
    if os.name == 'nt':
        import msvcrt
        print(prompt, end='', flush=True)
        password = ""
        while True:
            char = msvcrt.getch()
            if char == b'\r':  # Enter 鍵
                break
            elif char == b'\x08':  # Backspace
                if password:
                    password = password[:-1]
                    print('\b \b', end='', flush=True)
            elif char == b'\x03':  # Ctrl+C
                print()
                sys.exit(1)
            else:
                try:
                    password += char.decode('utf-8')
                    print('*', end='', flush=True)
                except UnicodeDecodeError:
                    pass
        print()
        return password
    else:
        # macOS / Linux：用標準庫 getpass，輸入時自動隱藏
        import getpass
        return getpass.getpass(prompt)

def test_login(username, password):
    """測試登入是否成功，使用現有的登入函數"""
    test_driver = None
    try:
        # 創建臨時瀏覽器進行登入測試
        test_chrome_options = Options()
        if os.name == 'nt':
            test_chrome_options.add_argument("--headless")
        else:
            test_chrome_options.add_argument("--headless")
        test_chrome_options.add_argument("--disable-gpu")
        test_chrome_options.add_argument("--log-level=3")
        test_chrome_options.add_argument("--disable-extensions")
        test_chrome_options.add_argument("--disable-dev-shm-usage")
        test_chrome_options.add_argument("--no-sandbox")
        test_chrome_options.add_experimental_option("excludeSwitches", ["enable-logging"])
        apply_chrome_binary_option(test_chrome_options)

        test_driver = create_webdriver(test_chrome_options, hide_windows_console=True)
        clear_last_login_error()

        login_url = f"{MOODLE_BASE_URL}/login/index.php?loginredirect=1"
        test_driver.get(login_url)
        user_el = WebDriverWait(test_driver, 10).until(
            EC.visibility_of_element_located((By.ID, "username"))
        )
        try:
            user_el.clear()
        except Exception:
            pass
        user_el.send_keys(username)

        pw_el = WebDriverWait(test_driver, 10).until(
            EC.visibility_of_element_located((By.ID, "password"))
        )
        try:
            pw_el.clear()
        except Exception:
            pass
        pw_el.send_keys(password)

        try:
            test_driver.execute_script("document.getElementById('loginbtn').click();")
        except Exception:
            test_driver.find_element(By.ID, "loginbtn").click()

        time.sleep(0.6)

        def _read_danger_alert_text():
            try:
                el = test_driver.find_element(By.CSS_SELECTOR, "div.alert.alert-danger[role='alert']")
                txt = (el.text or "").strip()
                return txt or "（登入失敗，但頁面未提供錯誤文字）"
            except Exception:
                return None

        def _classify(text: str) -> str:
            t = (text or "").strip()
            if not t:
                return "未知"
            if "驗證碼" in t or "captcha" in t.lower():
                return "驗證碼"
            if "帳號" in t or "學號" in t:
                if "密碼" in t or "password" in t.lower():
                    return "帳密"
                return "帳號"
            if "密碼" in t or "password" in t.lower():
                return "密碼"
            if "鎖" in t or "lock" in t.lower():
                return "帳號鎖定"
            return "登入"

        # 須測試兩次（避免跳轉/渲染中瞬間狀態）
        alert_text = None
        for i in range(2):
            alert_text = _read_danger_alert_text()
            if alert_text:
                break
            if i == 0:
                time.sleep(0.6)

        if alert_text:
            kind = _classify(alert_text)
            set_last_login_error(kind, alert_text)
            return False, kind, alert_text

        # 沒看到 danger alert：再用「是否仍在登入頁」做最後判斷
        try:
            test_driver.find_element(By.ID, "username")
            # 仍在登入頁但沒錯誤框，多半是網路/頁面載入異常
            return False, None, None
        except Exception:
            return True, None, None
    except Exception:
        kind, text = get_last_login_error()
        return False, kind, text
    finally:
        try:
            if test_driver:
                test_driver.quit()
        except Exception:
            pass

def setup_chrome():
    """在 macOS 上檢查並自動安裝 Chrome。"""
    if platform.system() != "Darwin":
        return True  # 非 macOS 不需要執行

    def _macos_major_version() -> int:
        try:
            ver = platform.mac_ver()[0] or ""
            parts = [int(p) for p in ver.split('.') if p.isdigit()]
            return parts[0] if parts else 0
        except Exception:
            return 0

    def _macos_version_tuple():
        try:
            ver = platform.mac_ver()[0] or ""
            parts = [int(p) for p in ver.split('.') if p.isdigit()]
            while len(parts) < 3:
                parts.append(0)
            return parts[0], parts[1], parts[2]
        except Exception:
            return 0, 0, 0

    def _parse_ver_tuple(v: str):
        try:
            return tuple(int(x) for x in (v or "").split('.') if x.isdigit())
        except Exception:
            return tuple()

    def _curl_download(url: str, out_path: str, *, timeout_sec: int = 420) -> bool:
        if not url or not out_path:
            return False
        try:
            cmd = [
                "curl", "--fail", "--location",
                "--retry", "3", "--retry-delay", "2",
                "--connect-timeout", "20",
                "--max-time", str(timeout_sec),
                "-o", out_path,
                url
            ]
            subprocess.run(cmd, check=True, capture_output=True, timeout=timeout_sec + 30)
            return os.path.exists(out_path) and os.path.getsize(out_path) > 0
        except Exception:
            return False

    def _fetch_json(url: str, *, timeout_sec: int = 20):
        try:
            import urllib.request
            with urllib.request.urlopen(url, timeout=timeout_sec) as resp:
                data = resp.read()
            return json.loads(data.decode('utf-8', errors='ignore'))
        except Exception:
            return None

    def _url_exists(url: str, *, timeout_sec: int = 20) -> bool:
        """用 curl HEAD 檢查 URL 是否存在（避免下載大檔前就失敗）。"""
        if not url:
            return False
        try:
            subprocess.run(
                [
                    "curl", "--fail", "--location", "--head",
                    "--connect-timeout", "10",
                    "--max-time", str(timeout_sec),
                    url,
                ],
                check=True,
                capture_output=True,
                timeout=timeout_sec + 5,
            )
            return True
        except Exception:
            return False

    def _install_cft_by_static_versions(versions_to_try, plat: str) -> bool:
        """不依賴 JSON：用固定版本清單逐一探測並安裝 Chrome for Testing。"""
        home_dir = os.path.expanduser("~")
        root_dir = os.path.join(home_dir, "chrome-for-testing")
        os.makedirs(root_dir, exist_ok=True)
        dl_dir = os.path.join(root_dir, "_downloads")
        os.makedirs(dl_dir, exist_ok=True)

        for version_str in versions_to_try:
            version_str = (version_str or "").strip()
            if not version_str:
                continue

            chrome_url = f"https://storage.googleapis.com/chrome-for-testing-public/{version_str}/{plat}/chrome-{plat}.zip"
            driver_url = f"https://storage.googleapis.com/chrome-for-testing-public/{version_str}/{plat}/chromedriver-{plat}.zip"

            # 先 HEAD 探測，避免浪費時間
            if not (_url_exists(chrome_url) and _url_exists(driver_url)):
                continue

            print(f"{BLUE}安裝 Chrome for Testing {version_str}（{plat}，static）{RESET}")
            chrome_zip = os.path.join(dl_dir, f"chrome-{plat}-{version_str}.zip")
            driver_zip = os.path.join(dl_dir, f"chromedriver-{plat}-{version_str}.zip")

            if not os.path.exists(chrome_zip):
                if not _curl_download(chrome_url, chrome_zip, timeout_sec=900):
                    continue
            if not os.path.exists(driver_zip):
                if not _curl_download(driver_url, driver_zip, timeout_sec=300):
                    continue

            # 清掉舊資料夾（避免混到舊版本）
            try:
                import shutil
                shutil.rmtree(os.path.join(root_dir, f"chrome-{plat}"), ignore_errors=True)
                shutil.rmtree(os.path.join(root_dir, f"chromedriver-{plat}"), ignore_errors=True)
            except Exception:
                pass

            try:
                with zipfile.ZipFile(chrome_zip, 'r') as zf:
                    zf.extractall(root_dir)
                with zipfile.ZipFile(driver_zip, 'r') as zf:
                    zf.extractall(root_dir)
            except Exception:
                continue

            chrome_bin = os.path.join(root_dir, f"chrome-{plat}", "Google Chrome for Testing.app", "Contents", "MacOS", "Google Chrome for Testing")
            driver_bin = os.path.join(root_dir, f"chromedriver-{plat}", "chromedriver")
            if not os.path.exists(chrome_bin) or not os.path.exists(driver_bin):
                continue

            try:
                subprocess.run(["chmod", "+x", driver_bin], check=False, timeout=5)
            except Exception:
                pass

            os.environ["CHROME_BINARY"] = chrome_bin
            os.environ["CHROMEDRIVER_PATH"] = driver_bin
            try:
                global _cached_driver_path
                _cached_driver_path = driver_bin
            except Exception:
                pass

            if not is_chrome_usable(chrome_bin):
                continue
            reset_chromedriver_cache()
            if can_start_chrome_webdriver(chrome_bin):
                print(f"{GREEN}✓ Chrome for Testing 安裝完成：{version_str}{RESET}")
                return True

        return False

    def _install_chrome_for_testing_major_cap(max_major: int) -> bool:
        """安裝 Chrome for Testing + chromedriver 到 ~/chrome-for-testing（不需 sudo）。

        會從 known-good 版本清單中挑選「主版號 <= max_major」的最新幾個版本嘗試。
        """
        arch = (platform.machine() or "").lower()
        plat = "mac-arm64" if "arm" in arch else "mac-x64"

        kjson_url = "https://googlechromelabs.github.io/chrome-for-testing/known-good-versions-with-downloads.json"
        payload = _fetch_json(kjson_url, timeout_sec=25)
        if not payload or not isinstance(payload, dict):
            # 有些網路環境會擋 googlechromelabs.github.io；改用 static 版本清單直接探測 storage
            try:
                print(f"{YELLOW}! 無法取得 Chrome for Testing 版本清單，改用備援版本探測下載（不依賴 {kjson_url}）{RESET}")
            except Exception:
                pass

            env_list = os.environ.get("MACOS11_CHROME_VERSIONS", "").strip()
            if env_list:
                versions_to_try = [v.strip() for v in env_list.split(',') if v.strip()]
            else:
                # 由新到舊嘗試（不保證全部存在；會先 HEAD 探測）
                versions_to_try = [
                    "128.0.6613.137",
                    "127.0.6533.120",
                    "126.0.6478.127",
                    "125.0.6422.142",
                    "124.0.6367.208",
                    "123.0.6312.123",
                    "122.0.6261.130",
                    "121.0.6167.185",
                    "120.0.6099.225",
                    "119.0.6045.200",
                    "118.0.5993.71",
                    "117.0.5938.150",
                    "116.0.5845.188",
                    "115.0.5790.171",
                    "114.0.5735.199",
                    "113.0.5672.93",
                ]

            # 套用 cap
            filtered = []
            for v in versions_to_try:
                vt = _parse_ver_tuple(v)
                if vt and vt[0] <= int(max_major):
                    filtered.append(v)
            if not filtered:
                filtered = versions_to_try

            return _install_cft_by_static_versions(filtered[:8], plat)

        versions = payload.get("versions") or []
        candidates = []
        for item in versions:
            v = (item or {}).get("version")
            if not v:
                continue
            vt = _parse_ver_tuple(v)
            if not vt:
                continue
            major = vt[0]
            if major <= int(max_major):
                candidates.append((vt, item))

        if not candidates:
            print(f"{RED}X 找不到可用的 Chrome for Testing 版本（cap={max_major}）{RESET}")
            return False

        candidates.sort(key=lambda x: x[0], reverse=True)
        candidates = candidates[:3]  # 只嘗試最新 3 個，避免花太久

        home_dir = os.path.expanduser("~")
        root_dir = os.path.join(home_dir, "chrome-for-testing")
        os.makedirs(root_dir, exist_ok=True)
        dl_dir = os.path.join(root_dir, "_downloads")
        os.makedirs(dl_dir, exist_ok=True)

        for vt, item in candidates:
            version_str = (item or {}).get("version") or ""
            downloads = (item or {}).get("downloads") or {}
            chrome_list = downloads.get("chrome") or []
            driver_list = downloads.get("chromedriver") or []

            chrome_url = ""
            driver_url = ""
            for ent in chrome_list:
                if (ent or {}).get("platform") == plat:
                    chrome_url = (ent or {}).get("url") or ""
                    break
            for ent in driver_list:
                if (ent or {}).get("platform") == plat:
                    driver_url = (ent or {}).get("url") or ""
                    break

            if not chrome_url or not driver_url:
                continue

            print(f"{BLUE}安裝 Chrome for Testing {version_str}（{plat}）{RESET}")

            chrome_zip = os.path.join(dl_dir, f"chrome-{plat}-{version_str}.zip")
            driver_zip = os.path.join(dl_dir, f"chromedriver-{plat}-{version_str}.zip")

            if not os.path.exists(chrome_zip):
                if not _curl_download(chrome_url, chrome_zip, timeout_sec=600):
                    print(f"{YELLOW}! 下載 Chrome for Testing 失敗，將嘗試較舊版本{RESET}")
                    continue
            if not os.path.exists(driver_zip):
                if not _curl_download(driver_url, driver_zip, timeout_sec=300):
                    print(f"{YELLOW}! 下載 chromedriver 失敗，將嘗試較舊版本{RESET}")
                    continue

            # 清掉舊資料夾（避免混到舊版本）
            try:
                import shutil
                shutil.rmtree(os.path.join(root_dir, f"chrome-{plat}"), ignore_errors=True)
                shutil.rmtree(os.path.join(root_dir, f"chromedriver-{plat}"), ignore_errors=True)
            except Exception:
                pass

            try:
                with zipfile.ZipFile(chrome_zip, 'r') as zf:
                    zf.extractall(root_dir)
                with zipfile.ZipFile(driver_zip, 'r') as zf:
                    zf.extractall(root_dir)
            except Exception as e:
                print(f"{YELLOW}! 解壓失敗，將嘗試較舊版本：{e}{RESET}")
                continue

            chrome_bin = os.path.join(root_dir, f"chrome-{plat}", "Google Chrome for Testing.app", "Contents", "MacOS", "Google Chrome for Testing")
            driver_bin = os.path.join(root_dir, f"chromedriver-{plat}", "chromedriver")

            if not os.path.exists(chrome_bin) or not os.path.exists(driver_bin):
                print(f"{YELLOW}! 安裝結果不完整，將嘗試較舊版本{RESET}")
                continue

            # 確保 chromedriver 可執行
            try:
                subprocess.run(["chmod", "+x", driver_bin], check=False, timeout=5)
            except Exception:
                pass

            # 設定環境變數，讓後續 Selenium 使用這組 Chrome/Driver
            os.environ["CHROME_BINARY"] = chrome_bin
            os.environ["CHROMEDRIVER_PATH"] = driver_bin
            try:
                global _cached_driver_path
                _cached_driver_path = driver_bin
            except Exception:
                pass

            # 驗證：至少能跑 --version，並嘗試建立 WebDriver session
            if not is_chrome_usable(chrome_bin):
                print(f"{YELLOW}! Chrome 無法啟動（--version 失敗），將嘗試較舊版本{RESET}")
                continue
            reset_chromedriver_cache()
            if can_start_chrome_webdriver(chrome_bin):
                print(f"{GREEN}✓ Chrome for Testing 安裝完成：{version_str}{RESET}")
                return True
            print(f"{YELLOW}! Chrome WebDriver 建立失敗，將嘗試較舊版本{RESET}")

        return False

    def is_chrome_usable(binary_path):
        """檢查 Chrome 執行檔是否可正常啟動。"""
        if not binary_path or not os.path.exists(binary_path):
            return False
        try:
            probe = subprocess.run(
                [binary_path, "--version"],
                capture_output=True,
                timeout=15,
                check=False
            )
            output = (probe.stdout or b"").decode("utf-8", errors="ignore")
            return probe.returncode == 0 and "Chrome" in output
        except Exception:
            return False

    def can_start_chrome_webdriver(binary_path):
        """實際驗證 Chrome + ChromeDriver 能否建立 session。"""
        if not binary_path:
            return False
        test_options = Options()
        if os.name == 'nt':
            test_options.add_argument("--headless=new")
        else:
            test_options.add_argument("--headless")
        test_options.add_argument("--disable-gpu")
        test_options.add_argument("--no-sandbox")
        test_options.add_argument("--disable-dev-shm-usage")
        test_options.binary_location = binary_path

        tmp_driver = None
        try:
            driver_path = get_chrome_driver_path()
            if driver_path:
                service = Service(driver_path, log_path=os.devnull)
                tmp_driver = webdriver.Chrome(options=test_options, service=service)
            else:
                tmp_driver = webdriver.Chrome(options=test_options)
            tmp_driver.get("about:blank")
            return True
        except Exception:
            return False
        finally:
            try:
                if tmp_driver:
                    tmp_driver.quit()
            except Exception:
                pass

    def reset_chromedriver_cache():
        """清除本機 chromedriver 快取，避免沿用錯誤版本。"""
        global _cached_driver_path
        _cached_driver_path = None
        try:
            cache_dir = os.path.join(os.path.expanduser("~"), ".wdm", "drivers", "chromedriver")
            if os.path.exists(cache_dir):
                import shutil
                shutil.rmtree(cache_dir, ignore_errors=True)
        except Exception:
            pass

    def wait_for_chrome_ready(max_wait_seconds=45, interval_seconds=3):
        """安裝後給系統一些時間完成索引/權限套用，再重試啟動檢查。"""
        deadline = time.time() + max_wait_seconds
        attempt = 0
        while time.time() < deadline:
            attempt += 1
            candidate = get_chrome_binary_path()
            if candidate and is_chrome_usable(candidate):
                reset_chromedriver_cache()
                if can_start_chrome_webdriver(candidate):
                    return True
            remaining = int(max(0, deadline - time.time()))
            print(f"{YELLOW}! Chrome 尚未就緒，{interval_seconds} 秒後重試（剩餘約 {remaining} 秒）...{RESET}")
            time.sleep(interval_seconds)
        return False
    
    # 先檢查 Chrome 是否存在且可用；若存在但損壞，仍需重裝。
    # 注意：在 macOS 11.x 上，/Applications 的 system Chrome 可能太新，呼叫一次就 crash 跳視窗。
    # 因此預設不主動 probe system Chrome，改優先走 Chrome for Testing（除非 PREFER_SYSTEM_CHROME=1）。
    mac_major = _macos_major_version()
    existing_binary = get_chrome_binary_path()
    system_chrome_bin = "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome"
    skip_probe_system_chrome = (
        mac_major and mac_major <= 11
        and existing_binary == system_chrome_bin
        and os.environ.get("PREFER_SYSTEM_CHROME", "0") != "1"
    )
    if not skip_probe_system_chrome:
        if existing_binary and is_chrome_usable(existing_binary):
            if can_start_chrome_webdriver(existing_binary):
                return True
            print(f"{YELLOW}! Chrome 可啟動，但 WebDriver 建立失敗，將自動修復{RESET}")
            reset_chromedriver_cache()
        if existing_binary and not is_chrome_usable(existing_binary):
            print(f"{YELLOW}! 偵測到 Chrome 但無法啟動，將自動重新下載與修復{RESET}")
    else:
        print(f"{YELLOW}! macOS 11.x 預設略過 system Chrome 探測（避免因版本過新而 crash），改用 Chrome for Testing{RESET}")

    # macOS 11.x（Big Sur）：不要裝最新 stable（可能不相容），改裝 Chrome for Testing（cap 主版號）。
    if mac_major and mac_major <= 11:
        cap_env = (os.environ.get("MACOS11_CHROME_MAJOR_CAP", "") or "").strip()
        if cap_env:
            cap = int(cap_env)
        else:
            # Big Sur 11.x（尤其 11.5.x）對新版本 Chrome 需求的系統 framework 可能不足，
            # 預設把 cap 壓低到較舊版本避免下載後必 crash。
            cap = 114

        mac_tuple = _macos_version_tuple()
        print(f"{YELLOW}! 偵測到 macOS {platform.mac_ver()[0]}，將下載相容的 Chrome for Testing（主版號 <= {cap}）{RESET}")
        if mac_tuple[0] == 11 and mac_tuple[1] <= 5 and not cap_env:
            print(f"{YELLOW}! 你的 macOS 版本偏舊（11.{mac_tuple[1]}），若仍遇到 Chrome crash，建議先更新到 11.7.x（或設定 MACOS11_CHROME_MAJOR_CAP 更低）{RESET}")
        ok = _install_chrome_for_testing_major_cap(cap)
        if ok:
            return True
        print(f"{RED}X 無法安裝相容的 Chrome（macOS 11.x）。建議：升級 macOS 或自行安裝舊版 Chrome。{RESET}")
        return False
    
    # 定義要執行的指令列表
    commands = [
        {
            "name": "下載 Chrome DMG",
            "cmd": 'curl --fail --location --retry 3 --retry-delay 2 --connect-timeout 20 --max-time 420 -o /tmp/googlechrome.dmg https://dl.google.com/chrome/mac/universal/stable/GGRO/googlechrome.dmg',
            "timeout": 480,
            "retries": 2,
        },
        {"name": "驗證 DMG 完整性", "cmd": 'hdiutil verify /tmp/googlechrome.dmg', "timeout": 180, "retries": 1},
        {"name": "移除舊版 Chrome", "cmd": 'sudo rm -rf "/Applications/Google Chrome.app"', "timeout": 120, "retries": 1},
        {"name": "掛載 Chrome DMG", "cmd": 'hdiutil attach /tmp/googlechrome.dmg -quiet', "timeout": 180, "retries": 1},
        {"name": "複製 Chrome 到應用程式資料夾", "cmd": 'sudo ditto "/Volumes/Google Chrome/Google Chrome.app" "/Applications/Google Chrome.app"', "timeout": 240, "retries": 1},
        {"name": "解除 macOS 安全性鎖定", "cmd": 'sudo xattr -cr "/Applications/Google Chrome.app"', "timeout": 120, "retries": 1},
    ]

    for step in commands:
        step_name = step["name"]
        cmd = step["cmd"]
        timeout_sec = step.get("timeout", 120)
        max_retries = max(1, int(step.get("retries", 1)))

        for attempt in range(1, max_retries + 1):
            try:
                print(f"{BLUE}{step_name}{RESET}", flush=True)
                subprocess.run(cmd, shell=True, check=True, capture_output=True, timeout=timeout_sec)
                break
            except subprocess.TimeoutExpired:
                if attempt < max_retries:
                    print(f"{YELLOW}! {step_name} 超時，準備重試（{attempt}/{max_retries}）{RESET}")
                    time.sleep(2)
                    continue
                print(f"{RED}✗ 超時{RESET}")
                print(f"   {RED}步驟：{step_name}（timeout={timeout_sec}s）{RESET}")
                return False
            except subprocess.CalledProcessError as e:
                if attempt < max_retries:
                    print(f"{YELLOW}! {step_name} 執行失敗，準備重試（{attempt}/{max_retries}）{RESET}")
                    time.sleep(2)
                    continue
                print(f"{RED}✗ 失敗{RESET}")
                print(f"   {RED}步驟：{step_name}{RESET}")
                print(f"   {RED}錯誤：{e.stderr.decode('utf-8', errors='ignore')}{RESET}")
                return False
            except Exception as e:
                print(f"{RED}✗ 異常{RESET}")
                print(f"   {RED}步驟：{step_name}{RESET}")
                print(f"   {RED}詳情：{e}{RESET}")
                return False

    print(f"{GREEN}Chrome 環境設置完成{RESET}")


    repaired_binary = get_chrome_binary_path()
    if not repaired_binary:
        print(f"{RED}X 找不到 Chrome 執行檔，請檢查安裝結果{RESET}")
        return False

    # 安裝後先等待一段時間，避免剛安裝完成時立即驗證失敗。
    if not wait_for_chrome_ready(max_wait_seconds=45, interval_seconds=3):
        print(f"{RED}X 等待後仍無法啟動 Chrome/ChromeDriver，將先退出{RESET}")
        return False

    return True

def append_marker_to_password_file(marker):
    """將標記追加到 password.txt 檔尾；若已存在則不重複寫入。"""
    if not marker:
        return
    normalized = marker.strip().lower()
    if not normalized:
        return

    lines = []
    raw_text = ""
    if os.path.exists(PASSWORD_FILE):
        with open(PASSWORD_FILE, 'r', encoding='utf-8') as f:
            raw_text = f.read()
            lines = raw_text.splitlines()

    if any(line.strip().lower() == normalized for line in lines):
        return

    with open(PASSWORD_FILE, 'a', encoding='utf-8') as f:
        if raw_text and not raw_text.endswith("\n"):
            f.write("\n")
        f.write(f"{marker.strip()}\n")

# 讀取帳號密碼
def load_credentials():
    """從 password.txt 讀取帳號密碼，如果不存在則讓使用者輸入並創建"""
    global IS_FIRST_TIME, AUTO_OPEN_NEW_ACTIVITY_FOLDERS, IS_BUILD_ENV  # 使用全域變數來追蹤狀態
    
    try:
        if os.path.exists(PASSWORD_FILE):
            with open(PASSWORD_FILE, 'r', encoding='utf-8') as f:
                lines = f.read().splitlines()
                if len(lines) >= 2:
                    username = lines[0].strip()
                    password = lines[1].strip()
                    if not username or not password:
                        print(f"\n{RED}{'='*60}{RESET}")
                        print(f"{RED}X 錯誤：password.txt 檔案內容不完整{RESET}")
                        input()
                        sys.exit(1)
                    markers = {line.strip().lower() for line in lines[2:] if line.strip()}
                    has_builderror = any(line.strip().lower() == BUILDERROR_MARKER for line in lines)
                    has_nonpop = any(m in markers for m in POPUP_DISABLE_MARKERS)
                    if has_builderror:
                        IS_FIRST_TIME = True
                        IS_BUILD_ENV = True
                    else:
                        IS_FIRST_TIME = False
                        IS_BUILD_ENV = False
                    AUTO_OPEN_NEW_ACTIVITY_FOLDERS = not has_nonpop
                    return username, password
                else:
                    print(f"\n{RED}{'='*60}{RESET}")
                    print(f"{RED}X 錯誤：password.txt 格式錯誤{RESET}")
                    print(f"{RED}{'='*60}{RESET}")
                    print(f"\n{YELLOW}檢查到檔案，但內容不夠兩行{RESET}")
                    print(f"\n📁 檔案位置：{PASSWORD_FILE}")
                    print(f"\n✅ 正確格式：")
                    print(f"   第 1 行：你的學號")
                    print(f"   第 2 行：你的密碼")
                    print(f"\n按 Enter 鍵離開...")
                    input()
                    sys.exit(1)
        else:
            # 第一次使用，讓使用者輸入帳號密碼
            IS_FIRST_TIME = True
            AUTO_OPEN_NEW_ACTIVITY_FOLDERS = True
            IS_BUILD_ENV = True
            print(f"{YELLOW}請輸入你的 Moodle 帳號密碼：{RESET}")
            
            while True:  # 持續輸入直到登入成功
                username = input(f"{PINK}學號：{RESET}").strip()
                while not username:
                    print(f"{RED}學號不能為空，請重新輸入{RESET}")
                    username = input(f"{PINK}學號：{RESET}").strip()
                
                password = get_password_input(f"{PINK}密碼：{RESET}")
                while not password:
                    print(f"{RED}密碼不能為空，請重新輸入{RESET}")
                    password = get_password_input(f"{PINK}密碼：{RESET}")
                
                # 測試登入
 
                ok, kind, text = test_login(username, password)

                if ok:

                    break
                else:
                    # 只有真的出現 danger alert 才輸出明確錯誤；否則避免誤判成帳密錯
                    if should_emit_login_failure_messages():
                        if text:
                            print(f"\n{RED}X 登入失敗（{kind or '未知'}）：{text}{RESET}")
                        else:
                            print(f"\n{RED}X 登入失敗：未偵測到錯誤提示（可能網路/系統異常），請重試{RESET}")
                    continue
            
            # 登入成功，創建 password.txt 檔案
            try:
                with open(PASSWORD_FILE, 'w', encoding='utf-8') as f:
                    f.write(f"{username}\n{password}\n")
                append_marker_to_password_file(BUILDERROR_MARKER)
                print(f"\n{BLUE}建置環境中{RESET}")
                return username, password
            except Exception as e:
                print(f"\n{RED}X 無法創建密碼檔案：{e}{RESET}")
                print(f"按 Enter 鍵離開...")
                input()
                sys.exit(1)
                
    except Exception as e:
        print(f"\n{RED}{'='*60}{RESET}")
        print(f"{RED}X 錯誤：讀取 password.txt 失敗{RESET}")
        print(f"{RED}{'='*60}{RESET}")
        print(f"\n錯誤訊息：{e}")
        print(f"\n可能原因：")
        print(f"   - 檔案正在被其他程式使用")
        print(f"   - 檔案編碼問題（請使用 UTF-8 編碼儲存）")
        print(f"   - 檔案讀取權限不足")
        print(f"\n按 Enter 鍵離開...")
        input()
        sys.exit(1)

# macOS 上自動檢查和設置 Chrome
if platform.system() == "Darwin":
    if not setup_chrome():
        sys.exit(1)

# 載入帳號密碼
USERNAME, PASSWORD = load_credentials()

def clear_builderror_marker(username, password):
    """首次建置完成後，移除第三行 builderror 標記。"""
    try:
        kept_markers = []
        if os.path.exists(PASSWORD_FILE):
            with open(PASSWORD_FILE, 'r', encoding='utf-8') as f:
                lines = f.read().splitlines()
            for line in lines[2:]:
                marker = line.strip()
                if marker and marker.lower() != BUILDERROR_MARKER:
                    kept_markers.append(marker)
        with open(PASSWORD_FILE, 'w', encoding='utf-8') as f:
            f.write(f"{username}\n{password}\n")
            for marker in kept_markers:
                f.write(f"{marker}\n")
    except Exception:
        pass

def is_activity(text):
    return not (text.startswith("第") and text.endswith("週")) and text.strip()

def clean_activity_name(text):
    """清理活動名稱中的換行字元和亂碼"""
    if not text:
        return text
    
    # 移除各種換行字元和不需要的空白字元
    cleaned = text.replace('\n', ' ').replace('\r', ' ').replace('\t', ' ')
    
    # 移除多餘的空格
    cleaned = ' '.join(cleaned.split())
    
    # 移除常見的亂碼後綴
    if '\n作業' in cleaned:
        cleaned = cleaned.replace('\n作業', '')
    
    # 再次清理空白字元
    cleaned = cleaned.strip()
    
    return cleaned

def extract_activity_assets_from_course_page(driver, activity_name):
    """使用 JavaScript 直接從課程頁取得活動資源，避免 stale element。"""
    script = """
    const targetName = arguments[0] || '';
    const normalize = (s) => (s || '').replace(/\s+/g, ' ').trim();
    const wanted = normalize(targetName);

    const items = Array.from(document.querySelectorAll('div.activity-item'));
    for (const item of items) {
        let currentName = item.getAttribute('data-activityname') || '';
        if (!currentName) {
            const instance = item.querySelector('span.instancename');
            currentName = instance ? instance.textContent : '';
        }
        currentName = normalize(currentName);
        if (currentName !== wanted) {
            continue;
        }

        const anchors = Array.from(item.querySelectorAll("a[href*='pluginfile.php']"))
            .map(a => a.href)
            .filter(Boolean);
        const images = Array.from(item.querySelectorAll("img[src*='pluginfile.php']"))
            .map(img => img.src)
            .filter(Boolean);
        const externals = Array.from(item.querySelectorAll("a[href]"))
            .map(a => a.href)
            .filter(href => href && !href.includes('pluginfile.php') && !href.startsWith('https://elearningv4.nuk.edu.tw'));

        return {
            found: true,
            anchors,
            images,
            externals
        };
    }

    return {
        found: false,
        anchors: [],
        images: [],
        externals: []
    };
    """
    try:
        result = driver.execute_script(script, activity_name)
        if isinstance(result, dict):
            result.setdefault('found', False)
            result.setdefault('anchors', [])
            result.setdefault('images', [])
            result.setdefault('externals', [])
            return result
    except Exception:
        pass
    return {'found': False, 'anchors': [], 'images': [], 'externals': []}

def load_submitted_assignments():
    """讀取已繳交作業記錄"""
    try:
        if os.path.exists(SUBMITTED_ASSIGNMENTS_FILE):
            with open(SUBMITTED_ASSIGNMENTS_FILE, 'r', encoding='utf-8') as f:
                return json.load(f)
    except:
        pass
    return {}

def save_submitted_assignments(submitted_dict):
    """儲存已繳交作業記錄"""
    try:
        with open(SUBMITTED_ASSIGNMENTS_FILE, 'w', encoding='utf-8') as f:
            json.dump(submitted_dict, f, ensure_ascii=False, indent=2)
    except Exception as e:
        print(f"{RED}警告：無法儲存已繳交作業記錄: {e}{RESET}")

def build_assignment_key(course_name, act_href):
    """用 course + assign id/url 產生穩定 key。"""
    assign_id = None
    try:
        q = parse_qs(urlparse(act_href).query)
        assign_id = (q.get('id') or [None])[0]
    except Exception:
        assign_id = None
    if assign_id:
        # id 在 Moodle 通常全站唯一，比課程名稱更穩（避免課程名變動導致快取對不上）
        return f"id:{assign_id}"
    # 後備：仍保留 course_name，避免非 assign id 類型連結互撞
    return f"{course_name}||url:{act_href}"

def has_submitted_record(records, course_name, act_href, assignment_key):
    if assignment_key in (records or {}):
        return True

    # 以 id 為主做比對（即使 course_name 不同也能對上）
    aid = None
    try:
        q = parse_qs(urlparse(act_href).query)
        aid = (q.get('id') or [None])[0]
        if aid and f"id:{aid}" in (records or {}):
            return True
    except Exception:
        pass

    for rec in (records or {}).values():
        if isinstance(rec, dict):
            # 舊資料：用 course+url
            if rec.get('course') == course_name and rec.get('url') == act_href:
                return True
            # 新資料：用 id
            try:
                rurl = rec.get('url') or ''
                q2 = parse_qs(urlparse(rurl).query)
                rid = (q2.get('id') or [None])[0]
                if aid and rid and str(aid) == str(rid):
                    return True
            except Exception:
                pass
            continue
    return False

def parse_due_datetime(due_date_str):
    """解析 '2026年 4月 2日 ... 23:59' 類型時間；失敗回傳 None。"""
    import re
    import datetime
    if not due_date_str:
        return None
    m = re.search(r"(\d{4})年\s*(\d{1,2})月\s*(\d{1,2})日.*?(\d{1,2}):(\d{2})", due_date_str)
    if not m:
        return None
    y, mo, d, h, mi = map(int, m.groups())
    return datetime.datetime(y, mo, d, h, mi)

def extract_due_date_str(driver_obj):
    """只取『到期』那一行，避免整頁文字污染。"""
    import re
    # 優先從 activity-dates 擷取
    try:
        due = driver_obj.execute_script("""
            const root = document.querySelector("div.activity-dates[data-region='activity-dates']")
                      || document.querySelector('div.activity-dates');
            if (!root) return '';
            const txt = (root.innerText || '').replace(/\r/g, '');
            const m = txt.match(/到期：\s*([^\n]+)/);
            return m ? m[1].trim() : '';
        """)
        if isinstance(due, str) and due.strip():
            return due.strip()
    except Exception:
        pass

    # 後備：找包含『到期』的區塊
    try:
        els = driver_obj.find_elements(By.XPATH, "//div[contains(., '到期：')] | //th[contains(text(), '截止時間')]/following-sibling::td")
        for el in els:
            text = (el.text or '').strip()
            m = re.search(r"到期：\s*([^\n\r]+)", text)
            if m:
                return m.group(1).strip()
            if "年" in text and "月" in text:
                return text.split('\n')[0].strip()
    except Exception:
        pass
    return ""

def load_pending_assignments():
    """讀取待繳作業快取文字檔。

    格式：course\tname\tdue_date_str\turl
    """
    pending = {}
    try:
        if not os.path.exists(PENDING_ASSIGNMENTS_FILE):
            return {}
        with open(PENDING_ASSIGNMENTS_FILE, 'r', encoding='utf-8') as f:
            for raw in f:
                line = (raw or '').rstrip('\n')
                if not line.strip():
                    continue
                parts = line.split('\t')
                if len(parts) < 4:
                    continue
                course, name, due_date_str, url = parts[0], parts[1], parts[2], parts[3]
                key = build_assignment_key(course, url)
                pending[key] = {
                    'course': course,
                    'name': name,
                    'url': url,
                    'due_date_str': due_date_str,
                    'due_date_obj': parse_due_datetime(due_date_str) or __import__('datetime').datetime.max,
                }
    except Exception:
        return {}
    return pending

def save_pending_assignments(pending_dict):
    try:
        lines = []
        for key, item in (pending_dict or {}).items():
            course = (item.get('course') or '').strip()
            name = (item.get('name') or '').strip()
            due = (item.get('due_date_str') or '').strip()
            url = (item.get('url') or '').strip()
            if not course or not name or not url:
                continue
            lines.append(f"{course}\t{name}\t{due}\t{url}")
        with open(PENDING_ASSIGNMENTS_FILE, 'w', encoding='utf-8') as f:
            f.write("\n".join(lines))
    except Exception:
        pass

def is_expired_assignment(item, now_dt):
    try:
        import datetime
        due_dt = item.get('due_date_obj')
        if not due_dt or due_dt == datetime.datetime.max:
            return False
        return due_dt < now_dt
    except Exception:
        return False

def check_assignments_inline(driver_obj, assignments, *, submitted_assignments, pending_cache, limit=ASSIGNMENT_CHECK_LIMIT):
    """使用主 driver 限量檢查作業狀態（不額外開新分頁/新 driver）。

    - 已在 submitted / pending_cache 內的會跳過
    - 檢查到未繳交：寫入 pending_cache
    - 檢查到已繳交：寫入 submitted_assignments 並從 pending_cache 移除
    """
    import datetime
    newly_submitted = {}
    checked = 0
    now_dt = datetime.datetime.now()

    # 去重 + 依原順序
    seen_keys = set()
    candidates = []
    for a in assignments or []:
        course = a.get('course') or ''
        url = a.get('url') or ''
        name = a.get('name') or ''
        if not course or not url or not name:
            continue
        key = build_assignment_key(course, url)
        if key in seen_keys:
            continue
        seen_keys.add(key)
        if key in (pending_cache or {}):
            continue
        if has_submitted_record(submitted_assignments, course, url, key):
            continue
        candidates.append((key, a))

    for key, a in candidates:
        if checked >= max(0, int(limit or 0)):
            break
        try:
            driver_obj.get(a['url'])
            time.sleep(0.2)

            status_text = ""
            try:
                status_cell = driver_obj.find_element(By.XPATH, "//th[contains(text(), '繳交狀態')]/following-sibling::td")
                status_text = (status_cell.text or '').strip()
            except Exception:
                status_text = ""

            # 已繳交判斷
            if status_text and any(k in status_text for k in ["已繳交", "已提交", "已送出"]):
                newly_submitted[key] = {
                    'course': a['course'],
                    'name': a['name'],
                    'url': a['url'],
                    'assignment_key': key,
                    'checked_date': time.strftime('%Y-%m-%d %H:%M:%S')
                }
                if key in pending_cache:
                    pending_cache.pop(key, None)
                checked += 1
                continue

            # 是否可繳交（或尚無繳交）
            has_submit_button = False
            try:
                submit_buttons = driver_obj.find_elements(By.XPATH, "//button[contains(text(), '繳交作業')]")
                if submit_buttons:
                    has_submit_button = True
            except Exception:
                pass

            if not has_submit_button:
                if status_text and ("尚無任何作業繳交" in status_text or "目前尚無" in status_text):
                    has_submit_button = True

            if has_submit_button:
                due_date_str = extract_due_date_str(driver_obj)
                due_dt = parse_due_datetime(due_date_str) or datetime.datetime.max
                item = {
                    'course': a['course'],
                    'name': a['name'],
                    'url': a['url'],
                    'due_date_str': due_date_str,
                    'due_date_obj': due_dt,
                }
                # 若已繳交或已過期，皆記錄為已處理（寫入 submitted_assignments 以後不要再點進去了）
                if is_expired_assignment(item, now_dt):
                    newly_submitted[key] = {
                        'course': a['course'],
                        'name': a['name'],
                        'url': a['url'],
                        'assignment_key': key,
                        'checked_date': time.strftime('%Y-%m-%d %H:%M:%S'),
                        'status': 'expired'
                    }
                    if key in pending_cache:
                        pending_cache.pop(key, None)
                else:
                    pending_cache[key] = item
            else:
                # 沒看到繳交按鈕也沒明確狀態：當作已繳交/無需繳交，避免一直卡在 pending
                newly_submitted[key] = {
                    'course': a['course'],
                    'name': a['name'],
                    'url': a['url'],
                    'assignment_key': key,
                    'checked_date': time.strftime('%Y-%m-%d %H:%M:%S')
                }
                if key in pending_cache:
                    pending_cache.pop(key, None)

            checked += 1
        except Exception:
            checked += 1
            continue

    if newly_submitted:
        submitted_assignments.update(newly_submitted)
    return submitted_assignments, pending_cache


def recheck_pending_assignments(driver_obj, *, pending_cache, submitted_assignments, limit=None):
    """對少量「已在 pending_cache」的作業做再驗證。

    原本 pending_cache 內的作業為了加速會被跳過檢查；若使用者已在網頁上完成繳交，
    可能會暫時殘留在未繳交清單。這裡用小額度重新檢查，將已繳交者移出 pending。
    """
    # limit=None 代表全部檢查（最可靠，避免已繳交仍殘留在未繳清單）
    if limit is None:
        limit = len(pending_cache or {})
    try:
        limit = max(0, int(limit or 0))
    except Exception:
        limit = 0
    if limit <= 0:
        return submitted_assignments, pending_cache

    import datetime
    now_dt = datetime.datetime.now()

    items = list((pending_cache or {}).items())
    # 越接近截止越優先檢查
    items.sort(key=lambda kv: (kv[1].get('due_date_obj') or datetime.datetime.max))

    checked = 0
    for key, item in items:
        if checked >= limit:
            break
        try:
            url = item.get('url')
            course = item.get('course') or ''
            name = item.get('name') or ''
            if not url:
                continue

            # 若已在 submitted 記錄中，直接移除
            stable_key = build_assignment_key(course, url)
            if has_submitted_record(submitted_assignments, course, url, stable_key):
                pending_cache.pop(key, None)
                continue

            driver_obj.get(url)
            time.sleep(0.15)

            status_text = ""
            try:
                status_cell = driver_obj.find_element(By.XPATH, "//th[contains(text(), '繳交狀態')]/following-sibling::td")
                status_text = (status_cell.text or '').strip()
            except Exception:
                status_text = ""

            if status_text and any(k in status_text for k in ["已繳交", "已提交", "已送出"]):
                submitted_assignments[stable_key] = {
                    'course': course,
                    'name': name,
                    'url': url,
                    'assignment_key': stable_key,
                    'checked_date': time.strftime('%Y-%m-%d %H:%M:%S')
                }
                pending_cache.pop(key, None)
            else:
                # 若抓不到狀態文字，改用「是否還能繳交」判斷：
                # 找不到「繳交作業」按鈕，通常代表已繳交/無需繳交，避免一直殘留在未繳清單。
                try:
                    has_submit_button = False
                    try:
                        submit_buttons = driver_obj.find_elements(By.XPATH, "//button[contains(text(), '繳交作業')]")
                        if submit_buttons:
                            has_submit_button = True
                    except Exception:
                        pass

                    if (not has_submit_button) and status_text and (
                        "尚無任何作業繳交" in status_text or "目前尚無" in status_text
                    ):
                        has_submit_button = True

                    if not has_submit_button:
                        submitted_assignments[stable_key] = {
                            'course': course,
                            'name': name,
                            'url': url,
                            'assignment_key': stable_key,
                            'checked_date': time.strftime('%Y-%m-%d %H:%M:%S')
                        }
                        pending_cache.pop(key, None)
                    else:
                        # 若已過期就移除（寫入 submitted_assignments 以避免殘留與重複驗證）
                        if is_expired_assignment(item, now_dt):
                            submitted_assignments[stable_key] = {
                                'course': course,
                                'name': name,
                                'url': url,
                                'assignment_key': stable_key,
                                'checked_date': time.strftime('%Y-%m-%d %H:%M:%S'),
                                'status': 'expired'
                            }
                            pending_cache.pop(key, None)
                except Exception:
                    # 若已過期就移除（寫入 submitted_assignments 以避免殘留與重複驗證）
                    if is_expired_assignment(item, now_dt):
                        submitted_assignments[stable_key] = {
                            'course': course,
                            'name': name,
                            'url': url,
                            'assignment_key': stable_key,
                            'checked_date': time.strftime('%Y-%m-%d %H:%M:%S'),
                            'status': 'expired'
                        }
                        pending_cache.pop(key, None)

            checked += 1
        except Exception:
            checked += 1
            continue

    return submitted_assignments, pending_cache


def check_assignments_background_early(course_hrefs_snapshot, username, password):
    """用獨立 WebDriver 在程式一開始就背景檢查未繳交作業。

    重要：不要共用主流程的 driver/all_tabs，避免多執行緒搶同一個 WebDriver 造成不穩定。
    """
    global empty_assignments, assignment_check_completed

    import datetime
    import re

    submitted_assignments = load_submitted_assignments()
    empty_assignments = []
    newly_submitted = {}

    assignment_check_completed = False
    bg_driver = None

    try:
        if os.name == 'nt':
            _kill_existing_chrome_processes()

        bg_options = Options()
        bg_options.add_argument("--headless")
        bg_options.add_argument("--disable-gpu")
        bg_options.add_argument("--log-level=3")
        bg_options.add_argument("--window-size=1920,1080")
        bg_options.add_argument("--disable-blink-features=AutomationControlled")
        bg_options.add_argument("--disable-background-timer-throttling")
        bg_options.add_argument("--disable-backgrounding-occluded-windows")
        bg_options.add_argument("--disable-renderer-backgrounding")
        bg_options.add_argument("--memory-pressure-off")
        bg_options.add_experimental_option("excludeSwitches", ["enable-logging", "enable-automation"])
        bg_options.page_load_strategy = 'eager'
        apply_chrome_binary_option(bg_options)

        bg_prefs = {
            "download.default_directory": BASE_DOWNLOAD_DIR,
            "plugins.always_open_pdf_externally": True,
            "profile.default_content_setting_values.images": 2,
            "profile.default_content_setting_values.media_stream": 2,
            "profile.managed_default_content_settings.images": 2,
            "profile.default_content_settings.popups": 0,
        }
        bg_options.add_experimental_option("prefs", bg_prefs)

        bg_driver = create_webdriver(bg_options, hide_windows_console=True)

        # 登入（避免依賴主流程的 driver）
        if not ensure_logged_in_retry_once(bg_driver, username, password, silent=True, max_retries=2):
            raise RuntimeError("背景作業檢查登入未成功")

        # 開始檢查課程作業
        for course_href in list(course_hrefs_snapshot or []):
            try:
                bg_driver.get(course_href)
                time.sleep(0.2)

                try:
                    course_name = bg_driver.find_element(By.CSS_SELECTOR, "h1.h2").text
                except Exception:
                    course_name = ""

                assignments_info = bg_driver.execute_script("""
                    return Array.from(document.querySelectorAll("div.activity-item a.aalink[href*='mod/assign/view.php']"))
                        .map(a => {
                            const span = a.querySelector('span.instancename');
                            return {
                                href: a.getAttribute('href') || a.href || '',
                                name: span ? span.textContent : ''
                            };
                        })
                        .filter(x => x.href && x.name);
                """) or []

                for assign_info in assignments_info:
                    try:
                        act_href = assign_info.get('href', '')
                        act_name = clean_activity_name(assign_info.get('name', ''))
                        if not act_href or not act_name:
                            continue

                        assignment_key = build_assignment_key(course_name, act_href)
                        if has_submitted_record(submitted_assignments, course_name, act_href, assignment_key):
                            continue
                        if has_submitted_record(newly_submitted, course_name, act_href, assignment_key):
                            continue

                        bg_driver.get(act_href)
                        bg_driver.execute_script("document.body.style.zoom='80%'")

                        has_submit_button = False
                        try:
                            WebDriverWait(bg_driver, 1).until(
                                EC.presence_of_element_located((By.XPATH, "//button[contains(text(), '繳交作業')] | //th[contains(text(), '繳交狀態')]"))
                            )
                        except Exception:
                            pass

                        try:
                            submit_buttons = bg_driver.find_elements(By.XPATH, "//button[contains(text(), '繳交作業')]")
                            if submit_buttons:
                                has_submit_button = True
                        except Exception:
                            pass

                        if not has_submit_button:
                            try:
                                status_cell = bg_driver.find_element(By.XPATH, "//th[contains(text(), '繳交狀態')]/following-sibling::td")
                                status_text = status_cell.text
                                if "尚無任何作業繳交" in status_text or "目前尚無" in status_text:
                                    has_submit_button = True
                            except Exception:
                                pass

                        if has_submit_button:
                            due_date_str = extract_due_date_str(bg_driver)
                            due_dt = parse_due_datetime(due_date_str) or datetime.datetime.max
                            empty_assignments.append({
                                'name': act_name,
                                'course': course_name,
                                'url': act_href,
                                'due_date_str': due_date_str,
                                'due_date_obj': due_dt,
                            })
                        else:
                            newly_submitted[assignment_key] = {
                                'course': course_name,
                                'name': act_name,
                                'url': act_href,
                                'assignment_key': assignment_key,
                                'checked_date': time.strftime('%Y-%m-%d %H:%M:%S')
                            }
                    except Exception:
                        continue
            except Exception:
                continue

        if newly_submitted:
            submitted_assignments.update(newly_submitted)
        save_submitted_assignments(submitted_assignments)
    finally:
        try:
            if bg_driver is not None:
                bg_driver.quit()
        except Exception:
            pass
        assignment_check_completed = True

def open_folder(path):
    """開啟檔案總管到指定資料夾"""
    try:
        if platform.system() == "Windows":
            # Windows 會根據系統設定自動決定開新視窗或新分頁
            os.startfile(path)
        elif platform.system() == "Darwin":  # macOS
            subprocess.run(["open", path])
        else:  # Linux
            subprocess.run(["xdg-open", path])
    except Exception as e:
        print(f"X 無法開啟資料夾: {e}")

def close_macos_terminal_and_exit(code=0):
    """在 macOS 嘗試關閉 Terminal 視窗，並強制結束行程。"""
    if platform.system() == "Darwin":
        try:
            subprocess.Popen(
                [
                    "osascript",
                    "-e",
                    'tell application "Terminal" to close (every window whose frontmost is true)'
                ],
                stdout=subprocess.DEVNULL,
                stderr=subprocess.DEVNULL
            )
        except Exception:
            pass

        # macOS：os._exit 會跳過 atexit，故先清理暫存下載資料夾
        try:
            _cleanup_temp_dl_dirs()
        except Exception:
            pass
        os._exit(code)

    # 非 macOS：避免 os._exit（會讓 Selenium/Chrome 還活著，導致 _temp_dl 刪不掉）
    try:
        # 關閉主 driver
        _drv = globals().get('driver')
        if _drv is not None:
            try:
                _drv.quit()
            except Exception:
                pass
    except Exception:
        pass

    # 關閉模擬器 driver + server（os._exit 不會跑 atexit）
    try:
        global _SIM_SERVER, _SIM_DRIVER
        try:
            if _SIM_DRIVER is not None:
                _SIM_DRIVER.quit()
        except Exception:
            pass
        try:
            if _SIM_SERVER is not None:
                _SIM_SERVER.shutdown()
        except Exception:
            pass
        try:
            if _SIM_SERVER is not None:
                _SIM_SERVER.server_close()
        except Exception:
            pass
    except Exception:
        pass

    try:
        _cleanup_temp_dl_dirs()
    except Exception:
        pass
    raise SystemExit(code)

# 載入舊活動名稱（改為解析為課程+活動 次數）
existing_activities = set()  # 保留舊行為兼容（只含活動名或檔名）
existing_activity_counts = defaultdict(int)  # key: (course_name, activity_name) -> count
existing_files = set()  # 已下載過的實際檔名集合

if os.path.exists(OUTPUT_FILE):
    current_course = None
    with open(OUTPUT_FILE, "r", encoding="utf-8") as f:
        for line in f:
            line = line.strip()
            if not line:
                continue
            # 解析課程標題格式: "課程名稱: {course_name}"
            if line.startswith("課程名稱:"):
                try:
                    current_course = line.split(':', 1)[1].strip()
                except:
                    current_course = None
                continue
            # 若看起來像活動名稱，加入計數
            if is_activity(line):
                existing_activities.add(line)
                if current_course:
                    key = (current_course, line)
                    existing_activity_counts[key] += 1

download_dir = BASE_DOWNLOAD_DIR

def create_course_folder(course_name):
    # 把課程名稱中的不合法字元移除，避免建立資料夾失敗
    safe_name = "".join(c for c in course_name if c.isalnum() or c in " _-")
    course_path = os.path.join(download_dir, safe_name)
    os.makedirs(course_path, exist_ok=True)
    return course_path

# ========== 在啟動 Chrome 前清理既有進程（Windows 專用） ==========
if os.name == 'nt':
    _kill_existing_chrome_processes()

chrome_options = Options()

# 選擇無頭模式：Windows 用傳統 --headless，避免 --headless=new 的不穩定性
if os.name == 'nt':
    chrome_options.add_argument("--headless")  # Windows：用傳統無頭模式
else:
    chrome_options.add_argument("--headless")  # macOS/Linux：無頭模式

# 核心穩定化參數
chrome_options.add_argument("--disable-gpu")  # 防止顯示錯誤
chrome_options.add_argument("--log-level=3")  # 降低日誌等級，避免雜訊
chrome_options.add_argument("--window-size=1920,1080")  # 設定解析度，避免元素渲染錯誤
chrome_options.add_argument("--disable-blink-features=AutomationControlled")  # 避免被偵測

# exe優化：減少記憶體使用和加速啟動
chrome_options.add_argument("--disable-background-timer-throttling")
chrome_options.add_argument("--disable-backgrounding-occluded-windows")
chrome_options.add_argument("--disable-renderer-backgrounding")
# 移除 TranslateUI,VizDisplayCompositor（可能導致崩潰）
chrome_options.add_argument("--memory-pressure-off")  # 減少記憶體壓力
chrome_options.add_experimental_option("excludeSwitches", ["enable-logging", "enable-automation"])  # 避免除錯訊息
chrome_options.page_load_strategy = 'eager'  # 加快頁面載入速度
apply_chrome_binary_option(chrome_options)

prefs = {
    "download.default_directory": download_dir,  # 下載到指定資料夾
    "plugins.always_open_pdf_externally": True,  # 避免 PDF 在 Chrome 裡直接開啟
    "profile.default_content_setting_values.images": 2,  # 禁用圖片加快速度
    "profile.default_content_setting_values.media_stream": 2,  # 禁用媒體串流
    "profile.managed_default_content_settings.images": 2,  # 徹底禁用圖片
    "profile.default_content_settings.popups": 0  # 禁用彈窗
}
chrome_options.add_experimental_option("prefs", prefs)
driver = create_webdriver(chrome_options, hide_windows_console=True)

def simulate_typing(driver, element_id, text):
    script = f"""
    var element = document.getElementById('{element_id}');
    element.focus();
    element.value = '';
    
    // 模擬逐字輸入
    var text = '{text}';
    for (let i = 0; i < text.length; i++) {{
        element.value += text[i];
        element.dispatchEvent(new KeyboardEvent('keydown', {{
            key: text[i],
            char: text[i],
            bubbles: true
        }}));
        element.dispatchEvent(new Event('input', {{ bubbles: true }}));
        element.dispatchEvent(new KeyboardEvent('keyup', {{
            key: text[i],
            char: text[i],
            bubbles: true
        }}));
    }}
    
    element.dispatchEvent(new Event('change', {{ bubbles: true }}));
    element.blur();
    """
    driver.execute_script(script)

def _looks_like_login_page(driver_obj) -> bool:
    """判斷目前頁面是否看起來是登入頁。

    不只看 URL（有時會被 redirect/帶參數），也嘗試看登入欄位是否存在。
    """
    try:
        cur = (getattr(driver_obj, "current_url", "") or "").lower()
        if "/login/" in cur or "loginredirect" in cur:
            return True
    except Exception:
        pass
    try:
        # 只要能找到 username 或 loginbtn，通常就是登入頁
        driver_obj.find_element(By.ID, "username")
        return True
    except Exception:
        pass
    try:
        driver_obj.find_element(By.ID, "loginbtn")
        return True
    except Exception:
        pass
    return False

def goto_my_courses_page(driver_obj, *, wait_seconds: int = 10) -> bool:
    """嘗試進入『我的課程』頁面。

    以 URL 直達為主（比點連結穩），並確認不會被導回登入頁。
    """
    targets = [
        f"{MOODLE_BASE_URL}/my/courses.php",
        f"{MOODLE_BASE_URL}/my/",
    ]
    for url in targets:
        try:
            driver_obj.get(url)
            if wait_seconds and wait_seconds > 0:
                try:
                    WebDriverWait(driver_obj, wait_seconds).until(lambda d: not _looks_like_login_page(d))
                except Exception:
                    pass
            if not _looks_like_login_page(driver_obj):
                return True
        except Exception:
            continue
    return False

def ensure_logged_in(driver, username, password, *, silent=False, max_retries=2) -> bool:
    """確保目前 driver 已登入 Moodle。

    - 若已登入：直接回傳 True
    - 若在登入頁/被重導：自動填入帳密並登入
    """
    login_url = f"{MOODLE_BASE_URL}/login/index.php?loginredirect=1"

    # 避免沿用上一次的錯誤訊息
    try:
        clear_last_login_error()
    except Exception:
        pass

    def _is_logged_in() -> bool:
        # 以直達 /my/ 為準：若仍被導向 login 則視為未登入
        try:
            # 不要在「登入前檢查」階段額外等待，避免拖慢每次啟動。
            return goto_my_courses_page(driver, wait_seconds=0)
        except Exception:
            return False

    if _is_logged_in():
        return True

    def _read_login_danger_alert_text():
        """只在登入頁出現紅色錯誤框時回傳其文字，否則回傳 None。"""
        try:
            el = driver.find_element(By.CSS_SELECTOR, "div.alert.alert-danger[role='alert']")
            txt = (el.text or "").strip()
            return txt or "（登入失敗，但頁面未提供錯誤文字）"
        except Exception:
            return None

    def _classify_login_error(text: str) -> str:
        t = (text or "").strip()
        if not t:
            return "未知"
        if "驗證碼" in t or "captcha" in t.lower():
            return "驗證碼"
        if "帳號" in t or "學號" in t:
            if "密碼" in t or "password" in t.lower():
                return "帳密"
            return "帳號"
        if "密碼" in t or "password" in t.lower():
            return "密碼"
        if "鎖" in t or "lock" in t.lower():
            return "帳號鎖定"
        return "登入"

    def _detect_login_error_twice():
        """為避免跳轉中的瞬間狀態，連續檢查兩次錯誤框。"""
        last_text = None
        for i in range(2):
            txt = _read_login_danger_alert_text()
            if txt:
                last_text = txt
                break
            if i == 0:
                time.sleep(0.6)
        if last_text:
            return _classify_login_error(last_text), last_text
        return None, None

    for attempt in range(max_retries):
        try:
            try:
                clear_last_login_error()
            except Exception:
                pass
            driver.get(login_url)
            user_el = WebDriverWait(driver, 10).until(
                EC.visibility_of_element_located((By.ID, "username"))
            )
            try:
                user_el.clear()
            except Exception:
                pass
            user_el.send_keys(username)

            # 密碼欄位用 simulate_typing（更貼近人類輸入）
            try:
                simulate_typing(driver, 'password', password)
            except Exception:
                try:
                    pw = driver.find_element(By.ID, "password")
                    pw.clear()
                    pw.send_keys(password)
                except Exception:
                    driver.execute_script("document.getElementById('password').value = arguments[0];", password)

            driver.execute_script("document.getElementById('loginbtn').click();")
            time.sleep(0.6)

            # 只有真的出現 danger alert 才視為「明確的登入錯誤」。
            kind, text = _detect_login_error_twice()
            if text:
                set_last_login_error(kind, text)
                if (not silent) and should_emit_login_failure_messages():
                    print(f"\n{RED}X 登入失敗（{kind}）：{text}{RESET}")
                return False

            # 等待不要再停留登入頁
            try:
                WebDriverWait(driver, 12).until(lambda d: not _looks_like_login_page(d))
            except Exception:
                pass

            # 登入後直接導向『我的課程』確認 session OK
            if goto_my_courses_page(driver, wait_seconds=10):
                return True
        except Exception as e:
            if (not silent) and should_emit_login_failure_messages():
                try:
                    print(f"{YELLOW}! 自動登入失敗，重試中... ({attempt+1}/{max_retries}){RESET}")
                except Exception:
                    pass
            if attempt < max_retries - 1:
                time.sleep(0.5)
                continue
            return False

    return False

def ensure_logged_in_retry_once(driver, username, password, *, silent=False, max_retries=2) -> bool:
    """登入失敗時，再額外重試一次（外層重試）。

    避免偶發網路/跳轉造成的單次失敗。
    """
    ok = ensure_logged_in(driver, username, password, silent=silent, max_retries=max_retries)
    if ok:
        return True
    if not silent:
        try:
            # print(f"{YELLOW}! 登入失敗，將再重試一次...{RESET}")
            pass
        except Exception:
            pass
    return ensure_logged_in(driver, username, password, silent=silent, max_retries=max_retries)

def login_to_moodle(driver):
    """執行登入流程"""
    # 舊函式保留相容性：統一走 ensure_logged_in
    ensure_logged_in(driver, USERNAME, PASSWORD, silent=True, max_retries=2)

def _relogin_best_effort(driver) -> bool:
    """主 driver 斷線重建後的輕量重新登入流程（不直接 sys.exit）。"""
    for _attempt in range(2):
        try:
            if not ensure_logged_in(driver, USERNAME, PASSWORD, silent=True, max_retries=2):
                continue
            wait_local = WebDriverWait(driver, 5)
            if not goto_my_courses_page(driver, wait_seconds=10):
                continue
            select_ongoing_courses_filter(driver, wait_local)
            if not _looks_like_login_page(driver):
                return True
        except Exception:
            continue
    return False

def _recreate_main_driver_or_raise():
    """重建主 driver，並嘗試重新登入。"""
    global driver
    try:
        driver.quit()
    except Exception:
        pass
    driver = create_webdriver(chrome_options, hide_windows_console=True)
    if not _relogin_best_effort(driver):
        raise RuntimeError("重新登入未成功（主瀏覽器已重建）。")
    return driver

def _open_course_tabs_with_retry(course_hrefs_list):
    """預先開啟所有課程分頁；遇到斷線/InvalidSessionId 會重建 driver 後重試。"""
    global driver
    max_attempts = 3
    last_exc = None
    for attempt in range(max_attempts):
        try:
            main_handle = driver.current_window_handle
            for idx, href in enumerate(course_hrefs_list, 1):
                if idx == 1:
                    driver.get(href)
                else:
                    driver.execute_script("window.open(arguments[0], '_blank');", href)
                time.sleep(0.05)
            return main_handle, driver.window_handles
        except (InvalidSessionIdException, WebDriverException) as e:
            last_exc = e
            msg = (str(e) or "").lower()
            should_retry = (
                "invalid session" in msg
                or "session deleted" in msg
                or "disconnected" in msg
                or "unable to receive message from renderer" in msg
            )
            if attempt < max_attempts - 1 and should_retry:
                try:
                    _recreate_main_driver_or_raise()
                except Exception:
                    time.sleep(0.5)
                continue
            raise
    if last_exc is not None:
        raise last_exc
    raise RuntimeError("開啟課程分頁失敗（未知原因）。")

def select_ongoing_courses_filter(driver, wait):
    """進入我的課程後，切換分組到「進行中」（TODO:若要下載全部則 Ctrl+h 將"進行中"取代為"過去"）"""
    debug = os.environ.get("DEBUG_COURSE_FILTER", "0").strip().lower() in {"1", "true", "yes", "on"}

    def _get_active_grouping_text() -> str:
        try:
            txt = driver.execute_script("""
                const btn = document.getElementById('groupingdropdown');
                if (!btn) return '';
                const span = btn.querySelector('[data-active-item-text]');
                if (span && span.textContent) return span.textContent.trim();
                const aria = btn.getAttribute('aria-label') || '';
                if (aria) return aria.trim();
                // 注意：btn.innerText 可能包含下拉選單文字，不可靠
                return (btn.textContent || '').trim();
            """)
            return (txt or "").strip()
        except Exception:
            return ""

    def _looks_like_ongoing(active_text: str) -> bool:
        t = (active_text or "").strip()
        if not t:
            return False
        # 避免把「過去」也包含在內的文字誤判
        return ("進行中" in t) and ("過去" not in t)

    try:
        # 先等 dropdown 出現（第一次執行/網路慢時很常還沒出現）
        try:
            wait.until(EC.presence_of_element_located((By.ID, "groupingdropdown")))
        except Exception:
            return False

        current_text = _get_active_grouping_text()
        if debug:
            print(f"{YELLOW}[DEBUG] groupingdropdown active: {current_text}{RESET}")

        if _looks_like_ongoing(current_text):
            return True

        dropdown_btn = wait.until(EC.element_to_be_clickable((By.ID, "groupingdropdown")))
        driver.execute_script("arguments[0].click();", dropdown_btn)
        time.sleep(0.2)

        option_clicked = driver.execute_script("""
            const btn = document.getElementById('groupingdropdown');
            if (!btn) return false;
            // 盡量只在這個 dropdown 的 menu 內找，避免誤點其他選單
            const root = btn.closest('.dropdown') || btn.parentElement;
            let menu = root ? root.querySelector('.dropdown-menu') : null;
            if (!menu) menu = document.querySelector('.dropdown-menu.show');
            if (!menu) return false;
            const items = Array.from(menu.querySelectorAll('a, button, .dropdown-item'));
            const norm = s => (s || '').replace(/\s+/g,' ').trim();
            const target = items.find(el => {
                const t = norm(el.innerText || el.textContent || '');
                return t === '進行中' || t.includes('進行中');
            });
            if (!target) return false;
            target.click();
            return true;
        """)

        if not option_clicked:
            if debug:
                print(f"{YELLOW}[DEBUG] cannot find/click '進行中' option{RESET}")
            return False

        # 等到 active text 真的變成「進行中」才算成功
        try:
            WebDriverWait(driver, 8).until(lambda d: _looks_like_ongoing(_get_active_grouping_text()))
        except Exception:
            pass

        after_text = _get_active_grouping_text()
        if debug:
            print(f"{YELLOW}[DEBUG] groupingdropdown after: {after_text}{RESET}")
        return _looks_like_ongoing(after_text)
    except Exception:
        return False

# 登入 + 進入「我的課程」
if not ensure_logged_in_retry_once(driver, USERNAME, PASSWORD, silent=False, max_retries=3):
    print(f"\n{RED}{'='*60}{RESET}")
    if should_emit_login_failure_messages():
        print(f"{RED}X 登入失敗：無法進入系統{RESET}")
    else:
        print(f"{RED}X 無法進入系統{RESET}")
    print(f"\n按 Enter 鍵離開...")
    input()
    driver.quit()
    sys.exit(1)

try:
    wait = WebDriverWait(driver, 5)
    # 不點連結，直接導向更穩
    ok_my = False
    for _i in range(3):
        if goto_my_courses_page(driver, wait_seconds=10):
            ok_my = True
            break
        # 若被導回 login，嘗試再登入一次
        ensure_logged_in(driver, USERNAME, PASSWORD, silent=True, max_retries=2)
    if not ok_my:
        raise RuntimeError("登入後仍無法進入 /my/ 或 /my/courses.php")
    # 第一次載入可能較慢：若切換失敗就再重試一次
    if not select_ongoing_courses_filter(driver, WebDriverWait(driver, 12)):
        try:
            driver.get(f"{MOODLE_BASE_URL}/my/courses.php")
        except Exception:
            pass
        select_ongoing_courses_filter(driver, WebDriverWait(driver, 12))
except Exception as e:
    print(f"\n{RED}{'='*60}{RESET}")
    print(f"{RED}X 登入後無法進入『我的課程』（已重試）{RESET}")
    print(f"\n錯誤訊息：{e}")
    print(f"\n按 Enter 鍵離開...")
    input()
    try:
        driver.quit()
    except Exception:
        pass
    sys.exit(1)

def snapshot_course_hrefs(driver):
    """用 JS 快照課程連結，降低前端版型變更造成的定位失敗。"""
    script = """
        const nodes = Array.from(document.querySelectorAll("a[href*='/course/view.php?id=']"));
        // 優先取「可見」的連結，避免抓到被隱藏/折疊的過去課程區塊
        const visible = nodes.filter(a => {
            try {
                return a && a.offsetParent !== null;
            } catch(e) { return false; }
        });
        const src = (visible.length ? visible : nodes);
        const hrefs = src
      .map(a => a.getAttribute('href') || a.href || '')
      .filter(Boolean)
      .map(h => h.split('#')[0]);
    return Array.from(new Set(hrefs));
    """
    try:
        result = driver.execute_script(script)
        if isinstance(result, list):
            return [h for h in result if isinstance(h, str) and h.startswith("http")]
    except Exception:
        pass
    return []

course_hrefs = []
for attempt in range(3):
    try:
        # 保險做法：若目前不在「我的課程」，直接導向固定頁面。
        if "my/courses.php" not in driver.current_url:
            driver.get("https://elearningv4.nuk.edu.tw/my/courses.php")

        wait = WebDriverWait(driver, 12)
        ok_filter = select_ongoing_courses_filter(driver, wait)
        if not ok_filter:
            # 若切換失敗，先刷新一次再試（避免抓到過去全部）
            try:
                driver.refresh()
            except Exception:
                pass
            ok_filter = select_ongoing_courses_filter(driver, WebDriverWait(driver, 12))
        if not ok_filter and os.environ.get("DEBUG_COURSE_FILTER", "0") not in {"1", "true", "yes", "on"}:
            print(f"{YELLOW}! 提醒：無法確認已切換到『進行中』，可能會抓到過去課程；可設 DEBUG_COURSE_FILTER=1 觀察{RESET}")
        course_hrefs = wait.until(lambda d: snapshot_course_hrefs(d))
        if course_hrefs:
            break
    except TimeoutException:
        if attempt < 2:
            continue
    except Exception:
        if attempt < 2:
            continue

if not course_hrefs:
    print(f"\n{RED}{'='*60}{RESET}")
    print(f"{RED}X 無法載入課程清單（逾時）{RESET}")
    print(f"{YELLOW}! 可能是 Moodle 回應較慢或頁面版型變更{RESET}")
    print(f"{YELLOW}! 請稍後重試，或手動確認『我的課程』頁面是否可正常開啟{RESET}")
    print(f"{RED}{'='*60}{RESET}")
    print(f"\n按 Enter 鍵離開...")
    input()
    driver.quit()
    sys.exit(1)

# 建置環境中：輸出課程數量（用於確認是否成功抓到課程清單）
try:
    if IS_BUILD_ENV:
        print(f"{BLUE}課程數量：{len(course_hrefs)}{RESET}")
except Exception:
    pass

# 啟動「模擬瀏覽器」本機轉發器：讓終端輸出的連結可直接點開並自動登入
start_simulator_server()

# ====== 作業檢查：改為併入主流程（不再額外開新 driver / 新分頁） ======
empty_assignments = []

all_output_lines = []
red_activities_to_print = []  # 暫存紅色活動（(name, link, course_name, week_header, course_path, course_url)）
failed_downloads = []  # 儲存下載失敗的連結

# 線程鎖，用於同步輸出和資料收集
output_lock = threading.Lock()
data_lock = threading.Lock()

# 預先為所有課程開啟分頁
main_window, all_tabs = _open_course_tabs_with_retry(course_hrefs)

# 在所有分頁中注入 JavaScript 開始數據提取
extraction_script = """
return (function() {
    // 文本清理函數
    function cleanActivityName(text) {
        if (!text) return text;
        
        // 移除各種換行字元和不需要的空白字元
        let cleaned = text.replace(/\\n/g, ' ')
                          .replace(/\\r/g, ' ')
                          .replace(/\\t/g, ' ');
        
        // 移除多餘的空格
        cleaned = cleaned.replace(/\\s+/g, ' ').trim();
        
        // 移除常見的亂碼後綴
        if (cleaned.includes('\\n作業')) {
            cleaned = cleaned.replace('\\n作業', '');
        }
        
        return cleaned.trim();
    }
    
    let data = {
        courseName: document.querySelector('h1.h2') ? document.querySelector('h1.h2').textContent : null,
        sections: []
    };
    
    if (!data.courseName) return data;
    
    let sections = document.querySelectorAll('li.section');
    sections.forEach((section, idx) => {
        let titleElem = section.querySelector('h3.sectionname');
        let weekText = titleElem ? titleElem.textContent.trim() : '';
        
        let activities = [];
        let activityItems = section.querySelectorAll('div.activity-item');
        
        activityItems.forEach(act => {
            let name = act.getAttribute('data-activityname');
            if (!name || !name.trim()) {
                let nameElem = act.querySelector('span.instancename');
                name = nameElem ? nameElem.textContent.trim() : null;
            }
            
            // 清理活動名稱
            name = cleanActivityName(name);
            
            if (name && !(name.startsWith('第') && name.endsWith('週')) && name.trim()) {
                let link = act.querySelector('a.aalink');
                let href = link ? link.getAttribute('href') : '（無連結）';
                
                let desc = '';
                let descElem = act.querySelector('.activity-description');
                if (descElem) {
                    desc = descElem.innerText || descElem.textContent || '';
                }
                
                activities.push({name: name, href: href, description: desc.trim()});
            }
        });
        
        if (activities.length > 0) {
            data.sections.push({
                index: idx,
                weekText: weekText,
                activities: activities
            });
        }
    });
    
    return data;
})();
"""

# 隨抓隨存：直接提取並儲存有效數據
tab_data_map = {}
for tab_handle in all_tabs:
    driver.switch_to.window(tab_handle)
    try:
        data = driver.execute_script(extraction_script)
        if data and data.get('courseName'):
            tab_data_map[tab_handle] = data
    except:
        pass  # 提取失敗則跳過該分頁

# 使用線程池並行處理提取到的數據
def process_extracted_data(tab_handle, data, href):
    """處理已提取的課程數據"""
    if not data or not data.get('courseName'):
        return None
    
    local_output = []
    local_red_activities = []
    local_assignments = []  # (course, name, url)
    from collections import defaultdict as _dd

    # 記錄在本次執行中每個 (course, activity) 已遇到的次數
    seen_counts_local = _dd(int)

    course_name = data['courseName']
    local_output.append(f"課程名稱: {course_name}\n")
    course_path = create_course_folder(course_name)
    
    # 準備課程活動記錄內容
    course_activity_log = []
    course_activity_log.append(f"課程名稱: {course_name}")
    course_activity_log.append(f"記錄時間: {time.strftime('%Y-%m-%d %H:%M:%S')}")
    course_activity_log.append("=" * 60)
    course_activity_log.append("")

    # 收集每週的紀錄區塊，最後再反轉輸出，讓周次大的排前面
    week_logs = []
    
    latest_with_content = None
    current_week_info = None  # 本週資訊
    next_week_info = None     # 下週資訊
    
    import datetime
    today = datetime.date.today()
    
    # 解析日期範圍的輔助函數
    def parse_week_dates(week_text):
        """從週次文字中解析日期範圍，例如 '01月 4日 - 01月 10日' 或 '2024-10-01 ~ 2024-10-07'"""
        try:
            # 格式1: '2024-10-01 ~ 2024-10-07' (有年份)
            if '~' in week_text and '-' in week_text.split('~')[0]:
                parts = week_text.split('~')
                start_str = parts[0].strip()
                end_str = parts[1].strip()
                start_date = datetime.datetime.strptime(start_str, '%Y-%m-%d').date()
                end_date = datetime.datetime.strptime(end_str, '%Y-%m-%d').date()
                return start_date, end_date
            
            # 格式2: '01月 4日 - 01月 10日' 或 '12月 28日 - 01月 3日' (沒有年份，需要推斷)
            if '-' in week_text or '~' in week_text:
                separator = '-' if '-' in week_text else '~'
                parts = week_text.split(separator)
                if len(parts) == 2:
                    # 提取月日
                    import re
                    start_match = re.search(r'(\d+)月\s*(\d+)日', parts[0])
                    end_match = re.search(r'(\d+)月\s*(\d+)日', parts[1])
                    
                    if start_match and end_match:
                        start_month = int(start_match.group(1))
                        start_day = int(start_match.group(2))
                        end_month = int(end_match.group(1))
                        end_day = int(end_match.group(2))
                        
                        # 推斷年份：假設課程在學年度內 (例如 2025-09 到 2026-01)
                        current_year = today.year
                        
                        # 判斷開始日期的年份
                        if start_month >= 9:  # 9-12月
                            # 如果今天的月份小於開始月份，說明開始月份是上一年
                            if today.month < start_month:
                                start_year = current_year - 1
                            else:
                                start_year = current_year
                        else:  # 1-8月
                            start_year = current_year
                        
                        # 判斷結束日期的年份
                        if end_month >= 9:  # 9-12月
                            end_year = start_year
                        else:  # 1-8月
                            # 如果開始月份是12月，結束月份是1月，則跨年
                            if start_month == 12 and end_month == 1:
                                end_year = start_year + 1
                            else:
                                end_year = start_year
                        
                        start_date = datetime.date(start_year, start_month, start_day)
                        end_date = datetime.date(end_year, end_month, end_day)
                        return start_date, end_date
        except Exception as e:
            print(f"    {RED}[錯誤] 日期解析失敗: {week_text}, {e}{RESET}")
        return None, None
    
    for section in data['sections']:
        idx = section['index']
        week_text = section['weekText']
        week_header = f"第{idx+1}週 ({week_text})"
        week_activity_infos = []
        
        # 判斷這週是本週、下週還是其他
        start_date, end_date = parse_week_dates(week_text)
        week_label = None
        if start_date and end_date:
            if start_date <= today <= end_date:
                week_label = '本週'
            elif start_date > today:
                days_diff = (start_date - today).days
                if 1 <= days_diff <= 7:
                    week_label = '下週'
        
        # 建立本週的紀錄區塊
        week_block = []
        week_block.append(week_header)
        week_block.append("-" * 40)
        
        for act_data in section['activities']:
            name = clean_activity_name(act_data['name'])  # 再次確保清理
            href_link = act_data['href']
            description = act_data.get('description', '')

            # 收集作業連結（不額外開頁，只做清單）
            if isinstance(href_link, str) and 'mod/assign/view.php' in href_link:
                local_assignments.append({
                    'course': course_name,
                    'name': name,
                    'url': href_link,
                })

            week_activity_infos.append((name, href_link))
            local_output.append(name)
            
            # 添加活動到週區塊
            week_block.append(f"  • {name}")
            week_block.append(f"    連結: {href_link}")

            # 以 (課程, 活動名) 為單位計數，若本次遇到次數超過 OUTPUT 中記錄的次數，視為新的
            key = (course_name, name)
            seen_counts_local[key] += 1
            existing_count = existing_activity_counts.get(key, 0)
            if seen_counts_local[key] > existing_count:
                local_red_activities.append((name, href_link, course_name, week_header, course_path, href, description))
                # 將 existing_activity_counts 增加以避免同一執行中重複標示同樣的項目
                existing_activity_counts[key] = existing_count + 1
                existing_activities.add(name)
        
        week_block.append("")
        
        # 保存週次資訊
        if week_activity_infos:
            week_info = (idx + 1, course_name, week_header, week_activity_infos, course_path, week_label)
            
            if week_label == '本週':
                current_week_info = week_info
            elif week_label == '下週':
                next_week_info = week_info
            
            latest_with_content = week_info
        
        local_output.append("")

        # 收集週區塊以便反轉排序
        week_logs.append(week_block)

    # 依周次由大到小輸出到活動記錄檔
    for block in reversed(week_logs):
        course_activity_log.extend(block)
    
    # 將活動記錄寫入課程資料夾中的文字檔
    activity_log_path = os.path.join(course_path, "課程活動記錄.txt")
    try:
        with open(activity_log_path, "w", encoding="utf-8") as f:
            f.write("\n".join(course_activity_log))
        os.utime(activity_log_path, None)  # 更新時間戳，確保檔案保持最新
    except Exception as e:
        print(f"{RED}! 無法寫入活動記錄: {course_name} - {e}{RESET}")
    
    return {
        'course_name': course_name,
        'output': local_output,
        'red_activities': local_red_activities,
        'assignments': local_assignments,
        'latest_with_content': latest_with_content,
        'current_week_info': current_week_info,
        'next_week_info': next_week_info,
        'failed_downloads': []
    }

# 使用線程池並行處理數據（真正的並行）
course_results = []  # 改為存 (idx, result) 以便排序
with ThreadPoolExecutor(max_workers=len(all_tabs)) as executor:
    future_to_tab = {
        executor.submit(process_extracted_data, tab_handle, tab_data_map[tab_handle], course_hrefs[idx]): (tab_handle, idx)
        for idx, tab_handle in enumerate(all_tabs)
    }
    
    # 誰先完成就先輸出誰的結果
    for future in as_completed(future_to_tab):
        result = future.result()
        tab_handle, idx = future_to_tab[future]
        if result:
            course_results.append((idx, result))  # 保存索引
            
            # 立即輸出黃色資訊（本週、下週、最新週的舊活動）
            # 如果是第一次使用，跳過輸出
            if not IS_FIRST_TIME:
                with output_lock:
                    weeks_to_show = []
                    
                    # 1. 本週
                    if result['current_week_info']:
                        weeks_to_show.append(result['current_week_info'])
                    
                    # 2. 下週（如果有）
                    if result['next_week_info']:
                        weeks_to_show.append(result['next_week_info'])
                    
                    # 3. 最新週（如果與本週/下週不重複）
                    if result['latest_with_content']:
                        latest_week_num = result['latest_with_content'][0]
                        current_week_num = result['current_week_info'][0] if result['current_week_info'] else None
                        next_week_num = result['next_week_info'][0] if result['next_week_info'] else None
                        
                        if latest_week_num == current_week_num:
                            # 最新週就是本週，不重複輸出但要註明
                            if result['current_week_info']:
                                info = result['current_week_info']
                                weeks_to_show[0] = (info[0], info[1], info[2], info[3], info[4], '本週（最新週）')
                        elif latest_week_num == next_week_num:
                            # 最新週就是下週，不重複輸出但要註明
                            for i, info in enumerate(weeks_to_show):
                                if info[5] == '下週':
                                    weeks_to_show[i] = (info[0], info[1], info[2], info[3], info[4], '下週（最新週）')
                        elif latest_week_num != current_week_num and latest_week_num != next_week_num:
                            # 最新週與本週和下週都不同，額外輸出
                            info = result['latest_with_content']
                            weeks_to_show.append((info[0], info[1], info[2], info[3], info[4], '最新週'))
                    
                    # 輸出所有需要顯示的週次
                    course_name_printed = False  # 追蹤課程名稱是否已輸出
                    for week_info in weeks_to_show:
                        if week_info and week_info[0] != 1: #TODO 排除用team的教授
                        # if week_info:  # 顯示所有週次，包括第一週
                            # week_info = (week_num, course_name, week_header, activities, course_path, week_label)
                            week_label = week_info[5] if len(week_info) > 5 else None
                            
                            # 如果是下週但沒有活動，跳過
                            if week_label == '下週' and not week_info[3]:
                                continue
                            
                            # 課程名稱只輸出一次（綠色）
                            if not course_name_printed:
                                print(f"\n{GREEN}{week_info[1]}{RESET}")
                                course_name_printed = True
                            
                            # 根據標籤選擇顏色和輸出格式（縮排2格）
                            if week_label:
                                # 處理本週（最新週）的情況
                                if week_label == '本週（最新週）':
                                    # 整體用最新週的紫色，但「本週」用黃橘色
                                    print(f"  {PURPLE}{week_info[2]} ({ORANGE}本週{PURPLE}（最新週）){RESET}")
                                # 處理下週（最新週）的情況
                                elif week_label == '下週（最新週）':
                                    # 整體用最新週的紫色，但「下週」用藍色
                                    print(f"  {PURPLE}{week_info[2]} ({BLUE}下週{PURPLE}（最新週）){RESET}")
                                # 一般情況
                                elif week_label == '本週':
                                    print(f"  {ORANGE}{week_info[2]} ({ORANGE}本週{ORANGE}){RESET}")
                                # 一般下週情況
                                elif week_label == '下週':
                                    print(f"  {BLUE}{week_info[2]} ({BLUE}下週{BLUE}){RESET}")
                                # 只是最新週
                                elif week_label == '最新週':
                                    print(f"  {PURPLE}{week_info[2]} ({PURPLE}最新週{PURPLE}){RESET}")
                                else:
                                    print(f"  {week_info[2]} ({week_label})")
                            else:
                                print(f"  {week_info[2]}")
                            
                            # 顯示所有活動（無論新舊）（縮排2格）
                            for name, href_link in week_info[3]:
                                # 再次確保活動名稱清理，移除換行字元和亂碼
                                clean_name = clean_activity_name(name)
                                display_link = make_simulator_open_link(href_link) if href_link else href_link
                                # 活動名稱維持黃色
                                print(f"  {YELLOW}{clean_name}{RESET} - {display_link}")
                            print("")
                             
            
            # 收集資料
            with data_lock:
                all_output_lines.extend(result['output'])
                red_activities_to_print.extend(result['red_activities'])
                failed_downloads.extend(result['failed_downloads'])

def remove_zone_identifier(filepath):
    """移除 Zone.Identifier 標記，避免「受保護的檢視」（僅 Windows）"""
    if os.name != 'nt':
        return  # macOS / Linux 無此機制
    try:
        zone_path = filepath + ":Zone.Identifier"
        os.remove(zone_path)
    except OSError:
        pass

def unblock_office_files_in_dir(root_dir):
    """最終保險：掃描資料夾內 Office/PDF 檔並解除封鎖。"""
    if os.name != 'nt' or not root_dir or not os.path.exists(root_dir):
        return
    office_exts = {".doc", ".docx", ".ppt", ".pptx", ".xls", ".xlsx", ".pdf"}
    for cur_root, _, files in os.walk(root_dir):
        for filename in files:
            _, ext = os.path.splitext(filename)
            if ext.lower() in office_exts:
                remove_zone_identifier(os.path.join(cur_root, filename))

def ensure_unique_filename(filepath):
    """
    檢查檔案是否已存在。若存在，自動改名為 filename (1).ext, (2).ext 等。
    返回可安全使用的檔案路徑，不會覆蓋既有檔案。
    """
    if not os.path.exists(filepath):
        return filepath
    
    dirname = os.path.dirname(filepath)
    basename = os.path.basename(filepath)
    
    # 分離檔名和副檔名
    if '.' in basename:
        name_part, ext_part = basename.rsplit('.', 1)
        ext_part = '.' + ext_part
    else:
        name_part = basename
        ext_part = ''

    # 嘗試找不重複的檔名
    counter = 1
    while counter <= 999:
        new_name = f"{name_part} ({counter}){ext_part}"
        new_path = os.path.join(dirname, new_name)
        if not os.path.exists(new_path):
            return new_path
        counter += 1
    
    # 如果 999 個都重複了（不太可能），就直接用原名
    return filepath
        
def wait_for_download(filename, download_path=None, timeout=300, ask_after=30, size_limit_mb=50):
    """
    等待下載完成。若超過 ask_after 秒或檔案超過 size_limit_mb MB，詢問是否繼續。
    
    返回值：
    - 檔案路徑：下載成功
    - None：跳過此檔案
    - 'SKIP_COURSE'：跳過整個課程的所有檔案
    """
    if download_path is None:
        download_path = download_dir
    # 第一次建置：30 秒自動跳過，不詢問
    if IS_FIRST_TIME:
        ask_after = 30
        timeout = 30
    file_path = os.path.join(download_path, filename)
    cr_path = file_path + ".crdownload"
    start = time.time()
    expected_root, expected_ext = os.path.splitext(filename)

    def _is_expected_variant(name):
        root, ext = os.path.splitext(name)
        if ext != expected_ext:
            return False
        if root == expected_root:
            return True
        if root.startswith(f"{expected_root} (") and root.endswith(")"):
            suffix = root[len(expected_root) + 2:-1]
            return suffix.isdigit()
        return False

    def _scan_existing_variants():
        existing = {}
        try:
            for name in os.listdir(download_path):
                if not _is_expected_variant(name):
                    continue
                path = os.path.join(download_path, name)
                try:
                    existing[path] = os.path.getmtime(path)
                except OSError:
                    continue
        except OSError:
            pass
        return existing

    existing_variants = _scan_existing_variants()

    已詢問時間 = False
    已詢問大小 = False
    last_progress_time = 0  # 記錄上次輸出進度的時間
    
    # 記錄開始時檔案的修改時間(如果存在)
    initial_mtime = None
    if os.path.exists(file_path):
        initial_mtime = os.path.getmtime(file_path)


    while True:
        # 支援 Chrome 自動改名（例如 filename (1).ext），避免誤判下載失敗。
        completed_candidates = []
        try:
            for name in os.listdir(download_path):
                if not _is_expected_variant(name):
                    continue
                candidate_path = os.path.join(download_path, name)
                candidate_cr = candidate_path + ".crdownload"
                if os.path.exists(candidate_cr):
                    continue
                try:
                    current_mtime = os.path.getmtime(candidate_path)
                except OSError:
                    continue

                previous_mtime = existing_variants.get(candidate_path)
                if previous_mtime is None or current_mtime > previous_mtime:
                    completed_candidates.append((current_mtime, candidate_path))
        except OSError:
            pass

        if completed_candidates:
            completed_candidates.sort(key=lambda x: x[0], reverse=True)
            return completed_candidates[0][1]

        elapsed = time.time() - start

        # 每 10 秒顯示一次進度
        if not IS_FIRST_TIME and elapsed - last_progress_time >= 10 and elapsed >= 10:
            print(f"已等待 {int(elapsed)} 秒")
            last_progress_time = elapsed

        # 超過指定等待秒數
        if elapsed > ask_after and not 已詢問時間:
            已詢問時間 = True
            if IS_FIRST_TIME:
                # 第一次建置：自動跳過，不詢問
                return None
            print(f"\n{YELLOW}下載已等待超過 {ask_after} 秒{RESET}")
            print(f"   檔案：{filename}")
            print(f"   - 繼續等待 (Enter)")
            print(f"   - 放棄此檔案 (輸入 d)")
            print(f"   - 放棄此課程所有檔案 (輸入 dd)")
            choice = input(f"   請選擇：").strip().lower()
            last_progress_time = elapsed
            if choice == "dd":
                return 'SKIP_COURSE'
            elif choice == "d":
                return None
        # 若部分下載的檔案存在 → 第一次建置不詢問，直接跳過大檔案
        if not 已詢問大小:
            try:
                for name in os.listdir(download_path):
                    if not _is_expected_variant(name):
                        continue
                    current_path = os.path.join(download_path, name)
                    size_mb = os.path.getsize(current_path) / (1024 * 1024)
                    if size_mb > size_limit_mb:
                        已詢問大小 = True
                        if IS_FIRST_TIME:
                            return None
                        choice = input(f"檔案超過 {size_limit_mb} MB，要繼續下載嗎？(y/n)：").strip().lower()
                        if choice != "y":
                            print("⏭️ 使用者選擇跳過。")
                            return None
                        break
            except OSError:
                pass

        # 檢查是否超過 timeout
        if elapsed > timeout:
            if not IS_FIRST_TIME:
                print(f"{RED}X 下載超時：{filename}{RESET}")
            raise TimeoutError(f"下載超時：{filename}")

        time.sleep(0.2)

def extract_filename_from_url(url):
    """從 URL 中提取檔名"""
    try:
        # 移除 query parameters
        url_without_params = url.split('?')[0]
        # 取得最後一段作為檔名
        filename = os.path.basename(url_without_params)
        # URL decode
        from urllib.parse import unquote
        filename = unquote(filename)
        return filename if filename else "unknown_file"
    except:
        return "unknown_file"

def should_skip_download_filename(filename):
    """硬編碼跳過特定檔名下載（不分副檔名）。"""
    try:
        base_name = os.path.basename(filename or "").strip().lower()
        stem = os.path.splitext(base_name)[0]
        if stem == "image10":
            return True
        if "downloads.htm" in base_name:
            return True
        return False
    except Exception:
        return False



def extract_file(file_path, dest_dir):
    temp_extract_dir = None

    def _move_extracted_files_no_overwrite(src_root: str, dst_root: str) -> int:
        moved = 0
        for cur_root, _, files in os.walk(src_root):
            for filename in files:
                src_path = os.path.join(cur_root, filename)
                try:
                    rel_path = os.path.relpath(src_path, src_root)
                except Exception:
                    rel_path = filename
                target_path = os.path.join(dst_root, rel_path)
                target_dir = os.path.dirname(target_path)
                try:
                    os.makedirs(target_dir, exist_ok=True)
                except Exception:
                    # 目的地路徑若遇到「檔案擋住資料夾」等情況，退回直接丟到根目錄
                    target_path = os.path.join(dst_root, os.path.basename(target_path))
                    target_dir = os.path.dirname(target_path)
                    os.makedirs(target_dir, exist_ok=True)

                target_path = ensure_unique_filename(target_path)
                shutil.move(src_path, target_path)
                moved += 1
        return moved

    try:
        os.makedirs(dest_dir, exist_ok=True)
        temp_extract_dir = tempfile.mkdtemp(prefix="_temp_extract_", dir=dest_dir)

        # 驗證檔案格式（檢查 magic bytes）
        with open(file_path, 'rb') as f:
            header = f.read(16)

        # ZIP 檔案的 magic bytes: 50 4B (PK)
        # RAR 檔案的 magic bytes: 52 61 72 21 (Rar!)
        # 7Z 檔案的 magic bytes: 37 7A BC AF 27 1C
        is_zip = header[:2] == b'PK'
        is_rar = header[:4] == b'Rar!' or header[:7] == b'\x52\x61\x72\x21\x1A\x07\x01'  # RAR 5.0
        is_7z = header[:6] == b'7z\xbc\xaf\x27\x1c'

        # 檢查檔案是否為 HTML（可能是下載錯誤頁面）
        is_html = header[:15].lower().startswith(b'<!doctype html') or header[:6].lower().startswith(b'<html')

        if is_html:
            print(f"   {YELLOW}! 檔案不是壓縮檔，而是 HTML 頁面，跳過解壓{RESET}")
            return False

        extracted_ok = False

        if file_path.endswith(".zip"):
            if not is_zip:
                print(f"   {YELLOW}! 檔案副檔名為 .zip 但不是有效的 ZIP 格式，跳過解壓{RESET}")
                return False

            # 處理 ZIP 檔案的中文檔名亂碼問題
            with zipfile.ZipFile(file_path, 'r') as zf:
                for info in zf.infolist():
                    try:
                        # 嘗試修正檔名編碼 (Windows 中文系統常用 cp437 -> gbk/big5)
                        fixed_filename = info.filename.encode('cp437').decode('big5')
                        info.filename = fixed_filename
                    except (UnicodeDecodeError, UnicodeEncodeError):
                        try:
                            fixed_filename = info.filename.encode('cp437').decode('gbk')
                            info.filename = fixed_filename
                        except Exception:
                            pass
                    zf.extract(info, temp_extract_dir)
            extracted_ok = True

        elif file_path.endswith(".rar"):
            if not is_rar:
                print(f"   {YELLOW}! 檔案副檔名為 .rar 但不是有效的 RAR 格式，跳過解壓{RESET}")
                return False

            # 方法1: 使用 7-Zip（最穩定，Windows 常見）
            seven_zip_paths = [
                r"C:\Program Files\7-Zip\7z.exe",
                r"C:\Program Files (x86)\7-Zip\7z.exe"
            ]
            for seven_zip in seven_zip_paths:
                if os.path.exists(seven_zip):
                    try:
                        result = subprocess.run(
                            [seven_zip, "x", file_path, f"-o{temp_extract_dir}", "-y"],
                            capture_output=True,
                            text=True,
                            timeout=60
                        )
                        if result.returncode == 0:
                            extracted_ok = True
                            break
                    except Exception as e:
                        print(f"{YELLOW}!  7-Zip 解壓失敗: {e}{RESET}")

            # 方法2: 優先使用 patool（更穩定，支持多種後端）
            if not extracted_ok and HAS_PATOOL:
                try:
                    patool.extract_archive(file_path, outdir=temp_extract_dir, verbosity=-1)
                    extracted_ok = True
                except Exception as e:
                    print(f"{YELLOW}!  patool 解壓失敗: {e}{RESET}")

            # 方法3: 使用 rarfile
            if not extracted_ok and HAS_RARFILE:
                try:
                    with rarfile.RarFile(file_path, 'r') as rf:
                        rf.extractall(temp_extract_dir)
                    extracted_ok = True
                except Exception as e:
                    print(f"{YELLOW}!  rarfile 解壓失敗: {e}{RESET}")

            if not extracted_ok:
                print(f"{YELLOW}!  無法解壓 RAR 檔案，已跳過: {os.path.basename(file_path)}{RESET}")
                print(f"   💡 7-Zip 已安裝但無法使用，請嘗試：")
                print(f"   1. 重新啟動終端或電腦")
                print(f"   2. 或手動安裝: winget install 7zip.7zip")
                return False

        elif file_path.endswith(".7z"):
            if not is_7z:
                print(f"   {YELLOW}! 檔案副檔名為 .7z 但不是有效的 7Z 格式，跳過解壓{RESET}")
                return False

            with py7zr.SevenZipFile(file_path, 'r') as sz:
                sz.extractall(temp_extract_dir)
            extracted_ok = True
        else:
            return False

        if not extracted_ok:
            return False

        _move_extracted_files_no_overwrite(temp_extract_dir, dest_dir)
        return True
    except zipfile.BadZipFile:
        print(f"   {YELLOW}! 無效的 ZIP 檔案格式，跳過解壓{RESET}")
        return False
    except Exception as e:
        print(f"X 解壓失敗: {os.path.basename(file_path)}, 原因: {e}")
        return False
    finally:
        try:
            if temp_extract_dir and os.path.exists(temp_extract_dir):
                shutil.rmtree(temp_extract_dir, ignore_errors=True)
        except Exception:
            pass

def create_session_with_cookies():
    """
    建立並設定 requests 會話，並帶上 Selenium 驅動的 cookies

    用途：保持登入狀態，下載受保護的資源
    返回：requests.Session 物件
    """
    cookies = driver.get_cookies()
    session = requests.Session()
    for cookie in cookies:
        session.cookies.set(cookie['name'], cookie['value'])
    return session

def try_download_google_drive(gdrive_url, dest_dir, base_name, session):
    """
    gdown for Drive files/folders, requests export for Docs/Sheets/Slides.
    Returns downloaded file/folder path, or None on failure.
    """
    import re as _re
    import gdown
    url = gdrive_url

    # --- Google Drive folder ---
    m_folder = _re.search(r'drive\.google\.com/drive/(?:u/\d+/)?folders/([a-zA-Z0-9_-]+)', url)
    if m_folder:
        fid = m_folder.group(1)
        folder_url = f"https://drive.google.com/drive/folders/{fid}"
        try:
            result = gdown.download_folder(folder_url, output=dest_dir, quiet=True, use_cookies=False)
            if result:
                return dest_dir
        except Exception:
            pass
        return None

    # --- Google Drive single file ---
    m = _re.search(r'drive\.google\.com/(?:file/d/|open\?id=)([a-zA-Z0-9_-]+)', url)
    if not m:
        m = _re.search(r'[?&]id=([a-zA-Z0-9_-]+)', url)
    if m:
        fid = m.group(1)
        file_url = f"https://drive.google.com/uc?id={fid}"
        try:
            out_path = gdown.download(file_url, output=dest_dir + "/", quiet=True, use_cookies=False, fuzzy=True)
            if out_path and os.path.exists(out_path):
                return out_path
        except Exception:
            pass
        return None

    def _save_stream(rsp, filepath):
        if rsp.status_code != 200:
            return False
        for chunk in rsp.iter_content(chunk_size=8192):
            pass  # consume to detect HTML error pages in first chunk
        rsp2 = session.get(rsp.url, stream=True, allow_redirects=True, timeout=30)
        if rsp2.status_code != 200:
            return False
        with open(filepath, 'wb') as f:
            for chunk in rsp2.iter_content(chunk_size=8192):
                f.write(chunk)
        return os.path.getsize(filepath) > 0

    def _export_stream(export_url, filepath):
        try:
            rsp = session.get(export_url, stream=True, allow_redirects=True, timeout=60)
            if rsp.status_code != 200:
                return False
            with open(filepath, 'wb') as f:
                for chunk in rsp.iter_content(chunk_size=8192):
                    f.write(chunk)
            return os.path.getsize(filepath) > 0
        except Exception:
            return False

    # --- Google Sheets ---
    m = _re.search(r'spreadsheets/d/([a-zA-Z0-9_-]+)', url)
    if m:
        sid = m.group(1)
        target_fp = os.path.join(dest_dir, f"{base_name}.xlsx")
        if os.path.exists(target_fp):
            return target_fp
        fp = ensure_unique_filename(target_fp)
        if _export_stream(f"https://docs.google.com/spreadsheets/d/{sid}/export?format=xlsx", fp):
            return fp
        return None

    # --- Google Docs ---
    m = _re.search(r'document/d/([a-zA-Z0-9_-]+)', url)
    if m:
        did = m.group(1)
        target_fp = os.path.join(dest_dir, f"{base_name}.docx")
        if os.path.exists(target_fp):
            return target_fp
        fp = ensure_unique_filename(target_fp)
        if _export_stream(f"https://docs.google.com/document/d/{did}/export?format=docx", fp):
            return fp
        return None

    # --- Google Slides ---
    m = _re.search(r'presentation/d/([a-zA-Z0-9_-]+)', url)
    if m:
        pid = m.group(1)
        target_fp = os.path.join(dest_dir, f"{base_name}.pptx")
        if os.path.exists(target_fp):
            return target_fp
        fp = ensure_unique_filename(target_fp)
        if _export_stream(f"https://docs.google.com/presentation/d/{pid}/export?format=pptx", fp):
            return fp
        return None

    return None

    # 🔴 先自動開啟所有紅色課程的資料夾
    # - 預設會開
    # - 若在 password.txt 第三行之後任一行加入「nonpop」（或常見誤植 nopop），則不會自動開啟
if red_activities_to_print and not IS_FIRST_TIME:
    try:
        distinct_courses = len({course_path for _name, _link, _course_name, _week_header, course_path, _course_url, _desc in red_activities_to_print})
    except Exception:
        distinct_courses = 0

    # 若 password.txt 沒寫 nonpop/nopop 且本次新資訊課程太多，詢問並可寫入 nonpop
    if distinct_courses >= 4 and AUTO_OPEN_NEW_ACTIVITY_FOLDERS:
        try:
            ans = input(
                f"{YELLOW}! 本次共有 {distinct_courses} 門課有最新資料。\n"
                f"是否關閉自動開啟資料夾（popup），並寫入 password.txt（nonpop）？ (y/N){RESET} "
            ).strip().lower()
            if ans in {"y", "yes"}:
                AUTO_OPEN_NEW_ACTIVITY_FOLDERS = False
                try:
                    append_marker_to_password_file(NONPOP_MARKER)
                except Exception:
                    pass
        except Exception:
            pass

    if AUTO_OPEN_NEW_ACTIVITY_FOLDERS:
        opened_folders = set()
        for name, link, course_name, week_header, course_path, course_url, description in red_activities_to_print:
            if course_path not in opened_folders:
                open_folder(course_path)
                opened_folders.add(course_path)
                time.sleep(0.2)  # 避免同時開啟太多視窗

    time.sleep(0.2)

# 輸出紅色活動資訊並下載
if red_activities_to_print:
    # 如果是第一次使用，跳過這些輸出和下載
    if not IS_FIRST_TIME:
        print("\n" + "="*60)
        print(f"{RED}🔻 以下為新增活動{RESET}")
        print("="*60 + "\n")
    
    total_downloaded_files = 0  # 追蹤總下載檔案數
    
    # 記錄已下載的檔案，避免重複下載（移到迴圈外，所有活動共享）
    downloaded_files = set()
    
    # 記錄需要移除 Zone.Identifier 的文件路徑
    files_to_unblock = []
    
    # 追蹤需要跳過的課程
    skipped_courses = set()
    
    for name, link, course_name, week_header, course_path, course_url, description in red_activities_to_print:
        # 如果這個課程已被標記為跳過，則跳過
        if course_name in skipped_courses:
            continue
        
        # 儲存課程頁卡片上的 activity-description
        if description and description.strip():
            safe_desc_name = "".join(c if c.isalnum() or c in " _-()（）" else "_" for c in name).strip()
            if not safe_desc_name:
                safe_desc_name = "activity"
            safe_desc_name = safe_desc_name[:80]
            desc_file_path = os.path.join(course_path, f"{safe_desc_name}_說明.txt")
            if not os.path.exists(desc_file_path):
                try:
                    with open(desc_file_path, 'w', encoding='utf-8') as df:
                        df.write(description.strip())
                except Exception as e:
                    print(f"{YELLOW}無法儲存活動說明: {e}{RESET}")
        
        # 如果是 URL 類型活動，先獲取實際的外部連結用於顯示
        display_link = link
        if "mod/url/view.php" in link:
            try:
                driver.get(link)
                time.sleep(0.2)
                # 用 JS 快照 href，避免動態 DOM 導致 stale element
                hrefs = driver.execute_script("""
                    return Array.from(document.querySelectorAll('div.urlworkaround a[href]'))
                        .map(a => a.getAttribute('href') || a.href)
                        .filter(Boolean);
                """) or []
                for actual_url in hrefs:
                    if actual_url and not actual_url.startswith("https://elearningv4.nuk.edu.tw"):
                        display_link = actual_url
                        break
            except:
                pass
            
        # 如果是第一次使用，跳過詳細輸出
        if not IS_FIRST_TIME:
            print(f"\n{YELLOW}━━━{RESET} {RED}{name}{RESET} {YELLOW}━━━{RESET}")
            print(f"{PINK}課程：{RESET}{LOWGREEN}{course_name}{RESET}")
            print(f"{PINK}週次：{RESET}{week_header}")
            print(f"{PINK}活動連結：{RESET}{ITALIC}{display_link}{RESET}")
            print(f"{PINK}儲存位置：{RESET}{ITALIC}{course_path}{RESET}\n")
        
        # 確保每次都設定下載路徑到專屬暫存資料夾，避免同名檔案直接覆蓋
        temp_dl_dir = os.path.join(course_path, "_temp_dl")
        os.makedirs(temp_dl_dir, exist_ok=True)
        _register_temp_dl_dir(temp_dl_dir)
        # 移除原有的清空暫存區邏輯，避免把上一次還沒下載完但這次需要接關的檔案刪掉

        if hasattr(driver, "execute_cdp_cmd"):
            driver.execute_cdp_cmd("Page.setDownloadBehavior", {
                "behavior": "allow",
                "downloadPath": os.path.abspath(temp_dl_dir)
            })

        # 🔽 點進活動頁面，抓取所有有 href 的下載連結並打開（會自動觸發下載）
        try:
            # 檢查是否有有效連結
            if link == "（無連結）" or not link or not link.startswith("http"):

                # 回到課程頁面，找出這個活動的圖片
                driver.get(course_url)
                wait = WebDriverWait(driver, 5)
                
                # 用 JS 一次提取活動資源，避免 DOM 更新造成 stale element
                try:
                    session = create_session_with_cookies()
                    assets = extract_activity_assets_from_course_page(driver, name)

                    if assets.get('found'):
                        anchors = assets.get('anchors', [])
                        images = assets.get('images', [])
                        externals = assets.get('externals', [])

                        for img_url in anchors:
                            filename = extract_filename_from_url(img_url)
                            if should_skip_download_filename(filename):
                                continue
                            try:
                                response = session.get(img_url, stream=True)
                                if response.status_code == 200:
                                    file_path = ensure_unique_filename(os.path.join(course_path, filename))
                                    with open(file_path, 'wb') as f:
                                        for chunk in response.iter_content(chunk_size=8192):
                                            f.write(chunk)
                                    remove_zone_identifier(file_path)
                                    downloaded_files.add(filename)
                                    existing_files.add(filename)
                                    total_downloaded_files += 1
                                else:
                                    print(f"{RED}X 下載失敗: HTTP {response.status_code}{RESET}")
                            except Exception as e:
                                print(f"{RED}X 下載失敗: {e}{RESET}")

                        # 若無帶連結的圖片，退而抓取 img[src] 縮圖
                        if not anchors:
                            for img_url in images:
                                filename = extract_filename_from_url(img_url)
                                if should_skip_download_filename(filename):
                                    continue
                                try:
                                    response = session.get(img_url, stream=True)
                                    if response.status_code == 200:
                                        file_path = ensure_unique_filename(os.path.join(course_path, filename))
                                        with open(file_path, 'wb') as f:
                                            for chunk in response.iter_content(chunk_size=8192):
                                                f.write(chunk)
                                        remove_zone_identifier(file_path)
                                        downloaded_files.add(filename)
                                        existing_files.add(filename)
                                        total_downloaded_files += 1
                                    else:
                                        print(f"{RED}X 下載失敗: HTTP {response.status_code}{RESET}")
                                except Exception as e:
                                    print(f"{RED}X 下載失敗: {e}{RESET}")

                        # 同時抓取外部連結（非 pluginfile.php 的 a[href]），存為捷徑
                        for href in externals:
                            safe_filename = "".join(c if c.isalnum() or c in " _-()（）" else "_" for c in name)
                            safe_filename = safe_filename[:100]
                            url_file = os.path.join(course_path, f"{safe_filename}.url")
                            with open(url_file, 'w', encoding='utf-8') as f:
                                f.write(f"[InternetShortcut]\n")
                                f.write(f"URL={href}\n")
                            total_downloaded_files += 1

                except StaleElementReferenceException as e:
                    # 拋出 stale element 異常，讓外層的重試機制處理
                    raise
                except Exception as e:
                    print(f"! 無法下載圖片: {e}")
                
                print()
                processed_successfully = True  # 標記該活動已成功處理
                continue
            
            # 判斷活動類型並處理
            # Case 1: 資源檔案 - 訪問頁面並找到下載連結
            if "mod/resource/view.php" in link:
                try:
                    
                    # 記錄下載目錄中現有的檔案及其修改時間
                    before_files = {}
                    if os.path.exists(temp_dl_dir):
                        for f in os.listdir(temp_dl_dir):
                            fpath = os.path.join(temp_dl_dir, f)
                            if os.path.isfile(fpath):
                                before_files[f] = os.path.getmtime(fpath)
                    
                    driver.get(link)
                    wait = WebDriverWait(driver, 5)
                    
                    # 偵測重導向：Moodle 有時直接把瀏覽器導向 pluginfile.php
                    if 'pluginfile.php' in driver.current_url:
                        redirect_url = driver.current_url
                        redirect_filename = extract_filename_from_url(redirect_url)
                        if should_skip_download_filename(redirect_filename):
                            continue
                        if redirect_filename and not redirect_filename.lower().endswith(('.htm', '.html')):
                            try:
                                session = create_session_with_cookies()
                                rsp = session.get(redirect_url, stream=True)
                                if rsp.status_code == 200:
                                    fp = ensure_unique_filename(os.path.join(course_path, redirect_filename))
                                    with open(fp, 'wb') as f:
                                        for chunk in rsp.iter_content(chunk_size=8192):
                                            f.write(chunk)
                                    remove_zone_identifier(fp)
                                    downloaded_files.add(redirect_filename)
                                    existing_files.add(redirect_filename)
                                    total_downloaded_files += 1
                            except Exception:
                                pass
                            continue
                    
                    # 在資源頁面中找到實際的下載連結（只抓主要內容區的連結）
                    download_links = driver.find_elements(By.CSS_SELECTOR, "div.resourceworkaround a[href*='pluginfile.php']")
                    # print(f"   找到 {len(download_links)} 個下載連結")
                    
                    # 檢查是否有新檔案出現或檔案被更新（可能是自動下載）
                    time.sleep(0.2)  
                    after_files = {}
                    if os.path.exists(temp_dl_dir):
                        for f in os.listdir(temp_dl_dir):
                            fpath = os.path.join(temp_dl_dir, f)
                            if os.path.isfile(fpath):
                                after_files[f] = os.path.getmtime(fpath)
                    
                    # 找出新增或更新的檔案
                    new_or_updated_files = []
                    for filename, mtime in after_files.items():
                        if filename not in before_files:
                            # 全新的檔案
                            new_or_updated_files.append(filename)
                        elif mtime > before_files[filename]:
                            # 檔案被更新（覆蓋）
                            new_or_updated_files.append(filename)
                    
                    
                    # 過濾掉 .crdownload 檔案，只看實際的新檔案
                    actual_new_files = [f for f in new_or_updated_files if not f.endswith('.crdownload')]
                    crdownload_files = [f for f in new_or_updated_files if f.endswith('.crdownload')]
                    
                    downloaded_in_this_activity = 0  # 本次活動下載的檔案數
                    
                    if actual_new_files:
                        for filename in actual_new_files:
                            if should_skip_download_filename(filename):
                                continue
                            
                            temp_file_path = os.path.join(temp_dl_dir, filename)
                            
                            if filename.lower().endswith(('.htm', '.html')):
                                try:
                                    os.remove(temp_file_path)
                                except:
                                    pass
                                continue
                            
                            import shutil
                            # 將檔案移出暫存區並確保檔名不衝突
                            file_path = ensure_unique_filename(os.path.join(course_path, filename))
                            try:
                                shutil.move(temp_file_path, file_path)
                            except Exception as e:
                                print(f"{RED}X 檔案移動失敗: {e}{RESET}")
                                continue
                            
                            remove_zone_identifier(file_path)
                            downloaded_files.add(filename)
                            existing_files.add(filename)
                            total_downloaded_files += 1
                            downloaded_in_this_activity += 1
                            
                            # 立即解壓縮檔案
                            if filename.endswith((".zip", ".rar", ".7z")):
                                actual_size = os.path.getsize(file_path)
                                if actual_size < 100:
                                    print(f"   {YELLOW}! 壓縮檔太小 ({actual_size} bytes)，可能損壞，跳過解壓{RESET}")
                                else:

                                    success = extract_file(file_path, course_path)
                                    if success:
                                        try:
                                            os.remove(file_path)
                                        except:
                                            pass
                                    else:
                                        print(f"   {YELLOW}! 解壓失敗，保留原始檔{RESET}")
                        
                        # 如果有自動下載的檔案，就不需要再找連結了
                        if downloaded_in_this_activity > 0:
                            continue
                    
                    if not download_links:
                        # 可能還有 .crdownload 正在下載
                        if crdownload_files:
                            # print(f"   發現 {len(crdownload_files)} 個正在下載的檔案")
                            # 等待 .crdownload 完成
                            for cr_file in crdownload_files:
                                base_filename = cr_file[:-11]  # 移除 .crdownload
                                if should_skip_download_filename(base_filename):
                                    continue
                                # print(f"   ⏳ 等待下載完成: {base_filename}")
                                temp_file_path = wait_for_download(base_filename, download_path=temp_dl_dir)
                                if temp_file_path and temp_file_path != 'SKIP_COURSE':
                                    if base_filename.lower().endswith(('.htm', '.html')):
                                        try:
                                            os.remove(temp_file_path)
                                        except:
                                            pass
                                        continue
                                    
                                    import shutil
                                    # 將已完成下載的檔案從暫存區移回
                                    file_path = ensure_unique_filename(os.path.join(course_path, base_filename))
                                    try:
                                        shutil.move(temp_file_path, file_path)
                                    except Exception as e:
                                        print(f"{RED}X 檔案移動失敗: {e}{RESET}")
                                        continue
                                    
                                    remove_zone_identifier(file_path)
                                    downloaded_files.add(base_filename)
                                    existing_files.add(base_filename)
                                    total_downloaded_files += 1
                                    
                                    # 解壓縮
                                    if base_filename.endswith((".zip", ".rar", ".7z")):
                                        actual_size = os.path.getsize(file_path)
                                        if actual_size >= 100:

                                            success = extract_file(file_path, course_path)
                                            if success:
                                                try:
                                                    os.remove(file_path)
                                                except:
                                                    pass
                            continue
                        
                        # 後備：掃描頁面內嵌的 pluginfile.php 資源（圖片/PDF inline 顯示時）
                        try:
                            inline_elems = driver.find_elements(By.CSS_SELECTOR,
                                "img[src*='pluginfile.php'], object[data*='pluginfile.php'], embed[src*='pluginfile.php'], a[href*='pluginfile.php']")
                            session = create_session_with_cookies()
                            for elem in inline_elems:
                                res_url = (elem.get_attribute("src") or
                                           elem.get_attribute("data") or
                                           elem.get_attribute("href"))
                                if not res_url:
                                    continue
                                res_filename = extract_filename_from_url(res_url)
                                if should_skip_download_filename(res_filename):
                                    continue
                                if not res_filename or res_filename.lower().endswith(('.htm', '.html')):
                                    continue
                                rsp = session.get(res_url, stream=True)
                                if rsp.status_code == 200:
                                    fp = ensure_unique_filename(os.path.join(course_path, res_filename))
                                    with open(fp, 'wb') as f:
                                        for chunk in rsp.iter_content(chunk_size=8192):
                                            f.write(chunk)
                                    remove_zone_identifier(fp)
                                    downloaded_files.add(res_filename)
                                    existing_files.add(res_filename)
                                    total_downloaded_files += 1
                        except Exception:
                            pass
                        continue
                    
                    for link_elem in download_links:
                        dl_href = link_elem.get_attribute("href")
                        filename = extract_filename_from_url(dl_href)
                        if should_skip_download_filename(filename):
                            continue
                        # print(f"   📎 發現文件: {filename}")
                        
                        # 過濾掉不需要的文件類型（如 downloads.htm）
                        if filename.lower().endswith(('.htm', '.html')):
                            # print(f"   ⏭️  跳過 HTML 文件: {filename}")
                            continue
                        
                        print(f"🔽 開始下載: {filename} (新增活動，覆蓋舊檔)")
                        
                        # 使用 requests 直接下載，確保檔案存到正確位置
                        file_path = None
                        try:
                            # 使用 session 以保持登入狀態
                            session = create_session_with_cookies()
                            
                            # 下載檔案
                            response = session.get(dl_href, stream=True)
                            response.raise_for_status()  # 確保狀態碼正常
                            
                            file_path = ensure_unique_filename(os.path.join(course_path, filename))
                            file_size = 0
                            with open(file_path, 'wb') as f:
                                for chunk in response.iter_content(chunk_size=8192):
                                    if chunk:
                                        f.write(chunk)
                                        file_size += len(chunk)
                            
                            # 檢查檔案大小
                            if file_size == 0 or not os.path.exists(file_path) or os.path.getsize(file_path) == 0:
                                print(f"{RED}X 下載失敗：檔案大小為 0{RESET}")
                                if os.path.exists(file_path):
                                    os.remove(file_path)
                                continue
                            
                            # print(f"{GREEN}✅ 下載完成：{filename} ({file_size / 1024:.1f} KB){RESET}")
                            remove_zone_identifier(file_path)
                            downloaded_files.add(filename)
                            existing_files.add(filename)
                            total_downloaded_files += 1
                            
                            # 立即解壓縮檔案
                            if filename.endswith((".zip", ".rar", ".7z")):
                                # 再次確認檔案大小
                                actual_size = os.path.getsize(file_path)
                                if actual_size < 100:  # 小於100 bytes的壓縮檔很可能是損壞的
                                    print(f"{YELLOW}! 壓縮檔太小 ({actual_size} bytes)，可能損壞，跳過解壓{RESET}")
                                else:

                                    success = extract_file(file_path, course_path)
                                    if success:
                                        os.remove(file_path)

                                    else:
                                        print(f"{YELLOW}   ! 解壓失敗，保留原始檔{RESET}")
                            
                        except Exception as e:
                            print(f"{RED}X 下載失敗: {e}{RESET}")
                            if file_path and os.path.exists(file_path):
                                os.remove(file_path)
                            
                except Exception as e:
                    print(f"{RED}X 下載失敗{RESET}")
                    print(f"   錯誤: {e}")
                    failed_downloads.append({
                        'name': name,
                        'course': course_name,
                        'url': link,
                        'filename': 'unknown'
                    })
                continue
            
            # Case 2: 資料夾或作業 - 進入頁面收集檔案
            if "mod/folder/view.php" in link or "mod/assign/view.php" in link:
                driver.get(link)
                wait = WebDriverWait(driver, 5)
                
  
                
                # 首先檢查當前頁面是否是作業頁面,並下載作業說明的附件
                if "mod/assign/view.php" in link:
                    # 先將作業說明文字存成 txt（包含 activity-description / activity-altcontent）。
                    try:
                        description_text = driver.execute_script("""
                            const targets = Array.from(document.querySelectorAll(
                                "div.activity-description, div.activity-altcontent.activity-description"
                            ));
                            const texts = targets
                                .map(el => (el.innerText || '').trim())
                                .filter(Boolean);
                            return texts.join('\n\n');
                        """) or ""
                        if description_text:
                            safe_desc_name = "".join(c if c.isalnum() or c in " _-()（）" else "_" for c in name).strip()
                            if not safe_desc_name:
                                safe_desc_name = "assignment"
                            safe_desc_name = safe_desc_name[:80]
                            desc_path = os.path.join(course_path, f"{safe_desc_name}_作業說明.txt")
                            with open(desc_path, 'w', encoding='utf-8') as df:
                                df.write(description_text)
                    except Exception:
                        pass
                    
                    # 提取作業說明中嵌入的影片連結（例如 YouTube VideoJS）
                    try:
                        import json as _json
                        # 用 JS 一次取出 data-setup-lazy，避免 WebElement stale
                        video_setups = driver.execute_script("""
                            return Array.from(document.querySelectorAll('div.activity-description [data-setup-lazy]'))
                                .map(el => el.getAttribute('data-setup-lazy'))
                                .filter(Boolean);
                        """) or []
                        saved_video_urls = set()
                        for setup_raw in video_setups:
                            try:
                                setup_data = _json.loads(setup_raw)
                                sources = setup_data.get("sources", [])
                                for src_item in sources:
                                    src_url = src_item.get("src", "")
                                    # HTML 中有時 & 被編碼為 &amp;，需還原
                                    src_url = src_url.replace("&amp;", "&")
                                    if src_url and src_url not in saved_video_urls:
                                        saved_video_urls.add(src_url)
                                        # 產生安全檔名
                                        safe_vname = "".join(c if c.isalnum() or c in " _-()（）" else "_" for c in name)
                                        safe_vname = safe_vname[:80]
                                        # 若有多段影片加序號區別
                                        vindex = len(saved_video_urls)
                                        suffix = f"_影片{vindex}" if vindex > 1 else "_影片"
                                        url_file = os.path.join(course_path, f"{safe_vname}{suffix}.url")
                                        with open(url_file, 'w', encoding='utf-8') as uf:
                                            uf.write(f"[InternetShortcut]\nURL={src_url}\n")
                                        total_downloaded_files += 1
                            except Exception:
                                pass
                    except Exception as e:
                        print(f"! 無法提取影片連結: {e}")
                
                # 用 JS 一次快照所有需要的連結，避免動態 DOM 造成 stale element。
                links_snapshot = driver.execute_script("""
                    const toHrefs = (nodes) => Array.from(nodes)
                        .map(n => n.getAttribute('href') || n.href || '')
                        .filter(Boolean);

                    const submissionLinks = toHrefs(
                        document.querySelectorAll("div[class*='summary_assignsubmission_file'] a[href]")
                    );
                    const pluginLinks = toHrefs(document.querySelectorAll("a[href*='pluginfile.php']"));
                    const forceLinks = toHrefs(document.querySelectorAll("a[href*='forcedownload=1']"));
                    const introAttachmentLinks = toHrefs(
                        document.querySelectorAll("div.activity-description a[href*='introattachment']")
                    );

                    return {
                        submissionLinks,
                        pluginLinks,
                        forceLinks,
                        introAttachmentLinks
                    };
                """) or {}

                submission_links = set(links_snapshot.get("submissionLinks", []))

                # 收集所有 pluginfile.php / forcedownload / introattachment 連結
                file_href_set = set()

                for f_href in links_snapshot.get("pluginLinks", []):
                    if f_href not in submission_links:
                        file_href_set.add(f_href)

                for f_href in links_snapshot.get("forceLinks", []):
                    if f_href not in submission_links and 'pluginfile.php' in f_href:
                        file_href_set.add(f_href)

                for intro_href in links_snapshot.get("introAttachmentLinks", []):
                    file_href_set.add(intro_href)
                
                
                # 處理收集到的檔案連結
                for f_href in file_href_set:
                    filename = extract_filename_from_url(f_href)
                    if should_skip_download_filename(filename):
                        continue
                    try:

                        
                        # 使用 session 直接下載
                        session = create_session_with_cookies()
                        response = session.get(f_href, stream=True)
                        
                        if response.status_code == 200:
                            file_path = ensure_unique_filename(os.path.join(course_path, filename))
                            with open(file_path, 'wb') as f:
                                for chunk in response.iter_content(chunk_size=8192):
                                    if chunk:
                                        f.write(chunk)
                            
                            remove_zone_identifier(file_path)
                            downloaded_files.add(filename)
                            existing_files.add(filename)
                            total_downloaded_files += 1
                        else:
                            print(f"{RED}X 下載失敗: HTTP {response.status_code}{RESET}")
                    except Exception as e:
                        print(f"{RED}X 下載失敗: {filename}{RESET}")
                        print(f"   錯誤: {e}")
                        failed_downloads.append({
                            'name': name,
                            'course': course_name,
                            'url': f_href,
                            'filename': filename
                        })
                continue  # 處理完資料夾/作業後,跳到下一個活動
            
            # Case 3: URL 類型活動 - 提取並顯示實際連結
            if "mod/url/view.php" in link:
                try:
                    driver.get(link)
                    time.sleep(0.2)
                    
                    # 用 JS 快照外部連結，避免 stale element
                    url_links = driver.execute_script("""
                        return Array.from(document.querySelectorAll('div.urlworkaround a[href]'))
                            .map(a => ({
                                href: a.getAttribute('href') || a.href || '',
                                text: (a.textContent || '').trim()
                            }))
                            .filter(x => x.href);
                    """) or []

                    if url_links:
                        for url_link in url_links:
                            actual_url = url_link.get("href", "")
                            link_text = url_link.get("text", "")
                            if actual_url and not actual_url.startswith("https://elearningv4.nuk.edu.tw"):
                                
                                # 清理檔案名稱，移除不合法字元
                                safe_filename = "".join(c if c.isalnum() or c in " _-()（）" else "_" for c in name)
                                # 限制檔案名稱長度，避免過長
                                if len(safe_filename) > 100:
                                    safe_filename = safe_filename[:100]
                                
                                # 檢查是否為 Google 雲端連結
                                is_gdrive = any(d in actual_url for d in (
                                    'drive.google.com', 'docs.google.com/spreadsheets',
                                    'docs.google.com/document', 'docs.google.com/presentation'))
                                if is_gdrive:
                                    try:
                                        session = create_session_with_cookies()
                                        fp = try_download_google_drive(actual_url, course_path, safe_filename, session)
                                        if fp:
                                            remove_zone_identifier(fp)
                                            existing_files.add(os.path.basename(fp))
                                            total_downloaded_files += 1
                                        else:
                                            url_file = os.path.join(course_path, f"{safe_filename}.url")
                                            with open(url_file, 'w', encoding='utf-8') as f:
                                                f.write(f"[InternetShortcut]\nURL={actual_url}\n")
                                            total_downloaded_files += 1
                                    except Exception as e:
                                        print(f"{RED}X 下載失敗: {e}{RESET}")
                                        url_file = os.path.join(course_path, f"{safe_filename}.url")
                                        with open(url_file, 'w', encoding='utf-8') as f:
                                            f.write(f"[InternetShortcut]\nURL={actual_url}\n")
                                        total_downloaded_files += 1
                                else:
                                    # 一般連結，儲存為 Windows 捷徑（可直接點擊開啟）
                                    url_file = os.path.join(course_path, f"{safe_filename}.url")
                                    with open(url_file, 'w', encoding='utf-8') as f:
                                        f.write(f"[InternetShortcut]\nURL={actual_url}\n")
                                    total_downloaded_files += 1

                        
                except Exception as e:
                    print(f"{RED}X 處理 URL 活動失敗: {e}{RESET}")
                continue
            
            # Case 4: 討論區類型活動 - 提取文字內容並儲存
            if "mod/forum/view.php" in link:
                try:
                    driver.get(link)
                    wait = WebDriverWait(driver, 5)
                    
                    # 提取討論區描述內容
                    description_text = ""
                    try:
                        description_divs = driver.find_elements(By.CSS_SELECTOR, "div.activity-description")
                        if description_divs:
                            description_text = description_divs[0].text.strip()
                    except:
                        pass
                    
                    if description_text:
                        # 清理檔案名稱
                        safe_filename = "".join(c if c.isalnum() or c in " _-()（）" else "_" for c in name)
                        if len(safe_filename) > 100:
                            safe_filename = safe_filename[:100]
                        
                        # 儲存為文字檔
                        txt_file = os.path.join(course_path, f"{safe_filename}.txt")
                        with open(txt_file, 'w', encoding='utf-8') as f:
                            f.write(f"{name}\n")
                            f.write("=" * 60 + "\n\n")
                            f.write(description_text)
                        total_downloaded_files += 1
                        existing_files.add(os.path.basename(txt_file))
                        
                        # 從文字中提取 URL 存為捷徑
                        import re as _re
                        _SCHOOL = 'elearningv4.nuk.edu.tw'
                        url_serial = 0
                        seen_turls = set()
                        # 掃描頁面 <a href> 外部連結
                        try:
                            desc_elem = driver.find_elements(By.CSS_SELECTOR, "div.activity-description")[0]
                            for el in desc_elem.find_elements(By.CSS_SELECTOR, "a[href]"):
                                href = el.get_attribute("href") or ""
                                if href.startswith("http") and _SCHOOL not in href and href not in seen_turls:
                                    seen_turls.add(href)
                        except Exception:
                            pass
                        # regex 補捉純文字 URL
                        for furl in _re.findall(r'https?://\S+', description_text):
                            furl = furl.rstrip('.,;:\'")') 
                            if furl and _SCHOOL not in furl and furl not in seen_turls:
                                seen_turls.add(furl)
                        # 儲存捷徑
                        for furl in seen_turls:
                            url_serial += 1
                            usuffix = f"_{url_serial}" if url_serial > 1 else ""
                            ufile = os.path.join(course_path, f"{safe_filename}{usuffix}.url")
                            try:
                                with open(ufile, 'w', encoding='utf-8') as uf:
                                    uf.write(f"[InternetShortcut]\nURL={furl}\n")
                                total_downloaded_files += 1
                            except Exception:
                                pass

                        
                except Exception as e:
                    print(f"{RED}X 處理討論區活動失敗: {e}{RESET}")
                continue
            
            # Case 5: 頁面類型活動 (mod/page) - 儲存文字內容與附件
            if "mod/page/view.php" in link:
                try:
                    safe_pname = "".join(c if c.isalnum() or c in " _-()（）" else "_" for c in name)[:100]
                    
                    # 進入頁面
                    driver.get(link)
                    session = create_session_with_cookies()
                    
                    # 儲存主內容區文字為 .txt（使用 div[role='main'] 精確取主體，排除頁首導覽）
                    try:
                        content_elem = driver.find_element(By.CSS_SELECTOR, "div[role='main']")
                        page_text = content_elem.text.strip()
                        if page_text:
                            txt_file = os.path.join(course_path, f"{safe_pname}.txt")
                            with open(txt_file, 'w', encoding='utf-8') as tf:
                                tf.write(page_text)
                            total_downloaded_files += 1
                    except Exception:
                        page_text = ""
                    
                    # 收集頁面中所有外部連結（<a href> + 文字中的 URL）
                    import re as _re
                    _SCHOOL = 'elearningv4.nuk.edu.tw'
                    external_urls = {}  # url -> base_name
                    
                    # 方式1：掃描 <a href> 外部連結
                    try:
                        main_elem2 = driver.find_element(By.CSS_SELECTOR, "div[role='main']")
                        ext_links = main_elem2.find_elements(By.CSS_SELECTOR, "a[href]")
                        for el in ext_links:
                            href = el.get_attribute("href") or ""
                            if (href.startswith("http") and
                                    _SCHOOL not in href and
                                    "pluginfile.php" not in href and
                                    "/theme_" not in href and
                                    href not in external_urls):
                                external_urls[href] = href
                    except Exception:
                        pass
                    
                    # 方式2：從文字中 regex 抓取 URL（補捉純文字貼上的連結）
                    if page_text:
                        for furl in _re.findall(r'https?://\S+', page_text):
                            furl = furl.rstrip('.,;:\'\")')
                            if furl and _SCHOOL not in furl and furl not in external_urls:
                                external_urls[furl] = furl
                    
                    # 儲存外部連結（Google Drive 嘗試下載，其餘存捷徑）
                    url_serial = 0
                    for furl in external_urls:
                        url_serial += 1
                        usuffix = f"_{url_serial}" if url_serial > 1 else ""
                        base_name = f"{safe_pname}{usuffix}"
                        is_gdrive = any(d in furl for d in (
                            'drive.google.com', 'docs.google.com/spreadsheets',
                            'docs.google.com/document', 'docs.google.com/presentation'))
                        try:
                            if is_gdrive:
                                fp = try_download_google_drive(furl, course_path, base_name, session)
                                if fp:
                                    remove_zone_identifier(fp)
                                    downloaded_files.add(os.path.basename(fp))
                                    existing_files.add(os.path.basename(fp))
                                    total_downloaded_files += 1
                                    continue
                            # 一般連結 or Google Drive 下載失敗 → 存捷徑
                            ufile = os.path.join(course_path, f"{base_name}.url")
                            with open(ufile, 'w', encoding='utf-8') as uf:
                                uf.write(f"[InternetShortcut]\nURL={furl}\n")
                            total_downloaded_files += 1
                        except Exception:
                            pass
                    
                    # 下載主內容區內的 pluginfile.php 附件與圖片（排除 theme/icon 等）
                    try:
                        main_elem = driver.find_element(By.CSS_SELECTOR, "div[role='main']")
                        page_res_elems = main_elem.find_elements(By.CSS_SELECTOR,
                            "a[href*='pluginfile.php'], img[src*='pluginfile.php'], "
                            "object[data*='pluginfile.php'], embed[src*='pluginfile.php']")
                    except Exception:
                        page_res_elems = []
                    seen_res_urls = set()
                    for pfe in page_res_elems:
                        pfl_href = (pfe.get_attribute("href") or
                                    pfe.get_attribute("src") or
                                    pfe.get_attribute("data"))
                        if not pfl_href or pfl_href in seen_res_urls:
                            continue
                        seen_res_urls.add(pfl_href)
                        # 排除布景主題 icon / logo（非課程內容）
                        if '/theme_' in pfl_href or '/theme/' in pfl_href:
                            continue
                        pfl_name = extract_filename_from_url(pfl_href)
                        if should_skip_download_filename(pfl_name):
                            continue
                        if not pfl_name or pfl_name.lower().endswith(('.htm', '.html')):
                            continue
                        try:
                            rsp = session.get(pfl_href, stream=True)
                            if rsp.status_code == 200:
                                fp = ensure_unique_filename(os.path.join(course_path, pfl_name))
                                with open(fp, 'wb') as f:
                                    for chunk in rsp.iter_content(chunk_size=8192):
                                        f.write(chunk)
                                remove_zone_identifier(fp)
                                downloaded_files.add(pfl_name)
                                existing_files.add(pfl_name)
                                total_downloaded_files += 1
                        except Exception:
                            pass
                    
                    # 提取 VideoJS 嵌入影片連結
                    import json as _json
                    # 用 JS 一次取出 data-setup-lazy，避免 WebElement stale
                    video_setups_p = driver.execute_script("""
                        return Array.from(document.querySelectorAll('[data-setup-lazy]'))
                            .map(el => el.getAttribute('data-setup-lazy'))
                            .filter(Boolean);
                    """) or []
                    saved_vurls = set()
                    for setup_raw in video_setups_p:
                        try:
                            setup_data = _json.loads(setup_raw)
                            for src_item in setup_data.get("sources", []):
                                src_url = src_item.get("src", "").replace("&amp;", "&")
                                if src_url and src_url not in saved_vurls:
                                    saved_vurls.add(src_url)
                                    vidx = len(saved_vurls)
                                    vsuffix = f"_影片{vidx}" if vidx > 1 else "_影片"
                                    vfile = os.path.join(course_path, f"{safe_pname}{vsuffix}.url")
                                    with open(vfile, 'w', encoding='utf-8') as vf:
                                        vf.write(f"[InternetShortcut]\nURL={src_url}\n")
                                    total_downloaded_files += 1
                                    print(f"🎬 {BLUE}影片連結: {src_url}{RESET}")
                        except Exception:
                            pass
                except Exception as e:
                    print(f"{RED}X 處理頁面活動失敗: {e}{RESET}")
                continue
            
            # Case 6: 其他未知類型 - 略過

            
        except StaleElementReferenceException:
            # stale element：DOM 自動更新導致元素失效
            # 對策：延遲後跳過，下一輪執行時會重試此活動
            time.sleep(2)
        except Exception as e:
            # print(f"{RED}X 處理活動時發生錯誤: {e}{RESET}")
            pass

    # 刪除本課程因下載而產生的暫存資料夾（確保結束後能清空）
    import shutil
    if os.path.exists(temp_dl_dir):
        shutil.rmtree(temp_dl_dir, ignore_errors=True)
    try:
        _TEMP_DL_DIRS.discard(temp_dl_dir)
    except Exception:
        pass

# 資源檔案使用 requests 直接下載，不會產生 .crdownload
# 資料夾/作業下載若有問題，wait_for_download() 會在當下處理

# 顯示下載失敗的連結
if failed_downloads and not IS_FIRST_TIME:
    print("\n" + "="*60)
    print(f"{RED}X 以下檔案下載失敗，請手動下載：{RESET}")
    print("="*60)
    for item in failed_downloads:
        print(f"\n📌 {item['name']}")
        print(f"   課程: {item['course']}")
        print(f"   檔名: {item['filename']}")
        print(f"   {BLUE}下載連結: {item['url']}{RESET}")

extracted_count = 0
failed_extract = []  # 記錄解壓失敗的檔案
for root, dirs, files in os.walk(download_dir):
    for file in files:
        filepath = os.path.join(root, file)
        if file.endswith((".zip", ".rar", ".7z")):
            success = extract_file(filepath, root)
            if success:
                os.remove(filepath)
                extracted_count += 1
            else:
                # 記錄失敗的檔案（特別是 RAR）
                if file.endswith(".rar"):
                    failed_extract.append(filepath)


if failed_extract and not IS_FIRST_TIME:
    print(f"\n{YELLOW}!  以下檔案因工具缺失而未解壓：{RESET}")
    for f in failed_extract:
        print(f"   - {os.path.basename(f)}")
    print(f"\n💡 建議安裝 patool（自動支持多種解壓工具）：")
    print(f"   {BLUE}pip install patool{RESET}")
    print(f"   或手動下載 UnRAR: https://www.rarlab.com/rar_add.htm")

# 按照課程名稱字母順序整理輸出並更新 output.txt
course_results.sort(key=lambda x: x[1]['course_name'])  # 按課程名稱字母排序
all_output_lines_sorted = []
for idx, result in course_results:
    all_output_lines_sorted.extend(result['output'])

with open(OUTPUT_FILE, "w", encoding="utf-8") as f:
    f.write("\n".join(all_output_lines_sorted))

# 如果是第一次使用，跳過互動選單並直接結束
if IS_FIRST_TIME:
    clear_builderror_marker(USERNAME, PASSWORD)
    print(f"\n{GREEN}環境建置完成{RESET}")
    print(f"\n{YELLOW}可在下次上課前再次執行此程式{RESET}")
    if platform.system() == "Darwin":
        time.sleep(1)
        close_macos_terminal_and_exit(0)
    a=input()
    # 程式結束前關閉所有分頁（在背景線程中執行避免卡頓）
    def cleanup_driver():
        try:
            driver.quit()
        except:
            pass
    
    cleanup_thread = threading.Thread(target=cleanup_driver, daemon=True)
    cleanup_thread.start()
    close_macos_terminal_and_exit(0)

import datetime

# 1) 從文字檔載入上次未繳交作業（加速）
submitted_assignments = load_submitted_assignments()
pending_cache = load_pending_assignments()

# 清掉已過期（不顯示），並且把它們加入已處理快取（submitted_assignments）
# 以避免未來程式又把它們當成新作業點進去驗證！
_now_dt = datetime.datetime.now()
_new_pending = {}
_newly_ignored = {}
for k, v in pending_cache.items():
    if is_expired_assignment(v, _now_dt):
        _newly_ignored[k] = {
            'course': v.get('course', ''),
            'name': v.get('name', ''),
            'url': v.get('url', ''),
            'assignment_key': k,
            'checked_date': time.strftime('%Y-%m-%d %H:%M:%S'),
            'status': 'expired'
        }
    else:
        _new_pending[k] = v

pending_cache = _new_pending
if _newly_ignored:
    submitted_assignments.update(_newly_ignored)
    try:
        save_submitted_assignments(submitted_assignments)
    except Exception:
        pass

# 先把「其實已繳交」的從 pending_cache 清掉（用 id 比對，避免課程名變動導致殘留）
try:
    for _k, _it in list(pending_cache.items()):
        _course = _it.get('course') or ''
        _url = _it.get('url') or ''
        _key = build_assignment_key(_course, _url)
        if has_submitted_record(submitted_assignments, _course, _url, _key):
            pending_cache.pop(_k, None)
except Exception:
    pass

# 2) 從本次「最新消息」抓到的課程頁資料中拿到作業清單（不額外開分頁）
all_assignments_found = []
for _idx, _res in course_results:
    for a in (_res.get('assignments') or []):
        if isinstance(a, dict) and a.get('url') and a.get('name'):
            all_assignments_found.append(a)

# 3) 限量檢查新作業（不在 pending/submitted 的），用主 driver 直接進入頁面判斷
try:
    # 預設：全部檢查（只會檢查「不在 pending/submitted」的新作業）。
    # 若要加速，可設定環境變數 ASSIGNMENT_CHECK_LIMIT。
    try:
        _env_check_limit = os.environ.get('ASSIGNMENT_CHECK_LIMIT', '').strip()
    except Exception:
        _env_check_limit = ''
    if _env_check_limit:
        _check_limit = int(_env_check_limit)
    else:
        _check_limit = len(all_assignments_found or [])

    submitted_assignments, pending_cache = check_assignments_inline(
        driver,
        all_assignments_found,
        submitted_assignments=submitted_assignments,
        pending_cache=pending_cache,
        limit=_check_limit
    )
    save_submitted_assignments(submitted_assignments)
except Exception:
    # 檢查失敗不阻擋主流程
    pass

# 3.5) 少量 recheck：避免「已繳交但仍殘留在未繳」
try:
    _env_limit = os.environ.get('PENDING_RECHECK_LIMIT', '').strip().lower()
    if not _env_limit:
        # 預設：全數 recheck（最可靠）
        recheck_limit = None
    elif _env_limit == 'all':
        recheck_limit = None
    elif _env_limit == 'auto':
        # 自動模式：pending 不多則全數，過多則只檢查最接近截止的前 N 筆
        _pending_n = len(pending_cache or {})
        recheck_limit = _pending_n if _pending_n <= 30 else 30
    else:
        recheck_limit = int(_env_limit)
except Exception:
    recheck_limit = None
try:
    submitted_assignments, pending_cache = recheck_pending_assignments(
        driver,
        pending_cache=pending_cache,
        submitted_assignments=submitted_assignments,
        limit=recheck_limit,
    )
    save_submitted_assignments(submitted_assignments)
except Exception:
    pass

# 4) 組出要顯示的未繳作業清單（只顯示未過期）
empty_assignments = list(pending_cache.values())
empty_assignments = [a for a in empty_assignments if not is_expired_assignment(a, _now_dt)]
empty_assignments.sort(key=lambda x: x.get('due_date_obj', datetime.datetime.max) if x.get('due_date_obj') else datetime.datetime.max)

# 如果沒有任何「未繳且未過期」作業：不顯示輸入編號提示，直接結束
if not empty_assignments:
    try:
        save_pending_assignments(pending_cache)
    except Exception:
        pass
    print(f"{GREEN}目前無須繳交作業{RESET}")
    close_macos_terminal_and_exit(0)

print(f"{PINK}開啟未繳交作業（可用空白分隔多個編號）：{RESET}\n")

items_list = []
current_idx = 1
ibxx = 0

if empty_assignments:
    for item in empty_assignments:
        items_list.append({'type': 'assignment', 'data': item})
        spacing = "  " if current_idx <= 9 else " "
        color = MIKU if ibxx % 2 == 0 else BBLUE
        due_str = f" {item['due_date_str']}" if item.get('due_date_str') else ""
        print(f"  {color}{current_idx}.{spacing}[{item.get('course','')}] {item['name']}{RED}{due_str}{RESET}")
        current_idx += 1
        ibxx += 1

# 5) 輸出完後，把不符合規定（未繳交且未過期）的作業寫回文字檔
save_pending_assignments(pending_cache)

choice = input(f"\n{PINK}請輸入編號: {RESET}").strip().lower()

if not choice:
    close_macos_terminal_and_exit(0)

choice_parts = choice.split()
selected_assignments = []

for part in choice_parts:
    if part.isdigit():
        idx = int(part) - 1
        if 0 <= idx < len(items_list):
            item = items_list[idx]
            selected_assignments.append(item['data'])

if selected_assignments:
    # 關閉 headless driver，改用可見模式
    try:
        driver.quit()
    except:
        pass
    
    # 重新啟動非 headless 模式
    chrome_options_visible = Options()
    chrome_options_visible.add_argument("--log-level=3")
    chrome_options_visible.add_experimental_option("excludeSwitches", ["enable-logging"])
    chrome_options_visible.add_argument("--disable-gpu")
    chrome_options_visible.add_argument("--disable-dev-shm-usage")
    chrome_options_visible.add_argument("--remote-debugging-pipe")
    chrome_options_visible.add_argument("--disable-software-rasterizer")
    if os.name == 'nt':
        chrome_options_visible.add_argument("--no-sandbox")       
        chrome_options_visible.add_argument("--disable-features=RendererCodeIntegrity")
    apply_chrome_binary_option(chrome_options_visible)
    try:
        driver = create_webdriver(chrome_options_visible, hide_windows_console=True)
    except WebDriverException as e:
        raise

    # 重新登入
    if not ensure_logged_in_retry_once(driver, USERNAME, PASSWORD, silent=False, max_retries=2):
        try:
            driver.quit()
        except Exception:
            pass
        sys.exit()

    for idx, assignment in enumerate(selected_assignments, 1):    
        if idx == 1:
            driver.get(assignment['url'])
        else:
            driver.execute_script(f"window.open(\"{assignment['url']}\", '_blank');")
            driver.switch_to.window(driver.window_handles[-1])    
        
        time.sleep(0.5)

        # 若已繳交（或老師允許重複繳交），不強制點「繳交作業」；只要能打開頁面並顯示狀態即可。
        try:
            status_text = ""
            try:
                status_cell = driver.find_element(By.XPATH, "//th[contains(text(), '繳交狀態')]/following-sibling::td")
                status_text = (status_cell.text or "").strip()
            except Exception:
                status_text = ""

            if status_text:
                # 常見：已繳交作業 / 已繳交 / 已提交
                if any(k in status_text for k in ["已繳交", "已提交", "已送出"]):
                    print(f"  {GREEN}✓ 已繳交{RESET} [{assignment.get('course','')}] {assignment.get('name','')}")
                    try:
                        _course = assignment.get('course') or ''
                        _url = assignment.get('url') or ''
                        _name = assignment.get('name') or ''
                        _akey = build_assignment_key(_course, _url)
                        submitted_assignments[_akey] = {
                            'course': _course,
                            'name': _name,
                            'url': _url,
                            'assignment_key': _akey,
                            'checked_date': time.strftime('%Y-%m-%d %H:%M:%S')
                        }
                        pending_cache.pop(_akey, None)

                        # 若 pending_cache 還是舊 key（例如 course+url），再用 URL / id 進一步清掉
                        try:
                            import re as _re
                            _aid = None
                            m = _re.search(r"[?&]id=(\d+)", _url)
                            if m:
                                _aid = m.group(1)

                            for _pk, _pv in list((pending_cache or {}).items()):
                                _pu = ((_pv or {}).get('url') or '').strip()
                                if _pu and _pu == _url:
                                    pending_cache.pop(_pk, None)
                                    continue
                                if _aid and _pu:
                                    m2 = _re.search(r"[?&]id=(\d+)", _pu)
                                    if m2 and m2.group(1) == _aid:
                                        pending_cache.pop(_pk, None)
                        except Exception:
                            pass
                    except Exception:
                        pass
                    continue

            # 仍可嘗試點擊「繳交作業」按鈕（未繳交時）
            submit_button = WebDriverWait(driver, 5).until(
                EC.element_to_be_clickable((By.XPATH, "//button[contains(text(), '繳交作業')]"))
            )
            submit_button.click()
            time.sleep(0.5)
        except Exception:
            # 不阻擋：保持頁面已開啟即可
            print(f"  {YELLOW}! 無法進入繳交介面（可能已繳交或版面不同），但已開啟作業頁面{RESET}")

    enter_pressed = threading.Event()
    def wait_for_enter_bg():
        input()
        enter_pressed.set()
    input_thread = threading.Thread(target=wait_for_enter_bg, daemon=True)
    input_thread.start()

    while True:
        try:
            driver.current_url
            if enter_pressed.is_set():
                break
            time.sleep(0.5)
        except:
            break

    try:
        driver.quit()
    except:
        pass

    # 將已繳交更新寫回快取（避免下次仍顯示未繳）
    try:
        save_submitted_assignments(submitted_assignments)
        save_pending_assignments(pending_cache)
    except Exception:
        pass
    sys.exit()

close_macos_terminal_and_exit(0)
