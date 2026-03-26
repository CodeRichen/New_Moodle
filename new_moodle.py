# pip install selenium webdriver-manager colorama requests py7zr patool rarfile gdown
# 不要在開啟Moodle網頁的狀態執行程式

import os
import sys
import ctypes

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

os.environ['WDM_LOG_LEVEL'] = '0' # 針對 webdriver-manager 的日誌屏蔽
os.environ['TF_CPP_MIN_LOG_LEVEL'] = '3'  #關閉 TensorFlow 的日誌
import subprocess
import platform
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import StaleElementReferenceException, SessionNotCreatedException, WebDriverException, TimeoutException
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

# 全域變數儲存 ChromeDriver 路徑，避免重複下載
_cached_driver_path = None

def get_chrome_driver_path():
    """獲取 ChromeDriver 路徑，包含重試機制和回退方案"""
    global _cached_driver_path

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
    """在 macOS 且找不到 Chrome 時，改用 Safari。"""
    force_chrome = os.environ.get("FORCE_CHROME", "0").strip().lower() in {"1", "true", "yes", "on"}
    return platform.system() == "Darwin" and not force_chrome and get_chrome_binary_path() is None

def _is_chrome_startup_issue(exc):
    """判斷是否屬於可透過更換 profile 重試的 Chrome 啟動問題。"""
    msg = str(exc)
    keywords = ["DevToolsActivePort", "Chrome failed to start", "session not created", "chrome not reachable"]
    return any(k in msg for k in keywords)

def _prepare_windows_chrome_options(chrome_options):
    """Windows 啟動穩定化：補齊必要參數。"""
    if os.name != 'nt' or chrome_options is None:
        return chrome_options

    existing_args = set(chrome_options.arguments)

    def _add_arg(arg):
        if arg not in existing_args:
            chrome_options.add_argument(arg)
            existing_args.add(arg)

    _add_arg("--remote-debugging-port=0")
    _add_arg("--disable-gpu")
    _add_arg("--disable-dev-shm-usage")
    _add_arg("--no-sandbox")
    _add_arg("--disable-features=RendererCodeIntegrity")
    _add_arg("--no-first-run")
    _add_arg("--no-default-browser-check")

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

    # Windows 偶發 DevToolsActivePort 問題時，自動換 profile 重試一次。
    if os.name == 'nt':
        try:
            return _create_chrome_driver()
        except WebDriverException as e:
            if _is_chrome_startup_issue(e):
                retry_options = Options()
                if chrome_options is not None:
                    for arg in chrome_options.arguments:
                        retry_options.add_argument(arg)
                    retry_options.page_load_strategy = chrome_options.page_load_strategy
                    retry_options.binary_location = chrome_options.binary_location
                    for key, value in chrome_options.experimental_options.items():
                        retry_options.add_experimental_option(key, value)
                _prepare_windows_chrome_options(retry_options)
                try:
                    return _create_chrome_driver(retry_options)
                except WebDriverException as second_error:
                    # 無固定 profile 仍無法啟動時，退回一次臨時 profile 做最後保底。
                    if _is_chrome_startup_issue(second_error):
                        fallback_options = Options()
                        for arg in retry_options.arguments:
                            if not arg.startswith("--user-data-dir="):
                                fallback_options.add_argument(arg)
                        fallback_options.page_load_strategy = retry_options.page_load_strategy
                        fallback_options.binary_location = retry_options.binary_location
                        for key, value in retry_options.experimental_options.items():
                            fallback_options.add_experimental_option(key, value)

                        fallback_dir = tempfile.mkdtemp(prefix="chrome_profile_fallback_", dir=BASE_DOWNLOAD_DIR)
                        fallback_options.add_argument(f"--user-data-dir={fallback_dir}")
                        return _create_chrome_driver(fallback_options)
                    raise
            raise

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

# 根據 BASE_DOWNLOAD_DIR 自動設定其他檔案路徑
OUTPUT_FILE = os.path.join(BASE_DOWNLOAD_DIR, "cless.txt")
SUBMITTED_ASSIGNMENTS_FILE = os.path.join(BASE_DOWNLOAD_DIR, "submitted_assignments.json")
PASSWORD_FILE = os.path.join(BASE_DOWNLOAD_DIR, "password.txt")
BUILDERROR_MARKER = "builderror"
NONPOP_MARKER = "nonpop"

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
    try:
        # 創建臨時瀏覽器進行登入測試
        test_chrome_options = Options()
        if os.name == 'nt':
            test_chrome_options.add_argument("--headless=new")
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
        
        # 使用現有的登入函數邏輯
        test_driver.get("https://elearningv4.nuk.edu.tw/login/index.php?loginredirect=1")
        WebDriverWait(test_driver, 10).until(
            EC.visibility_of_element_located((By.ID, "username"))
        ).send_keys(username)
        
        # 使用現有的 simulate_typing 函數邏輯
        password_script = f"""
        var element = document.getElementById('password');
        element.focus();
        element.value = '';
        
        // 模擬逐字輸入
        var text = '{password}';
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
        test_driver.execute_script(password_script)
        test_driver.execute_script("document.getElementById('loginbtn').click();")
        time.sleep(2)  # 等待登入處理
        
        # 檢查是否有錯誤警告
        try:
            error_alert = test_driver.find_element(By.CSS_SELECTOR, "div.alert.alert-danger")
            if "帳號不存在或密碼錯誤" in error_alert.text:
                test_driver.quit()
                return False
        except:
            pass
        
        # 檢查是否被重新導向到登入頁面（表示登入失敗）
        current_url = test_driver.current_url
        if "login" in current_url:
            test_driver.quit()
            return False
        
        # 登入成功
        test_driver.quit()
        return True
        
    except Exception as e:
        try:
            test_driver.quit()
        except:
            pass
        return False

def setup_chrome():
    """在 macOS 上檢查並自動安裝 Chrome。"""
    if platform.system() != "Darwin":
        return True  # 非 macOS 不需要執行

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
    existing_binary = get_chrome_binary_path()
    if existing_binary and is_chrome_usable(existing_binary):
        if can_start_chrome_webdriver(existing_binary):
            return True
        print(f"{YELLOW}! Chrome 可啟動，但 WebDriver 建立失敗，將自動修復{RESET}")
        reset_chromedriver_cache()
    if existing_binary and not is_chrome_usable(existing_binary):
        print(f"{YELLOW}! 偵測到 Chrome 但無法啟動，將自動重新下載與修復{RESET}")
    
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
    global IS_FIRST_TIME, AUTO_OPEN_NEW_ACTIVITY_FOLDERS  # 使用全域變數來追蹤狀態
    
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
                    has_nonpop = NONPOP_MARKER in markers
                    if has_builderror:
                        IS_FIRST_TIME = True
                    else:
                        IS_FIRST_TIME = False
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
 
                if test_login(username, password):

                    break
                else:
                    print(f"\n{RED}X 帳號不存在或密碼錯誤，請重新輸入{RESET}")
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
    os._exit(code)

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

chrome_options = Options()
if os.name == 'nt':
    chrome_options.add_argument("--headless=new")  # Windows 較穩定
else:
    chrome_options.add_argument("--headless")  # 無頭模式
chrome_options.add_argument("--disable-gpu")  # 防止顯示錯誤
chrome_options.add_argument("--log-level=3")  # 降低日誌等級，避免雜訊
chrome_options.add_argument("--window-size=1920,1080")  # 設定解析度，避免元素渲染錯誤
chrome_options.add_argument("--disable-extensions")  # 禁用擴充功能
chrome_options.add_argument("--disable-dev-shm-usage")  # 解決資源限制問題
chrome_options.add_argument("--no-sandbox")  # 加快啟動速度
chrome_options.add_argument("--disable-blink-features=AutomationControlled")  # 避免被偵測
# exe優化：減少記憶體使用和加速啟動
chrome_options.add_argument("--disable-background-timer-throttling")
chrome_options.add_argument("--disable-backgrounding-occluded-windows")
chrome_options.add_argument("--disable-renderer-backgrounding")
chrome_options.add_argument("--disable-features=TranslateUI,VizDisplayCompositor")
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

def login_to_moodle(driver):
    """執行登入流程"""
    driver.get("https://elearningv4.nuk.edu.tw/login/index.php?loginredirect=1")
    WebDriverWait(driver, 5).until(  # 減少等待時間
        EC.visibility_of_element_located((By.ID, "username"))
    ).send_keys(USERNAME)
    simulate_typing(driver, 'password', PASSWORD)
    driver.execute_script("document.getElementById('loginbtn').click();")
    time.sleep(0.2)  # 減少等待時間

def select_ongoing_courses_filter(driver, wait):
    """進入我的課程後，切換分組到「進行中」（TODO:若要下載全部則 Ctrl+h 將"進行中"取代為"過去"）"""
    try:
        # 先判斷當前是否已是「進行中」
        current_text = driver.execute_script("""
            const btn = document.getElementById('groupingdropdown');
            if (!btn) return '';
            const span = btn.querySelector('span[data-active-item-text]');
            return (span ? span.innerText : btn.innerText) || '';
        """)
        if "進行中" in (current_text or ""):
            return True

        dropdown_btn = wait.until(EC.element_to_be_clickable((By.ID, "groupingdropdown")))
        driver.execute_script("arguments[0].click();", dropdown_btn)
        time.sleep(0.2)

        option_clicked = driver.execute_script("""
            const candidates = Array.from(document.querySelectorAll('.dropdown-menu a, .dropdown-menu button, .dropdown-item'));
            const target = candidates.find(el => ((el.innerText || '').trim() === '進行中') || (el.innerText || '').includes('進行中'));
            if (!target) return false;
            target.click();
            return true;
        """)
        if option_clicked:
            time.sleep(0.2)
        return bool(option_clicked)
    except Exception:
        return False

# 登入重試機制（最多2次）

login_success = False
for attempt in range(2):
    
    login_to_moodle(driver)
    
    wait = WebDriverWait(driver, 5)  # 減少等待時間
    
    # 嘗試進入「我的課程」
    try:
        my_courses_link = wait.until(EC.presence_of_element_located((By.PARTIAL_LINK_TEXT, "我的課程")))
        my_courses_link.click()
        time.sleep(0.2)  # 減少等待時間
        select_ongoing_courses_filter(driver, wait)
        
        # 檢查是否被重定向回登入頁面
        current_url = driver.current_url
        if "login" in current_url:
            if attempt < 1:  # 還有重試機會
                continue
            else:  # 最後一次也失敗
                print(f"\n{RED}{'='*60}{RESET}")
                print(f"{RED}X 登入失敗：無法進入系統{RESET}")
                print(f"\n按 Enter 鍵離開...")
                input()
                driver.quit()
                sys.exit(1)
        else:
            # 登入成功
            login_success = True
            break
            
    except Exception as e:
        if attempt < 1:  # 還有重試機會
            continue
        else:  # 最後一次也失敗
            print(f"\n{RED}{'='*60}{RESET}")
            print(f"{RED}X 登入失敗：無法連接到 Moodle {RESET}")
            print(f"{RED}{'='*60}{RESET}")
            print(f"\n錯誤訊息：{e}")
            print(f"\n按 Enter 鍵離開...")
            input()
            driver.quit()
            sys.exit(1)

if not login_success:
    print(f"\n{RED}X 登入失敗{RESET}")
    print(f"按 Enter 鍵離開...")
    input()
    driver.quit()
    sys.exit(1)

def snapshot_course_hrefs(driver):
    """用 JS 快照課程連結，降低前端版型變更造成的定位失敗。"""
    script = """
    const nodes = Array.from(document.querySelectorAll("a[href*='/course/view.php?id=']"));
    const hrefs = nodes
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
        select_ongoing_courses_filter(driver, wait)
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

all_output_lines = []
red_activities_to_print = []  # 暫存紅色活動（(name, link, course_name, week_header, course_path, course_url)）
failed_downloads = []  # 儲存下載失敗的連結

# 線程鎖，用於同步輸出和資料收集
output_lock = threading.Lock()
data_lock = threading.Lock()

# 預先為所有課程開啟分頁
main_window = driver.current_window_handle  # 保存主視窗

for idx, href in enumerate(course_hrefs, 1):
    if idx == 1:
        # 第一個課程直接在當前分頁載入
        driver.get(href)
    else:
        # 其他課程開新分頁
        driver.execute_script(f"window.open('{href}', '_blank');")

# 獲取所有分頁的 handle
all_tabs = driver.window_handles

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
                                # 活動名稱維持黃色，網址改為白色（預設）
                                print(f"  {YELLOW}{clean_name}{RESET} - {href_link}")
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
        subprocess.run(
            ["powershell", "-Command", f"Unblock-File -Path '{filepath}'"],
            capture_output=True,
            timeout=5
        )
    except Exception:
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
    try:
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
                            # 嘗試 GBK 編碼
                            fixed_filename = info.filename.encode('cp437').decode('gbk')
                            info.filename = fixed_filename
                        except:
                            # 如果都失敗，保持原檔名
                            pass
                    
                    # 解壓縮單個檔案
                    zf.extract(info, dest_dir)
                    
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
                            [seven_zip, "x", file_path, f"-o{dest_dir}", "-y"],
                            capture_output=True,
                            text=True,
                            timeout=60
                        )
                        if result.returncode == 0:
                            return True
                    except Exception as e:
                        print(f"{YELLOW}!  7-Zip 解壓失敗: {e}{RESET}")
            
            # 方法2: 優先使用 patool（更穩定，支持多種後端）
            if HAS_PATOOL:
                try:
                    patool.extract_archive(file_path, outdir=dest_dir, verbosity=-1)
                    return True
                except Exception as e:
                    print(f"{YELLOW}!  patool 解壓失敗: {e}{RESET}")
            
            # 方法3: 使用 rarfile
            if HAS_RARFILE:
                try:
                    with rarfile.RarFile(file_path, 'r') as rf:
                        rf.extractall(dest_dir)
                    return True
                except Exception as e:
                    print(f"{YELLOW}!  rarfile 解壓失敗: {e}{RESET}")
            
            # 都失敗，顯示安裝提示
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
                sz.extractall(dest_dir)
        else:
            return False
        return True
    except zipfile.BadZipFile:
        print(f"   {YELLOW}! 無效的 ZIP 檔案格式，跳過解壓{RESET}")
        return False
    except Exception as e:
        print(f"X 解壓失敗: {os.path.basename(file_path)}, 原因: {e}")
        return False

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
if red_activities_to_print and not IS_FIRST_TIME and AUTO_OPEN_NEW_ACTIVITY_FOLDERS:
   

    
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
                                    files_to_unblock.append(file_path)
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
                                        files_to_unblock.append(file_path)
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
                                    files_to_unblock.append(fp)
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
                            
                            files_to_unblock.append(file_path)
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
                            # print(f"   ⏳ 發現 {len(crdownload_files)} 個正在下載的檔案")
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
                                    
                                    files_to_unblock.append(file_path)
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
                                    files_to_unblock.append(fp)
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
                            files_to_unblock.append(file_path)
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
                            
                            files_to_unblock.append(file_path)
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
                                            files_to_unblock.append(fp)
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
                                    files_to_unblock.append(fp)
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
                                files_to_unblock.append(fp)
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

    # 刪除各課程因下載而產生的暫存資料夾（如果裡面是空的，代表這堂課的所有檔案都處理完成了；如果還有檔案暫存中代表未完成，則保留至下次）
    for course_path in set(item[4] for item in red_activities_to_print):
        tmp_dir = os.path.join(course_path, "_temp_dl")
        if os.path.exists(tmp_dir):
            try:
                # 只有資料夾內完全沒有檔案或資料夾時才會被刪除
                if not os.listdir(tmp_dir):
                    os.rmdir(tmp_dir)
            except:
                pass

    # 所有下載完成後，統一移除 Zone.Identifier
    # print(f"\n📊 下載統計：共下載 {total_downloaded_files} 個檔案")
    # print(f"🔍 待解除封鎖的檔案數量：{len(files_to_unblock)}")
    
    
    if files_to_unblock:
        for file_path in files_to_unblock:
            try:
                remove_zone_identifier(file_path)
            except Exception as e:
                # 靜默失敗，不影響主流程
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

# 最終保險：確保 Word/PPT/PDF 檔案都已解除封鎖
def _run_unblock_in_background(root_dir):
    """背景執行檔案解除封鎖，避免阻塞互動提示。"""
    try:
        unblock_office_files_in_dir(root_dir)
    except Exception:
        pass

threading.Thread(
    target=_run_unblock_in_background,
    args=(download_dir,),
    daemon=True
).start()

# 按照課程名稱字母順序整理輸出並更新 output.txt
course_results.sort(key=lambda x: x[1]['course_name'])  # 按課程名稱字母排序
all_output_lines_sorted = []
for idx, result in course_results:
    all_output_lines_sorted.extend(result['output'])

with open(OUTPUT_FILE, "w", encoding="utf-8") as f:
    f.write("\n".join(all_output_lines_sorted))

# 定義背景檢查作業的函數
def check_assignments_background():
    global empty_assignments, assignment_check_completed
    from urllib.parse import urlparse, parse_qs

    def build_assignment_key(course_name, act_href):
        """用課程 + 作業 id 建立穩定鍵值，避免同名作業互相覆蓋。"""
        assign_id = None
        try:
            q = parse_qs(urlparse(act_href).query)
            assign_id = (q.get('id') or [None])[0]
        except Exception:
            assign_id = None

        if assign_id:
            return f"{course_name}||id:{assign_id}"
        return f"{course_name}||url:{act_href}"

    def has_submitted_record(records, course_name, act_href, assignment_key):
        """同時支援新舊格式：key 命中或同課程同網址視為已繳交。"""
        if assignment_key in records:
            return True
        for rec in records.values():
            if isinstance(rec, dict) and rec.get('course') == course_name and rec.get('url') == act_href:
                return True
        return False

    # 讀取已繳交作業記錄
    submitted_assignments = load_submitted_assignments()
    
    empty_assignments = []
    newly_submitted = {}  # 記錄新發現已繳交的作業
    
    # 使用已開啟的分頁遍歷所有課程
    for tab_handle in all_tabs:
        driver.switch_to.window(tab_handle)
        
        try:
            course_name = driver.find_element(By.CSS_SELECTOR, "h1.h2").text
            course_href = driver.current_url
        except:
            continue
        
        # 找出所有作業連結
        try:
            # 用 JS 一次快照所有作業 href/name，避免 stale element
            assignments_info = driver.execute_script("""
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
            
            # 迭代收集到的資訊
            for assign_info in assignments_info:
                try:
                    act_href = assign_info['href']
                    act_name = assign_info['name']
                    act_name = clean_activity_name(act_name)
                    
                    # 建立唯一識別鍵 (優先使用作業 id)
                    assignment_key = build_assignment_key(course_name, act_href)
                    
                    # 檢查是否在已繳交記錄中
                    if has_submitted_record(submitted_assignments, course_name, act_href, assignment_key):
                        continue
                    if has_submitted_record(newly_submitted, course_name, act_href, assignment_key):
                        continue
                    
                    # 沒有記錄，進入頁面檢查
                    driver.get(act_href)
                    wait = WebDriverWait(driver, 5)
                    driver.execute_script("document.body.style.zoom='80%'")
                    
                    # 檢查是否有「繳交作業」按鈕
                    has_submit_button = False
                    try:
                        WebDriverWait(driver, 1).until(
                            EC.presence_of_element_located((By.XPATH, "//button[contains(text(), '繳交作業')] | //th[contains(text(), '繳交狀態')]"))
                        )
                    except:
                        pass
                    
                    try:
                        submit_buttons = driver.find_elements(By.XPATH, "//button[contains(text(), '繳交作業')]")
                        if submit_buttons and len(submit_buttons) > 0:
                            has_submit_button = True
                    except:
                        pass
                    
                    if not has_submit_button:
                        try:
                            status_cell = driver.find_element(By.XPATH, "//th[contains(text(), '繳交狀態')]/following-sibling::td")
                            status_text = status_cell.text
                            if "尚無任何作業繳交" in status_text or "目前尚無" in status_text:
                                has_submit_button = True
                        except:
                            pass
                    
                    if has_submit_button:
                        # 未繳交
                        empty_assignments.append({
                            'name': act_name,
                            'course': course_name,
                            'url': act_href,
                            'tab_handle': tab_handle
                        })
                    else:
                        # 已繳交，加入記錄
                        newly_submitted[assignment_key] = {
                            'course': course_name,
                            'name': act_name,
                            'url': act_href,
                            'assignment_key': assignment_key,
                            'checked_date': time.strftime('%Y-%m-%d %H:%M:%S')
                        }
                    
                    driver.back()
                    time.sleep(0.1)
                    
                except Exception as e:
                    try:
                        driver.back()
                    except:
                        driver.switch_to.window(tab_handle)
                    time.sleep(0.1)
                    continue
        except Exception as e:
            driver.switch_to.window(tab_handle)
            continue
    
    # 更新已繳交作業記錄（無論有無新增，都儲存以確保檔案存在）
    if newly_submitted:
        submitted_assignments.update(newly_submitted)
    save_submitted_assignments(submitted_assignments)
    
    assignment_check_completed = True

# 在背景線程啟動作業檢查
empty_assignments = []
assignment_check_completed = False
assignment_check_thread = threading.Thread(target=check_assignments_background, daemon=True)
assignment_check_thread.start()

# 在結束前詢問是否要開啟任何課程資料夾
# 如果是第一次使用，跳過選擇並直接結束
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

print(f"{PINK}開啟以下課程的資料夾：{RESET}\n")


# 收集所有課程資料夾（直接使用已有的 course_results 數據）
all_courses = {}
for idx, result in course_results:
    course_name = result['course_name']
    course_path = create_course_folder(course_name)
    if course_name not in all_courses:
        all_courses[course_name] = course_path

# 顯示所有課程（有新活動的用紅色標記）
red_course_names = set()
for name, link, course_name, week_header, course_path, course_url, description in red_activities_to_print:
    red_course_names.add(course_name)
ibxx=0
for idx, (course_name, course_path) in enumerate(all_courses.items(), 1):
    if course_name in red_course_names:
        print(f"  {RED}{idx}. {course_name}{RESET}")
    else:
        if ibxx%2==0:
            print(f"  {MIKU}{idx}. {course_name}{RESET}")
        else:
            print(f"  {BBLUE}{idx}. {course_name}{RESET}")
    ibxx += 1
choice = input(f"\n{PINK}請輸入編號（或輸入 'u' 繳交作業，可用空白分隔多個編號）: {RESET}").strip().lower()


# 如果直接按 Enter（空輸入），立即結束程式
if not choice:
    close_macos_terminal_and_exit(0)

# 支援空白分隔多個編號
choice_parts = choice.split()

if choice == 'u':
    
    # 如果作業檢查還沒完成，等待完成
    if not assignment_check_completed:
        assignment_check_thread.join()  # 等待背景線程完成
        print(f"{GREEN}檢查完成！{RESET}\n")
    
    if empty_assignments:
        print(f"\n{YELLOW}找到 {len(empty_assignments)} 個未繳交作業{RESET}")
        
        # 列出所有未繳交作業
        print(f"\n{PINK}完整清單：{RESET}")
        ibxx=0
        for idx, item in enumerate(empty_assignments, 1):
            if ibxx%2==0:
                print(f" {BBLUE} {idx}. [{item['course']}] {item['name']}{RESET}")
            else:
                print(f" {MIKU} {idx}. [{item['course']}] {item['name']}{RESET}")
            ibxx += 1
        
        # 詢問要開啟哪些作業
        selection = input(f"{PINK}請輸入要開啟的作業編號（可個用空白分隔): {RESET}").strip().lower()
        
        # 按 Enter 直接結束（thread 已在 join 後儲存 submitted_assignments）
        if not selection:
            close_macos_terminal_and_exit(0)
        
        selected_assignments = []
        if selection == 'a':
            selected_assignments = empty_assignments
        else:
            indices = [int(x) - 1 for x in selection.split() if x.strip()]
            selected_assignments = [empty_assignments[i] for i in indices if 0 <= i < len(empty_assignments)]
        
        if selected_assignments:

            
            # 關閉 headless driver，改用可見模式
            driver.quit()
            
            # 重新啟動非 headless 模式
            chrome_options_visible = Options()
            chrome_options_visible.add_argument("--log-level=3")
            chrome_options_visible.add_experimental_option("excludeSwitches", ["enable-logging"])
            chrome_options_visible.add_argument("--disable-gpu")
            chrome_options_visible.add_argument("--disable-dev-shm-usage")
            chrome_options_visible.add_argument("--remote-debugging-port=0")
            if os.name == 'nt':
                chrome_options_visible.add_argument("--no-sandbox")
                chrome_options_visible.add_argument("--disable-features=RendererCodeIntegrity")
            apply_chrome_binary_option(chrome_options_visible)
            try:
                driver = create_webdriver(chrome_options_visible, hide_windows_console=True)
            except WebDriverException as e:
                # Windows 偶發 DevToolsActivePort 啟動失敗：改用臨時 profile 再重試
                if os.name == 'nt':
                    time.sleep(1)
                    chrome_options_visible_retry = Options()
                    chrome_options_visible_retry.add_argument("--log-level=3")
                    chrome_options_visible_retry.add_experimental_option("excludeSwitches", ["enable-logging"])
                    chrome_options_visible_retry.add_argument("--disable-gpu")
                    chrome_options_visible_retry.add_argument("--disable-dev-shm-usage")
                    chrome_options_visible_retry.add_argument("--remote-debugging-port=0")
                    chrome_options_visible_retry.add_argument("--no-sandbox")
                    chrome_options_visible_retry.add_argument("--disable-features=RendererCodeIntegrity")
                    submit_fallback_dir = tempfile.mkdtemp(prefix="chrome_profile_submit_fallback_", dir=BASE_DOWNLOAD_DIR)
                    chrome_options_visible_retry.add_argument(f"--user-data-dir={submit_fallback_dir}")
                    apply_chrome_binary_option(chrome_options_visible_retry)
                    driver = create_webdriver(chrome_options_visible_retry, hide_windows_console=True)
                else:
                    raise
            
            # 重新登入 (與第一次登入完全相同的流程，最多嘗試2次)
            login_success = False
            for login_attempt in range(2):
                try:
                    driver.get("https://elearningv4.nuk.edu.tw/login/index.php?loginredirect=1")
                    
                    WebDriverWait(driver, 10).until(
                        EC.visibility_of_element_located((By.ID, "username"))
                    ).send_keys(USERNAME)
                    
                    simulate_typing(driver, 'password', PASSWORD)
                    driver.execute_script("document.getElementById('loginbtn').click();")
                    
                    # 等待登入成功（確認導航成功）
                    time.sleep(0.2)
                    if "login" not in driver.current_url.lower():
                        login_success = True
                        break
                    else:
                        if login_attempt == 0:
                            time.sleep(1)
                except Exception as e:
                    if login_attempt == 0:
                        time.sleep(1)
            
            if not login_success:
                driver.quit()
                sys.exit()

            
            # 開啟選中的作業頁面並點擊「繳交作業」按鈕
            for idx, assignment in enumerate(selected_assignments, 1):
                
                # 在新分頁開啟作業
                if idx == 1:
                    driver.get(assignment['url'])
                else:
                    driver.execute_script(f"window.open('{assignment['url']}', '_blank');")
                    driver.switch_to.window(driver.window_handles[-1])
                
                time.sleep(0.5)
                
                # 點擊「繳交作業」按鈕
                try:
                    submit_button = WebDriverWait(driver, 5).until(
                        EC.element_to_be_clickable((By.XPATH, "//button[contains(text(), '繳交作業')]"))
                    )
                    submit_button.click()
                    time.sleep(0.5)
                except Exception as e:
                    print(f"   無法點擊按鈕: {e}")
        
            
            # 等待用戶按 Enter 或瀏覽器被關閉
            enter_pressed = threading.Event()
            
            def wait_for_enter():
                input()
                enter_pressed.set()
            
            input_thread = threading.Thread(target=wait_for_enter, daemon=True)
            input_thread.start()
            
            # 持續檢查瀏覽器狀態或 Enter 鍵
            while True:
                try:
                    # 檢查瀏覽器是否還活著
                    driver.current_url
                    # 檢查是否按了 Enter
                    if enter_pressed.is_set():
                        break
                    time.sleep(0.5)
                except:
                    # 瀏覽器已被關閉
                    break
            
            try:
                driver.quit()
            except:
                pass
            sys.exit()
    
    # 如果選擇 'u' 但沒有未繳交作業，繼續執行下面的程式碼
    else:
        print(f"\n{GREEN}所有作業都已繳交{RESET}")
        # 沒有作業需要處理，直接結束程式
        close_macos_terminal_and_exit(0)

elif choice_parts and all(token.isdigit() for token in choice_parts):
    opened = 0
    for part in choice_parts:
        idx_num = int(part)
        if all_courses and 1 <= idx_num <= len(all_courses):
            selected_course = list(all_courses.keys())[idx_num - 1]
            selected_path = all_courses[selected_course]
            open_folder(selected_path)
            opened += 1
    if opened == 0:
        print("輸入無效，未開啟任何資料夾")
    # 開啟資料夾後立即結束
    close_macos_terminal_and_exit(0)

# 程式結束前完整清理資源（優化 exe 執行效能）
def cleanup_resources():
    """完整清理所有資源，避免 exe 卡頓"""
    try:
        # 快速關閉 WebDriver，避免網路通信延遲
        if 'driver' in globals() and driver:
            # 使用更快速的關閉方式，避免atexit時的網路超時
            try:
                driver.service.stop()  # 直接停止服務
            except:
                pass
            try:
                driver.quit()
            except:
                pass
    except Exception:
        pass
    
    try:
        # 強制垃圾回收
        import gc
        gc.collect()
    except Exception:
        pass

# 註冊程式退出時的清理函數（僅在非強制退出時執行）
import atexit
atexit.register(cleanup_resources)

# 額外保險：背景線程清理（保留原有邏輯）
def cleanup_driver():
    try:
        if 'driver' in globals() and driver:
            driver.quit()
    except Exception:
        pass

cleanup_thread = threading.Thread(target=cleanup_driver, daemon=True)
cleanup_thread.start()