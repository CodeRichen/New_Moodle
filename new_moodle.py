# pip install selenium webdriver-manager colorama requests py7zr patool rarfile
# 不要在開啟Moodle網頁的狀態執行程式

import os
import sys
import ctypes

# ========== 使用 Windows API 設定終端機視窗為全螢幕 ==========
def maximize_console_window():
    """將終端機視窗最大化"""
    hwnd = ctypes.windll.kernel32.GetConsoleWindow()
    if hwnd:
        # SW_MAXIMIZE = 3
        ctypes.windll.user32.ShowWindow(hwnd, 3)

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
from selenium.webdriver.chrome.options import Options
import time
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

YELLOW = "\033[33m"
RED = "\033[31m"
BLUE = "\033[34m"
BBLUE = "\033[94m"
MIKU = "\033[36m"
GREEN = "\033[32m"
PURPLE = "\033[38;5;129m"  # 亮紫紅色/洋紅色
ORANGE = "\033[38;5;214m"  # 黃橘色
RESET = "\033[0m"
PINK = "\033[38;2;255;220;225m"

# ========== TODO 路徑設定區域（修改這裡可以改變所有檔案存放位置）==========
# 主要下載目錄 - 修改這裡就能改變所有檔案的存放位置
# 例如：r"D:\Moodle" 或 r"E:\課程資料"
BASE_DOWNLOAD_DIR = os.path.join(os.path.expanduser("~"), "Downloads", "class")
# ================================================================

# 根據 BASE_DOWNLOAD_DIR 自動設定其他檔案路徑
OUTPUT_FILE = os.path.join(BASE_DOWNLOAD_DIR, "cless.txt")
SUBMITTED_ASSIGNMENTS_FILE = os.path.join(BASE_DOWNLOAD_DIR, "submitted_assignments.json")
PASSWORD_FILE = os.path.join(BASE_DOWNLOAD_DIR, "password.txt")

# 確保主目錄存在
os.makedirs(BASE_DOWNLOAD_DIR, exist_ok=True)

# 初始化第一次使用標記
IS_FIRST_TIME = False

def get_password_input(prompt):
    """自定義密碼輸入函數，顯示星號"""
    import msvcrt
    print(prompt, end='', flush=True)
    password = ""
    while True:
        char = msvcrt.getch()
        if char == b'\r':  # Enter鍵
            break
        elif char == b'\x08':  # Backspace鍵
            if len(password) > 0:
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
                # 忽略無法解碼的字符
                pass
    print()  # 換行
    return password

def test_login(username, password):
    """測試登入是否成功，使用現有的登入函數"""
    try:
        # 創建臨時瀏覽器進行登入測試
        test_chrome_options = Options()
        test_chrome_options.add_argument("--headless")
        test_chrome_options.add_argument("--disable-gpu")
        test_chrome_options.add_argument("--log-level=3")
        test_chrome_options.add_argument("--disable-extensions")
        test_chrome_options.add_argument("--disable-dev-shm-usage")
        test_chrome_options.add_argument("--no-sandbox")
        test_chrome_options.add_experimental_option("excludeSwitches", ["enable-logging"])
        
        test_service = Service(log_path=os.devnull)
        test_service.creation_flags = subprocess.CREATE_NO_WINDOW
        test_driver = webdriver.Chrome(options=test_chrome_options, service=test_service)
        
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

# 讀取帳號密碼
def load_credentials():
    """從 password.txt 讀取帳號密碼，如果不存在則讓使用者輸入並創建"""
    global IS_FIRST_TIME  # 使用全域變數來追蹤是否為第一次使用
    
    try:
        if os.path.exists(PASSWORD_FILE):
            with open(PASSWORD_FILE, 'r', encoding='utf-8') as f:
                lines = f.read().splitlines()
                if len(lines) >= 2:
                    username = lines[0].strip()
                    password = lines[1].strip()
                    if not username or not password:
                        print(f"\n{RED}{'='*60}{RESET}")
                        print(f"{RED}❌ 錯誤：password.txt 檔案內容不完整{RESET}")
                        input()
                        sys.exit(1)
                    IS_FIRST_TIME = False
                    return username, password
                else:
                    print(f"\n{RED}{'='*60}{RESET}")
                    print(f"{RED}❌ 錯誤：password.txt 格式錯誤{RESET}")
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
                    print(f"\n{RED}❌ 帳號不存在或密碼錯誤，請重新輸入{RESET}")
                    continue
            
            # 登入成功，創建 password.txt 檔案
            try:
                with open(PASSWORD_FILE, 'w', encoding='utf-8') as f:
                    f.write(f"{username}\n{password}\n")
                print(f"\n{BLUE}建置環境中{RESET}")
                return username, password
            except Exception as e:
                print(f"\n{RED}❌ 無法創建密碼檔案：{e}{RESET}")
                print(f"按 Enter 鍵離開...")
                input()
                sys.exit(1)
                
    except Exception as e:
        print(f"\n{RED}{'='*60}{RESET}")
        print(f"{RED}❌ 錯誤：讀取 password.txt 失敗{RESET}")
        print(f"{RED}{'='*60}{RESET}")
        print(f"\n錯誤訊息：{e}")
        print(f"\n可能原因：")
        print(f"   - 檔案正在被其他程式使用")
        print(f"   - 檔案編碼問題（請使用 UTF-8 編碼儲存）")
        print(f"   - 檔案讀取權限不足")
        print(f"\n按 Enter 鍵離開...")
        input()
        sys.exit(1)

# 載入帳號密碼
USERNAME, PASSWORD = load_credentials()

def is_activity(text):
    return not (text.startswith("第") and text.endswith("週")) and text.strip()

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
        print(f"❌ 無法開啟資料夾: {e}")

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
chrome_options.add_argument("--headless")  # 無頭模式
chrome_options.add_argument("--disable-gpu")  # 防止顯示錯誤
chrome_options.add_argument("--log-level=3")  # 降低日誌等級，避免雜訊
chrome_options.add_argument("--window-size=1920,1080")  # 設定解析度，避免元素渲染錯誤
chrome_options.add_argument("--disable-extensions")  # 禁用擴充功能
chrome_options.add_argument("--disable-dev-shm-usage")  # 解決資源限制問題
chrome_options.add_argument("--no-sandbox")  # 加快啟動速度
chrome_options.add_argument("--disable-blink-features=AutomationControlled")  # 避免被偵測
chrome_options.add_experimental_option("excludeSwitches", ["enable-logging", "enable-automation"])  # 避免除錯訊息
chrome_options.page_load_strategy = 'eager'  # 加快頁面載入速度

prefs = {
    "download.default_directory": download_dir,  # 下載到指定資料夾
    "plugins.always_open_pdf_externally": True,  # 避免 PDF 在 Chrome 裡直接開啟
    "profile.default_content_setting_values.images": 2  # 禁用圖片加快速度
}
chrome_options.add_experimental_option("prefs", prefs)
service = Service(log_path=os.devnull)
service.creation_flags = subprocess.CREATE_NO_WINDOW
driver = webdriver.Chrome(options=chrome_options, service=service)

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
        
        # 檢查是否被重定向回登入頁面
        current_url = driver.current_url
        if "login" in current_url:
            if attempt < 1:  # 還有重試機會
                continue
            else:  # 最後一次也失敗
                print(f"\n{RED}{'='*60}{RESET}")
                print(f"{RED}❌ 登入失敗：無法進入系統{RESET}")
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
            print(f"{RED}❌ 登入失敗：無法連接到 Moodle {RESET}")
            print(f"{RED}{'='*60}{RESET}")
            print(f"\n錯誤訊息：{e}")
            print(f"\n按 Enter 鍵離開...")
            input()
            driver.quit()
            sys.exit(1)

if not login_success:
    print(f"\n{RED}❌ 登入失敗{RESET}")
    print(f"按 Enter 鍵離開...")
    input()
    driver.quit()
    sys.exit(1)

wait = WebDriverWait(driver, 10)
wait.until(EC.presence_of_all_elements_located((By.CSS_SELECTOR, "a.aalink.coursename")))
course_links = driver.find_elements(By.CSS_SELECTOR, "a.aalink.coursename")
course_hrefs = [link.get_attribute("href") for link in course_links]

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
            
            if (name && !(name.startsWith('第') && name.endsWith('週')) && name.trim()) {
                let link = act.querySelector('a.aalink');
                let href = link ? link.getAttribute('href') : '（無連結）';
                activities.push({name: name, href: href});
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
            name = act_data['name']
            href_link = act_data['href']

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
                local_red_activities.append((name, href_link, course_name, week_header, course_path, href))
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
        print(f"{RED}⚠️ 無法寫入活動記錄: {course_name} - {e}{RESET}")
    
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
                        if week_info and week_info[0] != 1:  # 不是第一週
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
                                # 活動名稱維持黃色，網址改為白色（預設）
                                print(f"  {YELLOW}{name}{RESET} - {href_link}")
                            print("")
                             
            
            # 收集資料
            with data_lock:
                all_output_lines.extend(result['output'])
                red_activities_to_print.extend(result['red_activities'])
                failed_downloads.extend(result['failed_downloads'])

def remove_zone_identifier(filepath):
    """移除 Windows Zone.Identifier 標記,避免「受保護的檢視」"""
    try:
        subprocess.run(
            ["powershell", "-Command", f"Unblock-File -Path '{filepath}'"],
            capture_output=True,
            timeout=5
        )
    except Exception as e:
        # 靜默失敗,不影響主流程
        pass
        
def wait_for_download(filename, download_path=None, timeout=300, ask_after=3, size_limit_mb=50):
    """
    等待下載完成。若超過 ask_after 秒或檔案超過 size_limit_mb MB，詢問是否繼續。
    
    返回值：
    - 檔案路徑：下載成功
    - None：跳過此檔案
    - 'SKIP_COURSE'：跳過整個課程的所有檔案
    """
    if download_path is None:
        download_path = download_dir
    file_path = os.path.join(download_path, filename)
    cr_path = file_path + ".crdownload"
    start = time.time()

    已詢問時間 = False
    已詢問大小 = False
    last_progress_time = 0  # 記錄上次輸出進度的時間
    
    # 記錄開始時檔案的修改時間(如果存在)
    initial_mtime = None
    if os.path.exists(file_path):
        initial_mtime = os.path.getmtime(file_path)

    print(f"⏳ 等待下載：{filename}")

    while True:
        # 檔案存在且無 .crdownload → 檢查是否為新下載的檔案
        if os.path.exists(file_path) and not os.path.exists(cr_path):
            # 如果之前檔案就存在,檢查修改時間是否改變
            if initial_mtime is not None:
                current_mtime = os.path.getmtime(file_path)
                if current_mtime <= initial_mtime:
                    # 檔案沒有更新,繼續等待
                    time.sleep(0.2)
                    elapsed = time.time() - start
                    if elapsed > 10:  # 等待超過10秒還沒新檔案,可能下載到其他地方
                        print(f"{RED}⚠️ 未偵測到新下載的檔案,可能已存在舊檔案{RESET}")
                        return None
                    continue
            
            # print(f"{GREEN}✅ 下載完成：{filename}{RESET}")
            return file_path

        elapsed = time.time() - start

        # 每 3 秒顯示一次進度
        if elapsed - last_progress_time >= 3 and elapsed >= 3:
            print(f"   已等待 {int(elapsed)} 秒…")
            last_progress_time = elapsed

        # 超過指定等待秒數 → 問是否繼續等
        if elapsed > ask_after and not 已詢問時間:
            print(f"\n{YELLOW}⚠️ 下載已等待超過 {ask_after} 秒{RESET}")
            print(f"   檔案：{filename}")
            print(f"   - 繼續等待 (Enter)")
            print(f"   - 放棄此檔案 (輸入 d)")
            print(f"   - 放棄此課程所有檔案 (輸入 dd)")
            choice = input(f"   請選擇：").strip().lower()
            已詢問時間 = True
            last_progress_time = elapsed  # 重置進度時間，避免詢問後立即輸出進度
            if choice == "dd":
                print(f"{YELLOW}⏭️ 已放棄此課程所有檔案下載{RESET}\n")
                return 'SKIP_COURSE'
            elif choice == "d":
                print(f"{YELLOW}⏭️ 已放棄此檔案下載{RESET}\n")
                return None
            print(f"   繼續等待下載...\n")

        # 若部分下載的檔案存在 → 檢查大小
        if os.path.exists(file_path) and not 已詢問大小:
            size_mb = os.path.getsize(file_path) / (1024 * 1024)
            if size_mb > size_limit_mb:
                choice = input(f"⚠️ 檔案超過 {size_limit_mb} MB，要繼續下載嗎？(y/n)：").strip().lower()
                已詢問大小 = True
                if choice != "y":
                    print("⏭️ 使用者選擇跳過。")
                    return None

        # 檢查是否超過 timeout
        if elapsed > timeout:
            print(f"{RED}❌ 下載超時：{filename}{RESET}")
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
            print(f"   {YELLOW}⚠️ 檔案不是壓縮檔，而是 HTML 頁面，跳過解壓{RESET}")
            return False
        
        if file_path.endswith(".zip"):
            if not is_zip:
                print(f"   {YELLOW}⚠️ 檔案副檔名為 .zip 但不是有效的 ZIP 格式，跳過解壓{RESET}")
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
                print(f"   {YELLOW}⚠️ 檔案副檔名為 .rar 但不是有效的 RAR 格式，跳過解壓{RESET}")
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
                        print(f"{YELLOW}⚠️  7-Zip 解壓失敗: {e}{RESET}")
            
            # 方法2: 優先使用 patool（更穩定，支持多種後端）
            if HAS_PATOOL:
                try:
                    patool.extract_archive(file_path, outdir=dest_dir, verbosity=-1)
                    return True
                except Exception as e:
                    print(f"{YELLOW}⚠️  patool 解壓失敗: {e}{RESET}")
            
            # 方法3: 使用 rarfile
            if HAS_RARFILE:
                try:
                    with rarfile.RarFile(file_path, 'r') as rf:
                        rf.extractall(dest_dir)
                    return True
                except Exception as e:
                    print(f"{YELLOW}⚠️  rarfile 解壓失敗: {e}{RESET}")
            
            # 都失敗，顯示安裝提示
            print(f"{YELLOW}⚠️  無法解壓 RAR 檔案，已跳過: {os.path.basename(file_path)}{RESET}")
            print(f"   💡 7-Zip 已安裝但無法使用，請嘗試：")
            print(f"   1. 重新啟動終端或電腦")
            print(f"   2. 或手動安裝: winget install 7zip.7zip")
            return False
            
        elif file_path.endswith(".7z"):
            if not is_7z:
                print(f"   {YELLOW}⚠️ 檔案副檔名為 .7z 但不是有效的 7Z 格式，跳過解壓{RESET}")
                return False
            
            with py7zr.SevenZipFile(file_path, 'r') as sz:
                sz.extractall(dest_dir)
        else:
            return False
        return True
    except zipfile.BadZipFile:
        print(f"   {YELLOW}⚠️ 無效的 ZIP 檔案格式，跳過解壓{RESET}")
        return False
    except Exception as e:
        print(f"❌ 解壓失敗: {os.path.basename(file_path)}, 原因: {e}")
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

    # 🔴 先自動開啟所有紅色課程的資料夾
if red_activities_to_print:
   

    
    opened_folders = set()
    for name, link, course_name, week_header, course_path, course_url in red_activities_to_print:
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
        print(f"{PINK}🔻 以下為新增活動{RESET}")
        print("="*60 + "\n")
    
    total_downloaded_files = 0  # 追蹤總下載檔案數
    
    # 記錄已下載的檔案，避免重複下載（移到迴圈外，所有活動共享）
    downloaded_files = set()
    
    # 記錄需要移除 Zone.Identifier 的文件路徑
    files_to_unblock = []
    
    # 追蹤需要跳過的課程
    skipped_courses = set()
    
    for name, link, course_name, week_header, course_path, course_url in red_activities_to_print:
        # 如果這個課程已被標記為跳過，則跳過
        if course_name in skipped_courses:
            continue
            
        # 如果是第一次使用，跳過詳細輸出
        if not IS_FIRST_TIME:
            print(f"\n{RED}━━━ {name} ━━━{RESET}")
            print(f"{PINK}課程：{course_name}{RESET}")
            print(f"週次：{week_header}")
            print(f"活動連結：{link}")
            print(f"儲存位置：{course_path}\n")
        
        # 確保每次都重新設定下載路徑到正確的課程資料夾（用於資料夾/作業的 Selenium 下載）
        driver.execute_cdp_cmd("Page.setDownloadBehavior", {
            "behavior": "allow",
            "downloadPath": os.path.abspath(course_path)
        })

        # 🔽 點進活動頁面，抓取所有有 href 的下載連結並打開（會自動觸發下載）
        try:
            # 檢查是否有有效連結
            if link == "（無連結）" or not link or not link.startswith("http"):

                # 回到課程頁面，找出這個活動的圖片
                driver.get(course_url)
                wait = WebDriverWait(driver, 5)
                
                # 找到包含此活動名稱的活動元素
                activities = driver.find_elements(By.CSS_SELECTOR, "div.activity-item")
                for act in activities:
                    act_name = act.get_attribute("data-activityname")
                    if not act_name:
                        try:
                            act_name = act.find_element(By.CSS_SELECTOR, "span.instancename").text.strip()
                        except:
                            continue
                    
                    if act_name == name:
                        # 找到對應的活動,檢查是否有圖片
                        try:
                            images = act.find_elements(By.CSS_SELECTOR, "img[src*='pluginfile.php']")
                            for img in images:
                                img_url = img.get_attribute("src")
                                filename = extract_filename_from_url(img_url)
                                
                                # 使用 requests 直接下載圖片（覆蓋同名檔案）
                                try:
                                    # 使用 session 以保持登入狀態
                                    session = create_session_with_cookies()
                                    
                                    # 下載圖片
                                    response = session.get(img_url, stream=True)
                                    if response.status_code == 200:
                                        file_path = os.path.join(course_path, filename)
                                        with open(file_path, 'wb') as f:
                                            for chunk in response.iter_content(chunk_size=8192):
                                                f.write(chunk)

                                        files_to_unblock.append(file_path)
                                        downloaded_files.add(filename)
                                        existing_files.add(filename)
                                        total_downloaded_files += 1
                                    else:
                                        print(f"{RED}❌ 下載失敗: HTTP {response.status_code}{RESET}")
                                except Exception as e:
                                    print(f"{RED}❌ 下載失敗: {e}{RESET}")
                        except Exception as e:
                            print(f"⚠️ 無法下載圖片: {e}")
                        break
                
                print()
                continue
            
            # 判斷活動類型並處理
            # Case 1: 資源檔案 - 訪問頁面並找到下載連結
            if "mod/resource/view.php" in link:
                try:
                    
                    # 記錄下載目錄中現有的檔案及其修改時間
                    before_files = {}
                    if os.path.exists(course_path):
                        for f in os.listdir(course_path):
                            fpath = os.path.join(course_path, f)
                            if os.path.isfile(fpath):
                                before_files[f] = os.path.getmtime(fpath)
                    
                    driver.get(link)
                    wait = WebDriverWait(driver, 5)
                    
                    # 在資源頁面中找到實際的下載連結（只抓主要內容區的連結）
                    download_links = driver.find_elements(By.CSS_SELECTOR, "div.resourceworkaround a[href*='pluginfile.php']")
                    # print(f"   找到 {len(download_links)} 個下載連結")
                    
                    # 檢查是否有新檔案出現或檔案被更新（可能是自動下載）
                    time.sleep(0.2)  
                    after_files = {}
                    if os.path.exists(course_path):
                        for f in os.listdir(course_path):
                            fpath = os.path.join(course_path, f)
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
                            if filename.lower().endswith(('.htm', '.html')):
                                print(f"   ⏭️  跳過 HTML 文件: {filename}")
                                file_to_remove = os.path.join(course_path, filename)
                                try:
                                    os.remove(file_to_remove)
                                except:
                                    pass
                                continue
                            
                            file_path = os.path.join(course_path, filename)
                            files_to_unblock.append(file_path)
                            downloaded_files.add(filename)
                            existing_files.add(filename)
                            total_downloaded_files += 1
                            downloaded_in_this_activity += 1
                            
                            # 立即解壓縮檔案
                            if filename.endswith((".zip", ".rar", ".7z")):
                                actual_size = os.path.getsize(file_path)
                                if actual_size < 100:
                                    print(f"   {YELLOW}⚠️ 壓縮檔太小 ({actual_size} bytes)，可能損壞，跳過解壓{RESET}")
                                else:

                                    success = extract_file(file_path, course_path)
                                    if success:
                                        os.remove(file_path)
                                        print(f"   ✅ 解壓完成並刪除原始檔")
                                    else:
                                        print(f"   {YELLOW}⚠️ 解壓失敗，保留原始檔{RESET}")
                        
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
                                # print(f"   ⏳ 等待下載完成: {base_filename}")
                                file_path = wait_for_download(base_filename, download_path=course_path)
                                if file_path and file_path != 'SKIP_COURSE':
                                    if base_filename.lower().endswith(('.htm', '.html')):
                                        # print(f"   ⏭️  跳過 HTML 文件: {base_filename}")
                                        try:
                                            os.remove(file_path)
                                        except:
                                            pass
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
                                                os.remove(file_path)
                                                # print(f"   {GREEN}✅ 解壓完成並刪除原始檔{RESET}")
                            continue
                        
                        # print(f"{YELLOW}⚠️  此資源頁面沒有附件，跳過{RESET}")
                        continue
                    
                    for link_elem in download_links:
                        dl_href = link_elem.get_attribute("href")
                        filename = extract_filename_from_url(dl_href)
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
                            
                            file_path = os.path.join(course_path, filename)
                            file_size = 0
                            with open(file_path, 'wb') as f:
                                for chunk in response.iter_content(chunk_size=8192):
                                    if chunk:
                                        f.write(chunk)
                                        file_size += len(chunk)
                            
                            # 檢查檔案大小
                            if file_size == 0 or not os.path.exists(file_path) or os.path.getsize(file_path) == 0:
                                print(f"{RED}❌ 下載失敗：檔案大小為 0{RESET}")
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
                                    print(f"{YELLOW}⚠️ 壓縮檔太小 ({actual_size} bytes)，可能損壞，跳過解壓{RESET}")
                                else:

                                    success = extract_file(file_path, course_path)
                                    if success:
                                        os.remove(file_path)

                                    else:
                                        print(f"{YELLOW}   ⚠️ 解壓失敗，保留原始檔{RESET}")
                            
                        except Exception as e:
                            print(f"{RED}❌ 下載失敗: {e}{RESET}")
                            if file_path and os.path.exists(file_path):
                                os.remove(file_path)
                            
                except Exception as e:
                    print(f"{RED}❌ 下載失敗{RESET}")
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
                
                print(f"📂 進入資料夾/作業頁面...")
                
                # 首先檢查當前頁面是否是作業頁面,並下載作業說明的附件
                if "mod/assign/view.php" in link:
                    # 這是作業頁面,先下載作業說明中的附件
                    intro_attachments = driver.find_elements(By.CSS_SELECTOR, "div.activity-description a[href*='pluginfile.php']")
                    for intro_link in intro_attachments:
                        intro_href = intro_link.get_attribute("href")
                        filename = extract_filename_from_url(intro_href)
                        
                        try:
                            print(f"📄 下載作業說明附件: {filename}")
                            intro_link.click()
                            wait = WebDriverWait(driver, 5)
                            file_path = wait_for_download(filename, download_path=course_path)
                            if file_path == 'SKIP_COURSE':
                                skipped_courses.add(course_name)
                                print(f"{YELLOW}⏭️ 跳過課程 {course_name} 的所有剩餘檔案{RESET}")
                                break
                            if file_path:
                                files_to_unblock.append(file_path)
                                downloaded_files.add(filename)
                                existing_files.add(filename)
                                total_downloaded_files += 1
                        except Exception as e:
                            print(f"{RED}❌ 下載失敗: {filename}{RESET}")
                            print(f"   錯誤: {e}")
                            failed_downloads.append({
                                'name': name,
                                'course': course_name,
                                'url': intro_href,
                                'filename': filename
                            })
                
                # 找出 submission 區塊內的所有連結（要排除 - 這是已提交的作業）
                submission_links = set()
                submission_blocks = driver.find_elements(By.CSS_SELECTOR, "div[class*='summary_assignsubmission_file']")
                for block in submission_blocks:
                    a_tags = block.find_elements(By.CSS_SELECTOR, "a[href]")
                    for a_tag in a_tags:
                        submission_links.add(a_tag.get_attribute("href"))

                # 收集所有 pluginfile.php 連結
                file_href_set = set()
                
                # 方法1: 找所有包含 pluginfile.php 的連結
                file_links = driver.find_elements(By.CSS_SELECTOR, "a[href*='pluginfile.php']")
                # print(f"   找到 {len(file_links)} 個 pluginfile.php 連結")
                
                for f in file_links:
                    f_href = f.get_attribute("href")
                    # 只排除已提交的檔案
                    if f_href not in submission_links:
                        file_href_set.add(f_href)
                        # print(f"   📎 收集連結: {extract_filename_from_url(f_href)}")
                
                # 方法2: 特別檢查 forcedownload 參數的連結
                force_download_links = driver.find_elements(By.CSS_SELECTOR, "a[href*='forcedownload=1']")
                # print(f"   找到 {len(force_download_links)} 個 forcedownload 連結")
                
                for f in force_download_links:
                    f_href = f.get_attribute("href")
                    if f_href not in submission_links and 'pluginfile.php' in f_href:
                        file_href_set.add(f_href)
                        # print(f"   📎 收集 forcedownload: {extract_filename_from_url(f_href)}")
                
                # 方法3: 檢查作業說明區域
                intro_attachments = driver.find_elements(By.CSS_SELECTOR, "div.activity-description a[href*='introattachment']")
                for intro_link in intro_attachments:
                    file_href_set.add(intro_link.get_attribute("href"))
                
                print(f"   總共收集到 {len(file_href_set)} 個唯一檔案連結")
                
                # 處理收集到的檔案連結
                for f_href in file_href_set:
                    filename = extract_filename_from_url(f_href)
                    try:
                        print(f"🔽 下載檔案: {filename}")
                        
                        # 使用 session 直接下載
                        session = create_session_with_cookies()
                        response = session.get(f_href, stream=True)
                        
                        if response.status_code == 200:
                            file_path = os.path.join(course_path, filename)
                            with open(file_path, 'wb') as f:
                                for chunk in response.iter_content(chunk_size=8192):
                                    if chunk:
                                        f.write(chunk)
                            
                            files_to_unblock.append(file_path)
                            downloaded_files.add(filename)
                            existing_files.add(filename)
                            total_downloaded_files += 1
                        else:
                            print(f"{RED}❌ 下載失敗: HTTP {response.status_code}{RESET}")
                    except Exception as e:
                        print(f"{RED}❌ 下載失敗: {filename}{RESET}")
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
                    print(f"🔗 訪問 URL 活動頁面...")
                    driver.get(link)
                    time.sleep(0.2)
                    
                    # 尋找實際的外部連結
                    url_links = driver.find_elements(By.CSS_SELECTOR, "div.urlworkaround a[href]")
                    
                    if url_links:
                        for url_link in url_links:
                            actual_url = url_link.get_attribute("href")
                            link_text = url_link.text
                            if actual_url and not actual_url.startswith("https://elearningv4.nuk.edu.tw"):
                                print(f"{BLUE}📎 找到外部連結: {link_text}{RESET}")
                                print(f"   {actual_url}")
                                
                                # 清理檔案名稱，移除不合法字元
                                safe_filename = "".join(c if c.isalnum() or c in " _-()（）" else "_" for c in name)
                                # 限制檔案名稱長度，避免過長
                                if len(safe_filename) > 100:
                                    safe_filename = safe_filename[:100]
                                
                                # 檢查是否為 Google Sheets 連結
                                if "docs.google.com/spreadsheets" in actual_url:
                                    try:
                                        print(f"📊 偵測到 Google Sheets，嘗試下載...")
                                        
                                        # 使用 session 以保持登入狀態
                                        session = create_session_with_cookies()
                                        
                                        # 提取 spreadsheet ID
                                        import re
                                        match = re.search(r'/spreadsheets/d/([a-zA-Z0-9-_]+)', actual_url)
                                        if match:
                                            sheet_id = match.group(1)
                                            
                                            # 嘗試下載為 Excel 格式
                                            export_url = f"https://docs.google.com/spreadsheets/d/{sheet_id}/export?format=xlsx"
                                            
                                            print(f"   下載 URL: {export_url}")
                                            response = session.get(export_url, stream=True)
                                            
                                            if response.status_code == 200:
                                                excel_file = os.path.join(course_path, f"{safe_filename}.xlsx")
                                                with open(excel_file, 'wb') as f:
                                                    for chunk in response.iter_content(chunk_size=8192):
                                                        f.write(chunk)
                                                print(f"{GREEN}✅ 已下載 Google Sheets 為: {os.path.basename(excel_file)}{RESET}")
                                                files_to_unblock.append(excel_file)
                                                total_downloaded_files += 1
                                                existing_files.add(os.path.basename(excel_file))
                                            else:
                                                print(f"{YELLOW}⚠️ 無法下載 Google Sheets (可能需要權限){RESET}")
                                                # 仍然儲存連結
                                                url_file = os.path.join(course_path, f"{safe_filename}_連結.txt")
                                                with open(url_file, 'w', encoding='utf-8') as f:
                                                    f.write(f"{name}\n")
                                                    f.write(f"連結: {actual_url}\n")
                                                    f.write(f"說明: {link_text}\n")
                                                print(f"{GREEN}✅ 已儲存連結到: {os.path.basename(url_file)}{RESET}")
                                                total_downloaded_files += 1
                                        else:
                                            print(f"{YELLOW}⚠️ 無法解析 Google Sheets ID{RESET}")
                                    except Exception as e:
                                        print(f"{RED}❌ 下載 Google Sheets 失敗: {e}{RESET}")
                                        # 發生錯誤時仍儲存連結
                                        url_file = os.path.join(course_path, f"{safe_filename}_連結.txt")
                                        with open(url_file, 'w', encoding='utf-8') as f:
                                            f.write(f"{name}\n")
                                            f.write(f"連結: {actual_url}\n")
                                            f.write(f"說明: {link_text}\n")
                                        print(f"{GREEN}✅ 已儲存連結到: {os.path.basename(url_file)}{RESET}")
                                        total_downloaded_files += 1
                                else:
                                    # 一般連結，儲存為文字檔
                                    url_file = os.path.join(course_path, f"{safe_filename}_連結.txt")
                                    with open(url_file, 'w', encoding='utf-8') as f:
                                        f.write(f"{name}\n")
                                        f.write(f"連結: {actual_url}\n")
                                        f.write(f"說明: {link_text}\n")
                                    print(f"{GREEN}✅ 已儲存連結到: {os.path.basename(url_file)}{RESET}")
                                    total_downloaded_files += 1
                    else:
                        print(f"{YELLOW}⚠️ 未找到外部連結{RESET}")
                        
                except Exception as e:
                    print(f"{RED}❌ 處理 URL 活動失敗: {e}{RESET}")
                continue
            
            # Case 4: 討論區類型活動 - 提取文字內容並儲存
            if "mod/forum/view.php" in link:
                try:
                    print(f"💬 訪問討論區頁面...")
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
                        
                        print(f"{GREEN}✅ 已儲存討論區內容到: {os.path.basename(txt_file)}{RESET}")
                        total_downloaded_files += 1
                        existing_files.add(os.path.basename(txt_file))
                    else:
                        print(f"{YELLOW}⚠️ 未找到討論區描述內容{RESET}")
                        
                except Exception as e:
                    print(f"{RED}❌ 處理討論區活動失敗: {e}{RESET}")
                continue
            
            # Case 5: 其他類型的活動 (如 page 等) - 無需下載
            print(f"ℹ️ 此活動類型無需下載檔案")
            
        except Exception as e:
            print(f"{RED}❌ 處理活動時發生錯誤: {e}{RESET}")

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
    print(f"{RED}❌ 以下檔案下載失敗，請手動下載：{RESET}")
    print("="*60)
    for item in failed_downloads:
        print(f"\n📌 {item['name']}")
        print(f"   課程: {item['course']}")
        print(f"   檔名: {item['filename']}")
        print(f"   {BLUE}🔗 下載連結: {item['url']}{RESET}")

extracted_count = 0
failed_extract = []  # 記錄解壓失敗的檔案
for root, dirs, files in os.walk(download_dir):
    for file in files:
        filepath = os.path.join(root, file)
        if file.endswith((".zip", ".rar", ".7z")):
            if not IS_FIRST_TIME:
                print(f"📦 解壓縮: {os.path.basename(file)}")

            success = extract_file(filepath, root)
            if success:
                os.remove(filepath)
                if not IS_FIRST_TIME:
                    print(f"   ✅ 完成並刪除原始檔")
                extracted_count += 1
            else:
                # 記錄失敗的檔案（特別是 RAR）
                if file.endswith(".rar"):
                    failed_extract.append(filepath)

if extracted_count > 0 and not IS_FIRST_TIME:
    print(f"\n{GREEN}✅ 解壓縮完成，共處理 {extracted_count} 個檔案{RESET}")

if failed_extract and not IS_FIRST_TIME:
    print(f"\n{YELLOW}⚠️  以下檔案因工具缺失而未解壓：{RESET}")
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

# 定義背景檢查作業的函數
def check_assignments_background():
    global empty_assignments, assignment_check_completed
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
            assign_links = driver.find_elements(By.CSS_SELECTOR, "div.activity-item a.aalink[href*='mod/assign/view.php']")
            
            # 收集所有作業的資訊
            assignments_info = []
            for link_elem in assign_links:
                try:
                    act_href = link_elem.get_attribute("href")
                    act_name = link_elem.find_element(By.CSS_SELECTOR, "span.instancename").text.strip()
                    if '\n作業' in act_name:
                        act_name = act_name.replace('\n作業', '').strip()
                    if act_href and act_name:
                        assignments_info.append({'href': act_href, 'name': act_name})
                except:
                    continue
            
            # 迭代收集到的資訊
            for assign_info in assignments_info:
                try:
                    act_href = assign_info['href']
                    act_name = assign_info['name']
                    
                    # 建立唯一識別鍵 (課程名稱 + 作業名稱)
                    assignment_key = f"{course_name}||{act_name}"
                    
                    # 檢查是否在已繳交記錄中
                    if assignment_key in submitted_assignments:
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
    
    # 更新已繳交作業記錄
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
    print(f"\n{GREEN}環境建置完成{RESET}")
    print(f"\n{YELLOW}可在下次上課前再次執行此程式{RESET}")
    
    # 程式結束前關閉所有分頁（在背景線程中執行避免卡頓）
    def cleanup_driver():
        try:
            driver.quit()
        except:
            pass
    
    cleanup_thread = threading.Thread(target=cleanup_driver, daemon=True)
    cleanup_thread.start()
    sys.exit(0)

print(f"{PINK}開啟以下課程的資料夾：{RESET}")


# 收集所有課程資料夾（直接使用已有的 course_results 數據）
all_courses = {}
for idx, result in course_results:
    course_name = result['course_name']
    course_path = create_course_folder(course_name)
    if course_name not in all_courses:
        all_courses[course_name] = course_path

# 顯示所有課程（有新活動的用紅色標記）
red_course_names = set()
for name, link, course_name, week_header, course_path, course_url in red_activities_to_print:
    red_course_names.add(course_name)
ibxx=0
for idx, (course_name, course_path) in enumerate(all_courses.items(), 1):
    if course_name in red_course_names:
        print(f"  {RED}{idx}. {course_name} (NEW){RESET}")
    else:
        if ibxx%2==0:
            print(f"  {MIKU}{idx}. {course_name}{RESET}")
        else:
            print(f"  {BBLUE}{idx}. {course_name}{RESET}")
    ibxx += 1
choice = input(f"\n{PINK}請輸入編號（或輸入 'u' 繳交作業，可用空白分隔多個編號）: {RESET}").strip().lower()

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
            driver = webdriver.Chrome(options=chrome_options_visible)
            
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
        print(f"\n{GREEN}✅ 太棒了！所有作業都已繳交！{RESET}")

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

# 程式結束前關閉所有分頁（在背景線程中執行避免卡頓）
def cleanup_driver():
    try:
        driver.quit()
    except:
        pass

cleanup_thread = threading.Thread(target=cleanup_driver, daemon=True)
cleanup_thread.start()