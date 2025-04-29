import os
import time
from datetime import datetime, timedelta
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.keys import Keys
from webdriver_manager.chrome import ChromeDriverManager

# === 【1】保存先と前日設定 ===
yesterday_str = (datetime.today() - timedelta(days=1)).strftime("%Y%m%d")
project_root = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))
csv_root_dir = os.path.join(project_root, "マイム日次CSV")
download_dir = os.path.join(csv_root_dir, yesterday_str)
new_filename = f"マイムレポート_{yesterday_str}.csv"

# フォルダがなければ作成
os.makedirs(download_dir, exist_ok=True)

# === 【2】Chrome設定 ===
chrome_options = Options()
prefs = {
    "download.default_directory": download_dir,
    "safebrowsing.enabled": "false"
}
chrome_options.add_experimental_option("prefs", prefs)

# ChromeDriver 起動
service = Service(ChromeDriverManager().install())
driver = webdriver.Chrome(service=service, options=chrome_options)

try:
    # === 【3】マイムログインページへ ===
    login_url = "https://dashboard.assistads.net/login/"
    driver.get(login_url)
    driver.maximize_window()
    user_id = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "id_username")))
    password = driver.find_element(By.ID, "id_password")
    user_id.send_keys("0275")
    password.send_keys("iRWSt6WP")
    login_button = driver.find_element(By.XPATH, '/html/body/form/button')
    driver.execute_script("arguments[0].click();", login_button)
    print("[INFO] ログイン完了")

    # === 【4】データページへ ===
    url = "https://dashboard.assistads.net/publisher/display/382/"
    driver.get(url)

    # === 【5】最下部までスクロール ===
    last_height = driver.execute_script("return document.body.scrollHeight")
    while True:
        driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
        time.sleep(2)
        new_height = driver.execute_script("return document.body.scrollHeight")
        if new_height == last_height:
            break
        last_height = new_height
    print("[INFO] ページ最下部に到達")

    # === 【6】Bの日付に前日を入力 ===
    year, month, day = yesterday_str[:4], yesterday_str[4:6], yesterday_str[6:]
    for date_xpath in ['//*[@id="id_start_date"]', '//*[@id="id_end_date"]']:
        date_fields = driver.find_elements(By.XPATH, date_xpath)
        field = date_fields[-1]  # 最後の要素を取得
        driver.execute_script("arguments[0].scrollIntoView(true);", field)
        time.sleep(1)
        field.click()
        field.clear()
        field.send_keys(year)
        field.send_keys(Keys.RIGHT)
        field.send_keys(month)
        field.send_keys(Keys.RIGHT)
        field.send_keys(day)
        field.send_keys(Keys.TAB)
    print("[INFO] 日付入力完了")

    # === 【7】CSV出力ボタンを押下 ===
    csv_button = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, '/html/body/form[2]/div[2]/div/button')))
    driver.execute_script("arguments[0].scrollIntoView(true);", csv_button)
    time.sleep(1)
    driver.execute_script("arguments[0].click();", csv_button)
    print("[INFO] CSVのダウンロードを開始しました")

    # === 【8】ダウンロード完了待機＆リネーム ===
    time.sleep(5)
    latest_file = None
    for _ in range(30):
        files = [f for f in os.listdir(download_dir) if f.endswith(".csv")]
        if files:
            latest_file = max(files, key=lambda x: os.path.getmtime(os.path.join(download_dir, x)))
            break
        time.sleep(1)

    if latest_file:
        old_path = os.path.join(download_dir, latest_file)
        new_path = os.path.join(download_dir, new_filename)
        if os.path.exists(new_path):
            os.remove(new_path)
        os.rename(old_path, new_path)
        print(f"[INFO] CSVを {new_filename} にリネーム完了")
    else:
        print("[WARNING] CSVファイルが見つかりませんでした")

except Exception as e:
    print(f"[ERROR] 処理中にエラー: {e}")

finally:
    driver.quit()
