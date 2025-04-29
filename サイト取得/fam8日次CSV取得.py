import os
import time
import pandas as pd
import chardet
from datetime import datetime, timedelta
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager

# === 【1】定数定義 ===
LOGIN_URL = "https://admin.fam-8.net/report/index.php"
ID = "admin"
PASSWORD = "fhC7UPJiforgKTJ8"
CSV_FILENAME_TEMPLATE = "fam8レポート_{date}.csv"

# === 【2】保存先：プロジェクトルート/マイム日次CSV/前日付/ ===
project_root = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))
csv_root_dir = os.path.join(project_root, "マイム日次CSV")
yesterday_str = (datetime.today() - timedelta(days=1)).strftime("%Y%m%d")
download_dir = os.path.join(csv_root_dir, yesterday_str)
os.makedirs(download_dir, exist_ok=True)

# === 【3】Chrome起動 ===
def setup_driver():
    options = webdriver.ChromeOptions()
    prefs = {"download.default_directory": download_dir}
    options.add_experimental_option("prefs", prefs)
    service = Service(ChromeDriverManager().install())
    driver = webdriver.Chrome(service=service, options=options)
    driver.get(LOGIN_URL)
    return driver

# === 【4】ログイン処理 ===
def login(driver):
    print("[INFO] fam8 ログイン開始")
    driver.find_element(By.XPATH, '//*[@id="topmenu"]/tbody/tr[2]/td/div[1]/form/div/table/tbody/tr[1]/td/input').send_keys(ID)
    driver.find_element(By.XPATH, '//*[@id="topmenu"]/tbody/tr[2]/td/div[1]/form/div/table/tbody/tr[2]/td/input').send_keys(PASSWORD)
    driver.find_element(By.XPATH, '//*[@id="topmenu"]/tbody/tr[2]/td/div[1]/form/div/table/tbody/tr[3]/td/input[2]').click()
    time.sleep(3)
    print("[INFO] fam8 ログイン完了")

# === 【5】CSVダウンロード ===
def download_csv(driver):
    print("[INFO] CSV ダウンロード開始")
    wait = WebDriverWait(driver, 15)
    wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="sidemenu"]/div[3]/a[4]/div'))).click()
    time.sleep(1)
    wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="display_modesummary_mode"]'))).click()
    time.sleep(1)
    wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="main_area"]/form/div[1]/table[2]/tbody/tr[1]/td/select[3]/option[2]'))).click()
    time.sleep(1)
    search_box = wait.until(EC.presence_of_element_located((By.XPATH, '//*[@id="main_area"]/form/div[1]/input[7]')))
    search_box.clear()
    search_box.send_keys("マイム")
    time.sleep(1)
    search_button = wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="main_area"]/form/div[1]/input[10]')))
    search_button.click()
    time.sleep(3)
    download_button = wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="topmenu"]/table/tbody/tr/td[4]/table/tbody/tr[2]/td[2]/input[1]')))
    download_button.click()
    print("[INFO] CSV ダウンロード指示完了")
    time.sleep(6)

# === 【6】CSVファイルをリネーム ===
def rename_csv():
    files = [f for f in os.listdir(download_dir) if f.endswith(".csv")]
    if not files:
        print("[ERROR] ダウンロードされたCSVが見つかりません")
        exit(1)

    latest_file = max(files, key=lambda x: os.path.getmtime(os.path.join(download_dir, x)))
    old_path = os.path.join(download_dir, latest_file)
    new_filename = CSV_FILENAME_TEMPLATE.format(date=yesterday_str)
    new_path = os.path.join(download_dir, new_filename)

    if os.path.exists(new_path):
        print(f"[WARNING] 既存の {new_filename} を上書きします")
        os.remove(new_path)

    os.rename(old_path, new_path)
    print(f"[INFO] CSVファイルを {new_filename} にリネーム完了")
    return new_path

# === 【7】文字化け防止のエンコーディング検出（検証用） ===
def detect_encoding(file_path):
    with open(file_path, "rb") as f:
        result = chardet.detect(f.read())
    return result["encoding"]

# === 【8】必要列抽出の前処理（検証用） ===
def preprocess_csv(csv_path):
    encoding = detect_encoding(csv_path)
    try:
        df = pd.read_csv(csv_path, skiprows=2, encoding=encoding)
    except UnicodeDecodeError:
        print("[WARNING] エンコーディングエラー発生、cp932 で再試行")
        df = pd.read_csv(csv_path, skiprows=2, encoding="cp932")

    df = df.iloc[:, [0, 6]]
    df.columns = ["広告枠ID", "Imp"]
    df = df[~df["広告枠ID"].astype(str).str.contains("total", case=False, na=False)]
    df = df[df["広告枠ID"].notna()]
    print(f"[DEBUG] 読み込んだCSVの行数: {len(df)}")
    return df

# === 【9】メイン ===
if __name__ == "__main__":
    driver = setup_driver()
    login(driver)
    download_csv(driver)
    driver.quit()
    latest_csv = rename_csv()
    processed_data = preprocess_csv(latest_csv)
    print("[INFO] fam8 処理完了")
