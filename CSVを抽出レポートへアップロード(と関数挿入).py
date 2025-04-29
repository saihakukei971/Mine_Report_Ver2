import os
import pandas as pd
import gspread
from google.oauth2.service_account import Credentials
from datetime import datetime, timedelta
import chardet

# === 【1】設定 ===
SPREADSHEET_URL = "https://docs.google.com/spreadsheets/d/1jnC9BG6Y3Jesx_i0FTGx_8zx04SgzbARGbEFynheshU"
SHEET_NAME = "抽出レポート"
SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]
CREDENTIALS_FILE = os.path.join(os.path.dirname(__file__), "mime-454407-591d0c95ff2f.json")

# === 【2】CSVファイルのパス構築（Zikken.py が サイトCSV取得/ にある前提）===
yesterday = (datetime.today() - timedelta(days=1)).strftime("%Y%m%d")
base_dir = os.path.abspath(os.path.dirname(__file__))
csv_dir = os.path.join(base_dir, "マイム日次CSV", yesterday)
fam8_csv = os.path.join(csv_dir, f"fam8レポート_{yesterday}.csv")
mime_csv = os.path.join(csv_dir, f"マイムレポート_{yesterday}.csv")

# === 【3】Google認証 ===
def authenticate_google():
    creds = Credentials.from_service_account_file(CREDENTIALS_FILE, scopes=SCOPES)
    return gspread.authorize(creds)

# === 【4】文字コード検出 ===
def detect_encoding(file_path):
    with open(file_path, "rb") as f:
        return chardet.detect(f.read())["encoding"]

# === 【5】fam8 CSV整形 ===
def preprocess_fam8_csv(csv_path):
    if not os.path.exists(csv_path):
        print(f"[WARNING] fam8 CSVが見つかりません: {csv_path}")
        return None
    print(f"[DEBUG] fam8 CSV パス: {csv_path}")
    encoding = detect_encoding(csv_path)
    try:
        df = pd.read_csv(csv_path, skiprows=2, encoding=encoding)
    except Exception:
        df = pd.read_csv(csv_path, skiprows=2, encoding="cp932")
    df = df.iloc[:, [0, 6]]
    df.columns = ["広告枠ID", "Imp"]
    df = df[df["広告枠ID"].notna()]
    df = df[~df["広告枠ID"].astype(str).str.contains("total", case=False, na=False)]
    return df

# === 【6】マイム CSV読み込み ===
def read_csv_auto_encoding(file_path):
    if not os.path.exists(file_path):
        print(f"[WARNING] マイムCSVが見つかりません: {file_path}")
        return None
    print(f"[DEBUG] マイム CSV パス: {file_path}")
    try:
        encoding = detect_encoding(file_path)
        return pd.read_csv(file_path, encoding=encoding)
    except Exception:
        return pd.read_csv(file_path, encoding="cp932")

# === 【7】アップロード処理 ===
def upload_to_sheet(client, df, start_col, label="不明なCSV"):
    if df is None or df.empty:
        print(f"[SKIP] {label}：データなし")
        return
    sheet = client.open_by_url(SPREADSHEET_URL).worksheet(SHEET_NAME)
    df = df.fillna("")
    start_cell = f"{start_col}4"
    col_count = len(df.columns)
    col_index = ord(start_col.upper()) - 64

    def get_column_letter(idx):
        letters = ""
        while idx > 0:
            idx, rem = divmod(idx - 1, 26)
            letters = chr(65 + rem) + letters
        return letters

    end_col = get_column_letter(col_index + col_count - 1)
    end_cell = f"{end_col}{3 + len(df)}"
    range_str = f"{start_cell}:{end_cell}"

    try:
        sheet.update(values=df.values.tolist(), range_name=range_str)
        print(f"[INFO] {label} → {range_str} にアップロード完了")
    except Exception as e:
        print(f"[ERROR] {label} アップロード失敗: {e}")

# === 【8】D列にC列参照の関数を入れる ===
def insert_fam8_d_column_formula(client):
    try:
        sheet = client.open_by_url(SPREADSHEET_URL).worksheet(SHEET_NAME)
        formulas = [[f'=IF(C{row}<>"" , C{row}*0.95, "")'] for row in range(4, 45)]
        sheet.update(range_name="D4:D44", values=formulas, value_input_option="USER_ENTERED")
        print("[INFO] D列に関数を挿入完了")
    except Exception as e:
        print(f"[ERROR] D列関数の挿入失敗: {e}")

# === 【9】メイン ===
def main():
# ★パス確認ログ
    print("[INFO] スプレッドシート アップロード処理開始")
    print(f"[DEBUG] fam8 CSV full path: {fam8_csv}")
    print(f"[DEBUG] mime CSV full path: {mime_csv}")
    print(f"[EXISTS CHECK] fam8_csv exists: {os.path.exists(fam8_csv)}")
    print(f"[EXISTS CHECK] mime_csv exists: {os.path.exists(mime_csv)}")

    client = authenticate_google()
    fam8_df = preprocess_fam8_csv(fam8_csv)
    upload_to_sheet(client, fam8_df, "B", label="fam8 CSV")
    insert_fam8_d_column_formula(client)
    mime_df = read_csv_auto_encoding(mime_csv)
    upload_to_sheet(client, mime_df, "F", label="マイム CSV")
    print("[INFO] 全アップロード完了")

# === 【10】実行 ===
if __name__ == "__main__":
    main()
