import gspread
from google.oauth2.service_account import Credentials
import requests
import datetime
from google.auth.transport.requests import Request
import json
#20250404に￥でエンコードエラーがでたので以下を追加
import sys
import io

sys.stdout.reconfigure(encoding='utf-8')
sys.stderr.reconfigure(encoding='utf-8')

# 出力のエンコーディングを変更
# sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
# 標準エラー出力のエンコーディングもUTF-8に変更（エラーログ用）
# sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8')
# ログファイルの書き込みにもUTF-8を強制
# log_file_path = 'your_log_file.log'
# with open(log_file_path, 'a', encoding='utf-8') as log_file:
#     log_file.write('ログ内容')
# print("ログの文字コード: UTF-8で表示")
#----------------------------------------

# **Googleスプレッドシートに接続**
def get_google_sheet():
    scope = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]
    creds = Credentials.from_service_account_file("mime-454407-591d0c95ff2f.json", scopes=scope)
    client = gspread.authorize(creds)
    sheet_id = "1jnC9BG6Y3Jesx_i0FTGx_8zx04SgzbARGbEFynheshU"  # 本番用のsheet_idに変更
    return client.open_by_key(sheet_id), creds


# **日付を和名で出力**
def format_date_japanese(date):
    days = ["月", "火", "水", "木", "金", "土", "日"]  # 修正！
    return date.strftime(f"%Y年%m月%d日 ({days[date.weekday()]})")

# **セルの値を日付型に変換**
def convert_to_date(value):
    if isinstance(value, (datetime.datetime, datetime.date)):
        return value
    try:
        return datetime.datetime.strptime(value, "%Y/%m/%d").date()
    except ValueError:
        return None

# **列番号をアルファベットに変換**
def column_letter(n):
    result = ""
    while n > 0:
        n -= 1
        result = chr(65 + (n % 26)) + result
        n //= 26
    return result

# **数値変換・エラー処理**
def convert_cell_value(val):
    if isinstance(val, (int, float)):
        return val
    val = str(val).strip()
    if "%" in val:
        try:
            return float(val.replace("%", "").replace(",", "")) / 100
        except ValueError:
            return ""
    if val in ["#DIV/0!", "#REF!", "#VALUE!", "#ERROR!"]:
        return ""
    try:
        return float(val.replace(",", ""))
    except ValueError:
        return val

# **パーセントフォーマット適用**
def set_percent_format(sheet, columns, row):
    sheet_id = sheet._properties["sheetId"]
    requests_payload = [{"repeatCell": {
        "range": {"sheetId": sheet_id, "startRowIndex": row - 1, "endRowIndex": row, "startColumnIndex": col - 1, "endColumnIndex": col},
        "cell": {"userEnteredFormat": {"numberFormat": {"type": "PERCENT", "pattern": "0.00%"}}},
        "fields": "userEnteredFormat.numberFormat"}} for col in columns]
    body = {"requests": requests_payload}
    access_token = sheet.spreadsheet.client.auth.token
    url = f"https://sheets.googleapis.com/v4/spreadsheets/{sheet.spreadsheet.id}:batchUpdate"
    headers = {"Authorization": f"Bearer {access_token}", "Content-Type": "application/json"}
    response = requests.post(url, headers=headers, data=json.dumps(body))
    if response.status_code != 200:
        raise Exception(f"[API ERROR] %フォーマット適用失敗: {response.status_code}: {response.text}")

# **通貨形式（¥0）を設定する関数 - 新規追加**
def set_currency_format(sheet, columns, row, creds=None):
    """特定の列を通貨形式（¥0）に設定する"""
    if creds and not creds.valid:
        creds.refresh(Request())
        
    sheet_id = sheet._properties["sheetId"]
    requests_payload = [{"repeatCell": {
        "range": {"sheetId": sheet_id, "startRowIndex": row - 1, "endRowIndex": row, "startColumnIndex": col - 1, "endColumnIndex": col},
        "cell": {"userEnteredFormat": {"numberFormat": {"type": "CURRENCY", "pattern": "¥0"}}},
        "fields": "userEnteredFormat.numberFormat"}} for col in columns]
    
    body = {"requests": requests_payload}
    access_token = sheet.spreadsheet.client.auth.token
    url = f"https://sheets.googleapis.com/v4/spreadsheets/{sheet.spreadsheet.id}:batchUpdate"
    headers = {"Authorization": f"Bearer {access_token}", "Content-Type": "application/json"}
    response = requests.post(url, headers=headers, data=json.dumps(body))
    
    if response.status_code != 200:
        raise Exception(f"[API ERROR] 通貨フォーマット適用失敗: {response.status_code}: {response.text}")
    else:
        columns_str = ", ".join([column_letter(col) for col in columns])
        print(f"[INFO] 通貨形式(¥0)を {columns_str}{row} に適用しました")

# **数値形式（0 1235）を設定する関数 - 新規追加**
def set_number_format_zero(sheet, column, row, creds=None):
    """特定のセルを数値形式（0 1235）に設定する"""
    if creds and not creds.valid:
        creds.refresh(Request())
        
    sheet_id = sheet._properties["sheetId"]
    requests_payload = [{
        "repeatCell": {
            "range": {"sheetId": sheet_id, "startRowIndex": row - 1, "endRowIndex": row, "startColumnIndex": column - 1, "endColumnIndex": column},
            "cell": {"userEnteredFormat": {"numberFormat": {"type": "NUMBER", "pattern": "0"}}},
            "fields": "userEnteredFormat.numberFormat"
        }
    }]
    
    body = {"requests": requests_payload}
    access_token = sheet.spreadsheet.client.auth.token
    url = f"https://sheets.googleapis.com/v4/spreadsheets/{sheet.spreadsheet.id}:batchUpdate"
    headers = {"Authorization": f"Bearer {access_token}", "Content-Type": "application/json"}
    response = requests.post(url, headers=headers, data=json.dumps(body))
    
    if response.status_code != 200:
        raise Exception(f"[API ERROR] 数値フォーマット適用失敗: {response.status_code}: {response.text}")
    else:
        print(f"[INFO] 数値形式(0)を {column_letter(column)}{row} に適用しました")

# **B列の数式を更新する関数 - 新規追加**
def update_b_column_formula(sheet, row, target_columns, creds=None):
    """B列の数式をISFORMULA関数を使った形式に更新する"""
    formula_parts = []
    
    for col in target_columns:
        col_letter = column_letter(col)
        formula_parts.append(f"ROUND(IF(ISFORMULA({col_letter}{row}), 0, {col_letter}{row}), 1)")
    
    formula = "=" + " + ".join(formula_parts)
    print(f"[DEBUG] 生成されたB列数式: {formula}")
    
    # B列に新しい数式を設定
    sheet.update_cell(row, 2, formula)
    print(f"[INFO] B{row}に新しい数式を設定しました")
    
    # 数値形式を設定（0 1235形式）
    set_number_format_zero(sheet, 2, row, creds)

# **合計対象となる列番号を取得する関数 - 新規追加**
def get_sum_target_columns():
    """合計対象となる列番号のリストを返す（I, O, U, AA, AG...）"""
    # I列から始まり、6列ごとに対象列がある（I, O, U, AA, ...）
    base_col = 9  # I列は9番目
    return [base_col + i * 6 for i in range(20)]  # 20個の対象列（I〜DS）

# **関数を生成**
def get_percent_column_indexes(start_col=5, formula_block=6, offset=2):
    return [start_col + i * formula_block + offset for i in range(31)]  # 20250417 GM列対応：31ブロック分に拡張

def generate_formulas(row):
    formulas = []
    for i in range(31):  # E〜GM列対応：6列ブロック × 31 = 186列分（1列空き補正含む）　20250417追加
        offset = i * 6
        e, f, g, h, i_col = column_letter(5 + offset), column_letter(6 + offset), column_letter(7 + offset), column_letter(8 + offset), column_letter(9 + offset)
        formulas.extend([
            f"=IFERROR(VLOOKUP({e}$1, '抽出レポート'!B:C, 2, 0), \"\")",
            f"=IFERROR(VLOOKUP({f}$1, '抽出レポート'!H:M, 4, 0), \"\")",
            f"=IFERROR(-({f}{row} - {e}{row}) / {e}{row}, \"\")",
            f"=IF(TRIM({g}{row}) * 1 >= 5%, \"補填対象\", \"\")",
            f"=IF({h}{row} = \"補填対象\", (VLOOKUP({e}$1, '抽出レポート'!$B:$D, 3, 0) / 1000) * {i_col}$1 - VLOOKUP({f}$1, '抽出レポート'!$H:$O, 6, FALSE), \"\")",
            ""
        ])
    return formulas

# **数式を強制挿入（batchUpdate USER_ENTERED）**
def update_with_user_entered_force(sheet, creds, start_col, row, formulas):
    if not creds.valid:
        creds.refresh(Request())

    access_token = creds.token
    spreadsheet_id = sheet.spreadsheet.id
    sheet_id = sheet._properties["sheetId"]

    cells = [{"userEnteredValue": {"formulaValue" if formula.startswith("=") else "stringValue": formula}} for formula in formulas]

    requests_payload = [{
        "updateCells": {
            "rows": [{"values": cells}],
            "fields": "userEnteredValue",
            "range": {
                "sheetId": sheet_id,
                "startRowIndex": row - 1,
                "endRowIndex": row,
                "startColumnIndex": start_col - 1,
                "endColumnIndex": start_col - 1 + len(formulas)
            }
        }
    }]

    url = f"https://sheets.googleapis.com/v4/spreadsheets/{spreadsheet_id}:batchUpdate"
    headers = {"Authorization": f"Bearer {access_token}", "Content-Type": "application/json"}
    response = requests.post(url, headers=headers, data=json.dumps({"requests": requests_payload}))

    if response.status_code != 200:
        raise Exception(f"[API ERROR] 数式挿入失敗: {response.status_code}: {response.text}")

# **メイン処理**
def main():
    sheet, creds = get_google_sheet()
    sheet = sheet.worksheet("日時レポート")
    report_sheet = sheet.spreadsheet.worksheet("抽出レポート")

    if not sheet or not report_sheet:
        print("[ERROR] 指定シートが見つかりません")
        return

    today = datetime.date.today()
    yesterday = today - datetime.timedelta(days=1)
    today_str, yesterday_str = today.strftime("%Y/%m/%d"), yesterday.strftime("%Y/%m/%d")

    print("[INFO] 本日の日付 (和名):", format_date_japanese(today))
    print("[INFO] 昨日の日付 (和名):", format_date_japanese(yesterday))

    data = sheet.col_values(4)
    target_row, yesterday_row = None, None

    for i in range(3, len(data)):
        cell_value = convert_to_date(data[i])
        if cell_value:
            cell_str = cell_value.strftime("%Y/%m/%d")
            if cell_str == today_str:
                target_row = i + 1
            elif cell_str == yesterday_str:
                yesterday_row = i + 1

    if not target_row:
        print(f"[ERROR] 今日の日付 {today_str} がD列に見つかりません")
        return
    if not yesterday_row:
        print(f"[ERROR] 昨日の日付 {yesterday_str} がD列に見つかりません")
        return

    print(f"[DEBUG] 今日の行: {target_row} / 昨日の行: {yesterday_row}")
    print(f"[処理対象] 関数 = {target_row}行 / 値貼付 = {yesterday_row}行")

    formulas = generate_formulas(target_row)
    print(f"[DEBUG] 挿入される数式の総数: {len(formulas)} 列")

    start_col, end_col = 5, 190  # E列（5）～GM列（191列目-1）に固定、20250417対応
    start_letter, end_letter = column_letter(start_col), column_letter(end_col)
    cell_range = f"{start_letter}{target_row}:{end_letter}{target_row}"

    update_with_user_entered_force(sheet, creds, start_col, target_row, formulas)
    print(f"[完了] 数式挿入 → 範囲: {cell_range}")

    # **関数セルのテキストカラーを白に変更**
    try:
        sheet_id = sheet._properties["sheetId"]
        requests_payload = [{
            "repeatCell": {
                "range": {
                    "sheetId": sheet_id,
                    "startRowIndex": target_row - 1,
                    "endRowIndex": target_row,
                    "startColumnIndex": 4,  # E列 (5-1)
                    "endColumnIndex": 190,  # GM列 (191-1) 20250417追加
                },
                "cell": {
                    "userEnteredFormat": {
                        "textFormat": {"foregroundColor": {"red": 1, "green": 1, "blue": 1}}  # テキスト色を白にする
                    }
                },
                "fields": "userEnteredFormat.textFormat"
            }
        }]

        url = f"https://sheets.googleapis.com/v4/spreadsheets/{sheet.spreadsheet.id}:batchUpdate"
        headers = {"Authorization": f"Bearer {creds.token}", "Content-Type": "application/json"}
        response = requests.post(url, headers=headers, data=json.dumps({"requests": requests_payload}))

        if response.status_code == 200:
            print(f"[INFO] 適用範囲: E{target_row}:EE{target_row} に白文字を適用")
        else:
            print(f"[ERROR] セルの白文字適用失敗: {response.status_code}: {response.text}")

    except Exception as e:
        print(f"[ERROR] テキストカラー変更処理中にエラー発生: {str(e)}")

    # **値貼り付け行のテキストカラーを黒に変更**
    try:
        sheet_id = sheet._properties["sheetId"]
        requests_payload = [{
            "repeatCell": {
                "range": {
                    "sheetId": sheet_id,
                    "startRowIndex": yesterday_row - 1,
                     "endRowIndex": yesterday_row,
                    "startColumnIndex": 4,  # E列 (5-1)
                    "endColumnIndex": 190,  # GM列 (191-1)　20250417追加
                    },
                "cell": {
                "userEnteredFormat": {
                "textFormat": {"foregroundColor": {"red": 0, "green": 0, "blue": 0}}  # テキスト色を黒にする
                    }
                },
                "fields": "userEnteredFormat.textFormat"
            }
        }]

        # 昨日の日付の行（値貼り付け行）に黒色を適用するAPIリクエストを送信
        url = f"https://sheets.googleapis.com/v4/spreadsheets/{sheet.spreadsheet.id}:batchUpdate"
        headers = {"Authorization": f"Bearer {creds.token}", "Content-Type": "application/json"}
        response = requests.post(url, headers=headers, data=json.dumps({"requests": requests_payload}))

        if response.status_code == 200:
            print(f"[INFO] 適用範囲: E{yesterday_row}:EE{yesterday_row} に黒文字を適用")
        else:
            print(f"[ERROR] セルの黒文字適用失敗: {response.status_code}: {response.text}")

    except Exception as e:
        print(f"[ERROR] テキストカラー変更処理中にエラー発生: {str(e)}")

    # **昨日の行に値を貼り付け**
    result_row = sheet.row_values(target_row)[start_col - 1:end_col]
    cleaned_values = [convert_cell_value(val) for val in result_row]
    paste_range = f"{start_letter}{yesterday_row}:{end_letter}{yesterday_row}"

    print(f"[DEBUG] 貼り付け範囲: {paste_range}")
    print(f"[DEBUG] 貼り付け内容: {cleaned_values}")
    sheet.update(range_name=paste_range, values=[cleaned_values])

    # **パーセント形式を適用**
    percent_columns = get_percent_column_indexes()
    set_percent_format(sheet, percent_columns, yesterday_row)

    # **新機能：通貨形式を適用 - I, O, U, AA, AG...列など**
    sum_target_columns = get_sum_target_columns()
    set_currency_format(sheet, sum_target_columns, yesterday_row, creds)

    # **新機能：B列の数式をISFORMULA関数を使った形式に更新**
    update_b_column_formula(sheet, yesterday_row, sum_target_columns, creds)
    
    # **B列の数式をISFORMULA関数を使った形式に更新 - 今日の行にも適用**
    update_b_column_formula(sheet, target_row, sum_target_columns, creds)

    print(f"[完了] 関数: {target_row}行 / 値貼付: {yesterday_row}行 → 正常終了")

if __name__ == "__main__":
    main()