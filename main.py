import os
import json
import io
import pandas as pd
from datetime import datetime
from playwright.sync_api import sync_playwright
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseUpload, MediaIoBaseDownload

# --- 登録者数をスクレイピング ---
def scrape_subscriber_count(channel_url: str) -> int:
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        page = browser.new_page()
        page.goto(channel_url)
        page.wait_for_timeout(5000)  # ページが完全に読み込まれるまで待機
        count_text = page.locator(".odometer-value").all_inner_texts()
        browser.close()
    return int("".join(count_text))


# --- ローカルのExcelファイルにデータ追加保存 ---
def save_to_excel(count: int, file_path: str):
    new_data = {
        "channel": "nanoha_youtube",
        "subscriber_count": count,
        "date": datetime.now().strftime("%Y/%m/%d")
    }

    if os.path.exists(file_path):
        df_existing = pd.read_excel(file_path)
        df = pd.concat([df_existing, pd.DataFrame([new_data])], ignore_index=True)
    else:
        df = pd.DataFrame([new_data])

    df.to_excel(file_path, index=False)


# --- Google Drive APIの認証 ---
def get_drive_service():
    service_account_info = json.loads(os.environ.get("GCP_SERVICE_ACCOUNT_JSON"))
    credentials = service_account.Credentials.from_service_account_info(service_account_info)
    return build("drive", "v3", credentials=credentials)


# --- Driveにアップロード（存在すれば更新） ---
def upload_to_drive(local_file_path: str, file_name: str):
    drive_service = get_drive_service()

    # 同名ファイルのID取得
    results = drive_service.files().list(
        q=f"name = '{file_name}' and trashed = false",
        fields="files(id, name)"
    ).execute()
    files = results.get("files", [])
    file_id = files[0]["id"] if files else None

    media_body = MediaIoBaseUpload(
        io.FileIO(local_file_path, "rb"),
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        resumable=True
    )

    if file_id:
        drive_service.files().update(
            fileId=file_id,
            media_body=media_body
        ).execute()
        print(f"✅ Updated existing file on Drive: {file_name}")
    else:
        file_metadata = {"name": file_name}
        drive_service.files().create(
            body=file_metadata,
            media_body=media_body,
            fields="id"
        ).execute()
        print(f"✅ Uploaded new file to Drive: {file_name}")


# --- メイン処理 ---
if __name__ == "__main__":
    CHANNEL_URL = "https://subscribercounter.com/fullscreen/UCryNrgY4lfJgYkhMNgwHPMg"
    OUTPUT_FILENAME = "nanoha_youtube.xlsx"

    count = scrape_subscriber_count(CHANNEL_URL)
    save_to_excel(count, OUTPUT_FILENAME)
    upload_to_drive(OUTPUT_FILENAME, OUTPUT_FILENAME)
