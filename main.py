from flask import Flask, request, jsonify
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload, MediaIoBaseDownload
import os
import uuid
import datetime
import requests
import pandas as pd
import io

# CONFIGURATION
FOLDER_ID = '1diAVIuJdsOQhLEQuFzie6QACakeOie25'
SHEET_ID = '1ZZHmuGyxgq6ISyTpqSupugrNczYiyLWQY6w5oEgWnwc'
MAKE_WEBHOOK_URL = 'https://hook.eu2.make.com/h6jsruunr7u01wobm995dj8wtcmafph8'
INVENTORY_FILE_ID = 'your-google-drive-file-id-here'  # 🔁 Replace this with your actual file ID

# AUTH
SCOPES = ['https://www.googleapis.com/auth/drive', 'https://www.googleapis.com/auth/spreadsheets']
creds = service_account.Credentials.from_service_account_file(
    'credentials.json', scopes=SCOPES
)

# INIT
app = Flask(__name__)
drive_service = build('drive', 'v3', credentials=creds)
sheets_service = build('sheets', 'v4', credentials=creds)

# ✅ New: Pull inventory from Google Drive
def get_latest_inventory_from_drive(file_id=INVENTORY_FILE_ID):
    try:
        request = drive_service.files().get_media(fileId=file_id)
        fh = io.BytesIO()
        downloader = MediaIoBaseDownload(fh, request)

        done = False
        while done is False:
            status, done = downloader.next_chunk()

        fh.seek(0)
        df = pd.read_excel(fh)
        return df

    except Exception as e:
        print("Error downloading inventory file:", e)
        return None

@app.route('/upload', methods=['POST'])
def upload():
    file = request.files['file']
    client_email = request.form.get('email')

    if not file or not client_email:
        return jsonify({'error': 'Missing file or email'}), 400

    # Save file temporarily
    filename = f"temp_{uuid.uuid4()}.xlsx"
    file.save(filename)

    # Upload to Google Drive
    file_metadata = {
        'name': file.filename,
        'parents': [FOLDER_ID]
    }
    media = MediaFileUpload(filename, mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
    uploaded = drive_service.files().create(body=file_metadata, media_body=media, fields='id').execute()
    os.remove(filename)

    file_id = uploaded.get('id')
    file_link = f"https://drive.google.com/file/d/{file_id}/view?usp=sharing"

    # Make file shareable
    drive_service.permissions().create(fileId=file_id, body={
        'role': 'reader',
        'type': 'anyone'
    }).execute()

    # Append to Google Sheet
    timestamp = datetime.datetime.now().isoformat()
    row = [[timestamp, client_email, file.filename, file_link, "pending"]]

    sheets_service.spreadsheets().values().append(
        spreadsheetId=SHEET_ID,
        range="Sheet1!A1:E1",
        valueInputOption="RAW",
        body={"values": row}
    ).execute()

    # ✅ Send to Make.com webhook
    payload = {
        "email": client_email,
        "file_link": file_link,
        "file_name": file.filename
    }

    try:
        requests.post(MAKE_WEBHOOK_URL, json=payload)
    except Exception as e:
        print("Failed to notify Make.com:", e)

    return jsonify({'message': 'Uploaded successfully', 'link': file_link})
@app.route('/test-inventory', methods=['GET'])
def test_inventory():
    df = get_latest_inventory_from_drive()
    if df is None:
        return jsonify({"error": "Could not read inventory file"}), 500

    # Show first 5 rows as a preview
    return jsonify(df.head(5).to_dict(orient="records"))
if __name__ == '__main__':
    app.run(host="0.0.0.0", port=10000)
