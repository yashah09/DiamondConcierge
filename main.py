from flask import Flask, request, jsonify
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload
import os
import uuid
import datetime
import requests  # ✅ Make webhook needs this

# CONFIGURATION
FOLDER_ID = '1diAVIuJdsOQhLEQuFzie6QACakeOie25'
SHEET_ID = '1ZZHmuGyxgq6ISyTpqSupugrNczYiyLWQY6w5oEgWnwc'
MAKE_WEBHOOK_URL = 'https://hook.eu2.make.com/h6jsruunr7u01wobm995dj8wtcmafph8'  # ✅ Your Make.com webhook

# AUTH
SCOPES = ['https://www.googleapis.com/auth/drive', 'https://www.googleapis.com/auth/spreadsheets']
creds = service_account.Credentials.from_service_account_file(
    'credentials.json', scopes=SCOPES
)

# INIT
app = Flask(__name__)
drive_service = build('drive', 'v3', credentials=creds)
sheets_service = build('sheets', 'v4', credentials=creds)

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

if __name__ == '__main__':
    app.run(host="0.0.0.0", port=10000)
