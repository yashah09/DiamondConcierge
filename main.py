from flask import Flask, request, jsonify
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload, MediaIoBaseDownload
import os
import uuid
import requests
import pandas as pd
import io
from openpyxl import Workbook, load_workbook
from openpyxl.styles import PatternFill, Font

# CONFIGURATION
FOLDER_ID = '1diAVIuJdsOQhLEQuFzie6QACakeOie25'
SHEET_ID = '1ZZHmuGyxgq6ISyTpqSupugrNczYiyLWQY6w5oEgWnwc'
MAKE_WEBHOOK_URL = 'https://hook.eu2.make.com/h6jsruunr7u01wobm995dj8wtcmafph8'
INVENTORY_FILE_ID = '1McHVVICDeeMRiA1fRU7inHmbSUCzeOD2'

# AUTH
SCOPES = ['https://www.googleapis.com/auth/drive', 'https://www.googleapis.com/auth/spreadsheets']
creds = service_account.Credentials.from_service_account_file(
    os.environ["GOOGLE_APPLICATION_CREDENTIALS"], scopes=SCOPES
)

drive_service = build('drive', 'v3', credentials=creds)
app = Flask(__name__)


def get_latest_inventory_from_drive():
    try:
        request_file = drive_service.files().get_media(fileId=INVENTORY_FILE_ID)
        fh = io.BytesIO()
        downloader = MediaIoBaseDownload(fh, request_file)
        done = False
        while not done:
            status, done = downloader.next_chunk()
        fh.seek(0)
        df = pd.read_excel(fh, engine='openpyxl')
        return df
    except Exception as e:
        print("Error reading inventory:", e)
        return None


def create_formatted_excel(df):
    filename = f"filtered_{uuid.uuid4().hex[:8]}.xlsx"
    df.to_excel(filename, index=False, startrow=3)
    wb = load_workbook(filename)
    ws = wb.active

    # Add summary row in F2–P2
    count = len(df)
    total_cts = round(df['Cts'].sum(), 2)
    avg_ppc = round(df['PPC'].mean(), 2) if not df['PPC'].isnull().all() else 0
    avg_discount = round(df['Discount'].mean(), 2) if not df['Discount'].isnull().all() else 0
    avg_total_value = round(df['Total Value'].mean(), 2) if not df['Total Value'].isnull().all() else 0

    summary_headers = ["Selection", "Count", "Cts", "Avg. PPC", "Avg. Disc", "Avg. Value"]
    summary_values = ["Selection", count, total_cts, avg_ppc, avg_discount, avg_total_value]

    for col_index, value in enumerate(summary_headers, start=6):
        cell = ws.cell(row=1, column=col_index)
        cell.value = value

    for col_index, value in enumerate(summary_values, start=6):
        cell = ws.cell(row=2, column=col_index)
        cell.value = value

    # Styles
    red_fill = PatternFill(start_color='8B0000', end_color='8B0000', fill_type='solid')
    pink_fill = PatternFill(start_color='F4CCCC', end_color='F4CCCC', fill_type='solid')
    white_bold = Font(color='FFFFFF', bold=True)
    black_bold = Font(color='000000', bold=True)
    black_normal = Font(color='000000', bold=False)

    for col in range(6, 6 + len(summary_headers)):
        ws.cell(row=1, column=col).fill = red_fill
        ws.cell(row=1, column=col).font = white_bold
        ws.cell(row=2, column=col).fill = pink_fill
        ws.cell(row=2, column=col).font = black_bold if col == 6 else black_normal

    for col in range(1, ws.max_column + 1):
        ws.cell(row=4, column=col).fill = red_fill
        ws.cell(row=4, column=col).font = white_bold

    wb.save(filename)
    return filename


@app.route('/generate', methods=['POST'])
def generate():
    data = request.get_json()
    filters = data.get("filters", {})
    email = data.get("email")

    if not email:
        return jsonify({"error": "Missing email"}), 400

    df = get_latest_inventory_from_drive()
    if df is None:
        return jsonify({"error": "Could not load inventory"}), 500

    # NOTE: Assume filtering is applied here properly
    filtered_df = df.copy()  # Placeholder for actual filtering logic
    if filtered_df.empty:
        return jsonify({"error": "No matching stones"}), 404

    filename = create_formatted_excel(filtered_df)

    file_metadata = {
        'name': filename,
        'parents': [FOLDER_ID]
    }
    media = MediaFileUpload(filename, mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
    uploaded = drive_service.files().create(body=file_metadata, media_body=media, fields='id').execute()
    os.remove(filename)

    file_id = uploaded.get('id')
    file_link = f"https://drive.google.com/file/d/{file_id}/view?usp=sharing"

    drive_service.permissions().create(fileId=file_id, body={
        'role': 'reader',
        'type': 'anyone'
    }).execute()

    payload = {
        "email": email,
        "file_link": file_link,
        "file_name": filename
    }

    try:
        requests.post(MAKE_WEBHOOK_URL, json=payload)
    except Exception as e:
        print("Failed to notify Make.com:", e)

    return jsonify({"message": "File generated and uploaded", "link": file_link})


if __name__ == '__main__':
    app.run(host='0.0.0.0', port=int(os.environ.get('PORT', 10000)))
