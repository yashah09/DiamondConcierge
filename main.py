from flask import Flask, request, jsonify
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload, MediaFileUpload
import os
import uuid
import requests
import pandas as pd
import io
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment

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

# INIT
app = Flask(__name__)
drive_service = build('drive', 'v3', credentials=creds)

def get_latest_inventory_from_drive():
    try:
        request_file = drive_service.files().get_media(fileId=INVENTORY_FILE_ID)
        fh = io.BytesIO()
        downloader = MediaIoBaseDownload(fh, request_file)
        done = False
        while not done:
            status, done = downloader.next_chunk()
        fh.seek(0)
        df = pd.read_excel(fh, engine="openpyxl")
        return df
    except Exception as e:
        print("Error reading inventory:", e)
        return None

def write_styled_excel(df_filtered, filename):
    df_filtered.to_excel(filename, index=False, startrow=3)
    wb = load_workbook(filename)
    ws = wb.active

    # Styles
    red_fill = PatternFill(start_color="9E0000", end_color="9E0000", fill_type="solid")
    pink_fill = PatternFill(start_color="FFCDCD", end_color="FFCDCD", fill_type="solid")
    white_bold = Font(color="FFFFFF", bold=True, name="Aptos Narrow")
    white_regular = Font(color="FFFFFF", bold=False, name="Aptos Narrow")
    pink_bold = Font(color="000000", bold=True, name="Aptos Narrow")
    pink_regular = Font(color="000000", bold=False, name="Aptos Narrow")
    center = Alignment(horizontal="center", vertical="center")

    # F1–P1 styling
    for col in range(6, 17):
        cell = ws.cell(row=1, column=col)
        cell.fill = red_fill
        cell.font = white_bold
        cell.alignment = center

    # F2–P2 styling
    for col in range(6, 17):
        cell = ws.cell(row=2, column=col)
        cell.fill = pink_fill
        cell.alignment = center
        if col == 6:
            cell.font = pink_bold
            cell.value = "Selection"
        else:
            cell.font = pink_regular

    # Row 3 headers
    for cell in ws[3]:
        cell.fill = red_fill
        cell.font = white_bold
        cell.alignment = center

    # Summary formulas
    row_end = 3 + len(df_filtered)
    ws["G2"] = f"=SUBTOTAL(3,B4:B{row_end})"
    ws["H2"] = f"=SUBTOTAL(9,E4:E{row_end})"
    ws["J2"] = f"=SUBTOTAL(9,M4:M{row_end})/H2"
    ws["K2"] = f"=SUBTOTAL(9,P4:P{row_end})/H2"
    ws["L2"] = f"=((K2/J2)-1)*100"
    ws["P2"] = f"=IF(G2<200,SUBTOTAL(9,P4:P{row_end}),0)"

    wb.save(filename)

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

    df_filtered = df.copy()  # replace this with filter logic
    if df_filtered.empty:
        return jsonify({"error": "No matching stones"}), 404

    filename = f"filtered_{uuid.uuid4().hex[:8]}.xlsx"
    write_styled_excel(df_filtered, filename)

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
