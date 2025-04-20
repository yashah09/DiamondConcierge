from flask import Flask, request, jsonify
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload
import os
import uuid
import requests
import pandas as pd
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils.dataframe import dataframe_to_rows

# CONFIGURATION
FOLDER_ID = '1diAVIuJdsOQhLEQuFzie6QACakeOie25'
INVENTORY_FILE_ID = '1McHVVICDeeMRiA1fRU7inHmbSUCzeOD2'
MAKE_WEBHOOK_URL = 'https://hook.eu2.make.com/h6jsruunr7u01wobm995dj8wtcmafph8'

SCOPES = ['https://www.googleapis.com/auth/drive']
creds = service_account.Credentials.from_service_account_file(
    os.environ["GOOGLE_APPLICATION_CREDENTIALS"], scopes=SCOPES
)

app = Flask(__name__)
drive_service = build('drive', 'v3', credentials=creds)

def get_latest_inventory_from_drive():
    try:
        request_file = drive_service.files().get_media(fileId=INVENTORY_FILE_ID)
        with open("temp_inventory.xlsx", "wb") as f:
            downloader = MediaFileUpload("temp_inventory.xlsx")
            request_file.execute(f)
        df = pd.read_excel("temp_inventory.xlsx", engine="openpyxl")
        os.remove("temp_inventory.xlsx")
        return df
    except Exception as e:
        print("Error reading inventory:", e)
        return None

def generate_summary(df):
    count = len(df)
    cts = round(df['Cts'].sum(), 2)
    rap_avg = round(df['RAP Price'].mean(), 0) if count else 0
    ppc_avg = round(df['PPC'].mean(), 0) if count else 0
    disc_avg = round(df['Discount'].mean(), 1) if count else 0
    total_value = round(df['Total Value'].sum(), 0)
    return ["Selection", "Stones", "Carat", "Rap Avg", "PPC Avg", "Avg Disc", "", "", "", "Total Value"], ["", count, cts, rap_avg, ppc_avg, disc_avg, "", "", "", total_value]

def write_to_excel(df):
    wb = Workbook()
    ws = wb.active

    header_row, values_row = generate_summary(df)

    ws.append([])
    ws.append(values_row)
    ws.append(header_row)

    for r in dataframe_to_rows(df, index=False, header=True):
        ws.append(r)

    red_fill = PatternFill(start_color="990000", end_color="990000", fill_type="solid")
    pink_fill = PatternFill(start_color="F4CCCC", end_color="F4CCCC", fill_type="solid")
    white_font = Font(color="FFFFFF", bold=True)
    bold_font = Font(bold=True)

    for cell in ws[2][5:15]:
        cell.fill = pink_fill
        cell.font = bold_font if cell.col_idx == 6 else Font(bold=False)
        cell.alignment = Alignment(horizontal="center")

    for cell in ws[3][5:15]:
        cell.fill = red_fill
        cell.font = white_font
        cell.alignment = Alignment(horizontal="center")

    for cell in ws[4]:
        cell.fill = red_fill
        cell.font = white_font
        cell.alignment = Alignment(horizontal="center")

    filename = f"filtered_{uuid.uuid4().hex[:8]}.xlsx"
    wb.save(filename)
    return filename

@app.route("/generate", methods=["POST"])
def generate():
    data = request.get_json()
    email = data.get("email")

    if not email:
        return jsonify({"error": "Missing email"}), 400

    df = get_latest_inventory_from_drive()
    if df is None:
        return jsonify({"error": "Could not load inventory"}), 500

    if df.empty:
        return jsonify({"error": "No matching stones"}), 404

    filename = write_to_excel(df)

    file_metadata = {
        'name': filename,
        'parents': [FOLDER_ID]
    }
    media = MediaFileUpload(filename, mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
    uploaded = drive_service.files().create(body=file_metadata, media_body=media, fields='id').execute()
    os.remove(filename)

    file_id = uploaded.get("id")
    drive_service.permissions().create(fileId=file_id, body={"role": "reader", "type": "anyone"}).execute()

    file_link = f"https://drive.google.com/file/d/{file_id}/view?usp=sharing"

    requests.post(MAKE_WEBHOOK_URL, json={
        "email": email,
        "file_link": file_link,
        "file_name": filename
    })

    return jsonify({"message": "File generated and uploaded", "link": file_link})

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", 10000)))
