from flask import Flask, request, jsonify
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload, MediaFileUpload
import os
import io
import uuid
import requests
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font

# CONFIGURATION
FOLDER_ID = '1diAVIuJdsOQhLEQuFzie6QACakeOie25'
MAKE_WEBHOOK_URL = 'https://hook.eu2.make.com/h6jsruunr7u01wobm995dj8wtcmafph8'
INVENTORY_FILE_ID = '1McHVVICDeeMRiA1fRU7inHmbSUCzeOD2'
SCOPES = ['https://www.googleapis.com/auth/drive']

# AUTH
creds = service_account.Credentials.from_service_account_file(
    os.environ["GOOGLE_APPLICATION_CREDENTIALS"], scopes=SCOPES
)
drive_service = build('drive', 'v3', credentials=creds)

app = Flask(__name__)

def get_latest_inventory_from_drive():
    try:
        request = drive_service.files().get_media(fileId=INVENTORY_FILE_ID)
        fh = io.BytesIO()
        downloader = MediaIoBaseDownload(fh, request)
        done = False
        while not done:
            _, done = downloader.next_chunk()
        fh.seek(0)
        df = pd.read_excel(fh, engine="openpyxl")
        return df
    except Exception as e:
        print("Error loading inventory:", e)
        return None

def summarize(df):
    return {
        "stones": int(df.shape[0]),
        "carats": round(df["Cts"].sum(), 2),
        "ppc_avg": round(df["PPC"].mean(), 2),
        "disc_avg": round(df["Discount"].mean(), 2),
        "rap_avg": round(df["RAP Price"].mean(), 2),
        "total_value": round(df["Total Value"].sum(), 2)
    }

def beautify_excel(path, summary):
    wb = load_workbook(path)
    ws = wb.active

    # Insert 2 rows to make space for summary + title
    ws.insert_rows(1, amount=2)

    # Format cells
    title_fill = PatternFill(start_color="8B0000", end_color="8B0000", fill_type="solid")
    white_bold = Font(color="FFFFFF", bold=True)
    pink_fill = PatternFill(start_color="F4CCCC", end_color="F4CCCC", fill_type="solid")
    black_normal = Font(color="000000", bold=False)

    # F1:P1 titles
    titles = ["", "Stones", "Carats", "", "Rap Avg", "PPC Avg", "Avg Disc", "", "", "", "Total Value"]
    for col, val in enumerate(titles, start=6):
        cell = ws.cell(row=1, column=col)
        cell.value = val if val else None
        cell.fill = title_fill
        cell.font = white_bold

    # F2:P2 values or formulas
    formulas = [
        "\"Selection\"",  # F2
        "=SUBTOTAL(3,B4:B3153)",  # G2
        "=SUBTOTAL(9,E4:E3153)",  # H2
        "",  # I2
        "=SUBTOTAL(9,M4:M3153)/H2",  # J2
        "=SUBTOTAL(9,P4:P3153)/H2",  # K2
        "=((K2/J2)-1)*100",  # L2
        "", "", "",  # M2-O2
        "=IF(G2<200,SUBTOTAL(9,P4:P3153),0)"  # P2
    ]
    for col, formula in enumerate(formulas, start=6):
        cell = ws.cell(row=2, column=col)
        cell.fill = pink_fill
        cell.font = Font(bold=(col == 6), name="Aptos Narrow", size=11, color="000000")
        if formula:
            cell.value = formula

    # A3:AF3 headers
    for col in range(1, ws.max_column + 1):
        cell = ws.cell(row=3, column=col)
        cell.fill = title_fill
        cell.font = Font(color="FFFFFF", name="Aptos Narrow", size=11, bold=False)

    wb.save(path)

@app.route("/generate", methods=["POST"])
def generate():
    data = request.get_json()
    email = data.get("email")
    filters = data.get("filters", {})
    df = get_latest_inventory_from_drive()

    if not email:
        return jsonify({"error": "Missing email"}), 400

    if df is None:
        return jsonify({"error": "Could not load inventory"}), 500

    # For testing: no filtering
    filtered_df = df.copy()

    if filtered_df.empty:
        return jsonify({"error": "No matching stones"}), 404

    summary = summarize(filtered_df)

    filename = f"filtered_{uuid.uuid4().hex[:6]}.xlsx"
    filtered_df.to_excel(filename, index=False)
    beautify_excel(filename, summary)

    metadata = {"name": filename, "parents": [FOLDER_ID]}
    media = MediaFileUpload(filename, mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
    file = drive_service.files().create(body=metadata, media_body=media, fields='id').execute()
    file_id = file.get("id")
    drive_service.permissions().create(fileId=file_id, body={"role": "reader", "type": "anyone"}).execute()

    link = f"https://drive.google.com/file/d/{file_id}/view?usp=sharing"
    os.remove(filename)

    try:
        requests.post(MAKE_WEBHOOK_URL, json={"email": email, "file_link": link, "file_name": filename})
    except:
        pass

    return jsonify({"message": "File filtered", "summary": summary, "link": link})

if __name__ == "__main__":
    app.run(host='0.0.0.0', port=int(os.environ.get("PORT", 10000)))
