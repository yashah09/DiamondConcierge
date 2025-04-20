from flask import Flask, request, jsonify
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload
import os
import uuid
import requests
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font
from openpyxl.utils import get_column_letter

# CONFIGURATION
FOLDER_ID = '1diAVIuJdsOQhLEQuFzie6QACakeOie25'
MAKE_WEBHOOK_URL = 'https://hook.eu2.make.com/h6jsruunr7u01wobm995dj8wtcmafph8'
INVENTORY_FILE_ID = '1McHVVICDeeMRiA1fRU7inHmbSUCzeOD2'

# AUTH
SCOPES = ['https://www.googleapis.com/auth/drive']
creds = service_account.Credentials.from_service_account_file(
    os.environ["GOOGLE_APPLICATION_CREDENTIALS"], scopes=SCOPES
)

drive_service = build('drive', 'v3', credentials=creds)
app = Flask(__name__)

def get_latest_inventory_from_drive():
    try:
        request_file = drive_service.files().get_media(fileId=INVENTORY_FILE_ID)
        with open("temp_inventory.xlsx", "wb") as f:
            downloader = request_file.execute()
            f.write(downloader)
        df = pd.read_excel("temp_inventory.xlsx", engine="openpyxl")
        os.remove("temp_inventory.xlsx")
        return df
    except Exception as e:
        print("Error reading inventory:", e)
        return None

def beautify_excel(path, summary):
    wb = load_workbook(path)
    ws = wb.active

    # Shift everything down by 3 rows to insert header + summary
    ws.insert_rows(1, amount=3)

    # Title row (F1:P1)
    title_fill = PatternFill(start_color="8B0000", end_color="8B0000", fill_type="solid")
    title_font = Font(color="FFFFFF", bold=True)
    for col in range(6, 17):
        cell = ws.cell(row=1, column=col)
        cell.fill = title_fill
        cell.font = title_font

    # Summary labels (F2:P2)
    label_fill = PatternFill(start_color="F4CCCC", end_color="F4CCCC", fill_type="solid")
    for i, label in enumerate(["Selection", "Stones", "Carat", "Rap Avg", "PPC Avg", "Avg Disc", "Total Value"]):
        cell = ws.cell(row=2, column=6+i)
        cell.value = label
        cell.fill = label_fill
        cell.font = Font(bold=(i == 0))

    # Summary values (F3:P3)
    for i, key in enumerate(["stones", "carats", "rap_avg", "ppc_avg", "disc_avg", "total_value"]):
        cell = ws.cell(row=3, column=7+i)
        cell.value = round(summary[key], 2) if isinstance(summary[key], float) else summary[key]
        cell.fill = label_fill

    # Apply same formatting to header (row 4)
    for col in range(1, ws.max_column + 1):
        cell = ws.cell(row=4, column=col)
        cell.fill = title_fill
        cell.font = title_font

    wb.save(path)

def filter_inventory(df, filters):
    df = df.copy()
    if "size_min" in filters and "size_max" in filters:
        df = df[(df["Cts"] >= filters["size_min"]) & (df["Cts"] <= filters["size_max"])]
    if "color_min" in filters and "color_max" in filters:
        color_order = ["D","E","F","G","H","I","J","K","L","M"]
        min_idx = color_order.index(filters["color_min"].upper())
        max_idx = color_order.index(filters["color_max"].upper())
        df = df[df["Color"].str.upper().isin(color_order[min_idx:max_idx+1])]
    if "clarity_min" in filters and "clarity_max" in filters:
        clarity_order = ["IF","VVS1","VVS2","VS1","VS2","SI1","SI2","SI3","I1","I2"]
        min_idx = clarity_order.index(filters["clarity_min"].upper())
        max_idx = clarity_order.index(filters["clarity_max"].upper())
        df = df[df["Clarity"].str.upper().isin(clarity_order[min_idx:max_idx+1])]
    if "certified" in filters and filters["certified"]:
        df = df[df["Lab Name"].notna()]
    return df.reset_index(drop=True)

@app.route("/generate", methods=["POST"])
def generate():
    data = request.get_json()
    email = data.get("email")
    filters = data.get("filters", {})

    if not email:
        return jsonify({"error": "Missing email"}), 400

    df = get_latest_inventory_from_drive()
    if df is None:
        return jsonify({"error": "Could not load inventory"}), 500

    filtered_df = filter_inventory(df, filters)
    if filtered_df.empty:
        return jsonify({"error": "No matching stones"}), 404

    summary = {
        "carats": filtered_df["Cts"].sum(),
        "stones": len(filtered_df),
        "disc_avg": filtered_df["Discount"].mean(),
        "ppc_avg": filtered_df["PPC"].mean(),
        "rap_avg": filtered_df["RAP Price"].mean(),
        "total_value": filtered_df["Total Value"].sum()
    }

    filename = f"filtered_{uuid.uuid4().hex[:6]}.xlsx"
    filtered_df.to_excel(filename, index=False)
    beautify_excel(filename, summary)

    # Upload to Drive
    metadata = {"name": filename, "parents": [FOLDER_ID]}
    media = MediaFileUpload(filename, mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
    file = drive_service.files().create(body=metadata, media_body=media, fields='id').execute()
    file_id = file.get("id")
    drive_service.permissions().create(fileId=file_id, body={"role": "reader", "type": "anyone"}).execute()

    link = f"https://drive.google.com/file/d/{file_id}/view?usp=sharing"
    os.remove(filename)

    # Send webhook
    try:
        requests.post(MAKE_WEBHOOK_URL, json={"email": email, "file_link": link, "file_name": filename})
    except:
        pass

    return jsonify({"message": "File filtered", "summary": summary, "link": link})

if __name__ == "__main__":
    app.run(host='0.0.0.0', port=int(os.environ.get('PORT', 10000)))
