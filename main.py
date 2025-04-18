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
INVENTORY_FILE_ID = '1McHVVICDeeMRiA1fRU7inHmbSUCzeOD2'

# AUTH
SCOPES = ['https://www.googleapis.com/auth/drive', 'https://www.googleapis.com/auth/spreadsheets']
creds = service_account.Credentials.from_service_account_file(
    os.environ["GOOGLE_APPLICATION_CREDENTIALS"], scopes=SCOPES
)

# INIT
app = Flask(__name__)
drive_service = build('drive', 'v3', credentials=creds)
sheets_service = build('sheets', 'v4', credentials=creds)

def get_latest_inventory_from_drive(file_id=INVENTORY_FILE_ID):
    try:
        request = drive_service.files().get_media(fileId=file_id)
        fh = io.BytesIO()
        downloader = MediaIoBaseDownload(fh, request)

        done = False
        while not done:
            status, done = downloader.next_chunk()

        fh.seek(0)
        df = pd.read_excel(fh)
        return df
    except Exception as e:
        print("Error downloading inventory file:", e)
        return None

def filter_inventory(df, filters):
    df = df.copy()
    shape_aliases = {"CU": "Cushion", "CB": "Cushion", "AS": "Asscher", "SQEM": "Asscher", "RD": "Round", "PS": "Pear", "OV": "Oval", "EM": "Emerald", "PR": "Princess"}
    if "shape" in filters:
        shape = filters["shape"].upper()
        match_shape = shape_aliases.get(shape, shape)
        df = df[df["Shape"].str.upper() == match_shape.upper()]
    if "certified" in filters:
        df = df[df["Lab Name"].notna()] if filters["certified"] else df[df["Lab Name"].isna()]
    if "size_min" in filters and "size_max" in filters:
        df = df[(df["Cts"] >= filters["size_min"]) & (df["Cts"] <= filters["size_max"])]
    if "color_min" in filters and "color_max" in filters:
        color_order = ["D", "E", "F", "G", "H", "I", "J", "K", "L", "M"]
        min_idx = color_order.index(filters["color_min"].upper())
        max_idx = color_order.index(filters["color_max"].upper())
        df = df[df["Color"].str.upper().isin(color_order[min_idx:max_idx+1])]
    if "clarity_min" in filters and "clarity_max" in filters:
        clarity_order = ["IF", "VVS1", "VVS2", "VS1", "VS2", "SI1", "SI2", "I1", "I2"]
        min_idx = clarity_order.index(filters["clarity_min"].upper())
        max_idx = clarity_order.index(filters["clarity_max"].upper())
        df = df[df["Clarity"].str.upper().isin(clarity_order[min_idx:max_idx+1])]
    if "fluorescence" in filters:
        df = df[df["Fluo."].str.upper().isin([f.upper() for f in filters["fluorescence"]])]
    if "cut" in filters:
        df = df[df["Cut"].str.upper().isin([c.upper() for c in filters["cut"]])]
    if "polish" in filters:
        df = df[df["Pol"].str.upper().isin([p.upper() for p in filters["polish"]])]
    if "symmetry" in filters:
        df = df[df["Sym"].str.upper().isin([s.upper() for s in filters["symmetry"]])]
    if "td_min" in filters and "td_max" in filters:
        df = df[(df["Total Depth"] >= filters["td_min"]) & (df["Total Depth"] <= filters["td_max"])]
    if "table_min" in filters and "table_max" in filters:
        df = df[(df["Table Size"] >= filters["table_min"]) & (df["Table Size"] <= filters["table_max"])]
    if "pavilion_max" in filters:
        df = df[df["Pavilllion Depth"] <= filters["pavilion_max"]]
    if "crown_max" in filters:
        df = df[df["Crown Height"] <= filters["crown_max"]]
    if "girdle" in filters:
        df = df[df["GirdleThickness Type"].str.upper().isin([g.upper() for g in filters["girdle"]])]
    if "culet" in filters:
        df = df[df["Culet"].str.upper().isin([c.upper() for c in filters["culet"]])]
    return df.reset_index(drop=True)

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

    filtered_df = filter_inventory(df, filters)
    if filtered_df.empty:
        return jsonify({"error": "No matching stones"}), 404

    filename = f"filtered_{uuid.uuid4().hex[:8]}.xlsx"
    filtered_df.to_excel(filename, index=False)

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
    app.run(host="0.0.0.0", port=10000)
