from flask import Flask, request, jsonify
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload, MediaFileUpload
import os
import uuid
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

    df_filtered = df.copy()

    # Shape aliases
    shape_aliases = {
        "CU": ["CU", "CB"],
        "CB": ["CU", "CB"],
        "AS": ["AS", "SQEM"],
        "SQEM": ["AS", "SQEM"],
        "RD": ["RD"],
        "EM": ["EM"],
        "PR": ["PR"],
        "OV": ["OV"],
        "PS": ["PS"],
        "RAD": ["RAD"]
    }

    # Color scale
    color_order = ["D", "E", "F", "G", "H", "I", "J", "K", "L", "M"]

    # Clarity scale
    clarity_order = ["IF", "VVS1", "VVS2", "VS1", "VS2", "SI1", "SI2", "SI3", "I1", "I2", "I3"]

    # Fluorescence map
    fluo_map = {
        "NONE": "NON", "NON": "NON",
        "FAINT": "FNT", "FNT": "FNT",
        "MEDIUM": "MED", "MED": "MED",
        "STRONG": "STG", "STG": "STG", "STR": "STG",
        "VERY STRONG": "VST", "VST": "VST"
    }

    # Apply filters
    if filters.get("certified") is True:
        df_filtered = df_filtered[df_filtered["Lab Name"].notna()]

    if "lab" in filters:
        df_filtered = df_filtered[df_filtered["Lab Name"].isin(filters["lab"])]

    if "shape" in filters:
        shape_vals = []
        for shp in filters["shape"]:
            shape_vals.extend(shape_aliases.get(shp.upper(), [shp.upper()]))
        df_filtered = df_filtered[df_filtered["Shape"].isin(shape_vals)]

    if "size_min" in filters:
        df_filtered = df_filtered[df_filtered["Cts"] >= filters["size_min"]]

    if "size_max" in filters:
        df_filtered = df_filtered[df_filtered["Cts"] <= filters["size_max"]]

    if "color_min" in filters and "color_max" in filters:
        min_idx = color_order.index(filters["color_min"].upper())
        max_idx = color_order.index(filters["color_max"].upper())
        allowed = color_order[min_idx:max_idx+1]
        df_filtered = df_filtered[df_filtered["Color"].isin(allowed)]

    if "clarity_min" in filters and "clarity_max" in filters:
        min_idx = clarity_order.index(filters["clarity_min"].upper())
        max_idx = clarity_order.index(filters["clarity_max"].upper())
        allowed = clarity_order[min_idx:max_idx+1]
        df_filtered = df_filtered[df_filtered["Clarity"].isin(allowed)]

    if "cut" in filters:
        df_filtered = df_filtered[df_filtered["Cut"].isin(filters["cut"])]

    if "polish" in filters:
        df_filtered = df_filtered[df_filtered["Pol"].isin(filters["polish"])]

    if "symmetry" in filters:
        df_filtered = df_filtered[df_filtered["Sym"].isin(filters["symmetry"])]

    if "fluorescence" in filters:
        values = [fluo_map.get(f.upper(), f.upper()) for f in filters["fluorescence"]]
        df_filtered = df_filtered[df_filtered["Fluo."].isin(values)]

    if "discount_min" in filters:
        df_filtered = df_filtered[df_filtered["Discount"] >= filters["discount_min"]]

    if "discount_max" in filters:
        df_filtered = df_filtered[df_filtered["Discount"] <= filters["discount_max"]]

    if "ppc_min" in filters:
        df_filtered = df_filtered[df_filtered["PPC"] >= filters["ppc_min"]]

    if "ppc_max" in filters:
        df_filtered = df_filtered[df_filtered["PPC"] <= filters["ppc_max"]]

    if "total_min" in filters:
        df_filtered = df_filtered[df_filtered["Total Value"] >= filters["total_min"]]

    if "total_max" in filters:
        df_filtered = df_filtered[df_filtered["Total Value"] <= filters["total_max"]]

    for col in ["Total Depth", "Table Size", "Crown Angle", "Crown Height", "Pavilion Angle", "Pavilllion Depth"]:
        min_key = col.lower().replace(" ", "_") + "_min"
        max_key = col.lower().replace(" ", "_") + "_max"
        if min_key in filters:
            df_filtered = df_filtered[df_filtered[col] >= filters[min_key]]
        if max_key in filters:
            df_filtered = df_filtered[df_filtered[col] <= filters[max_key]]

    if "girdle_type" in filters:
        df_filtered = df_filtered[df_filtered["GirdleThickness Type"].isin(filters["girdle_type"])]

    if "girdle_percent_min" in filters:
        df_filtered = df_filtered[df_filtered["GirdleThickness Percent"] >= filters["girdle_percent_min"]]

    if "girdle_percent_max" in filters:
        df_filtered = df_filtered[df_filtered["GirdleThickness Percent"] <= filters["girdle_percent_max"]]

    if "culet" in filters:
        df_filtered = df_filtered[df_filtered["Culet"].isin(filters["culet"])]

    if df_filtered.empty:
        return jsonify({"error": "No matching stones"}), 404

    filename = f"filtered_{uuid.uuid4().hex[:8]}.xlsx"
    df_filtered.to_excel(filename, index=False)

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
