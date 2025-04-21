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

# CONFIG
FOLDER_ID = '1diAVIuJdsOQhLEQuFzie6QACakeOie25'
MAKE_WEBHOOK_URL = 'https://hook.eu2.make.com/h6jsruunr7u01wobm995dj8wtcmafph8'
INVENTORY_FILE_ID = '1McHVVICDeeMRiA1fRU7inHmbSUCzeOD2'
SCOPES = ['https://www.googleapis.com/auth/drive']

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
        print("Error reading inventory:", e)
        return None


def filter_inventory(df, filters):
    df = df.copy()

    shape_map = {
        "CU": "Cushion", "CB": "Cushion", "AS": "Asscher", "SQEM": "Asscher",
        "RD": "Round", "PS": "Pear", "OV": "Oval", "EM": "Emerald", "PR": "Princess"
    }
    if "shape" in filters:
        target = shape_map.get(filters["shape"].upper(), filters["shape"])
        df = df[df["Shape"].fillna('').str.upper() == target.upper()]

    if filters.get("certified"):
        df = df[df["Lab Name"].notna()]

    if "size_min" in filters and "size_max" in filters:
        df = df[(df["Cts"] >= filters["size_min"]) & (df["Cts"] <= filters["size_max"])]

    if "color_min" in filters and "color_max" in filters:
        color_order = ["D", "E", "F", "G", "H", "I", "J", "K", "L", "M"]
        cmin = color_order.index(filters["color_min"].upper())
        cmax = color_order.index(filters["color_max"].upper())
        df = df[df["Color"].fillna('').str.upper().isin(color_order[cmin:cmax+1])]

    if "clarity_min" in filters and "clarity_max" in filters:
        clarity_order = ["IF", "VVS1", "VVS2", "VS1", "VS2", "SI1", "SI2", "SI3", "I1", "I2"]
        clmin = clarity_order.index(filters["clarity_min"].upper())
        clmax = clarity_order.index(filters["clarity_max"].upper())
        df = df[df["Clarity"].fillna('').str.upper().isin(clarity_order[clmin:clmax+1])]

    for key, col in [("cut", "Cut"), ("polish", "Pol"), ("symmetry", "Sym")]:
        if key in filters:
            df = df[df[col].fillna('').str.upper().isin([v.upper() for v in filters[key]])]

    if "fluorescence" in filters:
        fl_map = {
            "NON": ["NON", "NONE"],
            "FNT": ["FNT", "FAINT"],
            "MED": ["MED", "MEDIUM"],
            "STR": ["STR", "STG", "STRONG"],
            "VST": ["VST", "VERY STRONG"]
        }
        resolved = []
        for f in filters["fluorescence"]:
            resolved.extend(fl_map.get(f.upper(), [f]))
        df = df[df["Fluo."].fillna('').str.upper().isin([v.upper() for v in resolved])]

    if "discount_max" in filters:
        df = df[df["Discount"] <= filters["discount_max"]]
    if "discount_min" in filters:
        df = df[df["Discount"] >= filters["discount_min"]]

    if "ppc_max" in filters:
        df = df[df["PPC"] <= filters["ppc_max"]]
    if "ppc_min" in filters:
        df = df[df["PPC"] >= filters["ppc_min"]]

    return df.reset_index(drop=True)


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
    ws.insert_rows(1, amount=3)

    red_fill = PatternFill(start_color="8B0000", end_color="8B0000", fill_type="solid")
    pink_fill = PatternFill(start_color="F4CCCC", end_color="F4CCCC", fill_type="solid")
    white_bold = Font(color="FFFFFF", bold=True)
    black_font = Font(color="000000")

    summary_labels = ["", "Stones", "Carats", "", "Rap Avg", "PPC Avg", "Avg Disc", "", "", "", "Total Value"]
    summary_values = ["Selection", summary["stones"], summary["carats"], "", summary["rap_avg"], summary["ppc_avg"], summary["disc_avg"], "", "", "", summary["total_value"]]

    for i, val in enumerate(summary_labels):
        cell = ws.cell(row=1, column=6+i)
        cell.value = val if val else None
        cell.fill = red_fill
        cell.font = white_bold

    for i, val in enumerate(summary_values):
        cell = ws.cell(row=2, column=6+i)
        cell.value = val if val != "" else None
        cell.fill = pink_fill
        cell.font = Font(bold=(i == 0))

    for col in range(1, ws.max_column + 1):
        cell = ws.cell(row=3, column=col)
        cell.fill = red_fill
        cell.font = white_bold

    wb.save(path)


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


if __name__ == '__main__':
    app.run(host='0.0.0.0', port=int(os.environ.get("PORT", 10000)))
