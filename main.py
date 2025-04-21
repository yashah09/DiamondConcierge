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
from openpyxl.styles import PatternFill, Font, Border, Side
from openpyxl.utils import get_column_letter

# CONFIGURATION
FOLDER_ID = '17gPLesM8wfEbJCJsNbHXGsCRh9qw9ezx'
MAKE_WEBHOOK_URL = 'https://hook.us2.make.com/8p08qs7c5af35g9t1iy6nuvbplhmyqpq'
INVENTORY_FILE_ID = '1ZvrrL85WuKh37KJ-DT3drf31cQOS38Va'
SCOPES = ['https://www.googleapis.com/auth/drive']

# AUTH
creds = service_account.Credentials.from_service_account_file(
    os.environ["GOOGLE_APPLICATION_CREDENTIALS"], scopes=SCOPES
)
drive_service = build('drive', 'v3', credentials=creds)

# INIT
app = Flask(__name__)

@app.before_request
def catch_all_errors():
    print(f"🛬 Incoming request: {request.method} {request.path}")

def get_latest_inventory_from_drive():
    try:
        request_file = drive_service.files().get_media(fileId=INVENTORY_FILE_ID)
        fh = io.BytesIO()
        downloader = MediaIoBaseDownload(fh, request_file)
        done = False
        while not done:
            _, done = downloader.next_chunk()
        fh.seek(0)
        df = pd.read_excel(fh, engine="openpyxl")
        return df
    except Exception as e:
        print("Error loading inventory:", e)
        return None

def beautify_excel(path, summary):
    wb = load_workbook(path)
    ws = wb.active

    ws.insert_rows(1, amount=3)

    dark_red_fill = PatternFill(start_color="8B0000", end_color="8B0000", fill_type="solid")
    white_bold_font = Font(color="FFFFFF", bold=True, name="Aptos Narrow", size=11)
    pink_fill = PatternFill(start_color="F4CCCC", end_color="F4CCCC", fill_type="solid")
    black_font = Font(color="000000", name="Aptos Narrow", size=11)
    black_bold_font = Font(color="000000", bold=True, name="Aptos Narrow", size=11)
    border_style = Border(
        left=Side(border_style="thin", color="000000"),
        right=Side(border_style="thin", color="000000"),
        top=Side(border_style="thin", color="000000"),
        bottom=Side(border_style="thin", color="000000")
    )

    headers = ["", "Stones", "Carats", "", "Rap Avg", "PPC Avg", "Avg Disc", "", "", "", "Total Value"]
    for col, header in enumerate(headers, start=6):
        cell = ws.cell(row=1, column=col)
        cell.value = header if header else None
        cell.fill = dark_red_fill
        cell.font = white_bold_font
        cell.border = border_style

    ws["F2"] = "Selection"
    ws["G2"] = "=SUBTOTAL(3,B4:B4000)"
    ws["H2"] = "=SUBTOTAL(9,E4:E4000)"
    ws["J2"] = "=SUBTOTAL(9,M4:M4000)/H2"
    ws["K2"] = "=SUBTOTAL(9,O4:O4000)/H2"
    ws["L2"] = "=((K2/J2)-1)*100"
    ws["P2"] = "=IF(G2<200,SUBTOTAL(9,P4:P4000),0)"

    for col in range(6, 17):
        cell = ws.cell(row=2, column=col)
        cell.fill = pink_fill
        cell.font = black_bold_font if col == 6 else black_font
        cell.border = border_style

    for col in range(1, ws.max_column + 1):
        cell = ws.cell(row=3, column=col)
        cell.fill = dark_red_fill
        cell.font = white_bold_font
        cell.border = border_style

    for row in range(4, 4001):
        ws[f"E{row}"].number_format = "0.00"
        ws[f"N{row}"].number_format = "0.00"
        ws[f"O{row}"].number_format = "0"
        ws[f"P{row}"].number_format = "0"
        for col in range(1, ws.max_column + 1):
            ws.cell(row=row, column=col).border = border_style

    ws["H2"].number_format = "0.00"
    ws["J2"].number_format = "0"
    ws["K2"].number_format = "0"
    ws["L2"].number_format = "0.00"
    ws["P2"].number_format = "0"

    wb.save(path)

def filter_inventory(df, filters):
    df = df.copy()
    shape_map = {
        "CU": "Cushion", "CB": "Cushion", "AS": "Asscher", "SQEM": "Asscher",
        "RD": "Round", "PS": "Pear", "OV": "Oval", "EM": "Emerald", "PR": "Princess"
    }
    if "shape" in filters:
        target = shape_map.get(filters["shape"].upper(), filters["shape"])
        df = df[df["Shape"].fillna('').str.upper() == target.upper()]

    if "certified" in filters:
        df = df[df["Lab Name"].notna()] if filters["certified"] else df[df["Lab Name"].isna()]

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
        flour_map = {
            "NON": ["NONE", "NON"],
            "FNT": ["FAINT", "FNT"],
            "MED": ["MEDIUM", "MED"],
            "STR": ["STRONG", "STR", "STG"],
            "VST": ["VERY STRONG", "VST"]
        }
        valid_vals = []
        for code in filters["fluorescence"]:
            valid_vals.extend(flour_map.get(code.upper(), [code]))
        df = df[df["Fluo."].fillna('').str.upper().isin([v.upper() for v in valid_vals])]

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

@app.route("/generate", methods=["POST"])
def generate():
    try:
        print("🔥 /generate endpoint HIT")
        print("🧪 Raw payload:", request.data)

        data = request.get_json()
        if data is None:
            print("❌ No JSON received. Did you forget the Content-Type header?")
            return jsonify({"error": "No JSON received"}), 400

        print("📥 Incoming JSON:", data)

        email = data.get("email")
        filters = data.get("filters", {})

        if not email:
            print("❌ Email missing in request")
            return jsonify({"error": "Missing email"}), 400

        df = get_latest_inventory_from_drive()
        if df is None:
            print("❌ Could not load inventory")
            return jsonify({"error": "Could not load inventory"}), 500

        filtered_df = filter_inventory(df, filters)
        if filtered_df.empty:
            print("❌ No matching stones")
            return jsonify({"error": "No matching stones"}), 404

        filename = f"filtered_{uuid.uuid4().hex[:6]}.xlsx"
        filtered_df.to_excel(filename, index=False)
        beautify_excel(filename, summarize(filtered_df))

        metadata = {"name": filename, "parents": [FOLDER_ID]}
        media = MediaFileUpload(filename, mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
        file = drive_service.files().create(body=metadata, media_body=media, fields='id').execute()
        file_id = file.get("id")
        drive_service.permissions().create(fileId=file_id, body={"role": "reader", "type": "anyone"}).execute()

        link = f"https://drive.google.com/file/d/{file_id}/view?usp=sharing"
        os.remove(filename)

        print(f"✅ File uploaded: {link}")

        try:
            requests.post(MAKE_WEBHOOK_URL, json={"email": email, "file_link": link, "file_name": filename})
            print("📤 Webhook sent")
        except Exception as e:
            print("⚠️ Webhook error:", e)

        return jsonify({"message": "File filtered", "summary": summarize(filtered_df), "link": link})

    except Exception as e:
        print("❌ UNHANDLED ERROR in /generate:", str(e))
        return jsonify({"error": "Internal Server Error", "details": str(e)}), 500

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=int(os.environ.get("PORT", 10000)))
