from flask import Flask, request, jsonify
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload
import os
import io
import pandas as pd

# CONFIGURATION
INVENTORY_FILE_ID = '1McHVVICDeeMRiA1fRU7inHmbSUCzeOD2'
SCOPES = ['https://www.googleapis.com/auth/drive']

# AUTHENTICATION
creds = service_account.Credentials.from_service_account_file(
    os.environ["GOOGLE_APPLICATION_CREDENTIALS"], scopes=SCOPES
)
drive_service = build('drive', 'v3', credentials=creds)

# INIT
app = Flask(__name__)

def get_inventory_df():
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

def apply_filters(df, filters):
    df = df.copy()

    # SHAPE
    shape_map = {
        "CU": "Cushion", "CB": "Cushion", "AS": "Asscher", "SQEM": "Asscher",
        "RD": "Round", "PS": "Pear", "OV": "Oval", "EM": "Emerald", "PR": "Princess"
    }
    if "shape" in filters:
        target = shape_map.get(filters["shape"].upper(), filters["shape"])
        df = df[df["Shape"].fillna('').str.upper() == target.upper()]

    # CERTIFIED
    if "certified" in filters:
        df = df[df["Lab Name"].notna()] if filters["certified"] else df[df["Lab Name"].isna()]

    # SIZE
    if "size_min" in filters and "size_max" in filters:
        df = df[(df["Cts"] >= filters["size_min"]) & (df["Cts"] <= filters["size_max"])]

    # COLOR
    if "color_min" in filters and "color_max" in filters:
        color_order = ["D", "E", "F", "G", "H", "I", "J", "K", "L", "M"]
        cmin = color_order.index(filters["color_min"].upper())
        cmax = color_order.index(filters["color_max"].upper())
        df = df[df["Color"].fillna('').str.upper().isin(color_order[cmin:cmax+1])]

    # CLARITY
    if "clarity_min" in filters and "clarity_max" in filters:
        clarity_order = ["IF", "VVS1", "VVS2", "VS1", "VS2", "SI1", "SI2", "SI3", "I1", "I2"]
        clmin = clarity_order.index(filters["clarity_min"].upper())
        clmax = clarity_order.index(filters["clarity_max"].upper())
        df = df[df["Clarity"].fillna('').str.upper().isin(clarity_order[clmin:clmax+1])]

    # CUT / POL / SYM
    for key, col in [("cut", "Cut"), ("polish", "Pol"), ("symmetry", "Sym")]:
        if key in filters:
            df = df[df[col].fillna('').str.upper().isin([v.upper() for v in filters[key]])]

    # FLUORESCENCE
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

    # DISCOUNT
    if "discount_max" in filters:
        df = df[df["Discount"] <= filters["discount_max"]]
    if "discount_min" in filters:
        df = df[df["Discount"] >= filters["discount_min"]]

    # PPC
    if "ppc_max" in filters:
        df = df[df["PPC"] <= filters["ppc_max"]]
    if "ppc_min" in filters:
        df = df[df["PPC"] >= filters["ppc_min"]]

    # TD & TABLE
    if "td_min" in filters and "td_max" in filters:
        df = df[(df["Total Depth"] >= filters["td_min"]) & (df["Total Depth"] <= filters["td_max"])]
    if "table_min" in filters and "table_max" in filters:
        df = df[(df["Table Size"] >= filters["table_min"]) & (df["Table Size"] <= filters["table_max"])]

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

@app.route('/generate', methods=['POST'])
def generate():
    data = request.get_json()
    filters = data.get("filters", {})
    df = get_inventory_df()
    if df is None:
        return jsonify({"error": "Could not load inventory"}), 500

    filtered_df = apply_filters(df, filters)
    if filtered_df.empty:
        return jsonify({"error": "No matching stones"}), 404

    return jsonify({"message": "File filtered", "summary": summarize(filtered_df)})

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=int(os.environ.get("PORT", 10000)))
