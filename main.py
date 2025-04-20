from flask import Flask, request, jsonify
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload
import os
import uuid
import requests
import pandas as pd
import io

# CONFIGURATION
FOLDER_ID = '1diAVIuJdsOQhLEQuFzie6QACakeOie25'
MAKE_WEBHOOK_URL = 'https://hook.eu2.make.com/h6jsruunr7u01wobm995dj8wtcmafph8'
INVENTORY_FILE_ID = '1McHVVICDeeMRiA1fRU7inHmbSUCzeOD2'

# AUTH
SCOPES = ['https://www.googleapis.com/auth/drive']
creds = service_account.Credentials.from_service_account_file(
    os.environ["GOOGLE_APPLICATION_CREDENTIALS"], scopes=SCOPES
)

# INIT
app = Flask(__name__)
drive_service = build('drive', 'v3', credentials=creds)

def get_latest_inventory_from_drive():
    try:
        print("Downloading inventory from Drive...")
        request_file = drive_service.files().get_media(fileId=INVENTORY_FILE_ID)
        fh = io.BytesIO()
        downloader = MediaIoBaseDownload(fh, request_file)
        done = False
        while not done:
            status, done = downloader.next_chunk()
        fh.seek(0)
        df = pd.read_excel(fh, engine="openpyxl")
        print("Inventory downloaded and read into DataFrame.")
        return df
    except Exception as e:
        print("Error reading inventory:", e)
        return None

def filter_inventory(df, filters):
    df = df.copy()
    # Example logic to test download success
    if "certified" in filters:
        df = df[df["Lab Name"].notna()] if filters["certified"] else df[df["Lab Name"].isna()]
    return df.reset_index(drop=True)

@app.route('/generate', methods=['POST'])
def generate():
    try:
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

        # Create summary JSON (no Excel output)
        summary = {
            "stones": len(filtered_df),
            "carats": round(filtered_df["Cts"].sum(), 2),
            "rap_avg": round(filtered_df["RAP Price"].mean(), 2),
            "ppc_avg": round(filtered_df["PPC"].mean(), 2),
            "disc_avg": round(filtered_df["Discount"].mean(), 2),
            "total_value": round(filtered_df["Total Value"].sum(), 2)
        }

        return jsonify({"message": "File filtered", "summary": summary})

    except Exception as e:
        print("Exception in /generate:", e)
        return jsonify({"error": "Server error"}), 500

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=int(os.environ.get('PORT', 10000)))
