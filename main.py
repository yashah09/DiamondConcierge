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

def filter_inventory(df, filters):
    df = df.copy()

    shape_aliases = {
        "CU": "Cushion",
        "CB": "Cushion",
        "AS": "Asscher",
        "SQEM": "Asscher",
        "RD": "Round",
        "PS": "Pear",
        "OV": "Oval",
        "EM": "Emerald",
        "PR": "Princess"
    }

    if "shape" in filters:
        shape = filters["shape"].upper()
        match_shape = shape_aliases.get(shape, shape)
        df = df[df["Shape"].str.upper() == match_shape.upper()]

    if "certified" in filters:
        if filters["certified"]:
            df = df[df["Lab Name"].notna()]
        else:
            df = df[df["Lab Name"].isna()]

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
        df = df[df["Polish"].str.upper().isin([p.upper() for p in filters["polish"]])]

    if "symmetry" in filters:
        df = df[df["Symm"].str.upper().isin([s.upper() for s in filters["symmetry"]])]

    if "td_min" in filters and "td_max" in filters and "TD" in df:
        df = df[(df["TD"] >= filters["td_min"]) & (df["TD"] <= filters["td_max"])]

    if "table_min" in filters and "table_max" in filters and "Table" in df:
        df = df[(df["Table"] >= filters["table_min"]) & (df["Table"] <= filters["table_max"])]

    if "pavilion_max" in filters and "Pavilion" in df:
        df = df[df["Pavilion"] <= filters["pavilion_max"]]

    if "crown_max" in filters and "Crown" in df:
        df = df[df["Crown"] <= filters["crown_max"]]

    if "girdle" in filters and "Girdle" in df:
        df = df[df["Girdle"].str.upper().isin([g.upper() for g in filters["girdle"]])]

    if "culet" in filters and "Culet" in df:
        df = df[df["Culet"].str.upper().isin([c.upper() for c in filters["culet"]])]

    return df.reset_index(drop=True)
