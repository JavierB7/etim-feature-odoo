import xlwt
import xlrd
import requests
import sys

if len(sys.argv) < 4:
    print("Please provide file name and the client_id, client_secret for the ETIM API.")
    exit()

DATA_FILE_NAME = sys.argv[1]
ETIM_CLIENT_ID = sys.argv[2]
ETIM_CLIENT_SECRET = sys.argv[3]

# This script is to process a BMECat -> CSV file to get the features decoded with the ETIM API and
# generate another xls file as Odoo attributes.

API_TOKEN_URL = "https://etimauth.etim-international.com/connect/token"
API_URL = "https://etimapi.etim-international.com/api/v2/"
API_SEARCH_FEATURE_URL = f"{API_URL}Feature/Search"
API_VALUE_FEATURE_URL = f"{API_URL}Value/Search"
API_REQUEST_TOKEN_DATA = {
    "grant_type": "client_credentials",
    "scope": "EtimApi",
    "client_id": ETIM_CLIENT_ID,
    "client_secret": ETIM_CLIENT_SECRET
}

HEADERS_ROW = 0
FEATURES_FILE_NAME = "data.xls"
FEATURES_SHEET_NUMBER = 0

workbook_data = xlrd.open_workbook(DATA_FILE_NAME)
old_data_sheet = workbook_data.sheet_by_index(FEATURES_SHEET_NUMBER)

write_workbook = xlwt.Workbook()


def get_feature_and_value_codes(old_sheet):
    features = []
    values = []
    for row in range(old_sheet.nrows):
        if row == HEADERS_ROW:
            continue
        for col in range(old_sheet.ncols):
            if "FNAME" in old_sheet.cell_value(HEADERS_ROW, col):
                feature_code = old_sheet.cell_value(row, col)
                if feature_code:
                    feature_code = feature_code.strip()
                if feature_code != "-" and feature_code != "":
                    features.append(feature_code)
            if "FVALUE" in old_sheet.cell_value(HEADERS_ROW, col):
                value_code = old_sheet.cell_value(row, col)
                if value_code and isinstance(value_code, str) and "EV" in value_code:
                    value_code = value_code.strip()
                if value_code != "-" and value_code != "" and isinstance(value_code, str) and "EV" in value_code:
                    values.append(value_code)
    return list(set(features)), list(set(values))


def get_access_token():
    headers = {'Content-Type': 'application/x-www-form-urlencoded'}
    try:
        res = requests.post(API_TOKEN_URL, data=API_REQUEST_TOKEN_DATA, headers=headers, timeout=5)
        if res.status_code == 200:
            res_data = res.json()
            return res_data["access_token"]
    except requests.exceptions.ConnectionError as e:
        return False
    except requests.exceptions.Timeout as e:
        return False


def map_etim_data(data):
    return {
        entity["code"]: entity["description"] for entity in data
    }


def get_feature_data(token, features):
    data = {
        "from": 0,
        "size": 1000,
        "languagecode": "EN",
        "include": {
            "descriptions": True,
            "translations": True,
            "fields": ["Features"]
        },
        "filters": [
            {
                "code": "Feature",
                "values": features
            }
        ]
    }
    headers = {'Authorization': f"Bearer {token}", 'Content-Type': 'application/json'}
    try:
        res = requests.post(API_SEARCH_FEATURE_URL, json=data, headers=headers, timeout=5)
        if res.status_code == 200:
            res_data = res.json()
            return map_etim_data(res_data["features"])
    except requests.exceptions.ConnectionError as e:
        return False
    except requests.exceptions.Timeout as e:
        return False


def get_value_data(token, values):
    data = {
        "from": 0,
        "size": 1000,
        "languagecode": "EN",
        "include": {
            "descriptions": True,
            "translations": True,
            "fields": ["Features"]
        },
        "filters": [
            {
                "code": "Value",
                "values": values
            }
        ]
    }
    headers = {'Authorization': f"Bearer {token}", 'Content-Type': 'application/json'}
    try:
        res = requests.post(API_VALUE_FEATURE_URL, json=data, headers=headers, timeout=5)
        if res.status_code == 200:
            res_data = res.json()
            return map_etim_data(res_data["values"])
    except requests.exceptions.ConnectionError as e:
        return False
    except requests.exceptions.Timeout as e:
        return False


def get_values_per_feature(o_sheet, features, values_map):
    feature_values = {
        feature: [] for feature in features
    }
    for row in range(o_sheet.nrows):
        if row == HEADERS_ROW:
            continue
        for col in range(o_sheet.ncols):
            if "FNAME" in o_sheet.cell_value(HEADERS_ROW, col):
                feature_code = o_sheet.cell_value(row, col)
                if feature_code:
                    feature_code = feature_code.strip()
                if feature_code != "-" and feature_code != "":
                    value_code = o_sheet.cell_value(row, col + 1)
                    if (
                            value_code is not None and
                            value_code is not False and
                            value_code != ""
                    ):
                        if isinstance(value_code, str) and "EV" in value_code:
                            value_code = values_map[value_code.strip()]
                        feature_values[feature_code].append(value_code)
    return feature_values


def write_header_sheet(sheet, name):
    sheet.write(HEADERS_ROW, 0, "name")
    sheet.write(1, 0, name)
    sheet.write(HEADERS_ROW, 1, "create_variant")
    sheet.write(1, 1, "always")
    sheet.write(HEADERS_ROW, 2, "display_type")
    sheet.write(1, 2, "select")
    sheet.write(HEADERS_ROW, 3, "value_ids / value")


def write_values(o_sheet, features, features_map, values_map):
    feature_values = get_values_per_feature(o_sheet, features, values_map)
    for feature, values in feature_values.items():
        values = list(set(values))
        if len(values) == 0:
            continue
            print("Feature: ", feature, "has no values and will be ignored.")
        feature_sheet = write_workbook.add_sheet(feature)
        write_header_sheet(feature_sheet, features_map[feature_sheet.name])
        value_row = 1
        value_col = 3
        for value in values:
            feature_sheet.write(value_row, value_col, value)
            value_row = value_row + 1


feature_codes, value_codes = get_feature_and_value_codes(old_data_sheet)
access_token = get_access_token()
feature_data = get_feature_data(access_token, feature_codes)
value_data = get_value_data(access_token, value_codes)
write_values(old_data_sheet, feature_codes, feature_data, value_data)
write_workbook.save(FEATURES_FILE_NAME)
