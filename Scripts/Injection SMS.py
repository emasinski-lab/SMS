import pandas as pd
from openpyxl import load_workbook

def route_sms(concat_file, tp_df, output_file):
    df = pd.read_excel(concat_file)

    wb = load_workbook(output_file)

    # mapping msisdn -> sheet name
    msisdn_to_sheet = {}
    for _, row in tp_df.iterrows():
        sheet_name = f"{row['Identité']}-{row['MSISDN']}"[:31]
        msisdn_to_sheet[row["MSISDN"]] = sheet_name

    for _, r in df.iterrows():
        apt = str(r["MSISDN APT"])
        ape = str(r["MSISDN APE"])

        row_data = [
            r["ActionDate"],
            r["MSISDN APT"],
            r["MSISDN APE"],
            r["Preview"],
            r["Traduction"]
        ]

        # écrire dans APT
        if apt in msisdn_to_sheet:
            ws = wb[msisdn_to_sheet[apt]]
            ws.append(row_data)

        # écrire dans APE (si différent)
        if ape in msisdn_to_sheet and ape != apt:
            ws = wb[msisdn_to_sheet[ape]]
            ws.append(row_data)

    wb.save(output_file)