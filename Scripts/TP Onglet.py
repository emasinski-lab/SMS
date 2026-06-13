from openpyxl import Workbook

def create_tp_sheets(tp_df, output_file):
    wb = Workbook()
    wb.remove(wb.active)

    sheet_map = {}

    for _, row in tp_df.iterrows():
        sheet_name = f"{row['Identité']} - {row['MSISDN']}"

        # Excel limite 31 caractères
        sheet_name = sheet_name[:31]

        ws = wb.create_sheet(title=sheet_name)
        ws.append(["ActionDate", "MSISDN APT", "MSISDN APE", "Preview", "Traduction"])

        sheet_map[row["MSISDN"]] = ws

    wb.save(output_file)
    return sheet_map