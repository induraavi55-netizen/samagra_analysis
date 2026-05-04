import pandas as pd
import glob
import os

folder_path = r"D:\Automation\diksha2"   # update this

for file_path in glob.glob(os.path.join(folder_path, "*.xlsx")):
    print("Processing:", os.path.basename(file_path))

    try:
        xls = pd.ExcelFile(file_path, engine="openpyxl")
    except:
        print("  ❌ skipped, not a valid Excel file")
        continue

    sheet_names = xls.sheet_names

    with pd.ExcelWriter(file_path, engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
        for sheet in sheet_names:
            df = pd.read_excel(file_path, sheet_name=sheet, engine="openpyxl")

            # convert every credit cell to numeric, non-numeric becomes 0
            df = df.apply(pd.to_numeric, errors="coerce").fillna(0)

            # calculate total credit
            df["total_credit"] = df.sum(axis=1)

            # percentage
            df["percentage"] = (df["total_credit"] / 16) * 100

            df.to_excel(writer, sheet_name=sheet, index=False)

print("⚡ All files updated with total_credit and percentage.")
