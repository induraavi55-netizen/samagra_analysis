import pandas as pd
import glob
import os

folder_path = r"D:\Automation\diksha2"   # change this

for file_path in glob.glob(os.path.join(folder_path, "*.xlsx")):
    print("Processing:", os.path.basename(file_path))

    try:
        xls = pd.ExcelFile(file_path, engine="openpyxl")
    except Exception:
        print("  ❌ skipped (not a valid excel file):", os.path.basename(file_path))
        continue

    sheet_names = xls.sheet_names
    out = {}

    for sheet in sheet_names:
        df = pd.read_excel(file_path, sheet_name=sheet, engine="openpyxl", header=None)

        # Row 2 (index=1) contains real headers
        row2_headers = df.iloc[1].astype(str).fillna("")

        # get columns where row2 header ends with '_credit'
        keep_cols = [i for i, val in enumerate(row2_headers) if val.endswith("_credit")]

        if keep_cols:
            df_new = df.iloc[:, keep_cols]
        else:
            df_new = pd.DataFrame()  # if nothing matches, blank sheet (better than deleting everything)

        out[sheet] = df_new

    with pd.ExcelWriter(file_path, engine="openpyxl", mode="w") as writer:
        for sheet, d in out.items():
            d.to_excel(writer, sheet_name=sheet, index=False, header=False)

print("done (this time using row-2 to detect *_credit).")
