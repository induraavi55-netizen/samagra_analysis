import pandas as pd

file_path = "Raw data.xlsx"

xls = pd.ExcelFile(file_path)

with pd.ExcelWriter(file_path, engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:

    for sheet in xls.sheet_names:
        # Process only formatted sheets
        if not sheet.lower().endswith("_formatted"):
            continue

        df = pd.read_excel(file_path, sheet_name=sheet)

        # Identify credit columns dynamically
        credit_cols = [c for c in df.columns if c.lower().endswith("_credit")]

        if not credit_cols:
            continue  # safety net

        # Group and sum credits
        pivot = (
            df.groupby(
                ["SchoolName", "school id", "Grade", "Subject"],
                as_index=False
            )[credit_cols]
            .sum()
        )

        # Output sheet name
        output_sheet = sheet.replace("_formatted", "_credit_pivot")

        # Write pivot sheet
        pivot.to_excel(writer, sheet_name=output_sheet, index=False)
