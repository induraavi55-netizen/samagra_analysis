import pandas as pd

input_file = "Raw data.xlsx"

xls = pd.ExcelFile(input_file)

with pd.ExcelWriter(input_file, engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:

    for sheet in xls.sheet_names:
        df = pd.read_excel(input_file, sheet_name=sheet)

        # Clean column names (important for pattern matching)
        df.columns = df.columns.str.strip()

        # Identify credit columns
        credit_cols = [c for c in df.columns if c.lower().endswith("_credit")]

        # Required base columns
        base_cols = [
            "SchoolName",
            "Student LoginId",
            "Grade",
            "Subject"
        ]

        # Keep only existing required columns
        keep_cols = [c for c in base_cols if c in df.columns] + credit_cols

        df = df[keep_cols].copy()

        # Create new school id column from Student LoginId (first 6 digits)
        if "Student LoginId" in df.columns:
            df["school id"] = df["Student LoginId"].astype(str).str[:6]

        # Reorder columns (school id next to student id)
        final_cols = []
        for col in ["SchoolName", "Student LoginId", "school id", "Grade", "Subject"]:
            if col in df.columns:
                final_cols.append(col)

        final_cols += credit_cols

        df = df[final_cols]

        # Write formatted sheet
        df.to_excel(writer, sheet_name=f"{sheet}_formatted", index=False)
