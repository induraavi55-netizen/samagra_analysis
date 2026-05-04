import os
import glob
import pandas as pd

# -------------------------------------------------
# PATH SETUP
# -------------------------------------------------

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
folder_path = os.path.join(BASE_DIR, "data")

print("Looking in:", folder_path)

files = glob.glob(os.path.join(folder_path, "*.xlsx"))
print("Files found:", files)

# -------------------------------------------------
# Helper: Clean Headers
# -------------------------------------------------

def clean_headers(df):
    df.columns = [
        str(col).strip().upper() if pd.notna(col) else ""
        for col in df.columns
    ]
    return df


# -------------------------------------------------
# PROCESS FILES
# -------------------------------------------------

for file in files:

    # 🔥 Skip participation file
    if "regn vs ptn" in file.lower():
        continue

    print(f"\nProcessing file: {file}")

    xls = pd.ExcelFile(file)

    with pd.ExcelWriter(file, engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:

        for sheet in xls.sheet_names:

            print(f"  Sheet: {sheet}")

            try:
                df = pd.read_excel(file, sheet_name=sheet, header=1)
            except Exception as e:
                print(f"    Could not read sheet. Skipping. Error: {e}")
                continue

            df = clean_headers(df)

            print("    Columns:", df.columns.tolist())

            # ---------------------------------------
            # Detect Question Columns Safely
            # ---------------------------------------

            q_cols = [
                col for col in df.columns
                if col.startswith("Q") and col[1:].isdigit()
            ]

            if not q_cols:
                print("    No Q columns found. Skipping.")
                continue

            # Sort Q columns numerically
            q_cols = sorted(q_cols, key=lambda x: int(x[1:]))

            # Select only Q columns
            q_section = df[q_cols]

            # ---------------------------------------
            # Transpose
            # ---------------------------------------

            transposed = q_section.T.reset_index()
            transposed.rename(columns={"index": "Q"}, inplace=True)

            total_cols = transposed.shape[1]

            # Need at least Q + QID + QLVL + QLO
            if total_cols < 4:
                print("    Not enough metadata rows. Skipping.")
                continue

            # Build new columns dynamically
            base_meta_cols = ["Q", "QID", "QLVL", "QLO"]

            num_students = total_cols - 4

            std_cols = [f"STD{i}" for i in range(1, num_students + 1)]

            transposed.columns = base_meta_cols + std_cols

            # ---------------------------------------
            # SHEET NAMING (Case Safe)
            # ---------------------------------------

            name_lower = sheet.lower()

            if "english" in name_lower:
                new_sheet = "eng_lf"
            elif "math" in name_lower:
                new_sheet = "math_lf"
            elif "science" in name_lower or "sci" in name_lower:
                new_sheet = "sci_lf"
            else:
                safe_name = (
                    sheet.replace(" ", "_")
                         .replace("-", "")
                         .lower()
                )
                new_sheet = f"{safe_name}_lf"

            print(f"    Writing to sheet: {new_sheet}")

            # Excel sheet name max length = 31
            transposed.to_excel(
                excel_writer=writer,
                sheet_name=new_sheet[:31],
                index=False
            )

print("\nAll Grade files processed successfully.")