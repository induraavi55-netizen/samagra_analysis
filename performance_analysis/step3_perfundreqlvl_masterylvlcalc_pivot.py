import pandas as pd
import os
import glob

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
# Process Files
# -------------------------------------------------

for file in files:

    # 🔥 Skip participation file
    if "regn vs ptn" in file.lower():
        continue

    print(f"\nProcessing: {file}")

    file_name = os.path.basename(file).replace(".xlsx", "")
    grade_parts = [
        p for p in file_name.replace("-", " ").replace("_", " ").split()
        if p.isdigit()
    ]

    if not grade_parts:
        print("  Could not detect grade. Skipping.")
        continue

    current_grade = int(grade_parts[0])
    xl = pd.ExcelFile(file)

    # =========================================
    # STEP 1: Create _OP sheets from _LF sheets
    # =========================================
    with pd.ExcelWriter(file, engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:

        for sheet in xl.sheet_names:

            if sheet.lower().endswith("_lf"):

                df = pd.read_excel(file, sheet_name=sheet)
                df = clean_headers(df)

                student_cols = [
                    c for c in df.columns
                    if isinstance(c, str) and c.startswith("STD")
                ]

                if not student_cols or "QLVL" not in df.columns:
                    continue

                # Handle Pre-primary variations
                df["QLVL"] = (
                    df["QLVL"]
                    .astype(str)
                    .str.strip()
                    .str.replace("PRE PRIMARY", "0", case=False)
                    .str.replace("PRE-PRIMARY", "0", case=False)
                )

                df["QLVL"] = pd.to_numeric(df["QLVL"], errors="coerce")

                df[student_cols] = (
                    df[student_cols]
                    .apply(pd.to_numeric, errors="coerce")
                    .fillna(0)
                )

                result = df.groupby("QLVL")[student_cols].mean() * 100
                result = result.round(0).fillna(0).astype(int)

                result = result.T.reset_index().rename(columns={"index": "STUDENT"})

                output_sheet = sheet.upper().replace("_LF", "_OP")

                result.to_excel(
                    excel_writer=writer,
                    sheet_name=output_sheet[:31],
                    index=False
                )

    # Reload updated workbook
    xl = pd.ExcelFile(file)

    # =========================================
    # STEP 2: Mastery Level Calculation
    # =========================================
    with pd.ExcelWriter(file, engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:

        for sheet in xl.sheet_names:

            if sheet.lower().endswith("_op"):

                df = pd.read_excel(file, sheet_name=sheet)
                df = clean_headers(df)

                # Normalize possible column variants
                if "PRE-PRIMARY" in df.columns:
                    df.rename(columns={"PRE-PRIMARY": "0"}, inplace=True)

                # Identify numeric level columns safely
                level_cols = [
                    col for col in df.columns
                    if col.isdigit()
                ]

                level_cols = sorted([int(c) for c in level_cols])

                if not level_cols:
                    continue

                mastery_levels = []
                mastery_gaps = []

                for _, row in df.iterrows():

                    mastery = 0

                    for level in level_cols:

                        if level > current_grade:
                            continue

                        val = pd.to_numeric(row[str(level)], errors="coerce")

                        if pd.notna(val) and val >= 60:
                            mastery = max(1, level)
                        else:
                            break

                    mastery_levels.append(mastery)
                    mastery_gaps.append(current_grade - mastery)

                df["MASTERY_LEVEL"] = mastery_levels
                df["MASTERY_GAP"] = mastery_gaps

                ordered_cols = (
                    ["STUDENT"]
                    + [str(l) for l in level_cols]
                    + ["MASTERY_LEVEL", "MASTERY_GAP"]
                )

                ordered_cols = [c for c in ordered_cols if c in df.columns]

                df = df[ordered_cols]

                df.to_excel(
                    excel_writer=writer,
                    sheet_name=sheet[:31],
                    index=False
                )

    # Reload again
    xl = pd.ExcelFile(file)

    # =========================================
    # STEP 3: Cross-Subject Mastery Comparison
    # =========================================
    with pd.ExcelWriter(file, engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:

        mastery_data = []

        for sheet in xl.sheet_names:

            if sheet.lower().endswith("_op"):

                df = pd.read_excel(file, sheet_name=sheet)
                df = clean_headers(df)

                if "MASTERY_LEVEL" not in df.columns:
                    continue

                total_students = len(df)
                if total_students == 0:
                    continue

                subject_name = sheet.upper().replace("_OP", "")

                mastery_counts = df["MASTERY_LEVEL"].value_counts()

                for level, count in mastery_counts.items():
                    mastery_data.append({
                        "QLVL": level,
                        "SUBJECT": subject_name,
                        "COUNT": int(count),
                        "PERCENTAGE": round((count / total_students) * 100)
                    })

        if mastery_data:

            mastery_df = pd.DataFrame(mastery_data)

            pivot_df = mastery_df.pivot_table(
                index="QLVL",
                columns="SUBJECT",
                values=["COUNT", "PERCENTAGE"],
                fill_value=0
            )

            # Flatten columns properly
            pivot_df.columns = [
                f"{subject}_{metric}"
                for metric, subject in pivot_df.columns
            ]

            pivot_df = pivot_df.reset_index()

            pivot_df.to_excel(
                excel_writer=writer,
                sheet_name="MASTERY_COMPARISON",
                index=False
            )

print("All files processed successfully.")