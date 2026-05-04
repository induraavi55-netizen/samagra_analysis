import pandas as pd
import os
import glob

# -------------------------------------------------
# PATH SETUP
# -------------------------------------------------

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
folder_path = os.path.join(BASE_DIR, "data")

print("Looking in:", folder_path)

files = glob.glob(os.path.join(folder_path, "*.xlsx"))
print("Files found:", files)

# -------------------------------------------------
# PROCESS FILES
# -------------------------------------------------

for file in files:

    # Skip participation file completely
    if "regn vs ptn" in file.lower():
        continue

    print(f"\nProcessing file: {file}")

    xls = pd.ExcelFile(file)
    all_rows = []

    with pd.ExcelWriter(file, engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:

        for sheet in xls.sheet_names:

            name_lower = sheet.lower()

            # =========================================
            # SUBJECT SHEETS
            # =========================================
            if any(sub in name_lower for sub in ["english", "mathematics", "science"]):

                df = pd.read_excel(file, sheet_name=sheet, header=1)

                # 🔥 FORCE CLEAN HEADERS (case safe + no float crash)
                df.columns = [
                    str(col).strip().upper() if pd.notna(col) else ""
                    for col in df.columns
                ]

                # Identify Q columns safely
                score_cols = [
                    c for c in df.columns
                    if isinstance(c, str) and c.startswith("Q") and c[1:].isdigit()
                ]

                if not score_cols:
                    continue

                # Separate metadata + student rows
                meta_rows = df.iloc[:3].copy()
                student_rows = df.iloc[3:].copy()

                # Convert scores safely
                student_rows[score_cols] = (
                    student_rows[score_cols]
                    .apply(pd.to_numeric, errors="coerce")
                    .fillna(0)
                )

                student_rows["TOTAL_CREDIT"] = student_rows[score_cols].sum(axis=1)

                student_rows["PERFORMANCE_PERCENTAGE"] = (
                    (student_rows["TOTAL_CREDIT"] / len(score_cols)) * 100
                ).round(0).astype(int)

                df_updated = pd.concat([meta_rows, student_rows], ignore_index=True)
                df_updated.to_excel(
                    excel_writer=writer,
                    sheet_name=sheet,
                    index=False
                )

                # --------------------------------------
                # SCHOOL PIVOT DATA COLLECTION
                # --------------------------------------

                school_col = None
                for col in df.columns:
                    if col.replace(" ", "") == "SCHOOLNAME":
                        school_col = col
                        break

                if school_col:
                    student_rows["SUBJECT"] = sheet.strip()

                    all_rows.append(
                        student_rows[
                            ["SUBJECT", school_col, "PERFORMANCE_PERCENTAGE"]
                        ].rename(columns={school_col: "SCHOOLNAME"})
                    )

            # =========================================
            # LF SHEETS
            # =========================================
            elif name_lower.endswith("_lf"):

                df = pd.read_excel(file, sheet_name=sheet)

                df.columns = [
                    str(col).strip().upper() if pd.notna(col) else ""
                    for col in df.columns
                ]

                student_cols = [
                    col for col in df.columns
                    if isinstance(col, str) and col.startswith("STD")
                ]

                if student_cols and "QLO" in df.columns:

                    df[student_cols] = (
                        df[student_cols]
                        .apply(pd.to_numeric, errors="coerce")
                        .fillna(0)
                    )

                    df["TOTAL_CREDIT"] = df[student_cols].sum(axis=1)

                    df["PERFORMANCE_PERCENTAGE"] = (
                        df[student_cols].mean(axis=1) * 100
                    ).round(0).astype(int)

                    df.to_excel(
                        excel_writer=writer,
                        sheet_name=sheet,
                        index=False
                    )

                    # QLO pivot
                    if "Q" in df.columns:
                        df["QUESTION_NUMBER"] = df["Q"].astype(str)

                        pivot_qlo = df.groupby("QLO").agg(
                            Questions_List=("QUESTION_NUMBER", lambda x: ", ".join(x)),
                            Average_Performance=("PERFORMANCE_PERCENTAGE", "mean")
                        ).reset_index()

                        pivot_qlo["Average_Performance"] = (
                            pivot_qlo["Average_Performance"]
                            .round(0)
                            .astype(int)
                        )

                        base_name = sheet.lower().replace("_lf", "")
                        new_sheet_name = f"{base_name}_QLO-wise-perf"

                        pivot_qlo.to_excel(
                            excel_writer=writer,
                            sheet_name=new_sheet_name[:31],
                            index=False
                        )

        # =========================================
        # SCHOOL-WISE PIVOT
        # =========================================

        if all_rows:

            pivot_source = pd.concat(all_rows, ignore_index=True)

            required_cols = {"SUBJECT", "SCHOOLNAME", "PERFORMANCE_PERCENTAGE"}

            if required_cols.issubset(set(pivot_source.columns)):

                school_subject_pivot = (
                    pivot_source
                    .pivot_table(
                        index="SUBJECT",
                        columns="SCHOOLNAME",
                        values="PERFORMANCE_PERCENTAGE",
                        aggfunc="mean"
                    )
                    .round(0)
                    .fillna(0)
                    .astype(int)
                )

                school_subject_pivot.to_excel(
                    excel_writer=writer,
                    sheet_name="School-wise-perf-across-grades",
                    index=True
                )

print("All files processed successfully.")