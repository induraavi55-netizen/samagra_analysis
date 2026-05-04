import pandas as pd
import os
import glob
import re

# -------------------------------------------------
# PATH CONFIG
# -------------------------------------------------

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
data_folder = os.path.join(BASE_DIR, "data")
output_file = os.path.join(data_folder, "uploadable_data.xlsx")

grade_files = glob.glob(os.path.join(data_folder, "*.xlsx"))

participation_file = None
for f in grade_files:
    fname = os.path.basename(f).lower()
    if (
        "reg" in fname
        and ("ptn" in fname or "part" in fname)
        and not fname.startswith("~$")
    ):
        participation_file = f
        break

# -------------------------------------------------
# STORAGE CONTAINERS
# -------------------------------------------------

mastery_all = []
school_all = []
qlo_all = []
participation_school_all = []
participation_grade_all = []

# -------------------------------------------------
# HELPER: CLEAN HEADERS
# -------------------------------------------------


def clean_headers(df):
    df.columns = [
        str(col).strip().upper() if pd.notna(col) else "" for col in df.columns
    ]
    return df


# -------------------------------------------------
# LO CLEANER
# -------------------------------------------------


def clean_lo(text):
    text = str(text).strip().lower()
    text = re.sub(r"[\/\-]", " ", text)

    fluff_patterns = [
        r"\bthe student will be able to\b",
        r"\bunderstands and applies\b",
        r"\bunderstands and performs\b",
        r"\bidentify explain\b",
        r"\bgive explain\b",
    ]

    for pattern in fluff_patterns:
        text = re.sub(pattern, "", text)

    text = re.sub(r"[^\w\s]", "", text)
    text = re.sub(r"\s+", " ", text).strip()

    stopwords = {
        "and",
        "or",
        "the",
        "of",
        "to",
        "using",
        "like",
        "with",
        "through",
        "from",
        "in",
        "on",
        "as",
        "is",
        "are",
        "be",
        "able",
        "will",
    }

    words = [w for w in text.split() if w not in stopwords]

    return " ".join(words).title()


# -------------------------------------------------
# PROCESS GRADE FILES
# -------------------------------------------------

for file in grade_files:
    if participation_file and file == participation_file:
        continue

    file_name = os.path.basename(file).replace(".xlsx", "")
    grade_parts = [
        p for p in file_name.replace("-", " ").replace("_", " ").split() if p.isdigit()
    ]

    if not grade_parts:
        continue

    grade = int(grade_parts[0])
    xl = pd.ExcelFile(file)

    for sheet in xl.sheet_names:
        sheet_lower = sheet.lower()

        # ------------------ MASTERY ------------------
        if sheet_lower == "mastery_comparison":
            df = pd.read_excel(file, sheet_name=sheet)
            df = clean_headers(df)

            if "QLVL" not in df.columns:
                continue

            records = []

            for _, row in df.iterrows():
                qlvl = row["QLVL"]

                for col in df.columns:
                    if col.endswith("_COUNT"):
                        subject = col.replace("_COUNT", "")
                        count_val = row[col]

                        perc_col = f"{subject}_PERCENTAGE"

                        percentage_val = row.get(perc_col, 0)

                        records.append(
                            {
                                "GRADE": grade,
                                "QLVL": qlvl,
                                "SUBJECT": subject,
                                "COUNT": count_val,
                                "PERCENTAGE": percentage_val,
                            }
                        )

            df_long = pd.DataFrame(records)

            mastery_all.append(df_long)

        # ------------------ SCHOOL PERFORMANCE ------------------
        elif "school" in sheet_lower and "perf" in sheet_lower:
            df = pd.read_excel(file, sheet_name=sheet)
            df = clean_headers(df)

            if "SUBJECT" not in df.columns:
                continue

            df_long = df.melt(
                id_vars=["SUBJECT"],
                var_name="SCHOOLNAME",
                value_name="PERFORMANCE_PERCENTAGE",
            )

            df_long["GRADE"] = grade
            school_all.append(df_long)

        # ------------------ QLO ------------------
        elif "qlo" in sheet_lower:
            df = pd.read_excel(file, sheet_name=sheet)
            df = clean_headers(df)

            required_cols = {"QLO", "QUESTIONS_LIST", "AVERAGE_PERFORMANCE"}
            if not required_cols.issubset(set(df.columns)):
                continue

            subject_name = sheet.split("_")[0].upper()

            df["GRADE"] = grade
            df["SUBJECT"] = subject_name
            df["QLO_FULL"] = df["QLO"]
            df["QLO"] = df["QLO"].apply(clean_lo)

            qlo_all.append(df)

# -------------------------------------------------
# PROCESS PARTICIPATION FILE
# -------------------------------------------------

if participation_file and os.path.exists(participation_file):
    xl = pd.ExcelFile(participation_file)
    sheet_map = {s.lower(): s for s in xl.sheet_names}

    # ---------- SCHOOL WISE ----------
    for key in sheet_map:
        if "schl" in key:
            df = pd.read_excel(participation_file, sheet_name=sheet_map[key])
            df = clean_headers(df)

            required_cols = {
                "SCHOOL NAME",
                "PARTICIPATED",
                "NOT PARTICIPATED",
                "REGISTERED",
            }

            if required_cols.issubset(set(df.columns)):
                for col in ["PARTICIPATED", "NOT PARTICIPATED", "REGISTERED"]:
                    df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)

                df_long = df.melt(
                    id_vars=["SCHOOL NAME"],
                    value_vars=["PARTICIPATED", "NOT PARTICIPATED", "REGISTERED"],
                    var_name="TYPE",
                    value_name="COUNT",
                )

                participation_school_all.append(df_long)

    # ---------- SCHOOL GRADE WISE (DIRECT COPY) ----------
    target_sheet = None
    for key, original_sheet in sheet_map.items():
        clean_key = re.sub(r"[^a-z]", "", key)
        if "school" in clean_key and "grade" in clean_key:
            target_sheet = original_sheet
            break

    print(f"[DEBUG] Detected sheet name: {target_sheet}")

    if target_sheet:
        df = pd.read_excel(participation_file, sheet_name=target_sheet)

        cleaned_cols = []
        for col in df.columns:
            if pd.notna(col):
                c = str(col).replace("\xa0", " ")
                c = re.sub(r"\s+", " ", c).strip().upper()
                cleaned_cols.append(c)
            else:
                cleaned_cols.append("")

        df.columns = cleaned_cols
        print(f"[DEBUG] Cleaned column names: {cleaned_cols}")

        required_targets = [
            "SCHOOL NAME",
            "GRADE",
            "PARTICIPATED",
            "NOT PARTICIPATED",
            "REGISTERED",
            "PARTICIPATION %",
        ]

        column_mapping = {}
        missing_cols = set(required_targets)

        for col in df.columns:
            col_normalized = col.replace(" ", "")
            for target in required_targets:
                if col_normalized == target.replace(" ", ""):
                    column_mapping[col] = target
                    if target in missing_cols:
                        missing_cols.remove(target)
                    break

        print(f"[DEBUG] Mapped columns: {column_mapping}")

        df.rename(columns=column_mapping, inplace=True)

        if missing_cols:
            print(f"[DEBUG] Missing columns: {missing_cols}")
        else:
            print("[DEBUG] Condition PASSED. Appending to participation_grade_all.")
            participation_grade_all.append(df)


# -------------------------------------------------
# WRITE FINAL FILE
# -------------------------------------------------

with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
    if mastery_all:
        pd.concat(mastery_all, ignore_index=True)[
            ["GRADE", "QLVL", "SUBJECT", "COUNT", "PERCENTAGE"]
        ].to_excel(
            excel_writer=writer, sheet_name="MASTERY_COMPARISON_ALL", index=False
        )

    if school_all:
        pd.concat(school_all, ignore_index=True)[
            ["GRADE", "SUBJECT", "SCHOOLNAME", "PERFORMANCE_PERCENTAGE"]
        ].to_excel(
            excel_writer=writer, sheet_name="SCHOOL_PERFORMANCE_ALL", index=False
        )

    if qlo_all:
        pd.concat(qlo_all, ignore_index=True)[
            ["GRADE", "SUBJECT", "QLO_FULL", "QUESTIONS_LIST", "AVERAGE_PERFORMANCE"]
        ].to_excel(excel_writer=writer, sheet_name="QLO_PERFORMANCE_ALL", index=False)

    if participation_school_all:
        pd.concat(participation_school_all, ignore_index=True)[
            ["SCHOOL NAME", "TYPE", "COUNT"]
        ].to_excel(
            excel_writer=writer, sheet_name="PARTICIPATION_SCHOOL_ALL", index=False
        )

    if participation_grade_all:
        pd.concat(participation_grade_all, ignore_index=True)[
            [
                "SCHOOL NAME",
                "GRADE",
                "PARTICIPATED",
                "NOT PARTICIPATED",
                "REGISTERED",
                "PARTICIPATION %",
            ]
        ].to_excel(
            excel_writer=writer,
            sheet_name="PARTICIPATION_SCHOOL_GRADE_ALL",
            index=False,
        )

print("uploadable_data.xlsx created successfully.")
