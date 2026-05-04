import pandas as pd
import numpy as np
from pathlib import Path

# -----------------------------
# CONFIG
# -----------------------------

EXAM_GRADES = [1,2,3,4,5,6, 7, 8, 9, 10]

PARTICIPATING_SCHOOLS = [
    "Diksha Foundation KHEL Bihta",
    "Diksha Foundation Sarai Ranjan",
    "Diksha Foundation KHEL Patna"

]

DATA_DIR = Path("data")
FILE = DATA_DIR / "REG VS PART.xlsx"
SHEET = "Assessment Participation"

# -----------------------------
# LOAD RAW (TWO HEADER ROWS)
# -----------------------------

raw = pd.read_excel(FILE, sheet_name=SHEET, header=None)

print("\n===== DEBUG: RAW SHAPE =====")
print(raw.shape)

# real data starts from row index 2
df = raw.iloc[2:].copy()

# second row contains grade numbers
header_row = raw.iloc[1]

cols = list(header_row)

# fix first columns
cols[0] = "S.No"
cols[1] = "School Name"
cols[2] = "District"

# totals
cols[15] = "Total Registered"
cols[28] = "Total Participated"

# last two
cols[-2] = "Contact Name"
cols[-1] = "Contact Phone"

df.columns = cols

print("\n===== DEBUG: ASSIGNED COLUMNS =====")
for i, c in enumerate(df.columns):
    print(i, c)

# drop total row
df = df[df["School Name"].astype(str).str.lower() != "total"]

print("\n===== DEBUG: DF SHAPE AFTER DROP TOTAL =====")
print(df.shape)

print("\n===== DEBUG: SCHOOL NAMES =====")
print(df["School Name"].unique())

# -----------------------------
# IDENTIFY GRADE COLUMN INDEXES (ROBUST)
# -----------------------------

def safe_grade(col):
    """
    Try converting column header to grade number.
    Return None if not numeric.
    """
    try:
        return int(float(str(col).strip()))
    except:
        return None


# registered block positions
reg_idx_raw = list(range(3, 15))

# participated block positions
part_idx_raw = list(range(16, 28))

# Build safe maps
reg_grade_map = {}
part_grade_map = {}

for i in reg_idx_raw:
    grade = safe_grade(df.columns[i])
    if grade is not None and grade in EXAM_GRADES:
        reg_grade_map[i] = grade

for i in part_idx_raw:
    grade = safe_grade(df.columns[i])
    if grade is not None and grade in EXAM_GRADES:
        part_grade_map[i] = grade

# final cleaned indexes
reg_idx = list(reg_grade_map.keys())
part_idx = list(part_grade_map.keys())

print("\n===== DEBUG: REG IDX =====", reg_idx)
print("===== DEBUG: PART IDX =====", part_idx)

# -----------------------------
# FILTER SCHOOLS FIRST
# -----------------------------

df["School Name"] = df["School Name"].astype(str).str.strip()

school_filter = [s.lower() for s in PARTICIPATING_SCHOOLS]

df_filt = df[
    df["School Name"]
        .str.lower()
        .isin(school_filter)
]

# -----------------------------
# SCHOOL TOTALS (SAFE)
# -----------------------------

school_totals = df_filt[["School Name"]].copy()

school_totals["Registered"] = df_filt.iloc[:, reg_idx].sum(axis=1)
school_totals["Participated"] = df_filt.iloc[:, part_idx].sum(axis=1)

school_totals["Not Participated"] = (
    school_totals["Registered"] - school_totals["Participated"]
)

school_totals = school_totals[
    ["School Name", "Participated", "Not Participated", "Registered"]
]

print("\n===== DEBUG: SCHOOL TOTALS =====")
print(school_totals)

# -----------------------------
# OVERALL GRADE TOTALS
# -----------------------------

overall = pd.DataFrame({
    "Grade": [reg_grade_map[i] for i in reg_idx],
    "Registered": df_filt.iloc[:, reg_idx].sum().values,
    "Participated": df_filt.iloc[:, part_idx].sum().values,
})

overall["Registered"] = pd.to_numeric(overall["Registered"], errors="coerce").fillna(0)
overall["Participated"] = pd.to_numeric(overall["Participated"], errors="coerce").fillna(0)

overall["Participation %"] = (
    overall["Participated"] / overall["Registered"] * 100
).round(0).astype(int)

overall["Participation %"] = overall["Participation %"].astype(int)

print("\n===== DEBUG: OVERALL =====")
print(overall)

# -----------------------------
# SCHOOL × GRADE PIVOT
# -----------------------------

records = []

for idx, row in df_filt.iterrows():

    school = row["School Name"]

    for reg_i in reg_idx:

        grade = reg_grade_map[reg_i]

        # find corresponding participated column index
        part_i = None
        for k, v in part_grade_map.items():
            if v == grade:
                part_i = k
                break

        if part_i is None:
            continue

        registered = pd.to_numeric(row.iloc[reg_i], errors="coerce")
        participated = pd.to_numeric(row.iloc[part_i], errors="coerce")

        registered = 0 if pd.isna(registered) else int(registered)
        participated = 0 if pd.isna(participated) else int(participated)

        not_participated = registered - participated

        participation_pct = (
            round((participated / registered) * 100)
            if registered > 0 else 0
        )

        records.append({
            "School Name": school,
            "Grade": grade,
            "Participated": participated,
            "Not Participated": not_participated,
            "Registered": registered,
            "Participation %": participation_pct
        })

school_grade_df = pd.DataFrame(records)

school_grade_df = school_grade_df.sort_values(
    by=["School Name", "Grade"]
)

print("\n===== DEBUG: SCHOOL GRADE =====")
print(school_grade_df)

# -----------------------------
# WRITE BACK TO SAME FILE
# -----------------------------

with pd.ExcelWriter(FILE, engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
    school_totals.to_excel(writer, sheet_name="schl_wise", index=False)
    overall.to_excel(writer, sheet_name="grade_wise", index=False)
    school_grade_df.to_excel(writer, sheet_name="school_grade_wise", index=False)

print("\n✅ DONE. Sheets written:")
print(" - schl_wise")
print(" - grade_wise")
print(" - school_grade_wise")
