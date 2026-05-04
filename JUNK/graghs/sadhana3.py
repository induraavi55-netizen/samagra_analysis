import pandas as pd
import glob
import os

folder_path = r"D:\Automation\diksha2"
output_file = os.path.join(folder_path, "pivots_performance.xlsx")

records = []

files = glob.glob(os.path.join(folder_path, "*.xlsx"))

print("\nFOUND FILES:")
for f in files:
    print("  ->", os.path.basename(f))
print("\nPROCESSING...\n")

for file_path in files:
    # skip the output file if it already exists
    if os.path.basename(file_path).lower() == "pivots_performance.xlsx":
        print("skipped pivot file:", file_path)
        continue

    grade = os.path.splitext(os.path.basename(file_path))[0]

    try:
        xls = pd.ExcelFile(file_path, engine="openpyxl")
    except Exception as e:
        print("❌ broken file skipped:", os.path.basename(file_path), "|", e)
        continue

    print("✔", grade)

    for sheet in xls.sheet_names:
        try:
            df = pd.read_excel(file_path, sheet_name=sheet, engine="openpyxl")
        except:
            print(f"   ❌ cannot read sheet: {sheet}")
            continue

        if "percentage" not in df.columns:
            print(f"   ⚠ sheet {sheet} skipped (no percentage column)")
            continue

        pct = pd.to_numeric(df["percentage"], errors="coerce")
        avg_pct = pct.mean()

        print(f"   → {sheet}: {avg_pct}")

        records.append({
            "grade": grade,
            "subject": sheet,
            "avg_percentage": avg_pct
        })

print("\nCREATING PIVOT FILE...\n")

summary = pd.DataFrame(records)

pivot = summary.pivot_table(
    index="grade",
    columns="subject",
    values="avg_percentage",
    aggfunc="mean"
)

with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
    pivot.to_excel(writer, sheet_name="pivot", index=True)

print("🔥 DONE →", output_file)
