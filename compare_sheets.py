import pandas as pd
from tkinter import Tk, filedialog
import os


def pick_file(title):
    return filedialog.askopenfilename(
        title=title,
        filetypes=[("Excel files", "*.xlsx *.xlsm")]
    )


def clean_df(df):
    df = df.dropna(how="all")
    df = df.dropna(axis=1, how="all")
    df.columns = df.columns.astype(str).str.strip()

    for col in df.columns:
        df[col] = df[col].astype(str).str.strip()

    return df


def show_columns(label, df):
    print(f"\nColumns in {label}:")
    for col in df.columns:
        print(f"- {col}")


def get_staff_column(label, df):
    show_columns(label, df)

    while True:
        column = input(f"\nEnter the staff column name in {label}: ").strip()
        if column in df.columns:
            return column

        print(f"Column '{column}' was not found. Please type the name exactly as shown above.")


def clean_staff_list(df, staff_column):
    staff = df[[staff_column]].copy()
    staff.columns = ["STAFF"]
    staff["STAFF"] = staff["STAFF"].astype(str).str.strip()
    staff = staff[staff["STAFF"].ne("")]
    staff = staff[staff["STAFF"].str.lower().ne("nan")]
    return staff.drop_duplicates().sort_values("STAFF").reset_index(drop=True)


Tk().withdraw()

file1 = pick_file("Select FIRST Excel file")
file2 = pick_file("Select SECOND Excel file")

sheet1 = input("Enter sheet name in FIRST file: ")
sheet2 = input("Enter sheet name in SECOND file: ")

df1 = pd.read_excel(file1, sheet_name=sheet1)
df2 = pd.read_excel(file2, sheet_name=sheet2)

df1 = clean_df(df1)
df2 = clean_df(df2)

staff_col1 = get_staff_column("FIRST file", df1)
staff_col2 = get_staff_column("SECOND file", df2)

staff1 = clean_staff_list(df1, staff_col1)
staff2 = clean_staff_list(df2, staff_col2)

merged = staff1.merge(
    staff2,
    on="STAFF",
    how="outer",
    indicator=True
)

only_file1 = merged[merged["_merge"] == "left_only"].drop(columns=["_merge"])
only_file2 = merged[merged["_merge"] == "right_only"].drop(columns=["_merge"])
matched = merged[merged["_merge"] == "both"].drop(columns=["_merge"])
summary = pd.DataFrame({
    "CHECK": [
        "Same staff in both files?",
        "Staff count in first file",
        "Staff count in second file",
        "Matched staff count",
        "Only in first file count",
        "Only in second file count",
    ],
    "RESULT": [
        "YES" if only_file1.empty and only_file2.empty else "NO",
        len(staff1),
        len(staff2),
        len(matched),
        len(only_file1),
        len(only_file2),
    ]
})

output = os.path.join(
    os.path.dirname(file1),
    "Staff_Comparison_Report.xlsx"
)

with pd.ExcelWriter(output, engine="openpyxl") as writer:
    summary.to_excel(writer, sheet_name="SUMMARY", index=False)
    matched.to_excel(writer, sheet_name="MATCHED", index=False)
    only_file1.to_excel(writer, sheet_name="ONLY_IN_FILE_1", index=False)
    only_file2.to_excel(writer, sheet_name="ONLY_IN_FILE_2", index=False)

print("DONE")
print(summary.to_string(index=False))
print(output)
