import os, re
from datetime import datetime
from tkinter import Tk, filedialog
from openpyxl import load_workbook, Workbook

YEAR = "2026"
MONTHS = {1:"JAN",2:"FEB",3:"MAR",4:"APR",5:"MAY",6:"JUN",7:"JUL",8:"AUG",9:"SEP",10:"OCT",11:"NOV",12:"DEC"}

def clean(v):
    return "" if v is None else str(v).strip()

def row_text(row):
    return " ".join(clean(v) for v in row if clean(v)).upper()

def sheet_date(name):
    name = str(name).strip()
    if re.fullmatch(r"\d{8}", name) and name.endswith(YEAR):
        return datetime.strptime(name, "%d%m%Y")
    return None

def currency_from_title(text):
    t = text.upper()
    m = re.search(r"\(([^)]+)\)", t)
    if m:
        cur = m.group(1).strip().upper()
    elif "USD" in t:
        cur = "USD"
    elif "ZIG" in t:
        cur = "ZIG"
    elif "ZWG" in t:
        cur = "ZWG"
    elif "ZWL" in t:
        cur = "ZWL"
    else:
        cur = ""
    return "ZWL" if cur == "ZIG" else cur

def count_nonblank(row):
    return sum(1 for v in row if clean(v))

def is_section_row(row, rt):
    return count_nonblank(row) <= 6 and "TOTAL" not in rt and "TRANSACTION ID" not in rt and "BOOK" not in rt

def find_exact(row, names):
    names = {x.upper() for x in names}
    for i, v in enumerate(row):
        if clean(v).upper() in names:
            return i
    return None

def find_prefix(row, prefixes):
    prefixes = [p.upper() for p in prefixes]
    for i, v in enumerate(row):
        x = clean(v).upper()
        for p in prefixes:
            if x.startswith(p):
                return i
    return None

def headers_from(row, start):
    h = [clean(v) for v in row[start:]]
    while h and h[-1] == "":
        h.pop()
    return h if h else ["Transaction ID"]

def get_sheet(out_wb, cache, name, headers):
    if name not in cache:
        ws = out_wb.create_sheet(name)
        ws.append(headers + ["DATE", "CURRENCY"])
        cache[name] = {"ws": ws, "headers": headers}
    return cache[name]["ws"], cache[name]["headers"]

def append_data(ws, row, start, headers, date_value, currency):
    data = [clean(v) for v in row[start:start+len(headers)]]
    while len(data) < len(headers):
        data.append("")
    data.append(date_value)
    data.append(currency)
    ws.append(data)

def main():
    Tk().withdraw()
    file_path = filedialog.askopenfilename(
        title="Select workbook",
        filetypes=[("Excel files", "*.xlsx *.xlsm")]
    )
    if not file_path:
        print("No file selected.")
        return

    wb = load_workbook(file_path, read_only=True, data_only=True)
    out = Workbook()
    out.remove(out.active)

    cache = {}
    mm_count = 0
    sec_count = 0

    for sheet_name in wb.sheetnames:
        dt = sheet_date(sheet_name)
        if not dt:
            continue

        ws = wb[sheet_name]
        month = MONTHS[dt.month]
        date_value = dt.strftime("%Y-%m-%d")

        section = None
        currency = ""

        mm_headers = None
        mm_start = None
        sec_headers = None
        sec_start = None

        print("Processing:", sheet_name, date_value)

        for row in ws.iter_rows(values_only=True):
            row = list(row)
            rt = row_text(row)

            if not rt:
                continue

            if "TOTAL" in rt:
                continue

            if "LIABILIT" in rt:
                section = None
                currency = ""
                continue

            if is_section_row(row, rt) and "HTM" in rt and "ASSET" in rt:
                section = "HTM"
                currency = currency_from_title(rt)
                continue

            if is_section_row(row, rt) and "HTM" not in rt and ("ASSET" in rt or "PLACEMENT" in rt or "MONEY MARKET" in rt):
                section = "MM"
                currency = currency_from_title(rt)
                continue

            if section == "MM":
                h = find_exact(row, {"TRANSACTION ID", "TRANSACTIONID"})
                if h is not None:
                    mm_start = h
                    mm_headers = headers_from(row, mm_start)
                    continue

                tx = find_prefix(row, ["MM"])
                if tx is not None:
                    if mm_headers is None:
                        mm_start = tx
                        mm_headers = [f"Column_{i+1}" for i in range(len(row[mm_start:]))]

                    out_ws, headers = get_sheet(out, cache, f"MM_{month}", mm_headers)
                    append_data(out_ws, row, mm_start, headers, date_value, currency)
                    mm_count += 1

            elif section == "HTM":
                h = find_exact(row, {"BOOK"})
                if h is not None:
                    sec_start = h
                    sec_headers = headers_from(row, sec_start)
                    continue

                tx = find_prefix(row, ["9999"])
                if tx is not None:
                    if sec_headers is None:
                        sec_start = tx
                        sec_headers = [f"Column_{i+1}" for i in range(len(row[sec_start:]))]

                    out_ws, headers = get_sheet(out, cache, f"SECURITIES_{month}", sec_headers)
                    append_data(out_ws, row, sec_start, headers, date_value, currency)
                    sec_count += 1

    output = os.path.join(os.path.dirname(file_path), "Combined_Assets_By_Month_2026.xlsx")
    out.save(output)

    print("DONE")
    print("MM rows:", mm_count)
    print("SECURITIES rows:", sec_count)
    print(output)

if __name__ == "__main__":
    main()