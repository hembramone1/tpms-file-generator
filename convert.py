#!/usr/bin/env python3
"""
TPMS File Generator - Python CLI Converter
Converts BEML (.xls/.xlsx) and Caterpillar (CAT MineStar .csv/.xlsx) TPMS payload reports
into standardized CSV files.
"""

import os
import sys
import csv
import argparse
from datetime import datetime, time
import python_calamine

# 18-Column Headers for BEML Dumpers
BEML_HEADERS = [
    'TRUCK NO',
    'T_Date',
    'Time',
    'PayLoad',
    'Loading Time',
    'Hauling Time',
    'Stop Dump Time',
    'Return Time',
    'Return Stop Time',
    'Total Time',
    'Haul Travel Distance',
    'Return Travel Distance',
    'Total Distance',
    'MaxHaul Speed',
    'AvgHaul Speed',
    'MaxReturn Speed',
    'AvgReturn Speed',
    'Uploaded By'
]
CSV_HEADERS = BEML_HEADERS

# 16-Column Headers for CAT MineStar Dumpers
CAT_HEADERS = [
    'TIMEDATE',
    'PAYLOAD SMH',
    'SERIAL NUMBER',
    'PAYLOADTONS',
    'TRAVEL EMPTY TIME HHMMSS',
    'TRAVEL EMPTY DISTANCEMI',
    'STOPPED EMPTY TIME HHMMSS',
    'LOAD TIME HHMMSS',
    'STOPPED LOADED TIME HHMMSS',
    'LOADED TRAVEL TIME HHMMSS',
    'LOADED TRAVEL DISTANCEMI',
    'CYCLE TIME HHMMSS',
    'CYCLE DISTANCEMI',
    'LOADER PASSES',
    'FUEL USED GAL',
    'UPLOADED BY'
]

def format_time_val(val):
    if isinstance(val, (datetime, time)):
        return f"{val.hour}:{val.minute:02d}"
    return str(val) if val is not None else ""

def format_float_val(val):
    try:
        return f"{float(val):.2f}"
    except (ValueError, TypeError):
        return str(val) if val is not None else ""

def transform_cat_time(val):
    """Unpad hour in CAT time strings: 00:29:55 -> 0:29:55"""
    if not val:
        return ""
    parts = str(val).strip().split(':')
    if len(parts) == 3:
        try:
            h = str(int(parts[0]))
            return f"{h}:{parts[1]}:{parts[2]}"
        except ValueError:
            pass
    return str(val).strip()

def process_beml_file(file_path, month, output_dir):
    try:
        wb = python_calamine.CalamineWorkbook.from_path(file_path)
    except Exception as e:
        print(f"[ERROR] Could not open {file_path}: {e}")
        return False, 0, "", "", []

    # Locate target sheet (prefer Sheet2, or check sheet names)
    target_sheet = None
    for name in wb.sheet_names:
        if name.lower() == 'sheet2':
            target_sheet = name
            break
    if not target_sheet:
        target_sheet = wb.sheet_names[0]

    sheet = wb.get_sheet_by_name(target_sheet)
    data = sheet.to_python()

    # 1. Extract Serial Number
    serial_no = None
    for r in range(min(10, len(data))):
        row = data[r]
        for c in range(len(row)):
            if str(row[c]).strip().lower() == 'serial number':
                val = row[c + 2] if c + 2 < len(row) and row[c + 2] != '' else (row[c + 1] if c + 1 < len(row) else '')
                if val != '':
                    serial_no = str(int(val)) if isinstance(val, (int, float)) else str(val).strip()
                    break
        if serial_no:
            break

    # Fallback to filename prefix
    base_name = os.path.basename(file_path)
    if not serial_no:
        serial_no = base_name.split('_')[0].strip()

    # 2. Locate start row (first row where col 0 is '1' or 1.0)
    start_row = -1
    for r in range(min(25, len(data))):
        if len(data[r]) > 0 and str(data[r][0]).strip() in ('1', '1.0', 1, 1.0):
            start_row = r
            break

    if start_row == -1:
        print(f"[WARN] No cycle data found in {base_name}")
        return False, 0, "", "", []

    # 3. Transform rows
    csv_rows = []
    for r in range(start_row, len(data)):
        row = data[r]
        if not row or row[0] is None or str(row[0]).strip() == "":
            break
        try:
            int(float(row[0]))
        except (ValueError, TypeError):
            break

        # Date formatting DD/MM/YYYY
        d = row[1]
        if isinstance(d, datetime):
            c_date = d.strftime('%d/%m/%Y')
        elif ' ' in str(d):
            c_date = str(d).split()[0]
        else:
            c_date = str(d)

        # Time formatting HH:MM:SS
        t = row[2]
        c_time = t.strftime('%H:%M:%S') if isinstance(t, (datetime, time)) else str(t)

        # PayLoad
        c_payload = format_float_val(row[3])

        # Times
        c_load_time = format_time_val(row[4])
        c_haul_time = format_time_val(row[5])
        c_dump_time = format_time_val(row[6])
        c_ret_time = format_time_val(row[7])
        c_ret_stop = format_time_val(row[8])
        c_tot_time = format_time_val(row[9])

        # Distances & Speeds
        c_haul_dist = format_float_val(row[10])
        c_ret_dist = format_float_val(row[11])
        c_tot_dist = format_float_val(row[12])
        c_max_haul = format_float_val(row[13])
        c_avg_haul = format_float_val(row[14])
        c_max_ret = format_float_val(row[15])
        c_avg_ret = format_float_val(row[16])
        c_uploaded_by = ""

        csv_rows.append([
            serial_no,
            c_date,
            c_time,
            c_payload,
            c_load_time,
            c_haul_time,
            c_dump_time,
            c_ret_time,
            c_ret_stop,
            c_tot_time,
            c_haul_dist,
            c_ret_dist,
            c_tot_dist,
            c_max_haul,
            c_avg_haul,
            c_max_ret,
            c_avg_ret,
            c_uploaded_by
        ])

    # 4. Write output CSV
    output_filename = f"{serial_no}_{month.lower()}.csv"
    os.makedirs(output_dir, exist_ok=True)
    out_path = os.path.join(output_dir, output_filename)

    with open(out_path, 'w', encoding='utf-8', newline='') as f:
        f.write(','.join(BEML_HEADERS) + '\n')
        for r in csv_rows:
            f.write(','.join(r) + '\n')

    return True, len(csv_rows), out_path, serial_no, csv_rows, "BEML"

def process_cat_file(file_path, output_dir):
    try:
        with open(file_path, 'r', encoding='utf-8', errors='replace') as f:
            reader = list(csv.reader(f))
    except Exception as e:
        print(f"[ERROR] Could not read CAT CSV {file_path}: {e}")
        return False, 0, "", "", [], "CAT"

    base_name = os.path.basename(file_path)

    # 1. Extract Serial Number from header or filename
    serial_no = None
    for row in reader[:8]:
        row_str = ' '.join(row)
        if 'equipment :' in row_str.lower():
            parts = row_str.split(':')
            if len(parts) > 1:
                serial_no = parts[1].strip()
                break

    if not serial_no:
        # Fallback to filename: TruckPayload_PRB00914_...
        parts = base_name.split('_')
        if len(parts) > 1 and parts[0].lower() == 'truckpayload':
            serial_no = parts[1].strip()
        else:
            serial_no = base_name.split('.')[0].strip()

    # 2. Locate data start row (starts with 'Date and Time')
    start_row = -1
    for i, row in enumerate(reader):
        if row and row[0].strip().lower() == 'date and time':
            start_row = i + 1
            break

    if start_row == -1:
        print(f"[WARN] No CAT cycle data found in {base_name}")
        return False, 0, "", "", [], "CAT"

    # 3. Transform 16 CAT columns matching Test 2 output
    csv_rows = []
    for row in reader[start_row:]:
        if not row or len(row) < 16 or not row[0].strip():
            continue

        formatted = [
            row[0].strip(),
            row[1].strip(),
            serial_no,
            row[2].strip(),
            transform_cat_time(row[3]),
            row[4].strip(),
            transform_cat_time(row[5]),
            transform_cat_time(row[6]),
            transform_cat_time(row[7]),
            transform_cat_time(row[8]),
            row[9].strip(),
            transform_cat_time(row[10]),
            row[11].strip(),
            row[12].strip(),
            row[15].strip(),
            "" # Uploaded By
        ]
        csv_rows.append(formatted)

    # Output filename is strictly <serial>.csv matching benchmark
    out_filename = f"{serial_no}.csv"
    os.makedirs(output_dir, exist_ok=True)
    out_path = os.path.join(output_dir, out_filename)

    with open(out_path, 'w', encoding='utf-8', newline='') as f:
        f.write(','.join(CAT_HEADERS) + '\n')
        for r in csv_rows:
            f.write(','.join(r) + '\n')

    return True, len(csv_rows), out_path, serial_no, csv_rows, "CAT"

def process_file(file_path, month, output_dir, dumper_type="auto"):
    base_name = os.path.basename(file_path).lower()

    # Explicit format selection
    if dumper_type == "cat":
        return process_cat_file(file_path, output_dir)
    elif dumper_type == "beml":
        return process_beml_file(file_path, month, output_dir)

    # Auto-detection
    if base_name.endswith('.csv') or base_name.startswith('truckpayload_'):
        return process_cat_file(file_path, output_dir)

    if base_name.endswith(('.xls', '.xlsx')):
        return process_beml_file(file_path, month, output_dir)

    return process_cat_file(file_path, output_dir)

def main():
    parser = argparse.ArgumentParser(description="Convert TPMS .xls/.xlsx/.csv files to standardized CSV")
    parser.add_argument("input", help="Single TPMS file or folder containing TPMS files")
    parser.add_argument("-t", "--type", choices=["beml", "cat", "auto"], default="auto",
                        help="Dumper TPMS type: beml, cat, or auto (default: auto)")
    parser.add_argument("-m", "--month", default="september", help="Month name for BEML output filename (default: september)")
    parser.add_argument("-o", "--output", default=".", help="Output directory for CSV files (default: current directory)")
    parser.add_argument("-c", "--consolidate", dest="consolidate", action="store_true", default=None,
                        help="Generate a consolidated CSV containing all dumper data (default: true if multiple files)")
    parser.add_argument("--no-consolidate", dest="consolidate", action="store_false",
                        help="Do not generate a consolidated CSV")
    args = parser.parse_args()

    input_path = os.path.abspath(args.input)
    output_dir = os.path.abspath(args.output)
    month = args.month.strip()
    dumper_type = args.type.lower()

    files_to_process = []
    if os.path.isfile(input_path):
        files_to_process.append(input_path)
    elif os.path.isdir(input_path):
        for root, _, files in os.walk(input_path):
            for file in files:
                ext = file.lower()
                if ext.endswith(('.xls', '.xlsx', '.csv')):
                    files_to_process.append(os.path.join(root, file))
    else:
        print(f"[ERROR] Input path does not exist: {input_path}")
        sys.exit(1)

    if not files_to_process:
        print(f"[WARN] No .xls, .xlsx, or .csv files found in: {input_path}")
        sys.exit(0)

    print(f"\n========================================================")
    print(f" TPMS File Generator (CLI)")
    print(f" Input:  {input_path}")
    print(f" Output: {output_dir}")
    print(f" Type:   {dumper_type.upper()}")
    print(f" Month:  {month}")
    print(f" Files:  {len(files_to_process)}")
    print(f"========================================================\n")

    total_converted = 0
    total_cycles = 0
    all_serials = []
    all_rows = []
    types_seen = set()

    for idx, fpath in enumerate(files_to_process, 1):
        fname = os.path.basename(fpath)
        success, count, out_path, sn, rows, dtype = process_file(fpath, month, output_dir, dumper_type)
        if success:
            total_converted += 1
            total_cycles += count
            types_seen.add(dtype)
            if sn and sn not in all_serials:
                all_serials.append(sn)
            all_rows.extend(rows)
            print(f"[{idx}/{len(files_to_process)}] SUCCESS [{dtype}]: {fname} -> {os.path.basename(out_path)} ({count} cycles)")
        else:
            print(f"[{idx}/{len(files_to_process)}] FAILED:  {fname}")

    # Generate consolidated CSV if requested or if multiple files were converted
    should_consolidate = args.consolidate if args.consolidate is not None else (total_converted > 1)
    if should_consolidate and all_rows:
        is_all_cat = (types_seen == {"CAT"})
        headers = CAT_HEADERS if is_all_cat else BEML_HEADERS
        cap_month = month.capitalize()
        cons_filename = f"{cap_month}_{'_'.join(all_serials)}.csv" if all_serials else f"{cap_month}_consolidated.csv"
        cons_path = os.path.join(output_dir, cons_filename)
        with open(cons_path, 'w', encoding='utf-8', newline='') as f:
            f.write(','.join(headers) + '\n')
            for r in all_rows:
                f.write(','.join(r) + '\n')
        print(f"\n[CONSOLIDATED] Generated consolidated CSV:")
        print(f"  -> {os.path.basename(cons_path)} ({len(all_rows)} total cycles across {len(all_serials)} dumpers)")

    print(f"\nCompleted! Converted {total_converted}/{len(files_to_process)} files ({total_cycles} total cycles).")

if __name__ == '__main__':
    main()
