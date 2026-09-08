#!/usr/bin/env python3
"""
TPMS File Generator - Synthetic Report Generator (CLI)
Generates realistic BEML / Dumper TPMS payload reports (.xlsx) and corresponding (.csv)
based on Date, Shift, Dumper Serial Number, and optional trip count.
"""

import os
import sys
import argparse
import random
from datetime import datetime, timedelta
import pandas as pd
import openpyxl

CSV_HEADERS = [
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

XLS_HEADERS = [
    'Sl. No.',
    'Date',
    'Time',
    'PayLoad',
    'Loading Time',
    'Hauling Time',
    'Stop & Dump Time',
    'Return Time',
    'Return Stop Time',
    'Total Time',
    'Haul Travel Distance',
    'Return Travel Distance',
    'Total Distance',
    'Max.Haul Speed',
    'Avg.Haul Speed',
    'Max.Return Speed',
    'Avg.Return Speed'
]

XLS_UNITS = [
    '', '', '', 'Tons', 'mm:ss', 'mm:ss', 'mm:ss', 'mm:ss', 'mm:ss', 'mm:ss',
    'km', 'km', 'km', 'kmph', 'kmph', 'kmph', 'kmph'
]

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

def fmt_m_ss(seconds):
    m = seconds // 60
    s = seconds % 60
    return f"{m}:{s:02d}"

def fmt_000_ss(seconds):
    m = seconds // 60
    s = seconds % 60
    return f"{m:03d}:{s:02d}"

def fmt_hh_mm_ss(seconds):
    """Padded HH:MM:SS for CAT MineStar raw file"""
    h = seconds // 3600
    m = (seconds % 3600) // 60
    s = seconds % 60
    return f"{h:02d}:{m:02d}:{s:02d}"

def fmt_h_mm_ss(seconds):
    """Unpadded H:MM:SS for CAT standardized 16-column CSV"""
    h = seconds // 3600
    m = (seconds % 3600) // 60
    s = seconds % 60
    return f"{h}:{m:02d}:{s:02d}"

def generate_shift_records(serial_no, date_obj, shift_type, num_trips=None):
    """
    Generate realistic cycles for an 8-hour shift.
    shift_type: '1' (06:00-14:00), '2' (14:00-22:00), '3' (22:00-06:00)
    """
    # Shift definitions:
    # 1st Shift: 06:00 to 14:00 -> Production: 07:00 to 13:00
    # 2nd Shift: 14:00 to 22:00 -> Production: 15:00 to 21:00
    # 3rd Shift: 22:00 to 06:00 -> Production: 23:00 to 05:00 (next day)
    if str(shift_type) in ('1', '1st'):
        shift_label = '06:00 to 14:00'
        prod_start = datetime(date_obj.year, date_obj.month, date_obj.day, 7, random.randint(0, 3), random.randint(0, 59))
        prod_end = datetime(date_obj.year, date_obj.month, date_obj.day, 13, 0, 0)
    elif str(shift_type) in ('2', '2nd'):
        shift_label = '14:00 to 22:00'
        prod_start = datetime(date_obj.year, date_obj.month, date_obj.day, 15, random.randint(0, 3), random.randint(0, 59))
        prod_end = datetime(date_obj.year, date_obj.month, date_obj.day, 21, 0, 0)
    else:
        shift_label = '22:00 to 06:00'
        prod_start = datetime(date_obj.year, date_obj.month, date_obj.day, 23, random.randint(0, 3), random.randint(0, 59))
        next_day = date_obj + timedelta(days=1)
        prod_end = datetime(next_day.year, next_day.month, next_day.day, 5, 0, 0)

    if num_trips is None or num_trips <= 0:
        num_trips = random.randint(11, 14)

    finish_target = prod_end - timedelta(minutes=random.randint(10, 25))
    meal_sec = random.randint(25, 35) * 60

    curr_dt = prod_start
    records = []

    for i in range(1, num_trips + 1):
        # Payload: Gaussian around 53.0 tons, clipped between 43 and 63 tons
        raw_payload = random.gauss(53.0, 4.0)
        payload = round(max(42.0, min(64.0, raw_payload)) * 2) / 2 # Round to nearest 0.5

        # Distances
        haul_dist = round(random.uniform(1.60, 1.85), 2)
        ret_dist = round(haul_dist + random.uniform(-0.04, 0.04), 2)
        tot_dist = round(haul_dist + ret_dist, 2)

        # Speeds
        avg_haul_spd = round(random.uniform(14.2, 18.2), 2)
        max_haul_spd = round(avg_haul_spd * random.uniform(1.5, 1.8), 2)

        avg_ret_spd = round(random.uniform(15.2, 19.8), 2)
        max_ret_spd = round(avg_ret_spd * random.uniform(1.6, 1.9), 2)

        # Durations in seconds
        load_sec = random.randint(85, 210)
        haul_sec = int(round((haul_dist / avg_haul_spd) * 3600))
        dump_sec = random.randint(15, 38)
        ret_sec = int(round((ret_dist / avg_ret_spd) * 3600))
        ret_stop_sec = random.choice([0, 0, 0, random.randint(20, 100)])

        tot_sec = load_sec + haul_sec + dump_sec + ret_sec + ret_stop_sec

        # Check if cycle exceeds production end deadline (1 hour before shift end)
        cycle_finish = curr_dt + timedelta(seconds=tot_sec)
        if cycle_finish > prod_end:
            break

        c_date = curr_dt.strftime('%d/%m/%Y')
        c_time = curr_dt.strftime('%H:%M:%S')

        rec = {
            'sl_no': i,
            'serial_no': serial_no,
            'date': c_date,
            'time': c_time,
            'payload': f"{payload:.2f}",
            'load_time': fmt_m_ss(load_sec),
            'haul_time': fmt_m_ss(haul_sec),
            'dump_time': fmt_000_ss(dump_sec),
            'ret_time': fmt_m_ss(ret_sec),
            'ret_stop_time': fmt_000_ss(ret_stop_sec),
            'tot_time': fmt_000_ss(tot_sec),
            'haul_dist': f"{haul_dist:.2f}",
            'ret_dist': f"{ret_dist:.2f}",
            'tot_dist': f"{tot_dist:.2f}",
            'max_haul_spd': f"{max_haul_spd:.2f}",
            'avg_haul_spd': f"{avg_haul_spd:.2f}",
            'max_ret_spd': f"{max_ret_spd:.2f}",
            'avg_ret_spd': f"{avg_ret_spd:.2f}"
        }
        records.append(rec)

        # Advance time to next cycle:
        # In real mining operations, cycles run back-to-back (0 sec gap).
        # If trips >= 7, one mid-shift break (35-50 min) occurs at midpoint.
        if i == num_trips // 2 and num_trips >= 7:
            curr_dt = cycle_finish + timedelta(seconds=meal_sec)
        else:
            # 85% exactly 0 gap, 15% tiny 5-15s shovel queue
            tiny_gap = 0 if random.random() < 0.85 else random.randint(5, 15)
            curr_dt = cycle_finish + timedelta(seconds=tiny_gap)

    return records, shift_label

def build_excel_workbook(records, serial_no, date_obj, shift_label):
    """
    Build an openpyxl Workbook with Sheet2 matching BEML report specifications.
    """
    wb = openpyxl.Workbook()
    # Sheet2 as primary sheet
    ws2 = wb.active
    ws2.title = 'Sheet2'

    # Rows 1-6 Metadata
    ws2.append([]) # Row 1 (1-indexed) empty
    ws2.append(['BHARAT EARTH MOVERS LTD, INDIA']) # Row 2
    ws2.append(['Payload Report']) # Row 3
    
    # Row 4: Report Date & Serial Number
    r4 = [''] * 17
    r4[1] = 'Report Date'
    r4[2] = date_obj.strftime('%B %d,%Y')
    r4[12] = 'Serial Number'
    r4[14] = int(serial_no) if serial_no.isdigit() else serial_no
    ws2.append(r4)

    # Row 5: Period & Model Number
    r5 = [''] * 17
    r5[1] = 'Period'
    r5[2] = f"{date_obj.strftime('%d/%m/%Y')} to {date_obj.strftime('%d/%m/%Y')}"
    r5[12] = 'Model Number'
    r5[14] = 'BH60M   '
    ws2.append(r5)

    # Row 6: Shift Time & Customer
    r6 = [''] * 17
    r6[1] = 'Shift Time'
    r6[2] = shift_label
    r6[12] = 'Customer'
    r6[14] = 'MCL LAKHANPUR       '
    ws2.append(r6)

    ws2.append([]) # Row 7 empty

    # Row 8: Table Headers
    ws2.append(XLS_HEADERS)
    # Row 9: Table Units
    ws2.append(XLS_UNITS)
    ws2.append([]) # Row 10 empty

    # Row 11+ : Data rows
    for r in records:
        ws2.append([
            r['sl_no'],
            r['date'],
            r['time'],
            float(r['payload']),
            r['load_time'],
            r['haul_time'],
            r['dump_time'],
            r['ret_time'],
            r['ret_stop_time'],
            r['tot_time'],
            float(r['haul_dist']),
            float(r['ret_dist']),
            float(r['tot_dist']),
            float(r['max_haul_spd']),
            float(r['avg_haul_spd']),
            float(r['max_ret_spd']),
            float(r['avg_ret_spd'])
        ])

    # Add Sheet1 as secondary reference sheet
    ws1 = wb.create_sheet(title='Sheet1')
    ws1.append(['BHARAT EARTH MOVERS LTD, INDIA'])
    ws1.append(['Payload Report'])

    return wb

def build_csv_content(records):
    """
    Build standard 18-column CSV string.
    """
    lines = [','.join(CSV_HEADERS)]
    for r in records:
        row = [
            str(r['serial_no']),
            r['date'],
            r['time'],
            r['payload'],
            r['load_time'],
            r['haul_time'],
            r['dump_time'],
            r['ret_time'],
            r['ret_stop_time'],
            r['tot_time'],
            r['haul_dist'],
            r['ret_dist'],
            r['tot_dist'],
            r['max_haul_spd'],
            r['avg_haul_spd'],
            r['max_ret_spd'],
            r['avg_ret_spd'],
            '' # Uploaded By
        ]
        lines.append(','.join(row))
    return '\n'.join(lines) + '\n'

# =========================================================
# CAT DUMPERS GENERATOR
# =========================================================

def generate_cat_shift_records(serial_no, date_obj, shift_type, num_trips=None):
    """
    Generate realistic CAT MineStar cycles for an 8-hour shift.
    shift_type: '1' (06:00-14:00), '2' (14:00-22:00), '3' (22:00-06:00)
    """
    if str(shift_type) in ('1', '1st'):
        shift_label = '06:00 to 14:00'
        prod_start = datetime(date_obj.year, date_obj.month, date_obj.day, 7, random.randint(0, 3), random.randint(0, 59))
        prod_end = datetime(date_obj.year, date_obj.month, date_obj.day, 13, 0, 0)
    elif str(shift_type) in ('2', '2nd'):
        shift_label = '14:00 to 22:00'
        prod_start = datetime(date_obj.year, date_obj.month, date_obj.day, 15, random.randint(0, 3), random.randint(0, 59))
        prod_end = datetime(date_obj.year, date_obj.month, date_obj.day, 21, 0, 0)
    else:
        shift_label = '22:00 to 06:00'
        prod_start = datetime(date_obj.year, date_obj.month, date_obj.day, 23, random.randint(0, 3), random.randint(0, 59))
        next_day = date_obj + timedelta(days=1)
        prod_end = datetime(next_day.year, next_day.month, next_day.day, 5, 0, 0)

    if num_trips is None or num_trips <= 0:
        num_trips = random.randint(11, 14)

    meal_sec = random.randint(25, 35) * 60
    curr_dt = prod_start
    smh_delta = round(random.uniform(11500.0, 14000.0), 2)
    records = []

    for i in range(1, num_trips + 1):
        raw_payload = random.gauss(45.5, 4.5)
        payload = round(max(33.0, min(62.0, raw_payload)), 2)

        empty_dist = round(random.uniform(0.70, 1.05), 2)
        load_dist = round(random.uniform(0.75, 1.10), 2)
        tot_dist = round(empty_dist + load_dist, 2)

        empty_travel_sec = random.randint(210, 330)
        empty_stop_sec = random.randint(1, 10) if random.random() < 0.65 else random.randint(20, 60)
        load_sec = random.randint(68, 123)
        load_stop_sec = random.randint(25, 50)
        load_travel_sec = random.randint(220, 340)

        tot_cycle_sec = empty_travel_sec + empty_stop_sec + load_sec + load_stop_sec + load_travel_sec

        cycle_finish = curr_dt + timedelta(seconds=tot_cycle_sec)
        if cycle_finish > prod_end:
            break

        c_timedate = curr_dt.strftime('%d/%m/%Y %H:%M:%S')

        pass_count = 2 if payload < 35.0 else 4 if payload > 52.0 else random.choice([3, 4])
        fuel = round(1.9 + (tot_cycle_sec / 600) * 0.8 + tot_dist * 0.35 + random.uniform(-0.1, 0.1), 2)

        curr_smh_str = f"{smh_delta:.2f}"
        smh_delta += (tot_cycle_sec / 3600.0)

        load_loc = f"21.736{random.randint(200, 600)},83.816{random.randint(100, 400)},155.0"
        dump_loc = f"21.731{random.randint(500, 900)},83.823{random.randint(600, 900)},199.0"

        rec = {
            'timedate': c_timedate,
            'smh_delta': curr_smh_str,
            'payload': f"{payload:.2f}",
            'empty_travel_hh': fmt_hh_mm_ss(empty_travel_sec),
            'empty_stop_hh': fmt_hh_mm_ss(empty_stop_sec),
            'load_time_hh': fmt_hh_mm_ss(load_sec),
            'load_stop_hh': fmt_hh_mm_ss(load_stop_sec),
            'load_travel_hh': fmt_hh_mm_ss(load_travel_sec),
            'cycle_time_hh': fmt_hh_mm_ss(tot_cycle_sec),
            'empty_travel_h': fmt_h_mm_ss(empty_travel_sec),
            'empty_stop_h': fmt_h_mm_ss(empty_stop_sec),
            'load_time_h': fmt_h_mm_ss(load_sec),
            'load_stop_h': fmt_h_mm_ss(load_stop_sec),
            'load_travel_h': fmt_h_mm_ss(load_travel_sec),
            'cycle_time_h': fmt_h_mm_ss(tot_cycle_sec),
            'empty_dist': f"{empty_dist:.2f}",
            'load_dist': f"{load_dist:.2f}",
            'tot_dist': f"{tot_dist:.2f}",
            'pass_count': pass_count,
            'shift_count': i,
            'fuel': f"{fuel:.2f}",
            'load_loc': load_loc,
            'dump_loc': dump_loc
        }
        records.append(rec)

        if i == num_trips // 2 and num_trips >= 7:
            curr_dt = cycle_finish + timedelta(seconds=meal_sec)
        else:
            tiny_gap = 0 if random.random() < 0.85 else random.randint(2, 10)
            curr_dt = cycle_finish + timedelta(seconds=tiny_gap)

    return records, shift_label

def build_cat_minestar_raw(records, serial_no, date_obj):
    now = datetime.now()
    day_name = now.strftime('%a')
    mon_name = now.strftime('%b')
    date_range_str = date_obj.strftime('%d/%m/%Y')
    rep_gen_date_str = f"{day_name} {mon_name} {now.strftime('%d %Y %H:%M:%S')} GMT+0530 (India Standard Time)"

    lines = [
        ',,,,,MineStar Health Technician Toolbox',
        f',,,,,Truck Payload Detail Data(Date Range :{date_range_str} 00:00 AM-{date_range_str} 11:59 PM)',
        f',,,,,Report Generated Date :{rep_gen_date_str}',
        f',,,,,Equipment :{serial_no}',
        '',
        ',,,,, ',
        'Date and Time,SMH Delta,Payload Weight (ton),Empty Travel Time  (hh:mm:ss),Empty Travel Distance (mile),Empty Stop Time (hh:mm:ss),Load Time (hh:mm:ss),Load Stop Time (hh:mm:ss),Load Travel Time (hh:mm:ss),Load Travel Distance (mile),Total Cycle Time (hh:mm:ss),Total Distance Travelled (mile),Loader Pass Count,Shift Count,Operator ID,Fuel Burned (gal),Load Location,Dump Location'
    ]

    for r in records:
        lines.append(f"{r['timedate']},{r['smh_delta']},{r['payload']},{r['empty_travel_hh']},{r['empty_dist']},{r['empty_stop_hh']},{r['load_time_hh']},{r['load_stop_hh']},{r['load_travel_hh']},{r['load_dist']},{r['cycle_time_hh']},{r['tot_dist']},{r['pass_count']},{r['shift_count']},9.0,{r['fuel']},\"{r['load_loc']}\",\"{r['dump_loc']}\"")

    return '\n'.join(lines) + '\n'

def build_cat_csv_content(records, serial_no):
    lines = [','.join(CAT_HEADERS)]
    for r in records:
        row = [
            r['timedate'],
            r['smh_delta'],
            serial_no,
            r['payload'],
            r['empty_travel_h'],
            r['empty_dist'],
            r['empty_stop_h'],
            r['load_time_h'],
            r['load_stop_h'],
            r['load_travel_h'],
            r['load_dist'],
            r['cycle_time_h'],
            r['tot_dist'],
            str(r['pass_count']),
            r['fuel'],
            ''
        ]
        lines.append(','.join(row))
    return '\n'.join(lines) + '\n'

def main():
    parser = argparse.ArgumentParser(description="Generate synthetic TPMS report (.xlsx and .csv)")
    parser.add_argument("-t", "--type", choices=['beml', 'cat'], default='beml',
                        help="Dumper TPMS model type: beml or cat (default: beml)")
    parser.add_argument("-s", "--serial", default=None, help="Dumper / Truck Serial Number (default: 60611 for BEML, PRB00914 for CAT)")
    parser.add_argument("-d", "--date", default=None, help="Shift date (YYYY-MM-DD or DD/MM/YYYY, default: today)")
    parser.add_argument("--shift", choices=['1', '2', '3', '1st', '2nd', '3rd'], default="1", 
                        help="Shift (1: 06-14, 2: 14-22, 3: 22-06, default: 1)")
    parser.add_argument("-n", "--trips", type=int, default=None, help="Number of trips (optional, default: auto 11-15)")
    parser.add_argument("-m", "--month", default=None, help="Month suffix for CSV (default: derived from date)")
    parser.add_argument("-o", "--output", default=".", help="Output directory (default: current directory)")

    args = parser.parse_args()
    dumper_type = args.type.lower()

    # Parse date
    if args.date:
        for fmt in ('%Y-%m-%d', '%d/%m/%Y', '%d-%m-%Y'):
            try:
                date_obj = datetime.strptime(args.date.strip(), fmt)
                break
            except ValueError:
                continue
        else:
            print(f"[ERROR] Invalid date format: {args.date}. Use YYYY-MM-DD or DD/MM/YYYY.")
            sys.exit(1)
    else:
        date_obj = datetime.now()

    if args.serial:
        serial_no = args.serial.strip()
    else:
        serial_no = 'PRB00914' if dumper_type == 'cat' else '60611'

    out_dir = os.path.abspath(args.output)
    os.makedirs(out_dir, exist_ok=True)

    if dumper_type == 'cat':
        records, shift_label = generate_cat_shift_records(serial_no, date_obj, args.shift, args.trips)

        now = datetime.now()
        day_str = now.strftime('%a')
        mon_str = now.strftime('%b')
        ts_str = f"{day_str} {now.strftime('%d')} {mon_str} {now.strftime('%Y %H %M %S')}"
        raw_filename = f"TruckPayload_{serial_no}_{ts_str}.csv"
        csv_filename = f"{serial_no}.csv"

        raw_path = os.path.join(out_dir, raw_filename)
        csv_path = os.path.join(out_dir, csv_filename)

        # 1. Write Raw MineStar CSV
        raw_text = build_cat_minestar_raw(records, serial_no, date_obj)
        with open(raw_path, 'w', encoding='utf-8', newline='') as f:
            f.write(raw_text)

        # 2. Write Standard 16-Column Converted CSV
        csv_text = build_cat_csv_content(records, serial_no)
        with open(csv_path, 'w', encoding='utf-8', newline='') as f:
            f.write(csv_text)

        total_tons = sum(float(r['payload']) for r in records)

        print("\n========================================================")
        print(" TPMS Report Generator (Caterpillar MineStar Synthetic)")
        print(f" Dumper Type:    CAT")
        print(f" Dumper Serial:  {serial_no}")
        print(f" Shift:          {shift_label} ({args.shift})")
        print(f" Date:           {date_obj.strftime('%d/%m/%Y')}")
        print(f" Total Cycles:   {len(records)}")
        print(f" Total Payload:  {total_tons:.2f} Tons")
        print("--------------------------------------------------------")
        print(f" Created Raw:    {raw_path}")
        print(f" Created CSV:    {csv_path}")
        print("========================================================\n")

    else:
        month_name = args.month.strip().lower() if args.month else date_obj.strftime('%B').lower()
        records, shift_label = generate_shift_records(serial_no, date_obj, args.shift, args.trips)

        # Filenames
        timestamp_str = date_obj.strftime('%d_%m_%Y') + "_" + datetime.now().strftime('%H_%M_%S')
        xlsx_filename = f"{serial_no}_{timestamp_str}.xlsx"
        csv_filename = f"{serial_no}_{month_name}.csv"

        xlsx_path = os.path.join(out_dir, xlsx_filename)
        csv_path = os.path.join(out_dir, csv_filename)

        # 1. Write Excel (.xlsx)
        wb = build_excel_workbook(records, serial_no, date_obj, shift_label)
        wb.save(xlsx_path)

        # 2. Write CSV (.csv)
        csv_text = build_csv_content(records)
        with open(csv_path, 'w', encoding='utf-8', newline='') as f:
            f.write(csv_text)

        total_tons = sum(float(r['payload']) for r in records)

        print("\n========================================================")
        print(" TPMS Report Generator (BEML Synthetic)")
        print(f" Dumper Type:    BEML")
        print(f" Dumper Serial:  {serial_no}")
        print(f" Shift:          {shift_label} ({args.shift})")
        print(f" Date:           {date_obj.strftime('%d/%m/%Y')}")
        print(f" Total Cycles:   {len(records)}")
        print(f" Total Payload:  {total_tons:.2f} Tons")
        print("--------------------------------------------------------")
        print(f" Created Excel:  {xlsx_path}")
        print(f" Created CSV:    {csv_path}")
        print("========================================================\n")

if __name__ == '__main__':
    main()
