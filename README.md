# TPMS File Generator (PWA & CLI)

A Progressive Web App (PWA) and Python CLI tool designed to parse BEML / Dumper TPMS (Tyre/Truck Payload Monitoring System) `.xls` / `.xlsx` report files, arrange and clean cycle records, and export standardized CSV files named by dumper serial number and month (e.g., `60611_september.csv`).

---

- **Dual Dumper Support (BEML & CAT MineStar)**: Choose between **BEML Dumpers** (18-column standard format named `<serial>_<month>.csv`) and **CAT Dumpers** (16-column MineStar format named `<Equipment>.csv`).
- **Synthetic Report Generator (.xlsx & .csv)**: Generate brand-new, realistic BEML TPMS payload reports by simply picking the Date, Shift (1st, 2nd, or 3rd), Dumper Serial Number, and optional trip count.
- **PWA (Offline-First)**: Runs directly in Google Chrome, Microsoft Edge, or Android browsers. Once loaded or installed, works 100% offline without needing internet.
- **Single File or Entire Folder**: Accepts single files, multiple files, or an entire folder of `.xls` / `.xlsx` / `.csv` reports via drag-and-drop or file pickers.
- **Month Suffix Selection**: Easily select from all 12 months (defaults to lowercase as in `60611_september.csv`) or specify a custom suffix.
- **Automatic Serial Number Extraction**: Pulls truck serials from BEML metadata cells (`Serial Number: 60611`) or CAT MineStar headers (`Equipment :PRB00914`).
- **Interactive In-Browser Preview**: Preview all converted cycles and summary metrics (total tonnage, cycles, average payload) in a scrollable table before downloading, dynamically adapted to 18 BEML columns or 16 CAT columns.
- **Batch Export & Consolidated CSV**: Download individual converted CSV files, click **"Download Consolidated CSV"** to get a single unified CSV containing all dumpers (e.g. `August_60611_60645.csv`), or click **"Download All (ZIP)"** to get a complete archive containing all individual CSVs plus the consolidated file.
- **Python CLI Tools**: Includes `convert.py` for conversion (with `-t/--type` and `-c/--consolidate` options) and `generate.py` for automated synthetic report generation.

---

## Output Specifications

### 1. BEML Dumpers (18 Columns)
Generated filename: `<serial>_<month>.csv` (e.g. `60611_september.csv`).

| # | Column Header | Description | Format |
|---|---------------|-------------|--------|
| 1 | `TRUCK NO` | Dumper serial number | String (e.g. `60611`) |
| 2 | `T_Date` | Cycle Date | `DD/MM/YYYY` |
| 3 | `Time` | Cycle Start Time | `HH:MM:SS` |
| 4 | `PayLoad` | Measured payload weight | Metric Tons (`0.00`) |
| 5 | `Loading Time` | Duration of shovel loading | `M:SS` |
| 6 | `Hauling Time` | Duration of travel under load | `M:SS` |
| 7 | `Stop Dump Time` | Stop & dumping duration | `000:SS` or `M:SS` |
| 8 | `Return Time` | Empty return travel duration | `M:SS` |
| 9 | `Return Stop Time` | Empty return stop duration | `000:SS` or `M:SS` |
| 10 | `Total Time` | Total cycle time duration | `000:SS` or `M:SS` |
| 11 | `Haul Travel Distance` | Distance traveled loaded | Kilometers (`0.00`) |
| 12 | `Return Travel Distance` | Distance traveled empty | Kilometers (`0.00`) |
| 13 | `Total Distance` | Total round-trip distance | Kilometers (`0.00`) |
| 14 | `MaxHaul Speed` | Maximum speed loaded | km/h (`0.00`) |
| 15 | `AvgHaul Speed` | Average speed loaded | km/h (`0.00`) |
| 16 | `MaxReturn Speed` | Maximum speed empty | km/h (`0.00`) |
| 17 | `AvgReturn Speed` | Average speed empty | km/h (`0.00`) |
| 18 | `Uploaded By` | Upload operator placeholder | Blank |

### 2. CAT MineStar Dumpers (16 Columns)
Generated filename: `<Equipment>.csv` (e.g. `PRB00914.csv`).

| # | Column Header | Source in MineStar Input | Description / Transformation |
|---|---|---|---|
| 1 | `TIMEDATE` | `Date and Time` | Exact timestamp string (`DD/MM/YYYY HH:MM:SS`) |
| 2 | `PAYLOAD SMH` | `SMH Delta` | Machine hour meter delta (`0.00`) |
| 3 | `SERIAL NUMBER` | `Equipment :<SN>` | Equipment ID string (e.g. `PRB00914`) |
| 4 | `PAYLOADTONS` | `Payload Weight (ton)` | Measured payload in tons (`0.00`) |
| 5 | `TRAVEL EMPTY TIME HHMMSS` | `Empty Travel Time` | Unpadded hour (`H:MM:SS`, e.g. `0:29:55`) |
| 6 | `TRAVEL EMPTY DISTANCEMI` | `Empty Travel Distance` | Travel distance empty in miles |
| 7 | `STOPPED EMPTY TIME HHMMSS` | `Empty Stop Time` | Unpadded hour (`H:MM:SS`, e.g. `0:04:45`) |
| 8 | `LOAD TIME HHMMSS` | `Load Time` | Shovel loading time (`H:MM:SS`, e.g. `0:01:13`) |
| 9 | `STOPPED LOADED TIME HHMMSS`| `Load Stop Time` | Stop time under load (`H:MM:SS`, e.g. `0:00:40`)|
| 10 | `LOADED TRAVEL TIME HHMMSS` | `Load Travel Time` | Loaded hauling time (`H:MM:SS`, e.g. `0:03:50`)|
| 11 | `LOADED TRAVEL DISTANCEMI` | `Load Travel Distance` | Loaded travel distance in miles |
| 12 | `CYCLE TIME HHMMSS` | `Total Cycle Time` | Total cycle time duration (`H:MM:SS`) |
| 13 | `CYCLE DISTANCEMI` | `Total Distance Travelled`| Total travel distance in miles |
| 14 | `LOADER PASSES` | `Loader Pass Count` | Loader pass count integer |
| 15 | `FUEL USED GAL` | `Fuel Burned (gal)` | Fuel consumption in gallons |
| 16 | `UPLOADED BY` | Added column | Blank placeholder with trailing comma |

---

## How to Run the PWA

### Method A: One-Click Windows Launcher (Recommended)
Simply double-click:
```bat
start.bat
```
This automatically starts a local web server at `http://localhost:8080` and opens your default browser.

### Method B: Manual Command Line
```powershell
python -m http.server 8080
```
Then navigate to [http://localhost:8080](http://localhost:8080) in your web browser.

### Installing as a Desktop App (PWA)
1. Open [http://localhost:8080](http://localhost:8080) in Chrome or Edge.
2. Click the **"Install App"** button in the header, or click the install icon in your browser's address bar.
3. The app is now installed on your desktop with its own icon and window, and works completely offline!

---

## How to Use the Python CLI Converter

A standalone CLI utility `convert.py` is included for batch processing directly from terminal:

```powershell
# Convert BEML reports:
python convert.py "Test input" -t beml -m august -o "Test output"

# Convert CAT MineStar reports:
python convert.py "Test 2 input" -t cat -o "Test 2 output"

# Auto-detect format:
python convert.py "Test 2 input" -o "Test 2 output"
```

### CLI Options:
- `input`: Path to a single report file (`.xls`, `.xlsx`, `.csv`) or a directory containing reports.
- `-t, --type`: Dumper type selection: `beml`, `cat`, or `auto` (default: `auto`).
- `-m, --month`: Target month name for output filename (default: `september`).
- `-o, --output`: Directory where output `.csv` files will be saved (default: `.`).
- `-c, --consolidate`: Create a consolidated CSV containing all processed dumper rows concatenated under a single header named `<Month>_<serial1>_<serial2>_...csv` (e.g. `August_60611_60645.csv`). Default: automatically generated whenever multiple files are processed.
- `--no-consolidate`: Skip generating the consolidated CSV file.

### Shift Timings & Production Windows:
Mining dumpers enter production 1 hour after shift start and cease production 1 hour before shift end:
- **1st Shift**: `06:00 AM – 02:00 PM` $\rightarrow$ **Production Window**: `07:00 AM – 01:00 PM`
- **2nd Shift**: `02:00 PM – 10:00 PM` $\rightarrow$ **Production Window**: `03:00 PM – 09:00 PM`
- **3rd Shift**: `10:00 PM – 06:00 AM` $\rightarrow$ **Production Window**: `11:00 PM – 05:00 AM` (next day)

Cycles and mid-shift breaks are realistically distributed within this 6-hour production window.

A standalone CLI generator `generate.py` creates new synthetic reports for both **BEML** and **CAT** dumpers:

```powershell
# Generate CAT MineStar synthetic report (PRB00914, 12 trips, Shift 1):
python generate.py -t cat -s PRB00914 -d 2026-09-08 --shift 1 -n 12 -o "Test 2 output"

# Generate BEML synthetic report (60611, 14 trips, Shift 1):
python generate.py -t beml -s 60611 -d 2026-09-08 --shift 1 -n 14 -o "Test output"

# Generate CAT 2nd Shift (14:00 to 22:00) with auto trip count:
python generate.py -t cat -s PRB00937 -d 2026-09-08 --shift 2 -o "Test 2 output"
```

### CLI Generator Options:
- `-t, --type`: Dumper model type: `beml` or `cat` (default: `beml`).
- `-s, --serial`: Dumper Serial Number (e.g. `PRB00914` for CAT, `60611` for BEML; default: auto-selected by model).
- `-d, --date`: Shift date in `YYYY-MM-DD` or `DD/MM/YYYY` (default: today).
- `--shift`: Shift selection (`1` = 06:00-14:00, `2` = 14:00-22:00, `3` = 22:00-06:00, default: `1`).
- `-n, --trips`: Number of trips/cycles (optional; default: auto 11-14 trips).
- `-m, --month`: Target month name for BEML CSV (default: derived from date).
- `-o, --output`: Target output directory (default: current directory).

