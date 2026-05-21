# xl2csv

A Python tool for exporting Excel (.xlsx) cell ranges to CSV files, driven by a JSON configuration template.

---

## How it works

The script reads a `.xlsx` file, then for each **step** defined in the JSON template:

1. Opens the target Excel sheet
2. Extracts the **headers** range
3. Extracts the **data** range
4. Writes everything to a `.csv` file in the output folder

On success, the source `.xlsx` file is moved to `dir_out`. If an error occurs on any step, it is moved to `dir_rej`.

---

## Project structure

```
xl2csv/
├── xl2csv.py           # Main script
├── xl2csv.ini          # Configuration file (paths, separator)
├── template.json       # Export definitions (sheets, ranges, CSV names)
├── data_in/            # Input folder for xlsx files to process
├── data_out/           # Output folder (exported CSVs + processed xlsx)
├── work/               # Intermediate working folder
└── rej/                # Rejection folder (xlsx files with errors)
```

---

## Requirements

**Python 3.8+** and the following dependencies:

```bash
pip install openpyxl python-dotenv
```

---

## Configuration

### `xl2csv.ini`

Main configuration file, passed to the script via the `-cf` argument.

| Key | Description |
|---|---|
| `default_path` | Root path of the project (e.g. `C:/projects/xl2csv/`) |
| `dir_in` | Input folder for xlsx files (e.g. `data_in/`) |
| `dir_out` | Output folder for CSVs and processed xlsx files (e.g. `data_out/`) |
| `dir_work` | Temporary working folder (e.g. `work/`) |
| `dir_rej` | Rejection folder for files with errors (e.g. `rej/`) |
| `S_TEMPLATE_PATH` | Absolute path to the `template.json` file |
| `S_SEPARATOR` | ASCII code of the CSV separator (`59` = `;`, `44` = `,`, `9` = tab) |

Example:
```ini
default_path=C:/projects/xl2csv/
dir_in=data_in/
dir_out=data_out/
dir_work=work/
dir_rej=rej/
S_TEMPLATE_PATH=C:/projects/xl2csv/template.json
S_SEPARATOR=59
```

---

### `template.json`

Defines the exports to perform as a list of **steps**. Each step produces one CSV file.

```json
{
  "steps": [
    {
      "sheet_name": "SheetName",
      "export_name": "output_file.csv",
      "headers_dimension": {
        "col_start": 1,
        "col_end": 4,
        "row_start": 1,
        "row_end": 1
      },
      "data_dimension": {
        "col_start": 1,
        "col_end": 4,
        "row_start": 2,
        "row_end": 10
      }
    }
  ]
}
```

| Field | Description |
|---|---|
| `sheet_name` | Exact name of the sheet in the xlsx file |
| `export_name` | Name of the CSV file to generate |
| `headers_dimension` | Cell range containing the column headers |
| `data_dimension` | Cell range containing the data to export |

The `col_start`, `col_end`, `row_start`, `row_end` values use Excel numbering (column A = 1, row 1 = 1).

It is possible to define **multiple steps on the same sheet** with different ranges, to produce several distinct CSV files (see `clients.csv` and `commandes.csv` in the example template).

---

## Usage

```bash
python .\xl2csv.py -cf <path to .ini file> -fn "<path to .xlsx file>"
```

### Arguments

| Argument | Description |
|---|---|
| `-cf` / `--context_file` | Path to the `.ini` configuration file |
| `-fn` / `--input_file` | Path to the `.xlsx` file to process |

### Examples

```bash
# Windows
python .\xl2csv.py -cf .\xl2csv.ini -fn "C:\data\my_file.xlsx"

# Linux / macOS
python ./xl2csv.py -cf ./xl2csv.ini -fn "/data/my_file.xlsx"
```

---

## Output

- CSV files are generated in `{default_path}{dir_out}` using the configured separator.
- The source `.xlsx` file is **moved** to `{default_path}{dir_out}` after processing.
- If an error occurs on any step, the `.xlsx` is moved to `{default_path}{dir_rej}`.
- Logs are printed to the console with timestamp and level (`INFO`, `ERROR`).

---

## Logs

The script produces structured console logs, for example:

```
(+0200)2026-04-17T10:23:01.042: INFO: xl2csv.py: STARTING...
(+0200)2026-04-17T10:23:01.043: INFO: xl2csv.py: Openning my_file.xlsx
(+0200)2026-04-17T10:23:01.120: INFO: xl2csv.py: Ecriture du fichier : data_out/clients.csv - OK
```
