# research-tools-week1
This repository contains the first-week experiment for the Data Security and Privacy Protection course.

`code/demo.py` is a runnable program which will output "Hello, GitHub!".

THIS IS A PARAGRAPH.

## XLSX to JSON statistics tool

`code/xlsx_to_json.py` converts a two-dimensional `.xlsx` worksheet to JSON. The
first row is used as field names. The output contains the data records and table,
column, type, empty-cell, and numeric statistics.

Install the dependency:

```powershell
python -m pip install -r requirements.txt
```

Run the converter from the project root:

```powershell
python code\xlsx_to_json.py input.xlsx output.json
```

The output argument is optional. Without it, the JSON file is created beside the
input file. For a workbook with multiple worksheets, select one by name:

```powershell
python code\xlsx_to_json.py input.xlsx output.json --sheet "Sheet2"
```

Run the automated tests:

```powershell
python -m unittest discover -s tests -v
```
