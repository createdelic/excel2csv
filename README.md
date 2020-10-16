# excel2csv

Convert worksheets in an Excel spreadsheet to CSV files.

Created to help with game localization. The localizer can work with a giant, friendly spreadsheet in Excel, which this tool can convert into CSV file(s) that are easy to parse by a game.

## How to build

```
python -m PyInstaller excel2csv.py
python -m PyInstaller excel2singlecsv.py
python -m PyInstaller excel2wc.py
```

After building, check the `dist` folder for the EXE files.

## Scripts

`excel2csv` - convert each worksheet inside an Excel spreadsheet into a CSV file

`excel2singlecsv` - convert an Excel spreadsheet into a single CSV file

`excel2wc` - count the number of words in an Excel spreadsheet

## How to Use

Run each command with `--help` or see the `examples` folder.
