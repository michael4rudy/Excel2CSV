# Excel2CSV
- Converts excel files types (.xlsx) to .csv from working directory through subdirectories (files contained in nested folders)
- Duplicates file before converting file type to preserve original data - file is converted to csv file type in Excel2CSV_Output
- Output folder contains converted file types for all folders/subfolders processed.

## Executing Script
1. Copy and paste Excel2CSV into working directory containing files/subdirectories
2. cd into directory containing Excel2CSV (do not cd into Excel2CSV)
3. Run the following command: 'python3 Excel2CSV/converter.py'
4. Excel2CSV_Output folder will appear in working directory 

## General Purpose
- Converts .xlsx file types to .csv for programs that process unicode encoded files. 
- Portability allows users to process excel files by simply copying and pasting Excel2CSV into directories.

## Further Builds
- Create similar function to process .docx to .txt with UTF-8 encoding.
