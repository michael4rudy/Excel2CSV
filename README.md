# Excel2CSV
- Converts excel files types (.xlsx) to .csv from working directory through subdirectories (files contained in nested folders)
- Duplicates file before converting file type to preserve original data - file is converted to csv file type in Excel2CSV_Output
- Output folder contains converted file types for all folders/subfolders processed.
- CSV files are created for each sheet contained in an excel file.
- CSV files are generated according to their sheet name (random integer is added at the end of filename to prevent duplicated sheet name errors)

## Executing Script
1. Copy and paste Excel2CSV into working directory containing files/subdirectories
2. cd into directory containing Excel2CSV (do not cd into Excel2CSV)
3. Run the following command: 'python3 Excel2CSV/converter.py'
4. Excel2CSV_Output folder will appear in working directory 

## General Purpose
- Converts .xlsx file types to .csv for programs that process unicode encoded files. 
- Portability allows users to process excel files by simply copying and pasting Excel2CSV into directories.

## Encountering Error Messages
<img width="530" alt="Screen Shot 2021-03-21 at 2 56 09 PM" src="https://user-images.githubusercontent.com/80732776/111917290-94e7bb00-8a55-11eb-80bd-5897dda2f020.png">

- If interpreter encounters an exception, it checks the except blocks associated with that try block.
- Wrote three except blocks for errors that I found while testing on both WindowsOS and MacOS -- (e.g. corrupted files / encoding errors)
- This allows script to continue converting files after encountering an error (append to except blocks if neccessary)

## DEMO

- Copy and paste Excel2CSV-main into any directory. Script will walk through each folder and scan for files that end with '.xlsx'
<img width="769" alt="Screen Shot 2021-03-21 at 1 38 17 PM" src="https://user-images.githubusercontent.com/80732776/111915010-b2fbee00-8a4a-11eb-90d6-7a934da9d40e.png">

- Open terminal and cd into directory containing Excel2CSV-main (do not cd into the main folder)
- Run the following command: python3 Excel2CSV-main/converter.py
- 
![Screen Shot 2021-03-21 at 1 42 03 PM](https://user-images.githubusercontent.com/80732776/111915152-39b0cb00-8a4b-11eb-985a-0d8586361f88.png)
 
 ![Screen Shot 2021-03-21 at 1 43 02 PM](https://user-images.githubusercontent.com/80732776/111915186-5cdb7a80-8a4b-11eb-8f5e-dcd17e0a9ca5.png)

- Output folder will appear in the working directory

<img width="776" alt="Screen Shot 2021-03-21 at 1 43 25 PM" src="https://user-images.githubusercontent.com/80732776/111915196-6a910000-8a4b-11eb-929d-f228563e623f.png">

<img width="743" alt="Screen Shot 2021-03-21 at 1 44 49 PM" src="https://user-images.githubusercontent.com/80732776/111915251-9ca26200-8a4b-11eb-917a-d7661e294e5f.png">

## TO-DO
- Display the number of files converted and the number of files that were skipped due to error.
- Create similar function to process .docx to .txt with UTF-8 encoding.


