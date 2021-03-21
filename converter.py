import os
import shutil
import sys
import xlrd
import csv
import random

"""
Steps to execute script:
1. Copy and paste Excel2CSV into Directory containing files/folders.
2. cd into directory that contains Excel2CSV (do not cd into Batch Convert)
3. Run the following command: python3 Excel2CSV/converter.py 
4. Dumped files will appear Excel2CSV_Output in the working directory.


By Michael A. Rudy

"""

## directory names to hold copy of paths that are xlsx or docx
main_directory = "Excel2CSV"
excel_directory = "Excel2CSV_dump"
output_directory = "Excel2CSV_Output"



##current working directory
cwd = os.getcwd()


#new connections 
output_dump = os.path.join(cwd, output_directory)
excel_dump = os.path.join(output_dump, excel_directory)
main_folder = os.path.join(cwd, main_directory)


#scan directory to see if folders exist - overwrite if needed --> this prevents errors if user is runner script multiple times in the same directory
def directory_scanner():
    if not os.path.exists(output_dump):
        os.mkdir(output_dump)
    else:
        shutil.rmtree(output_dump)
        os.makedirs(output_dump)

    if not os.path.exists(excel_dump):
        os.mkdir(excel_dump)
    else:
        shutil.rmtree(excel_dump)
        os.makedirs(excel_dump)
    


#collect all xlsx through sub-directories from cwd
def xlsx_scrub():
    for dir, subdir, files in os.walk('.'):
        for file in files:
            if file.endswith('.xlsx'):
                #if the file is not in output folder, copy it and move. This prevents errors if user has duplicated files that are nested in sub-directories.
                if os.path.basename(file) not in os.listdir(excel_dump):
                    print("Processing "+file)
                    path = os.path.join(dir, file)
                    shutil.copy(path, excel_dump)
                    #rename files if whitespaces are found to prevent errors during conversion
                    for filename in os.listdir(excel_dump):
                        if " " in filename:
                            os.rename(os.path.join(excel_dump,filename),os.path.join(excel_dump, filename.replace(' ', '_').lower()))            
                else:
                    continue
   
def xlsx_to_csv_converter():
    for dir, subdir, files in os.walk(excel_dump):
        for file in files:
            try:
                
                print("Converting "+file)
                path_to_excel_file = os.path.join(dir, file) 
                wb = xlrd.open_workbook(path_to_excel_file, on_demand=True)                   
                #creates a csv file for each sheet according to their sheet name 
                for x in wb.sheet_names():
                    sh = wb.sheet_by_name(x)
                    os.chdir(excel_dump)
                    #add randomint to the end of filename to duplicate sheet name errors when processing.
                    new_csv = open(x+str(random.randint(0,50))+'.csv', 'w')
                    wr = csv.writer(new_csv, quoting=csv.QUOTE_ALL)
                    for rownum in range(sh.nrows):
                        wr.writerow(sh.row_values(rownum))
                    new_csv.close()
            
            #XLRDErrors may be encountered if file being processed is not true .xlsx file --> this may occur if file has .xlsx exentsion but is not a true excel file
            except xlrd.XLRDError as error:
                print("Failed to convert  "+file+" - "+str(error))
                continue

            #removes original .xlsx file from output folder --> end results contains unique csv files
            os.remove(file)
                
def main():
    funcs = [directory_scanner(), xlsx_scrub(), xlsx_to_csv_converter()]
    for func in funcs:
        try:
            func
        except: 
            continue
            

if __name__ == "__main__":
    main()

