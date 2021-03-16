import os
import shutil


"""
Steps to execute script:
1. Copy and paste BatchConvert into Directory containing files/folders.
2. cd into directory that contains BatchConvert (do not cd into Batch Convert)
3. Run the command python3 ./BatchCovert/converter.py 
4. Dumped files will appear Batch Output in that directory.

"""

## directory names to hold copy of paths that are xlsx or docx
excel_directory = "xlsx_to_csv_dump"
docx_directory = "docx_to_txt_dump"
output_directory = "Batch Convert Output"

##current working directory
cwd = os.getcwd()


#new connections 
output_dump = os.path.join(cwd, output_directory)
excel_dump = os.path.join(output_dump, excel_directory)
word_dump = os.path.join(output_dump, docx_directory)



#scan directory to see if folder exists --> then create
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

    if not os.path.exists(word_dump):
        os.mkdir(word_dump)
    else:
        shutil.rmtree(word_dump)
        os.makedirs(word_dump)
    


#find all xlsx through sub-folders
def xlsx_scrub():
    for dir, subdir, files in os.walk('.'):
        for file in files:
            if file.endswith('.xlsx'):
                if os.path.basename(file) not in os.listdir(excel_dump):
                    print("Processing "+file)
                    path = os.path.join(dir, file)
                    shutil.copy(path, excel_dump)
                else:
                    continue
    
#convert all xlsx from scrub to csv in dedicated directory
def xlsx_to_csv_converter():
    for dir, subdir, files in os.walk(output_directory):
            for file in files:
                if file.endswith('.xlsx'):
                    paths = os.path.join(dir, file)
                    root = os.path.splitext(paths)[0]
    
                    os.rename(paths, root + '.csv')
    print("Excel files processed..")

def docx_scrub():
    for dir, subdir, files in os.walk('.'):
        for file in files:
            if file.endswith('.docx'):
                if os.path.basename(file) not in os.listdir(word_dump):
                    print("Processing "+file)
                    path = os.path.join(dir, file)
                    shutil.copy(path, word_dump)
                else:
                    continue
   
def docx_to_txt_converter():
    for dir, subdir, files in os.walk(output_directory):
        for file in files:
            if file.endswith('.docx'):
                paths = os.path.join(dir, file)
                root = os.path.splitext(paths)[0]
                os.rename(paths, root + '.text')
    print("Docx files processed..")        
    
def main():
    funcs = [directory_scanner(), xlsx_scrub(), xlsx_to_csv_converter(), docx_scrub(), docx_to_txt_converter()]
    for func in funcs:
        try:
            func
        except ValueError:
            continue

if __name__ == "__main__":
    main()
   