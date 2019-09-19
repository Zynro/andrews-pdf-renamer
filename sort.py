import os
import traceback
from os import walk
from os import listdir
from docx import Document

def sorting(L):
    splitup = L.split('-')
    return splitup[2], splitup[0], splitup[1]

def main():
    print("=======================Welcome to a Simple PDF Sorting/Indexing Script=======================")
    file_list = []
    try:
        for (dirpath, dirnames, filenames) in walk(os.getcwd()):
            file_list = [filenames.strip() for filenames in filenames]
            break

        file_dict = {}

        for file in file_list:
            if ".pdf" not in file:
                continue
            file_stripped = file.replace(".pdf", "").strip()
            ind = file_stripped.split(" ")
            file_dict[file] = ind[-1] + '.pdf'
        if not file_dict:
            input('No PDF files found to rename, press Enter to exit...')
            return
        else:
            print(f"I found {len(file_dict)} files that can be sorted. Would you like to continue? y/n")
            x = input()
            if x == "n":
                print("Press Enter to exit...")
                return
            start_index = input("Please enter an index to start at, enter 0 for default:\n")
            if start_index:
                try:
                    start_index = int(start_index)
                except Exception:
                    print("Has to be a number or nothing, exiting.")
                    input()
                    return
    except Exception as e:
        traceback.print_exc()
        input("Press Enter to exit...")
        return
    print("=========================================Now Sorting...=========================================")
    try:
        file_sorted_list = [key for key, value in sorted(file_dict.items(), key=lambda x: sorting(x[1]))]
    except Exception as e:
        traceback.print_exc()
        input("Press Enter to exit...")
        return
    if not start_index or start_index == 0:
        start_index = 1
    try:
        file_sorted_dict = {}
        for x in range(start_index, len(file_sorted_list)+start_index):
            if len(str(x)) == 1:
                value = f"00{x}"
            elif len(str(x)) == 2:
                value = f"0{x}"
            else:
                value = x
            file_name = file_sorted_list[x - start_index].split(" ")
            for y in file_name[0]:
                if not y.isdigit():
                    break
                del file_name[0]
                break
            file_name = " ".join(file_name)
            file_sorted_dict[file_sorted_list[x - start_index]] = f"{value} {file_name}"
    except Exception as e:
        traceback.print_exc()
        input("Press Enter to exit...")
        return

    #for key in file_sorted_dict:
    #   print(key, file_sorted_dict[key])
    pause_check = 1
    try:
        for file in file_sorted_dict:
            if pause_check == 150:
                input("Paused due to 150 files reached, press Enter to continue...")
                pause_check = 1
            os.rename(file, file_sorted_dict[file])
            print(f"{file} has been renamed to {file_sorted_dict[file]}")
            pause_check += 1
    except Exception as e:
        traceback.print_exc()
        input("Press Enter to exit...")
        return
    print("===================================================================================")
    print("===================================================================================")


    print("Would you like to undo? y/n")
    x = input()
    pause_check = 1
    if x == "y":
        for file in file_sorted_dict:
            if pause_check == 150:
                input("Paused due to 150 files reached, press Enter to continue...")
                pause_check = 1
            os.rename(file_sorted_dict[file], file)
            pause_check += 1
        input("All changes reverted.\nPress Enter to exit.")
    else:
        document = Document()
        table = document.add_table(rows = 1, cols = 4)
        heading_cells = table.rows[0].cells
        heading_cells[0].text = 'Index #'
        heading_cells[1].text = 'File'
        heading_cells[2].text = 'Party'
        heading_cells[3].text = 'Date'
        for index in file_sorted_dict.keys():
            split_list = file_sorted_dict[index].split(" ")
            file_index = split_list[0]
            file_client = split_list[1]
            file_date = split_list[-1]
            file_date = file_date.split(".")[0]
            file_name = " ".join(split_list[2:len(split_list)-1])

            cells = table.add_row().cells
            cells[0].text = str(file_index)
            cells[1].text = file_name
            cells[2].text = file_client
            cells[3].text = file_date

        document.save('Index.docx')
        print("===================================================================================")
        input("Document 'Index' has been generated and a table has been exported.\nPress Enter to exit...")



if __name__ == '__main__':
    main()