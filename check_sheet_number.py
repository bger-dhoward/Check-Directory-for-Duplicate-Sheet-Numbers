import os
import csv
import time
import tkinter as tk
from tkinter import filedialog

os.system('color')

class bcolors: 
    # from https://stackoverflow.com/questions/287871/how-do-i-print-colored-text-to-the-terminal 
    # per comments at above, make sure to import os, then set os.system('color') to get these to work!
    HEADER = '\x1b[95m' 
    OKBLUE = '\x1b[94m'
    OKCYAN = '\x1b[96m'
    OKGREEN = '\x1b[92m'
    WARNING = '\x1b[93m'
    FAIL = '\x1b[91m'
    ENDC = '\x1b[0m'
    BOLD = '\x1b[1m'
    UNDERLINE = '\x1b[4m'



def proceed_or_quit():
    response = input(f'Press {bcolors.WARNING}Enter{bcolors.ENDC} to proceed, or {bcolors.WARNING}Q{bcolors.ENDC} to quit: ')
    if response.upper() == "Q":
        print('Thanks for stopping by!')
        time.sleep(3)
        quit()
    elif len(response) > 0:
        print(f'Invalid response: {response}')
        proceed_or_quit()

def process_directory_files():
    root = tk.Tk()
    root.withdraw()

    directory = filedialog.askdirectory(title='Select Directory')
    root.destroy()

    timestamp = time.strftime('%Y-%m-%d %H_%M')
    folder_name = directory.split('/')[-1]
    data_filename = f"{folder_name} Current Files {timestamp}.csv"

    os.chdir(directory)

    pdfs = [file for file in os.listdir() if file.endswith('.pdf')]

    sheet_nums = {}
    export_data = [['Sheet Number', 'File Name', 'Modified (Epochal)', 'Date Modified', 'Status', 'Single/Multiple Files']]
    duplicate_count = 0

    for file in pdfs:
        sheet_number = file.split(' ')[0]
        file_modified = os.path.getmtime(file)
        file_modified_date = time.ctime(file_modified)
        if sheet_number not in sheet_nums:
            sheet_nums[sheet_number] = []
        sheet_nums[sheet_number].append([sheet_number, file, file_modified, file_modified_date])

    for sheet_number in sorted(sheet_nums):
        files = sorted(sheet_nums[sheet_number], key=lambda sheet: sheet[1])
        files[0].append('Current')
        for file in files[1:]:
            file.append('Superseded - Delete')
        if len(files) > 1:
            for file in files:
                file.append('Multiple')
            duplicate_count += 1            
        else:
            files[0].append('Single')

        for file in files:
            export_data.append(file)

        #export_data.append([])

    with open(data_filename, 'w', newline='') as file:
        writer = csv.writer(file)
        for row in export_data:
            writer.writerow(row)
    
    print(f'\nDirectory: {bcolors.OKGREEN}{directory}{bcolors.ENDC}\nChecked {bcolors.FAIL}{len(pdfs)}{bcolors.ENDC} PDF files, found {bcolors.FAIL}{duplicate_count}{bcolors.ENDC} sheet numbers with duplicates.\nResults saved to file: {bcolors.OKGREEN}{data_filename}{bcolors.ENDC}.\n')

test_styles = f"""
{bcolors.HEADER}header{bcolors.ENDC}
{bcolors.OKBLUE}okblue{bcolors.ENDC}
{bcolors.OKCYAN}okcyan{bcolors.ENDC}
{bcolors.OKGREEN}okgreen{bcolors.ENDC}
{bcolors.WARNING}warning{bcolors.ENDC}
{bcolors.FAIL}fail{bcolors.ENDC}
{bcolors.ENDC}endc{bcolors.ENDC}
{bcolors.BOLD}bold{bcolors.ENDC}
{bcolors.UNDERLINE}underline{bcolors.ENDC}
"""

starting_text = f"""{bcolors.OKGREEN}
##############################################
#    Check Directory for Duplicate Sheets    #
############################################## {bcolors.ENDC}

This tool will check all PDF files within a directory for duplicate sheet numbers
based on the filename up to first Space character. 

{bcolors.OKGREEN}Instructions:{bcolors.ENDC}
1. Select Directory to be checked for duplicate sheets.
2. Results will be saved to a timestamped .csv file within the selected directory.
3. Press Enter to select another directory, or Q to quit.

{bcolors.OKGREEN}Using the results file:{bcolors.ENDC}
1. Open the results .csv file in Excel.
2. Select the entire range by selecting the top left cell, then pressing
   {bcolors.WARNING}CTRL-Right Arrow{bcolors.ENDC} then {bcolors.WARNING}CTRL-Down Arrow{bcolors.ENDC}.
3. Click the {bcolors.OKCYAN}Format as Table{bcolors.ENDC} button in Excel to make the data a table.
4. Filter the results using the Sheet Number, Current, or Single/Multiple Files columns.
5. Manually delete or archive the duplicate files as needed.
6. {bcolors.BOLD}(Optional){bcolors.ENDC} Enjoy your time saved.

For questions or troubleshooting, contact Dan Howard (dhoward@ballinger.com)

Ready? Let's go!

"""

if __name__ == "__main__":
    print(starting_text)
    while True:
        proceed_or_quit()
        process_directory_files()






    
