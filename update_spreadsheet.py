# Author: Michael Cope
# Date: 3-31-20
# Purpose: Adds resume/job application files from a specified directory to a spreadsheet.

# Works with Ubuntu, but unsure if it works with Windows, Mac, or other OS.

# This File assumes that a worksheet already exists in the specified location and has the first row of the first and
# second column occupied with "Name", and "Date" respectively

# For this app to function correctly, there are two statements you must change to reflect where your job application
# folder is, and where the spreadsheet file is at. These lines have a comment above them stating this.

import openpyxl
import os
import re
from datetime import datetime


def main():
    # Runs the updating function
    update_spreadsheet()
    # Additional functionality could be added here as well, perhaps to backup the workbook to Dropbox or Google Drive,


if __name__ == '__main__':
    main()


def update_spreadsheet():

    # Here is the path for the spreadsheet file that will store our job applications history
    # FILL IN THIS LINE
    workbook_path = "../../PATH/TO/SPREADSHEET/GOES/HERE.xlsx"
    wb = openpyxl.load_workbook(workbook_path)

    # You can set the specific sheet that you want to save the data to or the last active sheet.
    # Only one of the two following lines should be uncommented.
    # job_app_sheet = wb["Sheet 4"]
    job_app_sheet = wb.active

    # grabs names of the resumes/applications that are already in the spreadsheet to prevent duplication, or rewriting
    # the same data multiple times.
    old_apps = []
    for i in range(1, job_app_sheet.max_row+1):
        old_apps.append(job_app_sheet.cell(row=i, column=1).value)

    # Removes the first cell as it will be occupied by the column title.
    old_apps = old_apps[1:]

    # Finds the first row that is empty, and sets index 1 past it to avoid overwriting previous entries.
    new_index = job_app_sheet.max_row + 1

    # Here is the path for the resume/job application directory.
    # FILL IN THIS LINE
    job_apps_path = "../../PATH/TO/JOB/APPLICATIONS/DIRECTORY/GOES/HERE"

    # Here is where the worksheet is updated.
    for app_name in os.listdir(job_apps_path):
        # First step is to clean the file names in the resume folder
        cleaned_app_name = file_extension_remover(app_name)

        # Next, each clean name is compared against the ones already in the excel sheet.
        if cleaned_app_name not in old_apps:

            # If it isn't in the spreadsheet, the cleaned name gets added, along with the last modified date and time.
            job_app_sheet.cell(row=new_index, column=1).value = cleaned_app_name

            full_path = os.path.join(job_apps_path, app_name)

            date_time = datetime.fromtimestamp(os.path.getmtime(full_path))
            human_time = date_time.strftime("%m/%d/%Y, %H:%M:%S")

            job_app_sheet.cell(row=new_index, column=2).value = human_time
            # Moves index of spreadsheet down to avoid overwriting.
            new_index += 1

    # At the end of the cycle, the worksheet is saved to preserve changes and is closed.
    wb.save(workbook_path)
    wb.close()


def file_extension_remover(file_name):
    # This util function removes the file extension from the file name if it is present. Otherwise the name is returned.

    # More specific patterns if you want to parse files with multiple "." in the name
    # pattern1 = re.compile(r'\.pdf')
    # pattern2 = re.compile(r'\.doc')
    # pattern3 = re.compile(r'\.docx')
    # pattern4 = re.compile(r'\.odt')

    # Searches for any periods followed by one or more chars
    pattern5 = re.compile(r'\..+')

    pattern_match = pattern5.search(file_name)

    # Checks if a match was found.
    # If so, it returns the file name up to the beginning of the match.
    if bool(pattern_match):
        return file_name[:pattern_match.span()[0]]

    else:
        return file_name


