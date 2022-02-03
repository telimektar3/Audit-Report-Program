# Audit Report v. 0.1
# Timothy Goode (telimektar3)

# Batch reads and processes an auditing tool to determine the average audit score by discipline, and then outputs 
# a file that includes the averages by discipline.
from tkinter import Tk, filedialog
import openpyxl
import os
from datetime import datetime as dt
from openpyxl import Workbook

# Creates filename including todays date
mask = '%d%m%Y'
dte = dt.now().strftime(mask)
fname = "/Audit_Averages_Report_{}.xlsx".format(dte)
ffname = "Audit_Averages_Report_{}.xlsx".format(dte)

# Selects a folder containing the audit files necessary to be read using a visual interface
root = Tk() # pointing root to Tk() to use it as Tk() in program.
root.withdraw() # Hides small tkinter window.
root.attributes('-topmost', True) # Opened windows will be active. above all windows despite of selection.
open_folder = filedialog.askdirectory() # Returns opened path as str

    # print(open_folder) # testing that it takes in the selected folder

# Creates list of each Excel file in the selected folder location
files_to_rip = []
for filename in os.listdir(open_folder):
    if filename.endswith(".xlsx") and filename != ffname: # excludes non-.xlsx files and a generated report file
        f = os.path.join(open_folder, filename)
        # checking if it is a file
        if os.path.isfile(f):
            files_to_rip.append(f)
    # print(files_to_rip)

# Validates and imports the desired data from the list of files in the folder and stores in appropriate discipline
social_work = []
activity_therapy = []
ics = []
transitional_services = []
forensic_counseling = []
files_processed = 0
social_work_incomplete = []
activity_incomplete = []
ics_incomplete = []
transition_incomplete = []
forensic_incomplete = []



for file in files_to_rip: # looks at each file in the folder
    files_processed += 1
    wb = openpyxl.load_workbook(filename = file, data_only = True) # open the file
    ws = wb['AUDIT TOOL'] # select the necessary sheet
    if ws['E241'].value != "Incomplete":
        social_work.append(ws['E241'].value) # append the necessary values to social work
    else:
        social_work_incomplete.append(file)
    if ws['E242'].value != "Incomplete":
        activity_therapy.append(ws['E242'].value) # append values to activity therapy
    else:
        activity_incomplete.append(file)
    if ws['E243'].value != "Incomplete":
        ics.append(ws['E243'].value) # append values to ICS
    else:
        ics_incomplete.append(file)
    if ws['E244'].value != "Incomplete":
        transitional_services.append(ws['E244'].value) # append values to Transitional Services
    else: 
        transition_incomplete.append(file)
    if ws['E245'].value != "Incomplete":
        forensic_counseling.append(ws['E245'].value)
    else:
        forensic_incomplete.append(file)


    
# Maths the data into averages by discipline
# get sum of all numbers in lists
social_sum = sum(social_work)
activity_sum = sum(activity_therapy)
ics_sum = sum(ics)
transition_sum = sum(transitional_services)
forensic_sum = sum(forensic_counseling)
# get number of numbers in list
social_length = len(social_work)
activity_length = len(activity_therapy)
ics_length = len(ics)
transition_length = len(transitional_services)
forensic_length = len(forensic_counseling)
# get average percentage for each discipline
social_avg = (social_sum / social_length) * 100
activity_avg = (activity_sum / activity_length) * 100
ics_avg = (ics_sum / ics_length) * 100
transition_avg = (transition_sum / transition_length) * 100
forensic_avg = (forensic_sum / forensic_length) * 100




# Outputs the data into a new Excel Spreadsheet in the previously selected folder
wb = Workbook()
dest_filename = open_folder + fname
ws = wb.active
ws.title = "Avg Audit Scores"

# assign names to needed cells
ws['A2'] = "Social Work"
ws['B2'] = social_avg
ws['A3'] = "Activity Therapy"
ws['B3'] = activity_avg
ws['A4'] = "Psychology/Counseling"
ws['B4'] = ics_avg
ws['A5'] = "Transitional Services"
ws['B5'] = transition_avg
ws['A6'] = "Forensic Counseling"
ws['B6'] = forensic_avg
ws['C2'] = "Files Processed"
ws['D2'] = files_processed
ws['F2'] = "Social Work files incomplete:"
ws['G2'] = str(social_work_incomplete)
ws['F3'] = "Activity Therapy files incomplete:"
ws['G3'] = str(activity_incomplete)
ws['F4'] = "Psychology/Counseling files incomplete:"
ws['G4'] = str(ics_incomplete)
ws['F5'] = "Transitional Services files incomplete:"
ws['G5'] = str(transition_incomplete)
ws['F6'] = "Forensic Counseling files incomplete:"
ws['G6'] = str(forensic_incomplete)

# Saves the workbook
wb.save(open_folder + fname)


# Pop-up stating that the operation has been carried out with "Ok" box
#  NO. Could implement later.