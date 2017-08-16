#Creates the system of folders which the insepction reports and detail reports will be organized into. First organizes by subcompany,
#then by individual project

import os, openpyxl

linuxPath = '/home/gmclaughlin/Python/Safety Project'
windowsPath = 'C:\\Users\\gmclaughlin\\Safety Project'

os.makedirs(linuxPath+'/Demo')
os.makedirs(linuxPath+'/JDC')
os.makedirs(linuxPath+'/NewRoads')
os.makedirs(linuxPath+'/SiteCrew')
os.makedirs(linuxPath+'/Other')

wb = openpyxl.load_workbook('ProjectSummary.xlsx')
sheet = wb.active

homePath = linuxPath
demPath = linuxPath+'/Demo'
JDCPath = linuxPath+'/JDC'
newRoadsPath =  linuxPath+'/NewRoads'
sitePath = linuxPath+'/SiteCrew'
otherPath = linuxPath+'/Other'

def makeSheet(compPath, cellObj):       #use when back on windows
    os.chdir(compPath)
    projectPath = compPath+"/%s" % cellObj.value
    print(projectPath)
    os.mkdir(projectPath)
    os.chdir(projectPath)
    wb = openpyxl.Workbook()
    sheet = wb.active
    sheet.cell(row=1, column=1).value = "Inspection Date"
    sheet.cell(row=1, column=2).value = "System Date"
    sheet.cell(row=1, column=3).value = "Inspection Type"
    sheet.cell(row=1, column=4).value = "Project"
    sheet.cell(row=1, column=5).value = "Observer"
    sheet.cell(row=1, column=6).value = "Obs Count"
    sheet.cell(row=1, column=7).value = "Unsafe Count"
    sheet.cell(row=1, column=8).value = "All Corrected?" 
    wb.save(cellObj.value+"-Inspection Reports.xlsx")
    
for cellObj in sheet['A']:
    if cellObj.value.find("DEM") != -1:
        makeSheet(demPath, cellObj)

    elif cellObj.value.find("JDC") != -1:
        makeSheet(JDCPath, cellObj)

    elif cellObj.value.find("NEW") != -1:
        makeSheet(newRoadsPath, cellObj)

    elif cellObj.value.find("Site") != -1:
        makeSheet(sitePath, cellObj)

    else:
        makeSheet(otherPath, cellObj)
