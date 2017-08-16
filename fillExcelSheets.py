#Takes an unordered list of 5000+ inspections and puts them into the correct project folder that was created by makeFolders.py

import os, openpyxl

wb = openpyxl.load_workbook('allInspections.xlsx')
sheet = wb.active

linuxPath = '/home/gmclaughlin/Python/Safety Project'
windowsPath = 'C:\\Users\\gmclaughlin\\Safety Project'

for cellObj in sheet['D']:
    for folderName, subfolders, filenames in os.walk(linuxPath):
        #print('The current folder is ' + folderName)
        os.chdir(folderName)
        #for subfolder in subfolders:
            #print('SUBFOLDER OF ' + folderName + ': ' + subfolder)
            #os.chdir("/home/gmclaughlin/Python/Safety Project/"+subfolder)
        for filename in filenames:
            #print('FILE INSIDE ' + folderName + ': '+ filename)
            if cellObj.value+"-Inspection Reports.xlsx" == filename:
                
                if cellObj.value.find("DEM") != -1:
                    os.chdir(linuxPath+"/Demo/"+cellObj.value)
                elif cellObj.value.find("JDC") != -1:
                    os.chdir(linuxPath+"/JDC/"+cellObj.value)
                elif cellObj.value.find("NEW") != -1:
                    os.chdir(linuxPath+"/NewRoads/"+cellObj.value)
                elif cellObj.value.find("Site") != -1:
                    os.chdir(linuxPath+"/SiteCrew/"+cellObj.value)
                else:
                    os.chdir(linuxPath+"/Other/"+cellObj.value)
                print(os.getcwd())
                print("found match: %s and %s" % (cellObj.value, filename))
                tempBook = openpyxl.load_workbook(cellObj.value+"-Inspection Reports.xlsx")
                tempSheet = tempBook.active
                tempSheet.cell(row=((tempSheet.max_row)+1), column=1).value = sheet.cell(row=cellObj.row, column=1).value
                tempSheet.cell(row=((tempSheet.max_row)), column=2).value = sheet.cell(row=cellObj.row, column=2).value
                tempSheet.cell(row=((tempSheet.max_row)), column=3).value = sheet.cell(row=cellObj.row, column=3).value
                tempSheet.cell(row=((tempSheet.max_row)), column=4).value = sheet.cell(row=cellObj.row, column=4).value
                tempSheet.cell(row=((tempSheet.max_row)), column=5).value = sheet.cell(row=cellObj.row, column=5).value
                tempSheet.cell(row=((tempSheet.max_row)), column=6).value = sheet.cell(row=cellObj.row, column=6).value
                tempSheet.cell(row=((tempSheet.max_row)), column=7).value = sheet.cell(row=cellObj.row, column=7).value
                tempSheet.cell(row=((tempSheet.max_row)), column=8).value = sheet.cell(row=cellObj.row, column=8).value
                tempBook.save(cellObj.value+"-Inspection Reports.xlsx")
        #print('')
