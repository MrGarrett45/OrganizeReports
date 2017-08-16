import os, openpyxl, time, shutil
from selenium import webdriver
from selenium.webdriver.common.keys import Keys

wb = openpyxl.load_workbook('ProjectSummary.xlsx')
sheet = wb.active

browser = webdriver.Firefox()
browser.get('https://safetynet.predictivesolutions.com/CRMApp/default_login.jsp?loginZoneID=10459&originalHostName=jdc.predictivesolutions.com')

userElem = browser.find_element_by_id('username')
userElem.send_keys('temp')
passElem = browser.find_element_by_id('password')
passElem.send_keys('temp')
passElem.submit()

time.sleep(3)
linkElem = browser.find_element_by_link_text('Reports')
linkElem.click()
time.sleep(2)
linkElem = browser.find_element_by_link_text('Detail Report')
linkElem.click()
time.sleep(4)

def pdfToFolder(projectName):
    os.chdir('/home/gmclaughlin/Downloads')
    if projectName.find("DEM") != -1:
        shutil.move('/home/gmclaughlin/Downloads/Detail Report - Basic.pdf','/home/gmclaughlin/Python/Safety Project/Demo/%s/%s-Detail Report.pdf' % (projectName, projectName))

    elif projectName.find("JDC") != -1:
        shutil.move('/home/gmclaughlin/Downloads/Detail Report - Basic.pdf','/home/gmclaughlin/Python/Safety Project/JDC/%s/%s-Detail Report.pdf' % (projectName, projectName))

    elif projectName.find("NEW") != -1:
        shutil.move('/home/gmclaughlin/Downloads/Detail Report - Basic.pdf','/home/gmclaughlin/Python/Safety Project/NewRoads/%s/%s-Detail Report.pdf' % (projectName, projectName))

    elif projectName.find("Site") != -1:
        shutil.move('/home/gmclaughlin/Downloads/Detail Report - Basic.pdf','/home/gmclaughlin/Python/Safety Project/SiteCrew/%s/%s-Detail Report.pdf' % (projectName, projectName))

    else:
        shutil.move('/home/gmclaughlin/Downloads/Detail Report - Basic.pdf','/home/gmclaughlin/Python/Safety Project/Other/%s/%s-Detail Report.pdf' % (projectName, projectName))

finsihedFlag = False
addValue = 0
counter = 0
for cellObj in sheet['A']:
    if cellObj.value != 'Project' and cellObj.value != 'JDC-Winchester HS Enabling (CONSIG':

        linkElem = browser.find_element_by_name('clear') #clear existing settings
        linkElem.click()
        time.sleep(4)

        linkElem = browser.find_element_by_name('showSafeAndUnsafeDetails') #select all reports
        linkElem.click()
        time.sleep(1)

        linkElem = browser.find_element_by_name('showImages') #show images in reports
        linkElem.click()
        time.sleep(1)

        linkElem = browser.find_element_by_name('datePickerRadio')
        linkElem.click()
        time.sleep(1)

        projectElem = browser.find_elements_by_xpath("//input[@type='text']") #find and use text fields
        print(cellObj.value)
        #projectElem = browser.find_element_by_xpath("//input[4]")
        #time.sleep(2)
        #projectElem[5+addValue].clear()
        projectElem[5+addValue].send_keys('01/01/2010')
        time.sleep(1)
        #projectElem[6+addValue].clear()
        projectElem[6+addValue].send_keys('08/15/2017')
        time.sleep(1)
        projectElem[8+addValue].clear()                   #this is the project name box
        projectElem[8+addValue].send_keys(cellObj.value)
        time.sleep(1)
        projectElem[8+addValue].send_keys(Keys.ENTER)
        time.sleep(3)

        linkElem = browser.find_element_by_xpath("//input[@type='submit']") #submit request for report
        linkElem.click()
        time.sleep(10)

        linkElem = browser.find_element_by_name('pdf') #download as PDF
        linkElem.click()
        time.sleep(70)
        addValue = 1

        pdfToFolder(cellObj.value)

    counter = counter + 1
