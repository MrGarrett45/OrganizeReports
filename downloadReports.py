import os, openpyxl, time
from selenium import webdriver

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
time.sleep(2)
linkElem = browser.find_element_by_name('showSafeAndUnsafeDetails')
linkElem.click()
time.sleep(2)

projectElem = browser.find_elements_by_xpath("//input[@type='text']")
print(type(projectElem))
#projectElem = browser.find_element_by_xpath("//input[4]")
projectElem[5].clear()
projectElem[5].send_keys('01/01/2010')
projectElem[6].clear()
projectElem[6].send_keys('08/15/2017')
projectElem[8].clear()                   #this is the project name box
projectElem[8].send_keys('test')
