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
time.sleep(1)
linkElem = browser.find_element_by_link_text('Detail Report')
linkElem.click()
time.sleep(1)
linkElem = browser.find_element_by_name('showSafeAndUnsafeDetails')
linkElem.click()
time.sleep(1)

projectElem = browser.find_element_by_id('label_1502743776874_4313')
projectElem.send_keys('test')
