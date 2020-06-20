from selenium.webdriver.firefox.options import Options as FirefoxOptions
from selenium import webdriver
import xlsxwriter
import sys
import os
from time import sleep

def resource_path(relative_path):
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.dirname(__file__)
    return os.path.join(base_path, relative_path)



options = FirefoxOptions()
driver = webdriver.Firefox(options=options, executable_path=resource_path('./driver/geckodriver.exe'))
# driver = webdriver.Firefox(options=options, executable_path="C:\webdrivers\geckodriver.exe")
event_id = input('Enter event id: ')
excel_name = input('Enter excel name: ')
email = input('Enter your email: ')
password = input('Enter your password: ')
driver.get('https://www.facebook.com/events/' + event_id)

status_sheet = excel_name + ".xlsx"
workbook = xlsxwriter.Workbook(status_sheet)
sheet = workbook.add_worksheet("Going")
sheet2 = workbook.add_worksheet("Interested")

sheet.write(0, 0, 'Name')
sheet.write(0, 1, 'Link')
sleep(2)
email_id = driver.find_element_by_id('email')
email_id.send_keys(email)
password_id = driver.find_element_by_id('pass')
password_id.send_keys(password)
log = driver.find_element_by_id('loginbutton')
log.click()

input("Enter any key after you done scroll going list")

person = driver.find_elements_by_xpath('//span[@class="_h24 _h25"]')

sheet.write(0, 0, 'Going')
counter = 1
for elem in person:
    link_person = elem.find_element_by_xpath('..')
    sheet.write(counter, 0, elem.text)
    sheet.write(counter, 1, link_person.get_attribute("href"))
    counter += 1

input("Enter any key after you done scroll interested list")

person = driver.find_elements_by_xpath('//span[@class="_h24 _h25"]')
sheet2.write(0, 0, 'Interested')
counter = 1
for elem in person:
    link_person = elem.find_element_by_xpath('..')
    sheet2.write(counter, 0, elem.text)
    sheet2.write(counter, 1, link_person.get_attribute("href"))
    counter += 1

workbook.close()
