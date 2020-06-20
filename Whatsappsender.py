import time
start = time.time()
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains

from datetime import datetime
from time import sleep
import random
from random import randint

from pandas import DataFrame
import pandas as pd
import xlsxwriter

driver = webdriver.Firefox(executable_path="C:\webdrivers\geckodriver.exe")  #determine the firefox driver
driver.get("https://web.whatsapp.com") #open Whatsapp web
filename = r'contact.xlsx' #getting the excel file with the contact
df = pd.read_excel(filename)
names = list(df['name']) #turn the contact column to list

time_now = str(datetime.today())[0:16]
time_now_fixed = time_now.replace(":", ".")
status_sheet = "Sender " + time_now_fixed + ".xlsx"
workbook = xlsxwriter.Workbook(status_sheet)
sheet = workbook.add_worksheet("Status")

print("Enter your message after greeting: ")
print("When you're done, enter a single period on a line by itself.")
msg = [] #each line user input add as term to the msg list
while True:
    print("> ", end="")
    line = input()
    if line == ".":
        break
    msg.append(line)
    

count = int(input("times to send: "))
input("Enter anything after scanning QR code")

greeting_list = ["Hi ", "Hello ", "Good day "]
first_line = 0 #help determine if line is the fitst line of the message
row = 0
for i in names: #loop run for every contact in the excel file
    try:
        head, sep, tail = i.partition(' ')
        first_name = head #take only the first name of contact

        contact_seacrch_box = driver.find_element_by_xpath('//div[@data-tab="3"][@class="_2S1VP copyable-text selectable-text"]')
        contact_seacrch_box.send_keys(i) #selecting the contact search box and search for contact
        sleep(randint(2,3))
        user = driver.find_element_by_xpath('//span[@class="_1wjpf _3NFp9 _3FXB1"][@title="{}"]'.format(i))
        user.click() #clicking on the contact
        contact_seacrch_box .clear()

        for x in range(count): #run as many times user want to send message to each contact
            greeting = random.choice(greeting_list)
            msg_box = driver.find_element_by_xpath('//div[@data-tab="1"][@class="_2S1VP copyable-text selectable-text"]')
            msg_box.click() # clicking on the message box inside contact whatsapp
            for line in msg: #create in order to send muliline message and in order to tell when putt greeting + frist name
                ActionChains(driver).key_down(Keys.SHIFT).send_keys(Keys.ENTER).key_up(Keys.SHIFT).perform()
                if first_line > 0:
                    msg_box.send_keys(line)
                else:
                    msg_box.send_keys(greeting + first_name + ", " + line)
                    first_line =+1
            first_line = 0
            sleep(randint(1,2))
            send_button = driver.find_element_by_xpath('//span[@data-icon ="send"]')
            send_button.click() #clicking on the send button
            sleep(randint(1,2))
        sheet.write(row, 0, i)
        sheet.write(row, 1, "Succeed")
        row += 1
    except:
        contact_seacrch_box.clear()
        sheet.write(row, 0, i)
        sheet.write(row, 1, "Failed")
        sleep(randint(1, 2))
        row += 1


workbook.close()
end = time.time()
print (end - start) 
