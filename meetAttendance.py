# NOTE: This code works only with google chrome
# Add chrome browser to your system's path
# Run chrome.cmd file from this repository to start chrome in debugging mode

from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from collections import defaultdict
import pandas as pd
import re
import datetime
import os
from time import sleep

opt = Options()
opt.add_experimental_option("debuggerAddress", "127.0.0.1:9222")
driverLoc = ''
if os.path.exists('chromedriverloc.txt'):
    with open('chromedriverloc.txt', 'r') as f:
        driverLoc = f.readline()
else:
    driverLoc = input("Enter the path of chrome driver (This is only for the first time, details entered will \
be saved in chromedriverloc.txt, do not delete the file) \n>>> ")
    with open('chromedriverloc.txt', 'a') as f:
        f.write(driverLoc)
driver = webdriver.Chrome(driverLoc, options=opt)

date = datetime.date.today()
d = date.strftime('%d-%m-%Y')
date = date.strftime('%d/%m/%Y')


lis = defaultdict(list)

try:
    driver.find_element_by_xpath(
        r'//*[@id="ow3"]/div[1]/div/div[8]/div[3]/div[6]/div[3]/div/div[2]/div[3]/span/span'
        ).click() # Chat button 
    print("\nChats open")
    sleep(2)

except:
    try:
        driver.find_element_by_xpath(
            r'//*[@id="ow3"]/div[1]/div/div[8]/div[3]/div[3]/div/div[2]/div[2]/div[1]/div[2]'
            ).click() # Chat button 
        sleep(2)
    except:
        pass
    
    print('\nChats open')

finally:
    x = driver.find_elements_by_class_name('GDhqjd') # Fetches container of the participants
    for i in x:
        lis['name'].append(i.find_element_by_class_name('YTbUzc').text) # Fetches the name of the participant 
        lis['replies'].append(i.find_element_by_class_name('Zmm6We').text) # Fetches the chats sent by that participant in that container

for i, j in enumerate(lis['replies']):
    x = j.split("\n")
    x = ' '.join(x)
    lis['replies'][i] = x

texpath = input(
    "\nEnter a path to store the chat history, please add back slash at the end (Example: D:/anoop/py/):\n>>> "
    )
with open(texpath + f'Meet_Chat_History_{d}.txt', 'a') as f:
    for i, j in zip(lis['name'], lis['replies']):
        f.write(i + ':-- ' + j + '\n')

att = []
secretKey = input("\nEnter the secret key given to students: \n>>> ")

with open(texpath + f'Meet_Chat_History_{d}.txt', 'r') as f:
    x = f.readlines()

attPath = input(
    '\nEnter the path of attendance sheet (File must be of xlsx type only. Example: "D:/anoop/py/test2.xlsx"):\n>>> '
    )
print(f"\nReading xlsx file from {attPath}")

df = pd.read_excel(attPath, engine='openpyxl')
df.drop(df.columns[df.columns.str.contains('unnamed', case=False)], axis=1, inplace=True)
print(f'\nCreating a column named "{date}"..')
df[f'{date}'] = ['ab']*len(df['name'])

for i in x:
    if secretKey in i:
        x = re.split(':-- ', i)
        att.append(x[0])

print("\nParticipants identified with secret key (Note: This count is with repetition): ", len(att))
print("\nMarking attendance...")

for i, j in enumerate(df['name']):
    for z in att:
        if j.lower() in z.lower():
            df[f'{date}'][i] = 'present'

df.to_excel(attPath)

print(f"\nAttendance marked and saved in {attPath}")
print('\nTotal participants attended today (Note:This count is without repetition): ',\
      sum([1 for i in df[f'{date}'] if i.lower() == 'present']))
driver.close()
x = input("\npress enter to exit...")
