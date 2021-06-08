from selenium import webdriver
import time
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
import itertools
import pandas as pd
import os
import pyautogui
import win32clipboard as clip
import win32con
from io import BytesIO
from PIL import Image

def copyImageToClipBoard(image):
    image = Image.open(image)
    output = BytesIO()
    image.convert('RGB').save(output, 'BMP')
    data = output.getvalue()[14:]
    output.close()
    clip.OpenClipboard()
    clip.EmptyClipboard()
    clip.SetClipboardData(win32con.CF_DIB, data)
    clip.CloseClipboard() 

def sendMessage(number,body="",attach_jpg=None):
    # number = "8310258040"
    Search = wait.until(EC.element_to_be_clickable((By.XPATH,"//div[@class='_2_1wd copyable-text selectable-text']")))
    time.sleep(2)
    Search.send_keys(number)
    time.sleep(2)
    pyautogui.press("down")
    time.sleep(2)
    Message = driver.find_elements_by_xpath('//*[@id="main"]/footer/div[1]/div[2]/div/div[2]')[0]
    time.sleep(1)

    Message.send_keys(body)
    time.sleep(1)
    Message.send_keys(Keys.ENTER)

    if attach_jpg is not None:
        copyImageToClipBoard(attach_jpg)
        time.sleep(1)
        Message.send_keys(Keys.CONTROL,"v")
        time.sleep(2)
    pyautogui.press('enter')
    

driver = webdriver.Chrome(executable_path="chromedriver")
wait = WebDriverWait(driver, 60)
print("Chrome opened successfully!")

Web_whatsapp = 'https://web.whatsapp.com/'  # To go to whatsapp web
driver.get(Web_whatsapp)
print("Accessing Whatsapp web")

df = pd.read_excel("Summer Camp 2021.xlsx",sheet_name="Form Responses 1")
dfAttendance = pd.read_excel('zoomAttendance.xlsx')
attendents = dfAttendance['Roll Number'].to_list()
input("Press Enter to continue...")

for i,roll in enumerate(attendents):
    # if int(roll)<351: continue
    strRoll = "GPSC2021_" + str(roll).zfill(3)
    photoName = "certificates/" + strRoll + ".jpg"
    message = 'Greetings From Geeta Pariwar!! The previous certificates have some errors so kindly delete them. Attaching the corrected certificate.'
    phoneNo = df['Whatsapp Mobile Number'][df['Roll No.']==strRoll].to_list()
    if phoneNo:
        sendMessage(phoneNo[0],body=message,attach_jpg=photoName)
        print ("Done:",photoName,i)
    else:
        print ("Error: phoneNo not found",phoneNo)

