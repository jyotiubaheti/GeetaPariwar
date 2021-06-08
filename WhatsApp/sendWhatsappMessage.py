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

def getAlternateNames(df,nameCol,phoneCol):
    Names = df[nameCol].tolist()
    alternateNames = []
    for i,name in enumerate(Names):
        newSet = set(df[nameCol][df[phoneCol]==df[phoneCol].iloc[i]].to_list())
        alternateNames.append(list(newSet))
    df['Alternate Names'] = alternateNames
    return df

def findName(name,driver,Search):
    found= False
    Search.send_keys(name)
    time.sleep(1)
    try:
        element = driver.find_element_by_xpath("//span[@title='" + name + "']")
        element.click() 
        found = True
    except:
        Search.clear()
        # print ("Not Found:",name)
        found = False
    return found

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

def copyTextToClipBoard(text):
    clip.OpenClipboard()
    clip.EmptyClipboard()
    clip.SetClipboardText(text)
    clip.CloseClipboard()

def sendMessage(names=[],driver=None,body="",attach_jpg=None,bodyImage=""):
    # names = ["Nikhil JioB","Nikhil Jio"]
    Search = wait.until(EC.element_to_be_clickable((By.XPATH,"//div[@class='_2_1wd copyable-text selectable-text']")))
    time.sleep(2)

    isFound = False
    for name in names:
        time.sleep(1)
        isFound = findName(name,driver,Search)
        if isFound: break

    if not isFound: 
        print ("Not Found Names",names)
        return
    time.sleep(3)

    Message = driver.find_elements_by_xpath('//*[@id="main"]/footer/div[1]/div[2]/div/div[2]')[0]
    time.sleep(1)

    lines = body.split("\n")
    for line in lines:
        Message.send_keys(line)
        Message.send_keys(Keys.SHIFT + Keys.ENTER)
        
    time.sleep(1)
    Message.send_keys(Keys.ENTER)

    if attach_jpg is not None:
        copyImageToClipBoard(attach_jpg)
        time.sleep(1)
        Message.send_keys(Keys.CONTROL,"v")
        time.sleep(2)
        imageMessage = driver.find_elements_by_xpath("//div[@class='_2_1wd copyable-text selectable-text']")[0]
        time.sleep(1)
        imageMessage.send_keys(bodyImage)
        time.sleep(1)
    pyautogui.press('enter')
    

###############################################################################################
fileName = "GP June July Enrollment.xlsx"
sheetName = "Form Responses 1"
isRollNumber = True
rollNumberColIndex = ord("H") - ord("A")
nameColIndex = ord("C") - ord("A")
message = "rollNumberMessage.txt"
phoneColIndex = ord("E") - ord("A")
photoName = None
messageImage = "imageMessage.txt"
###############################################################################################


driver = webdriver.Chrome(executable_path="chromedriver")
wait = WebDriverWait(driver, 60)
print("Chrome opened successfully!")

Web_whatsapp = 'https://web.whatsapp.com/'  # To go to whatsapp web
driver.get(Web_whatsapp)
print("Accessing Whatsapp web")

df = pd.read_excel(fileName,sheet_name=sheetName)
input("Press Enter to continue...")

messageFile = open(message,"rb")
messageText = messageFile.read().decode("utf-8")
messageFile.close()

imageFile = open(messageImage,"rb")
imageText = imageFile.read().decode("utf-8")
imageFile.close()

rollCol = df.columns[rollNumberColIndex]
phoneCol = df.columns[phoneColIndex]
nameCol = df.columns[nameColIndex]

df = getAlternateNames(df,rollCol,phoneCol)
for i in range(df[df.columns[0]].count()):
    bodyMessage = messageText
    if isRollNumber: 
        bodyMessage = bodyMessage.replace("<>","*" + df[rollCol].iloc[i] + "*")
        bodyMessage = bodyMessage.replace("<N>","*" + df[nameCol].iloc[i] + "*")
    names = df['Alternate Names'].iloc[i]

    if names:
        sendMessage(names,driver,body=bodyMessage,attach_jpg=photoName,bodyImage=imageText)
        print ("Done:",names,i)
    else:
        print ("Error: names is empty")

