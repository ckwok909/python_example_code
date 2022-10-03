from datetime import date
from datetime import datetime
import PyPDF2
import textract
import numpy
import pyperclip
import pyautogui
import time
import re
from difflib import SequenceMatcher
import os
from pdf2image import convert_from_path
import pytesseract
from pytesseract import Output
import ntpath
import shutil
import cv2
from PIL import Image
from tabula import read_pdf
import math
import json
from selenium import webdriver
from selenium.webdriver import Chrome
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.by import By
import psutil
from pywinauto import Desktop, Application
import win32com.client
import win32gui
from pathlib import Path
import zipfile


# Image read

image_path = "C:\\Users\\Jason\\PycharmProjects\\Mansfield\\request_file\\activate_image_selectors.png"
image_path_1 = "C:\\Users\\Jason\\Documents\\test.png"
crop_length = [250, 265, 500, 300]
image = Image.open(image_path)
image = image.crop((crop_length[0], crop_length[1], crop_length[2], crop_length[3]))
image.save(image_path_1)
time.sleep(1)
original_image = cv2.imread(image_path_1)
gray_image = cv2.cvtColor(original_image, cv2.COLOR_BGR2GRAY)
cv2.imwrite(image_path_1, gray_image)
time.sleep(3)
pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'
if os.path.exists(image_path_1):
    image_file = image_path_1
    if str(image_file).endswith(".png"):
        img = cv2.imread(image_file)
        data = pytesseract.image_to_string(img, output_type=Output.DICT)
print(data)
address = "8\nHyde Green\nStevenage\nHertfordshire\nSG2 9XU"
address_split = address.split("\n")
data = data['text'].replace("\n\n", "\n").strip()
found_address = data.split("\n")
count_start = False
count = 0
for add in found_address:
    similar = SequenceMatcher(None, add.lower(), address.lower()).ratio()
    if not count_start:
        if similar >= 0.3:
            count_start = True
    if count_start:
        count += 1
    if similar > 0.6 and add.startswith(address_split[0]):
        break

# PDF
"""
pdfFileObj = open('C:\\Users\\Jason\\Documents\\Customer\\Mansfield\\Broker Details.pdf', 'rb')

# creating a pdf reader object
pdfReader = PyPDF2.PdfFileReader(pdfFileObj)

# printing number of pages in pdf file
print(pdfReader.numPages)

# creating a page object
pageObj = pdfReader.getPage(0)

# extracting text from page
print(pageObj.extractText())

# closing the pdf file object
pdfFileObj.close()
"""
# Read PDF
"""
text = textract.process('C:\\Users\\Jason\\Documents\\Customer\\Mansfield\\Broker Details.pdf')
print(str(text))
"""
# RE
"""
texts = ""ing to be completed via a
Limited Company?
Yes
Name of Business Moulton Houses Limited
Registered address 168
Church Road
Hove
East Sussex
BN3 2DL
Correspondence Address 13
Chatsworth Road
Eccles
Manchester
Lancashire
M30 9DZ
Standard Industrial Classification Code
(SIC Code)
68100
Business Telephone Number 07515534358
Registration Number 07283513
Date of Incorporation 14/06/2010
Please provide names of any persons
involved within the business including full
names, position in business
(Shareholder, Director, Secretary) and
their percentage shareholding.
Alexander Moulton - Director - 50% Shareholder Richard Moulton - Director - 25%
Shareholder Stephen Moulton - Director - 25% Shareholder
How long has the business been
trading?
12 years
Has the business created any fixed
and/or floating charges and/or
debentures? Please provide full details
Outstanding: 2 charges over properties owned by the company - 1 Low Ferney Green,
LA23 3ES (holiday let) --- £282,000 @ 2% I/O - 45 Lawrence Road, BN3 5QE (BTL) --
- £376,000 @ 2% I/O Charges are registered to clients' own company, Moulton
Consortium Ltd, for IHT tax efficiency To be redeemed: 1 charge over property owned
but being sold (completion imminent - buyer has mortgage offer etc) - 9 Marlborough
Road, M30 9DE --- £115,900 @ 3.59% I/O Kensington charge to be redeemed on
sale ALL KEYED TO PORTFOLIO SECTION
Intermediary Fee
""
text = re.findall(r"Limited Company[0-9a-zA-Z\s\+\&\-\?\£\r\n\\\/\(\)\'\"\!\@\%\,\.\:\;]+", texts)
print(text)
"""
# Read PDF
"""
os.startfile("C:\\Users\\Jason\\Documents\\Customer\\Mansfield\\Residential Cases Old\\Residential Cases v1\\Abdin & Abdullah 20526983 02.11.2021\\Application Form\\App Form.pdf")
time.sleep(2)
pyautogui.doubleClick(874, 433)
time.sleep(2)
pyautogui.hotkey('ctrl', 'a')
time.sleep(2)
pyautogui.hotkey('ctrl', 'c')
time.sleep(5)
all_texts = pyperclip.paste()
texts = []
for process in psutil.process_iter():
    if process.name() == "Acrobat.exe":
        os.system("TASKKILL /F /IM Acrobat.exe")
number_of_applicant_arr = re.findall("Type of applicant", all_texts)
text_check = ["Application Form Details"]
match = ["Mortgage Application", "Application Information"]
for x, num in enumerate(number_of_applicant_arr):
    text_check.append("Personal Details")
    text_check.append("Employment Details")
    text_check.append("Income")
    text_check.append("Credit Declarations")
    match.append("Personal Details")
    match.append("Current Employment")
    match.append("Income Type")
    match.append("Credit Declarations")
    if x == 0:
        text_check.append("Property To Be Mortgaged")
        match.append("Description")
for z, txt in enumerate(text_check):
    data = str(all_texts).split(txt, 1)
    if len(data) >= 2:
        all_texts = data[1]
        texts.append(data[0])
    if z == len(text_check) - 1:
        texts.append(data[1])
matches = []
for n, txt in enumerate(texts):
    if txt.strip().startswith(match[n]):
        matches.append(True)
    else:
        matches.append(False)
print(matches)
for txt in texts:
    print(txt)
"""
# Windows title
"""
for x in pyautogui.getAllWindows():  
    print(x.title)   
"""
# Kill Process
"""
for process in psutil.process_iter():
    if process.name() == "Acrobat.exe":
        os.system("TASKKILL /F /IM Acrobat.exe")
"""
# Others
"""
app = Application(backend="uia").connect(title="File Explorer")
main_dlg = app.window(class_name='CabinetWClass')
main_dlg.set_focus()
time.sleep(1)
main_dlg.print_control_identifiers()
list_view = main_dlg.child_window(title="Downloads", control_type="TreeItem").click_input()
time.sleep(2)
"""
"""
date_moved = "05/03/1900"
today_date = date(date.today().year, date.today().month, date.today().day)
if date_moved and date_moved != "None":
    date_moved_split = date_moved.split('/')
    check_date = date(int(date_moved_split[2]), int(date_moved_split[1]), int(date_moved_split[0]))
    correct_month = today_date.month - check_date.month
    correct_year = today_date.year - check_date.year
    if correct_month < 0:
        correct_year -= 1
        correct_month = 12 - abs(correct_month)

print(correct_year, correct_month)
"""
text = """
Sort by: AR Relationship descending
Effective from
Sort by: Effective from descending
HL Partnership Limited 303397 Full 10 May 2019
Results per page
10
20
50
100
"""