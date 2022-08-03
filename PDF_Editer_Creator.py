import time
import pandas as pd
from datetime import date
from datetime import timedelta
import win32api, win32con
from pynput.keyboard import Key, Controller
import random
import numpy as np



XML_DIR = "./Company_Info.xlsx"
SLEEP_TIME_BETWEEN_TYPETHIS = 0.05
SLEEP_TIME_BETWEEN_LETTER = 0.01


def click(x,y):
    win32api.SetCursorPos((x,y))
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN,x,y,0,0)
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP,x,y,0,0)
    time.sleep(0.25)

# Type's in string
def typeThis(toType):
    time.sleep(SLEEP_TIME_BETWEEN_TYPETHIS)
    for letter in toType:
        keyboard.type(letter)
        time.sleep(SLEEP_TIME_BETWEEN_LETTER)
    time.sleep(SLEEP_TIME_BETWEEN_TYPETHIS)
    # keyboard.press(Key.enter)
    # keyboard.release(Key.enter)
    # time.sleep(SLEEP_TIME_BETWEEN_TYPETHIS)



# Start
# =============================================================================
keyboard = Controller()

df = pd.read_excel(XML_DIR)

company_list = []
website_list = []
position_list = []

for company in df['Company']:
    company_list.append(company)
    website_list.append(company+'.com')

random.shuffle(company_list)

for position in df['Position']:
    if 'nan' in str(position):
        break
    position_list.append(position)

random.shuffle(position_list)

company_info_matrix = np.array([['Company', 'Website', 'Position']])

for company in company_list:
    for position in position_list:
        to_append = np.array([[company, company+'.com', position]])
        company_info_matrix = np.append(company_info_matrix, to_append, axis=0)



company_info_matrix = np.delete(company_info_matrix, 0, axis=0)
np.random.shuffle(company_info_matrix)


date_to_start = date(2020, 1, 25)
date_now = date_to_start + timedelta(days=7)

for index in range(54):
    # file_explorer_icon_x = 1440
    # file_explorer_icon_y = 940
    
    # click(file_explorer_icon_x, file_explorer_icon_y)
    
    file_x = 1000
    file_y = 570
    
    click(file_x, file_y)
    click(file_x, file_y)
    time.sleep(1)
    
    date_location_x = 1020
    date_location_y = 275-55
    click(date_location_x, date_location_y)
    typeThis(date_now.strftime("%m/%d/%y"))
    
    cell_x = 520
    cell_y = 570-55
    
    for row in range(4):
        date_row = date_now - timedelta(days=5) + timedelta(days=row)
        for col in range(3):
            click(cell_x+150*col, cell_y+30*row)
            time.sleep(0.25)
            click(cell_x+150*col, cell_y+30*row)
            if col == 0:
                typeThis(date_row.strftime("%m/%d/%y"))
                time.sleep(0.25)
            else:
                typeThis(company_info_matrix[row, col-1])
                time.sleep(0.25)
        
    # Saves file
    # click(1580, 100)
    # -----------
    click(20, 40)
    click(20, 135)
    # -----------
    time.sleep(0.25)
    typeThis("UB-106-A--" + str(date_now) + ".pdf")
    # # Click save
    # click(1000, 660)
    # Click save
    click(525, 410)
    # # Click exit
    # click(1680, 20)
    click(1680, 15)
    time.sleep(0.5)
    
    for ii in range(4):
        company_info_matrix = np.delete(company_info_matrix, 0, axis=0)
    
    date_now = date_now + timedelta(days=7)


