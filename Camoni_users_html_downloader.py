# -*- coding: cp1255 -*-   # sets the coding to hebrew and not gibrish

from bs4 import BeautifulSoup
import os
import time
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from datetime import datetime

from webdriver_manager.chrome import ChromeDriverManager

CHROME_DRIVER_PATH = "C:\\Users\\ronke\\Desktop\\פרויקט גמר\\chromedriver.exe"
DEFAULT_URL = "https://www.camoni.co.il/411788/"
DOWNLOAD_PATH = "C:\\Users\\ronke\\Desktop\\פרויקט גמר\\parsing\\users_download_with_selenium\\"
SOURCE_FILE_NAMES = "C:\\Users\\ronke\\Desktop\\פרויקט גמר\\parsing\\411788"
LOG_FILE_PATH = DOWNLOAD_PATH + "log_file_users_download_" + str(datetime.now().strftime("%d-%m-%y")) + ".txt"
users_count = 0


def double_print(my_str):
    global log_file
    print(my_str)
    log_file.write(my_str)


log_file = open(LOG_FILE_PATH, 'w')
double_print("\n" + str(datetime.now().strftime("%d/%m/%y %H:%M:%S")) + " - Starting getting users file")
double_print("-----------------------------------------------------------------------------------------------------------------")

chrome_options = Options()
chrome_options.add_argument("--headless")
#browser = webdriver.Chrome(executable_path=CHROME_DRIVER_PATH, options=chrome_options)
browser = webdriver.Chrome(ChromeDriverManager().install(), options=chrome_options)

user_files = os.listdir(SOURCE_FILE_NAMES)

for file_name in user_files[8148:]:
    try:
        url = DEFAULT_URL + file_name
        browser.get(url)
        content = browser.page_source

        with open(DOWNLOAD_PATH + file_name + ".html", "w", encoding='utf-8') as file:
            file.write(str(content))
            double_print(str(datetime.now().strftime("%d/%m/%y - %H:%M:%S")) + " - created successfully file: " + file_name)

        users_count += 1
        time.sleep(0.5)

    except Exception as e:
        double_print(file_name + " failed because: " + str(e))


browser.close()

double_print("-----------------------------------------------------------------------------------------------------------------")
double_print("\n" + str(datetime.now().strftime("%d/%m/%y - %H:%M:%S")) + " - Downloaded " + str(users_count) + " files!")
log_file.close()