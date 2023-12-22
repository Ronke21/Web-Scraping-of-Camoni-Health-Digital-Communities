# -*- coding: cp1255 -*-   # sets the coding to hebrew and not gibrish
import requests
from datetime import datetime
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager

CHROME_DRIVER_PATH = "C:\\Users\\ronke\\Desktop\\פרויקט גמר\\chromedriver.exe"
DEFAULT_URL = "https://www.camoni.co.il/411804/"
DOWNLOAD_PATH  ="C:\\Users\\ronke\\Desktop\\פרויקט גמר\\parsing\\new_posts\\"
LOG_FILE_PATH = DOWNLOAD_PATH + "\log_file_posts_download_" + str(datetime.now().strftime("%d-%m-%y")) + ".txt"
posts_count = 0


def double_print(my_str):
    global log_file
    print(my_str)
    log_file.write(my_str + "\n")


log_file = open(LOG_FILE_PATH, 'a')
double_print("\n" + str(datetime.now().strftime("%d/%m/%y %H:%M:%S")) + " - Starting getting users file")
double_print(
    "-----------------------------------------------------------------------------------------------------------------")

chrome_options = Options()
chrome_options.add_argument("--headless")
browser = webdriver.Chrome(ChromeDriverManager().install(), options=chrome_options)

FROM = 18790
TO = 580000
for page_number in range(FROM, TO):  # range(FROM, TO):
    url = DEFAULT_URL + str(page_number)

    try:
        browser.get(url)
        response = browser.page_source
    except Exception as e:
        double_print(str(page_number) + " failed because: " + str(e))
        continue

    if "p404" not in response:
        try:
            with open(DOWNLOAD_PATH + str(page_number) + ".html", "wb") as file:  # save as
                file.write(response.encode())
        except Exception as e:
            double_print(str(page_number) + " failed because: " + str(e))
            continue

        double_print(
            str(datetime.now().strftime("%d/%m/%y - %H:%M:%S")) + " - created successfully file: " + str(page_number))
        posts_count += 1
    else:
        double_print(str(datetime.now().strftime("%d/%m/%y - %H:%M:%S")) + " - failed file # " + str(
            page_number) + " status code: 404")

    # time.sleep(0.5)
    if ((page_number-FROM) % 100) == 0:
        log_file.flush()

double_print(
    "-----------------------------------------------------------------------------------------------------------------")
double_print(
    "\n" + str(datetime.now().strftime("%d/%m/%y - %H:%M:%S")) + " - Downloaded " + str(posts_count) + " files!")
log_file.close()
