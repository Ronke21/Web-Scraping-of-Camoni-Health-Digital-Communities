# -*- coding: cp1255 -*-   # sets the coding to hebrew and not gibrish
import requests
import time
from datetime import datetime

CHROME_DRIVER_PATH = "C:\\Users\\ronke\\Desktop\\פרויקט גמר\\chromedriver.exe"
DEFAULT_URL = "https://www.camoni.co.il/411804/"
DOWNLOAD_PATH = "C:\\Users\\ronke\\Desktop\\פרויקט גמר\\parsing\\posts_download_with_selenium\\"
LOG_FILE_PATH = DOWNLOAD_PATH + "log_file_posts_download_" + str(datetime.now().strftime("%d-%m-%y")) + ".txt"
posts_count = 0


def double_print(my_str):
    global log_file
    print(my_str)
    log_file.write(my_str + "\n")


log_file = open(LOG_FILE_PATH, 'a')
double_print("\n" + str(datetime.now().strftime("%d/%m/%y %H:%M:%S")) + " - Starting getting users file")
double_print(
    "-----------------------------------------------------------------------------------------------------------------")

"""
525000-571066 V
10000-50000 V - check if real posts
100000-140000 V
500000-525000
"""
FROM = 570000
TO = 580000
for page_number in range(FROM, TO):  # range(FROM, TO):
    url = DEFAULT_URL + str(page_number)

    try:
        response = requests.get(url)
    except Exception as e:
        double_print(str(page_number) + " failed because: " + str(e))
        continue

    if response.status_code == 200:
        try:
            with open(DOWNLOAD_PATH + str(page_number) + ".html", "wb") as file:  # save as
                file.write(response.content)
        except Exception as e:
            double_print(str(page_number) + " failed because: " + str(e))
            continue

        double_print(
            str(datetime.now().strftime("%d/%m/%y - %H:%M:%S")) + " - created successfully file: " + str(page_number))
        posts_count += 1
    else:
        double_print(str(datetime.now().strftime("%d/%m/%y - %H:%M:%S")) + " - failed file # " + str(
            page_number) + " status code: " + str(response.status_code))

    # time.sleep(0.5)
    if (page_number - FROM) % 100 == 0:  # save log file every 100 iterations
        log_file.flush()

double_print(
    "-----------------------------------------------------------------------------------------------------------------")
double_print(
    "\n" + str(datetime.now().strftime("%d/%m/%y - %H:%M:%S")) + " - Downloaded " + str(posts_count) + " files!")
log_file.close()
