# -*- coding: cp1255 -*-   # sets the coding to hebrew and not gibrish
import os
from time import strftime, localtime

import requests
import xlwt
from bs4 import BeautifulSoup

OUTPUT_FOLDER = u'C:\\Users\\ronke\\Desktop\\פרויקט גמר\\parsing\\'


community_number = 1
book = xlwt.Workbook(encoding="utf-8")  # Excel file initialize:
sheet1 = book.add_sheet(u'Camoni')  # cell_overwrite_ok=True


def get_community_details(file_to_parse):

    page = file_to_parse.content
    soup = BeautifulSoup(page, 'html.parser')

    communities = soup.findAll('div', attrs={'class': "nameCommu"})
    communities_num = soup.findAll('div', attrs={'class': "sumFriends"})
    community_friends = []
    for i in range(len(communities_num)):
        txt = communities_num[i].text
        txt = txt.split()
        num = txt[0].replace(',', '')
        community_friends.append(num)
    return communities, community_friends


def set_style(border_num, boldness):
    style = xlwt.XFStyle()
    font = xlwt.Font()
    font.bold = boldness
    style.font = font

    borders = xlwt.Borders()
    borders.left = border_num
    borders.right = border_num
    borders.top = border_num
    borders.bottom = border_num
    if border_num > 0:
        style.borders = borders

    return style


def main():
    global community_number

    style = set_style(6, True)

    sheet1.write(0, 0, u'Serial Number', style)
    sheet1.write(0, 1, u'Community Name', style)
    sheet1.write(0, 2, u'Members number', style)

    style = set_style(1, False)

    print("today:", strftime("%d-%b-%Y - %H:%M", localtime()))
    CamoniFileName = OUTPUT_FOLDER + "Camoni_COMMUNITIES" + "__" + strftime("%d-%b-%Y-%H;%M", localtime()) + ".xls"
    # ---

    url = "https://www.camoni.co.il/411785"
    try:
        response = requests.get(url)
    except Exception as e:
        print(str(e))

    communities, community_friends = get_community_details(response)
    for i in range(len(communities)):
        sheet1.write(i+1, 0, community_number, style)
        sheet1.write(i+1, 1, communities[i].text, style)
        sheet1.write(i+1, 2, community_friends[i], style)
        community_number += 1


    print(CamoniFileName)
    book.save(CamoniFileName)
    print("end")


if __name__ == '__main__':
    main()
