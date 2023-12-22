# -*- coding: cp1255 -*-   # sets the coding to hebrew and not gibrish
from time import strftime, localtime
import xlwt
from bs4 import BeautifulSoup

OUTPUT_FOLDER = u'C:\\Users\\ronke\\Desktop\\פרויקט גמר\\parsing\\'

community_number = 1
book = xlwt.Workbook(encoding="utf-8")  # Excel file initialize:
sheet1 = book.add_sheet(u'Camoni')  # cell_overwrite_ok=True


def get_groups_details(file_to_parse):

    page = file_to_parse
    soup = BeautifulSoup(page, 'html.parser')

    groups = soup.findAll('a', attrs={'class': "nameGroup"})
    groups_num = soup.findAll('div', attrs={'class': "numAdmin"})
    group_about = soup.findAll('div', attrs={'class': "subTitle"})

    return groups, groups_num, group_about


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

    sheet1.write(0, 0, u'Group Name', style)
    sheet1.write(0, 1, u'Members number', style)
    sheet1.write(0, 2, u'About', style)

    style = set_style(1, False)

    print("today:", strftime("%d-%b-%Y - %H:%M", localtime()))
    CamoniFileName = OUTPUT_FOLDER + "Camoni_GROUPS" + "__" + strftime("%d-%b-%Y-%H;%M", localtime()) + ".xls"
    # ---

    file = open("C:\\Users\\ronke\\Desktop\\פרויקט גמר\\parsing\\DATA\\camoni_groups.html", 'rb')
    response = file.read()

    groups, groups_num, group_about = get_groups_details(response)
    for i in range(len(groups)):
        sheet1.write(i+1, 0, groups[i].text, style)
        sheet1.write(i+1, 1, groups_num[i].text, style)
        sheet1.write(i+1, 2, group_about[i].text, style)
        community_number += 1

    print(CamoniFileName)
    book.save(CamoniFileName)
    print("saved ", str(community_number), "groups!" "\nend")


if __name__ == '__main__':
    main()
