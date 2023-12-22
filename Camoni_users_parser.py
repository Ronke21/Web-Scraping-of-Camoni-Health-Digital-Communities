# -*- coding: cp1255 -*-   # sets the coding to hebrew and not gibrish
import os
from time import strftime, localtime
import xlwt
from bs4 import BeautifulSoup

#INPUT_DIR = "C:\\Users\\ronke\\Desktop\פרויקט גמר\\\parsing\\users_download_with_selenium"
INPUT_DIR = "C:\\Users\\ronke\\Desktop\\פרויקט גמר\\parsing\\users_download_new_try2\\"

OUTPUT_DIR = u'C:\\Users\\ronke\\Desktop\\פרויקט גמר\\parsing\\'
SUCCESS = 0
FAILURE = -1
DEFAULT_URL = "https://www.camoni.co.il/411788/"

book = xlwt.Workbook(encoding="utf-8")  # Excel file initialize:
sheet1 = book.add_sheet(u'Camoni')  # cell_overwrite_ok=True

gender_dict = {
    "זכר": "Male",
    "נקבה": "Female",
    "Deleted user 404": "Deleted user 404",
    "": "Empty"
}


class User:

    def __init__(self):
        self.Name = ""
        self.About_me = ""
        self.Gender = ""
        self.Age = ""
        self.Status = ""
        self.Join_date = ""
        self.Community_num = ""
        self.Communities = []
        self.Photo = ""


def get_user_details(user_file_to_parse):
    user = User()
    soup: BeautifulSoup
    try:
        # fp = open(user_file_to_parse, "r", errors='replace')  # errors ignore helps in case fo encoding problems. instead can "rb" read bytes
        fp = open(user_file_to_parse,
                  "rb")  # errors ignore helps in case fo encoding problems. instead can "rb" read bytes

        page = fp.read()
        soup = BeautifulSoup(page, 'html.parser')

        if "p404" in str((soup.find('title')).text):  # if page returned 404 http response - do not put in table
            user.Name = str(user_file_to_parse.decode()).rpartition('\\')[2].rpartition('.')[0]
            user.About_me = "Deleted user 404"
            user.Gender = "Deleted user 404"
            user.Age = "-1"
            user.Status = "Deleted user 404"
            user.Join_date = "-1"
            user.Community_num = "0"
            user.Communities = []
            user.Photo = "-1"
            return True, user

    except Exception as e:
        return False, e

    try:
        # rstrip removes \n in the end of string
        user.Name = (soup.find('h2', attrs={'class': "userName"})).get_text().rstrip()
    except Exception as e:
        user.Name = "ERROR"
        print(str(e))

    try:
        user.About_me = (soup.find('div', attrs={'class': "aboutMeText"})).get_text()
    except Exception as e:
        user.About_me = "ERROR"
        print(str(e))

    details = soup.findAll('div', attrs={'class': "personalInfoWrap"})
    if len(details) == 0:
        user.Name = str(user_file_to_parse.decode()).rpartition('\\')[2].rpartition('.')[0]
        user.About_me = "Empty user"
        user.Gender = "Empty user"
        user.Age = "-1"
        user.Status = "Empty user"
        user.Join_date = "-1"
        user.Community_num = "0"
        user.Communities = []
        user.Photo = "-1"


    for item in details:

        title = item.find('div', attrs={'class': "personalInfoTitle"}).get_text()
        try:
            content = item.find('div', attrs={'class': "personalInfoValue"}).get_text()
        except Exception as e:
            content = "ERROR"
            print(str(e))

        if "הצטרפתי" in title:
            user.Join_date = content.split(" ")[0]  # take date without time
        elif "גיל" in title:
            user.Age = content
        elif "מין" in title:
            user.Gender = gender_dict[content]
        elif "אני" in title:
            user.Status = content

    try:
        communities_soup = soup.findAll('a', attrs={'class': "communityLink"})
        for link in communities_soup:
            user.Communities.append(link.get_text())
        user.Community_num = len(user.Communities)
    except Exception as e:
        print(str(e))

    try:
        photo = soup.find('img', attrs={
            'class': "userProfilePic"})  # return tag of userProfilePic. contains dictionary with src as image source
        photo_url = photo.attrs.get('src')
        if "default_male" in photo_url or "default_female" in photo_url:
            user.Photo = "no"
        else:
            user.Photo = "yes"
    except Exception as e:
        print(str(e))

    return True, user


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


def write_to_file(user, style, date_format, users_number, file):
    file_name = str(file.decode()).rpartition('\\')[2].rpartition('.')[0]
    click = DEFAULT_URL + file_name

    sheet1.write(users_number, 0, str(users_number), style)
    sheet1.write(users_number, 1, file_name, style)
    sheet1.write(users_number, 2, xlwt.Formula('HYPERLINK("%s";"Link")' % click), style)  # write as hyperlink to url
    sheet1.write(users_number, 3, user.Name, style)
    sheet1.write(users_number, 4, user.About_me, style)
    sheet1.write(users_number, 5, user.Join_date, date_format)
    sheet1.write(users_number, 6, user.Age, style)
    sheet1.write(users_number, 7, user.Gender, style)
    sheet1.write(users_number, 8, user.Status, style)
    sheet1.write(users_number, 9, user.Photo, style)
    sheet1.write(users_number, 10, user.Community_num, style)
    sheet1.write(users_number, 11, ''.join(str(item) for item in user.Communities), style)


def main():
    users_number = 0  # number of parsed discussions (presents the current open line in the excel file
    users_fails_number = 0

    reg_style = set_style(6, True)
    sheet1.write(0, 0, u'Serial Number', reg_style)
    sheet1.write(0, 1, 'file', reg_style)
    sheet1.write(0, 2, 'URL', reg_style)
    sheet1.write(0, 3, 'Name', reg_style)
    sheet1.write(0, 4, u'About me', reg_style)
    sheet1.write(0, 5, u'Joining date', reg_style)
    sheet1.write(0, 6, u'Age', reg_style)
    sheet1.write(0, 7, u'Gender', reg_style)
    sheet1.write(0, 8, u'Status', reg_style)
    sheet1.write(0, 9, 'picture', reg_style)
    sheet1.write(0, 10, u'Community number', reg_style)
    sheet1.write(0, 11, u'Communities', reg_style)

    print("today:", strftime("%d-%b-%Y - %H:%M", localtime()))
    CamoniFileName = OUTPUT_DIR + "Camoni_USERS" + "__" + strftime("%d-%b-%Y-%H;%M", localtime()) + ".xls"
    # ---

    directory = os.fsencode(INPUT_DIR)

    reg_style = set_style(1, False)
    date_format = set_style(1, False)
    date_format.num_format_str = 'dd/mm/yy'

    failed_users = []

    for path, subdirs, files in os.walk(directory):
        for name in files:
            file = os.path.join(path, name)
            stat, user = get_user_details(file)
            if stat is True:
                users_number += 1
                write_to_file(user, reg_style, date_format, users_number, file)
            else:
                print(str(name.decode()) + "failed because: " + str(user))
                users_fails_number += 1
                failed_users.append(name.decode())
            if users_number % 500 == 0:
                print("parsed", users_number, "files")

    print(CamoniFileName)
    sheet1.write(users_number + 1, 1, "\n")
    sheet1.write(users_number + 1, 2, "\n")
    sheet1.write(users_number + 1, 7, "\n")
    book.save(CamoniFileName)

    print("parsed", users_number, "users!")
    print("failed", users_fails_number, "users!")

    print("end")
    print(failed_users)


if __name__ == '__main__':
    main()
