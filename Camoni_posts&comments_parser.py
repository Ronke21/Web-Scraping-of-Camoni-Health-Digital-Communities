# -*- coding: cp1255 -*-   # sets the coding to hebrew and not gibrish
import datetime
import os
import random
import time
from time import strftime, localtime
from bs4 import BeautifulSoup
import xlsxwriter

DEFAULT_URL = "https://www.camoni.co.il/411804/"
OUTPUT_FOLDER = u'C:\\Users\\ronke\\Desktop\\פרויקט גמר\\parsing\\'
INPUT_DIR = "C:\\Users\\ronke\\Desktop\\פרויקט גמר\\parsing\\new_posts_with_selenium"

FAILURE = -1
COMMUNITIES = ['אוסטאופורוזיס', 'אטופיק דרמטיטיס', 'אירוע מוחי', 'בני משפחה מטפלים', 'בעיות גדילה', 'גמילה מעישון',
               'דיכאון וחרדה', 'הידרדניטיס סופורטיבה HS', 'הפרעות אכילה', 'השמנה', 'זכויות החולה', 'חוט שדרה',
               'טראומה מטרור ומלחמה', 'טרשת נפוצה', 'כאב', 'כליות ודיאליזה', 'לחץ דם', 'מושתלים', 'מחלות לב',
               'מיאלומה נפוצה', 'מפרקים', 'ניוון מקולרי גילי (AMD)', 'נשימה', 'סוכרת סוג-1', 'סוכרת סוג-2',
               'סיוגרן (sjogren)', 'סרטן', 'סרטן המעי הגס', 'סרטן הריאות', 'סרטן השד', 'סרטן השחלות',
               'סרטן שלפוחית השתן', 'ער"ן - עזרה ראשונה נפשית', 'פיברומיאלגיה', 'פסוריאזיס', 'קדחת ים תיכונית (FMF)',
               'קרוהן וקוליטיס', 'קשב וריכוז',
               'שלפוחית רגיזה', 'תרופות ורוקחות', 'ער"ן - עזרה ראשונה נ...', 'קדחת ים תיכונית (FMF...', 'דלקות מפרקים',
              'נזאל פוליפוזיס', 'הידרדניטיס סופורטיבה...', 'ניוון מקולרי גילי (A...']

class Talkback:
    def __init__(self):
        self.Content = ""
        self.Publisher = ""
        self.Date = ""
        self.Time = ""
        self.Word_number_in_post = 0
        self.Char_number_in_post = 0
        self.Likes_num = 0


class Full_Post:
    def __init__(self):
        self.Number = ""
        self.ID = ""
        self.Date = ""
        self.Time = ""
        self.Group = ""
        self.Group_type = ""
        self.Author = ""
        self.Title = ""
        self.Content = ""
        self.Word_number_in_post = 0
        self.Char_number_in_post = 0
        self.Talkback_list = []
        self.Talkbacks_num = 0
        self.URL = ""
        self.Likes_num = 0


def get_full_post_details(file_to_parse, post_number):
    post = Full_Post()
    post.ID = (file_to_parse[-11:-5]).replace('\\', '')  # take end of path as file name

    fp = open(file_to_parse, "rb")
    page = fp.read()
    soup = BeautifulSoup(page, 'html.parser')

    for elem in soup.find_all(["a", "p", "div", "h3", "br"]):  # create new line in text objects
        elem.append('\n')

    # ______________________________________________________________________________________________________________________________________________________________
    EmptyPage = soup.find('div', attrs={'class': "error-num"})
    if EmptyPage is not None:
        print(file_to_parse, "not Found", EmptyPage.text)
        post.Number = post_number - 1
        return FAILURE

    stuff = soup.find('a', attrs={'class': "communityLink"})
    if stuff is None:  # not community
        return FAILURE
    post.Group = stuff.text.rstrip('\n')
    if post.Group == '':
        post.Group = "קבוצה שנסגרה"

    if post.Group in COMMUNITIES:
        post.Group_type = "community"
    else:
        post.Group_type = "group"

    content = soup.find('div', attrs={'class': "wrappContent"})
    stuff = content.find('div', attrs={'class': "postTitle"})
    if stuff is None:
        post.Title = ""
    else:
        post.Title = stuff.text
    stuff = content.find('div', attrs={'class': "MessageSubject"})
    if stuff is None:
        stuff = content.find('div', attrs={'class': "Bg_text"})
        if stuff is None:
            post.Content = "ERROR"
        else:
            post.Content = stuff.text
    else:
        post.Content = stuff.text

    # stuff = content.find('a', attrs={'class': "postMemberName"})
    stuff = content.find('div', attrs={'class': "postDetail"})

    if stuff is None:
        stuff = content.find('img', attrs={'alt': "Anonymouse"})
    if stuff is None:
        post.Author = "ERROR"
    else:
        post.Author = stuff.text

    stuff = content.find('div', attrs={'class': "postDetail"})
    if stuff is None:
        post.Author = "ERROR"
    else:
        post.Author = stuff.text

    try:
        stuff = content.findAll('div', attrs={'class': "postDetail"})
        post_time = stuff[1].text
        times = get_correct_time(post_time, file_to_parse)

        post.Date = str(times[0])
        post.Time = str(times[1])

    except Exception as e:
        post.Time = "ERROR"
        post.Date = "ERROR"
        print("error with date or time because: ", str(e))

    post.Word_number_in_post = len(post.Content.split())
    post.Char_number_in_post = len(post.Content)

    talckbacks_num = stuff[2].text.split()
    post.Talkbacks_num = talckbacks_num[0]

    stuff = soup.findAll('div', attrs={'class': "iComm"})
    for comment in stuff:
        current_talkback = Talkback()
        try:
            comment_data = comment.find('div', attrs={'class': "readMoreFullText"})
            if comment_data is None:
                comment_data = comment.find('div', attrs={'class': "textCom enableEmoticons"})
            if comment_data is None:
                comment_data = comment.find('div', attrs={'class': "textCom enableEmoticons1"})

            if comment_data is not None:
                current_talkback.Content = comment_data.text
        except Exception as e:
            current_talkback.Content = "ERROR" + str(e)

        try:
            comment_writer = comment.find('a', attrs={'class': "postMemberName"})
            if comment_writer is None:
                comment_writer = comment.find('div', attrs={'class': "nameComm"})

            current_talkback.Publisher = comment_writer.text
        except Exception as e:
            current_talkback.Publisher = "ERROR" + str(e)

        try:
            comment_time = comment.find('div', attrs={'class': "commentTime"}).text
            comment_times = get_correct_time(comment_time, file_to_parse)
            current_talkback.Date = str(comment_times[0])
            current_talkback.Time = str(comment_times[1])
        except Exception as e:
            current_talkback.Date = "ERROR" + str(e)
            current_talkback.Time = "ERROR" + str(e)

        try:
            comment_likes = comment.find('span', attrs={'class': "likeTitle"}).text
            if len(comment_likes) == 0:
                current_talkback.Likes_num = 0
            else:
                current_talkback.Likes_num = int(comment_likes)
        except Exception as e:
            current_talkback.Likes_num = -1
            print(str(e))

        current_talkback.Word_number_in_post = len(current_talkback.Content.split())
        current_talkback.Char_number_in_post = len(current_talkback.Content)

        post.Talkback_list.append(current_talkback)
    # -------

    post.URL = DEFAULT_URL + post.ID

    try:
        post_likes = soup.find('span', attrs={'class': "likeTitle"}).text
        if len(post_likes) == 0:
            post.Likes_num = 0
        else:
            post.Likes_num = int(post_likes)
    except Exception as e:
        post.Likes_num = -1
        print(str(e))

    return post


def get_correct_time(time, file):
    if "אתמול" in time:
        modification_time_epoch = os.path.getmtime(file)
        modification_time_day = datetime.datetime.fromtimestamp(modification_time_epoch).strftime('%d/%m/%Y')
        times = (time.replace("אתמול ב:", modification_time_day)).split()
        times[0] = times[0][:-4] + times[0][-2:]  # make year only 2 digits
    elif "לפני" in time:
        time_list = time.split()
        subtract = int(time_list[1])
        modification_time_epoch = os.path.getmtime(file)
        modification_time_day = datetime.datetime.fromtimestamp(modification_time_epoch)
        good_date = (modification_time_day - datetime.timedelta(days=subtract)).strftime('%d/%m/%Y')
        good_date = good_date[:-4] + good_date[-2:]  # make year only 2 digits
        rand_hour = datetime.time(random.randrange(23), random.randrange(59)).strftime('%H:%M')
        times = [good_date, rand_hour]
    elif "בשעה האחרונה" in time:
        modification_time_epoch = os.path.getmtime(file)
        modification_time_day = datetime.datetime.fromtimestamp(modification_time_epoch).strftime('%d/%m/%Y %H:%M')
        times = modification_time_day.split()
        times[0] = times[0][:-4] + times[0][-2:]  # make year only 2 digits
    else:
        times = time.split()
    return times


def create_file(file_type):
    CAMONI_FILE_NAME = OUTPUT_FOLDER + "Camoni_" + str(file_type) + "__" + strftime("%d-%b-%Y-%H;%M",
                                                                                    localtime()) + ".xlsx"
    book = xlsxwriter.Workbook(CAMONI_FILE_NAME, {
        'strings_to_urls': False})  # prevent exception of "since it exceeds Excel's limit of 65,530 URLS per worksheet"
    sheet_full = book.add_worksheet(file_type + '_FULL')  # cell_overwrite_ok=True.
    sheet_communities = book.add_worksheet(file_type + '_COMMUNITIES')  # cell_overwrite_ok=True.
    sheet_groups = book.add_worksheet(file_type + '_GROUPS')  # cell_overwrite_ok=True.

    style = book.add_format({'bold': True, 'font_color': 'black'})
    print("created file: ", CAMONI_FILE_NAME)
    return book, sheet_full, sheet_communities, sheet_groups, style


def create_soup_posts_list():
    parsed_post_number = 0

    directory = os.fsencode(INPUT_DIR)
    posts_list = []

    for path, subdirs, files in os.walk(directory):
        for name in sorted(files):
            file = os.path.join(path, name)
            try:
                post = get_full_post_details(file.decode(), parsed_post_number)
            except Exception as e:
                print("error parsing with: " + str(name) + "because: " + str(e))
                continue
            if post == FAILURE:
                print("error parsing with: " + str(name))
                continue

            posts_list.append(post)
            parsed_post_number += 1

            if parsed_post_number % 500 == 0:  # save  file every 100 iterations
                print("-----------------------------------------------------------------parsed files: " + str(
                    parsed_post_number) + "---------------------------------------------------------------------")

    return posts_list


def create_headlines_posts_file(posts_sheet, style):
    posts_sheet.write(0, 0, u'Page Number', style)
    posts_sheet.write(0, 1, u'URL', style)
    posts_sheet.write(0, 2, u'Date', style)
    posts_sheet.write(0, 3, u'Time', style)
    posts_sheet.write(0, 4, u'Community', style)
    posts_sheet.write(0, 5, u'Group/Community', style)
    posts_sheet.write(0, 6, u'Post publisher', style)
    posts_sheet.write(0, 7, u'Title', style)
    posts_sheet.write(0, 8, u'Post content', style)
    posts_sheet.write(0, 9, u'Post words number', style)
    posts_sheet.write(0, 10, u'Post chars number', style)
    posts_sheet.write(0, 11, u'Post likes', style)
    posts_sheet.write(0, 12, u'Comment number', style)
    comment_idx = 1
    c1 = u'comment user #'
    c2 = u'comment #'
    for column in range(13, 20, 2):
        posts_sheet.write(0, column, c1 + str(comment_idx), style)
        posts_sheet.write(0, column + 1, c2 + str(comment_idx), style)
        comment_idx += 1


def create_headlines_comments_file(comments_sheet, style):
    comments_sheet.write(0, 0, u'Page Number', style)
    comments_sheet.write(0, 1, u'URL', style)
    comments_sheet.write(0, 2, u'Date', style)
    comments_sheet.write(0, 3, u'Time', style)
    comments_sheet.write(0, 4, u'Community', style)
    comments_sheet.write(0, 5, u'Group/Community', style)
    comments_sheet.write(0, 6, u'Comment author', style)
    comments_sheet.write(0, 7, u'Comment', style)
    comments_sheet.write(0, 8, u'Post/Comment', style)
    comments_sheet.write(0, 9, u'Word count', style)
    comments_sheet.write(0, 10, u'Char count', style)
    comments_sheet.write(0, 11, u'Likes count', style)


def write_full_post_to_file(post, line_number, sheet, post_number):
    sheet.write(line_number, 0, post.ID)
    sheet.write(line_number, 1, post.URL)
    sheet.write(line_number, 2, post.Date)
    sheet.write(line_number, 3, post.Time)
    sheet.write(line_number, 4, post.Group)
    sheet.write(line_number, 5, post.Group_type)
    sheet.write(line_number, 6, post.Author)
    sheet.write(line_number, 7, post.Title)
    sheet.write_string(line_number, 8, str(post.Content))
    sheet.write_number(line_number, 9, int(post.Word_number_in_post))
    sheet.write_number(line_number, 10, int(post.Char_number_in_post))
    sheet.write_number(line_number, 11, post.Likes_num)
    sheet.write_number(line_number, 12, int(post.Talkbacks_num))

    iField = 13
    i = 0

    for talkback in post.Talkback_list:
        sheet.write(line_number, iField + i, talkback.Publisher)
        content = talkback.Content
        if len(content) > 32767:  # maximum length of cell in xlsx is 32767
            content = content[:32766]
        sheet.write_string(line_number, iField + i + 1, str(content))
        i += 2


def write_comments_to_file(post, sheet, comment_number):
    # write original post
    sheet.write_number(comment_number, 0, int(post.ID))  # u'זיהוי'
    sheet.write(comment_number, 1, post.URL)  # u'זיהוי'
    sheet.write(comment_number, 2, post.Date)
    sheet.write(comment_number, 3, post.Time)
    sheet.write(comment_number, 4, post.Group)
    sheet.write(comment_number, 5, post.Group_type)
    sheet.write(comment_number, 6, post.Author)
    content = post.Content
    if len(content) > 32767:  # maximum length of cell in xlsx is 32767
        content = content[:32766]
    sheet.write_string(comment_number, 7, content)
    sheet.write(comment_number, 8, "post")
    sheet.write_number(comment_number, 9, post.Word_number_in_post)
    sheet.write_number(comment_number, 10, post.Char_number_in_post)
    sheet.write_number(comment_number, 11, post.Likes_num)

    comment_number += 1

    # write comments
    for i in range(int(post.Talkbacks_num)):
        comment = post.Talkback_list[i]
        sheet.write_number(comment_number, 0, int(post.ID))  # u'זיהוי'
        sheet.write(comment_number, 1, post.URL)
        sheet.write(comment_number, 2, comment.Date)
        sheet.write(comment_number, 3, comment.Time)
        sheet.write(comment_number, 4, post.Group)
        sheet.write(comment_number, 5, post.Group_type)
        sheet.write(comment_number, 6, comment.Publisher)
        text = comment.Content
        if len(text) > 32767:  # maximum length of cell in xlsx is 32767
            text = text[:32766]
        sheet.write_string(comment_number, 7, text)
        sheet.write(comment_number, 8, "comment")
        sheet.write_number(comment_number, 9, comment.Word_number_in_post)
        sheet.write_number(comment_number, 10, comment.Char_number_in_post)
        sheet.write_number(comment_number, 11, comment.Likes_num)

        comment_number += 1

    return comment_number


def main():
    posts_book, posts_sheet_full, posts_sheet_communities, posts_sheet_groups, style = create_file("POSTS")
    create_headlines_posts_file(posts_sheet_full, style)
    create_headlines_posts_file(posts_sheet_communities, style)
    create_headlines_posts_file(posts_sheet_groups, style)

    comments_book, comments_sheet_full, comments_sheet_communities, comments_sheet_groups, style = create_file(
        "COMMENTS")
    create_headlines_comments_file(comments_sheet_full, style)
    create_headlines_comments_file(comments_sheet_communities, style)
    create_headlines_comments_file(comments_sheet_groups, style)

    posts_list = create_soup_posts_list()
    # sorted_post_list = sorted(posts_list, key=lambda post: post.ID)
    sorted_post_list = sorted(posts_list, key=lambda post: datetime.datetime.strptime(post.Date, '%d/%m/%y'))

    post_number = 1
    comment_number = 1
    post_community_line = 1
    post_group_line = 1
    comment_community_line = 1
    comment_group_line = 1

    for post in sorted_post_list:
        try:
            write_full_post_to_file(post, post_number, posts_sheet_full, post_number)
            if post.Group_type == "community":
                write_full_post_to_file(post, post_community_line, posts_sheet_communities, post_number)
                post_community_line += 1
            else:
                write_full_post_to_file(post, post_group_line, posts_sheet_groups, post_number)
                post_group_line += 1

            post_number += 1
        except Exception as e:
            print("error writing with: " + str(post.ID) + "because: " + str(e))
            continue

        if post_number % 500 == 0:  # save  file every x iterations
            print("-----------------------------------------------------------------written posts: " + str(
                post_number) + "---------------------------------------------------------------------")

        try:
            comment_number = write_comments_to_file(post, comments_sheet_full, comment_number)
            if post.Group_type == "community":
                comment_community_line = write_comments_to_file(post, comments_sheet_communities,
                                                                comment_community_line)
            else:
                comment_group_line = write_comments_to_file(post, comments_sheet_groups, comment_group_line)
        except Exception as e:
            print("error writing with: " + str(post.ID) + "because: " + str(e))
            continue

        if comment_number % 500 == 0:  # update  file every x iterations
            print("-----------------------------------------------------------------written comments: " + str(
                comment_number) + "---------------------------------------------------------------------")

    posts_book.close()  # and save file
    comments_book.close()  # and save file

    print("written", comment_number - 1, "comments!")
    print("from", post_number - 1, "posts!")
    print("end")


if __name__ == '__main__':
    start_time = time.time()
    main()
    duration = (time.time() - start_time)
    print("Duration: ", datetime.timedelta(seconds=duration))
