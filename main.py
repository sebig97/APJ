import os
import csv
import time
import tabula
import datetime
import pandas as pd
from m import homePage
from collections import defaultdict


if __name__ == '__main__':
    # constructor with random parameters
    temp = homePage("1", "2", "3", "4")
    urls, merchants, users, passwords = temp.read_csv()

    portal = []
    for c in range(0, len(urls)):
        portal.append(homePage(str(urls[c]), str(merchants[c]), str(users[c]), str(passwords[c])))

    # JCB
    portal[1].get_chrome_driver()
    portal[1].test_load_home_page()
    portal[1].login_JCB()
    portal[1].search_JCB_page()
    time.sleep(2)
    portal[1].tearDown()

    # get name of the pdf downloaded from JCB
    entries = os.listdir('C:\DownloadFiles\\')
    print("Name of the pdf = ", entries[0])
    time.sleep(1)

    # VERITRANS
    portal[2].get_chrome_driver()
    portal[2].test_load_home_page()
    portal[2].login_veritrans()

    # path to the pdf downloaded from JCB
    # file = "C:\DownloadFiles\\" + str(entries[0])
    # print("PATH to the pdf = ", file)
    #
    # # read the pdf
    # tables = tabula.read_pdf(file, pages="all", multiple_tables=True)
    #
    # # output all the tables in the PDF to a CSV
    # tabula.convert_into(file, "test.csv", pages="all")

    columns = defaultdict(list)  # each value in each column is appended to a list

    with open('test.csv') as f:
        reader = csv.reader(f)
        next(reader)
        for row in reader:
            for (i, v) in enumerate(row):
                columns[i].append(v)

    dates = []
    counter = 0
    for j in range(0, len(columns)):
        for i in range(1, len(columns[j])):
            if columns[j][i] == 'F':
                length = len(str(columns[j + 1][i]))
                last_two_digits = str(columns[j + 1][i])[length - 2] + str(columns[j + 1][i])[length - 1]

                last_two_digits = last_two_digits.lstrip()
                month = columns[j + 1][i].split('/')[0]
                today = datetime.datetime.now()

                begin_date = str(today.year) + "/" + str(month) + "/" + str(last_two_digits)
                print("Begin date: ", begin_date)

                enddate = pd.to_datetime(
                    str(month) + "/" + str(last_two_digits + "/" + str(today.year))) + pd.DateOffset(days=5)
                print("End date:   ", enddate.strftime("%Y/%m/%d"))

                if begin_date not in dates:
                    dates.append(begin_date)

                    portal[2].search_veritrans_page(str(begin_date), str(enddate.strftime("%Y/%m/%d")))
                    time.sleep(3)
                    # rename file
                    portal[2].rename_file("JCB_2020.csv", r"C:\DownloadFiles\\", 60)

                    # upload to sharepoint
                    portal[2].send_excel_to_powerautomate(r'C:\DownloadFiles\JCB_2020.CSV')

                    #move to sharepoint
                    portal[2].move_to_sharepoint()

                    #rename the csv from sharepoint
                    portal[2].rename_from_sharepoint(str(today.year) + "-" + str(month) + "-" + str(last_two_digits), enddate.strftime("%Y-%m-%d"))

                    # delete all files
                    portal[2].delete_all_files(r'C:\DownloadFiles\*.csv')
                else:
                    print("This date already exists in the list")

    print(dates)

    portal[2].delete_all_files(r'C:\DownloadFiles\*')
    portal[2].tearDown()





