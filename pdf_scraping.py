import csv
import tabula
from collections import defaultdict
import datetime
import pandas as pd


# #path to pdf
# file = "C:\DownloadFiles\\test.pdf"
#
# #read the pdf
# tables = tabula.read_pdf(file, pages="all", multiple_tables=True)
#
# # output all the tables in the PDF to a CSV
# tabula.convert_into(file, "test.csv", pages="all")


columns = defaultdict(list) # each value in each column is appended to a list

with open('test.csv') as f:
    reader = csv.reader(f)
    next(reader)
    for row in reader:
        for (i,v) in enumerate(row):
            columns[i].append(v)

counter = 0
for j in range(0, len(columns)):
    for i in range(1,len(columns[j])):
        if columns[j][i] == 'F':
            length = len(str(columns[j+1][i]))
            last_two_digits = str(columns[j+1][i])[length-2] + str(columns[j+1][i])[length-1]

            last_two_digits = last_two_digits.lstrip()
            month = columns[j+1][i].split('/')[0]
            today = datetime.datetime.now()

            begin_date = str(today.year) + "/" + str(month) + "/" + str(last_two_digits)
            print("Begin date: ", begin_date)

            enddate = pd.to_datetime(str(month) + "/" + str(last_two_digits + "/" + str(today.year))) + pd.DateOffset(days=5)
            print("End date:   ", enddate.strftime("%Y/%m/%d"))







