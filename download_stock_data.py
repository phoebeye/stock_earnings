import sys
import os
# import urllib2
# from urllib3.contrib import pyopenssl
import csv
import json
import datetime
from dateutil.relativedelta import relativedelta
import calendar
import requests
import os
import xlsxwriter

from bs4 import BeautifulSoup
# date_time = datetime.datetime(2017,03,01,6,0)

def main(argv):
    working_dir = '/Users/Phoebe/Documents/data_analysis/stk_price_repo'
    workbook_name = datetime.datetime.now().strftime ("%m-%d-%Y")
    output_filename = os.path.join(working_dir, workbook_name + '_stock_analysis.xlsx')
    workbook = xlsxwriter.Workbook(output_filename)
    for num in range(22,23):
        date_time = datetime.datetime(2017,5,num,6,0)
        get_stickers(date_time,workbook)
    print 'closing workbook save now'
    workbook.close()

def get_stickers(date_time,workbook):
    stk_obj = {}
    # convert readable time into timestamp
    epoch = calendar.timegm(date_time.timetuple());
    print epoch, 'epoch'
    wiki = "https://www.zacks.com/includes/classes/z2_class_calendarfunctions_data.php?calltype=eventscal&date=" + str(epoch) +"&type=1"
    headers = {'User-agent':'Mozilla/5.0'}
    page = requests.get(wiki)

    page_obj = page.json()
    stickers_list = [];
    for line in page_obj['data']:

        soup = BeautifulSoup(line[0], "html.parser")
        # find out sticker and turn unicode into strings
        sticker = str(soup.find_all('a')[0]['rel'][0])
        # append sticker into stickers list
        stickers_list.append(sticker)
        earning_time = line[3]
        if earning_time == '--':
            earning_time = 'none'
        elif earning_time == 'amc':
            earning_time = 'after-mkt'
        elif earning_time == 'bmo':
            earning_time = 'before-mkt'

        stk_obj[sticker] = earning_time

    # convert stickers_list into a join strings for further price queries
    stickers_joint_str = ','.join(stickers_list)
    get_prices(stickers_joint_str, date_time, workbook, stk_obj)

def get_prices(stickers, date_time, workbook, stk_obj):
    data_arr = []
    dayOfWk = date_time.weekday()
    # get datetime of previous 3 weekday and the next weekday
    if any ([dayOfWk == 0, dayOfWk == 1, dayOfWk == 2 ]):
        subtraction = 5
    else:
        subtraction = 3

    if dayOfWk == 4:
        addition = 3
    else:
        addition = 1

    # startDate = date_time - datetime.timedelta(subtraction)
    # endDate = date_time + datetime.timedelta(addition)
    prevThreeMon = date_time+relativedelta(months=-3)
    print prevThreeMon, '3 month ago'
    startDate = prevThreeMon.replace(hour=6, minute=0, second=0)
    # add 1 day for endDate for after market price changes
    endDate = (date_time + datetime.timedelta(1)).replace(hour=6, minute=0, second=0)
    print date_time, startDate, endDate, '------get date'
    wiki ='https://www.quandl.com/api/v3/datatables/WIKI/PRICES.json?qopts.export=true&ticker=' + stickers +'&date.gte=' + startDate.strftime('%Y-%m-%d') + '&date.lte=' + endDate.strftime('%Y-%m-%d') +'&qopts.columns=ticker,date,open,close&api_key=ZpmYKWhBq1kYPmxDHePC'
    print wiki
    page_content = os.popen("curl " + '"'+wiki+'"').read()
    page_obj =json.loads(page_content)
    print page_obj, 'page obj'
    download_url = page_obj['datatable_bulk_download']['file']['link']
    print download_url
    if not download_url:
        return
    root_name = os.path.splitext(os.path.basename(download_url))[0]
    out_zip = root_name+".zip"
    out_csv = root_name+".csv"
    download_cmd = ' '.join(['curl','-o','"'+out_zip+'"','"'+download_url+'"'])
    print download_cmd
    os.system(download_cmd)
    unzip_cmd = ' '.join(['unzip',out_zip])
    os.system(unzip_cmd)
    with open(out_csv, 'rb') as csvfile:
        csv_reader = csv.reader(csvfile, delimiter=',')
        row_counter = 0
        for row in csv_reader:
            if row_counter == 0:
                row_counter+=1
                continue
            else:
                row.append(stk_obj[row[0]])
                data_arr.append(row)
    print data_arr
    create_worksheet(data_arr, date_time, workbook)

def create_worksheet(data, date_time, workbook):
    # Create a workbook and add a worksheet.
    sheet_name = date_time.strftime('%Y-%m-%d')
    worksheet = workbook.add_worksheet(sheet_name)
    print sheet_name
    print "sheet_name"
    # print data

    # Add a bold format to use to highlight cells.
    bold = workbook.add_format({'bold': 1})

    # Add a number format for cells with money.
    money_format = workbook.add_format({'num_format': '0.00'})

    # Add an Excel date format.
    date_format = workbook.add_format({'num_format': 'mm-dd-yy'})

    # Adjust the column width.
    worksheet.set_column(1, 1, 15)

    # Write some data headers.
    worksheet.write('A1', 'ticker')
    worksheet.write('B1', 'date')
    worksheet.write('C1', 'open')
    worksheet.write('D1', 'close')
    worksheet.write('E1', 'earning_time')

    # Start from the first cell. Rows and columns are zero indexed.
    row = 1
    col = 0

    # Iterate over the data and write it out row by row.
    for ticker, date_str, open_price, close_price, earning_time in data:
        date = datetime.datetime.strptime(date_str, "%Y-%m-%d")
        worksheet.write_string(row, col,     ticker)
        worksheet.write_datetime(row, col + 1, date, date_format)
        worksheet.write_number(row, col + 2, float(open_price), money_format)
        worksheet.write_number(row, col + 3, float(close_price), money_format)
        worksheet.write_string(row, col + 4, earning_time)

        row += 1

    # Write a total using a formula.
    # worksheet.write(row, 0, 'Total')
    # worksheet.write(row, 1, '=SUM(B1:B4)')


if __name__ == '__main__':
    main(sys.argv[1:])
