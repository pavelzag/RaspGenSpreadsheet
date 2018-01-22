import calendar
import gspread
import math
from oauth2client.service_account import ServiceAccountCredentials
from dbconnector import get_time_spent
from logger import logging_handler
from datetime import timedelta, date
from time import strptime

months_list = []
scope = ['https://spreadsheets.google.com/feeds']
credentials = ServiceAccountCredentials.from_json_keyfile_name('creds.json', scope)
gc = gspread.authorize(credentials)
worksheet = gc.open('generator_stats.xlsx')
worksheet_list = worksheet.worksheets()
time_spent_list = {}


def daterange(start_date, end_date):
    for n in range(int ((end_date - start_date).days)):
        yield start_date + timedelta(n)


def get_end_date_per_month(year, month):
    return date(year, month, calendar.monthrange(year, month)[1])


def get_last_day(year, month):
    return calendar.monthrange(year, month)[1]


def get_start_date_and_end_date(year, month):
    start_date = date(year, month, 1)
    end_date = get_end_date_per_month(year, month)
    return start_date, end_date

if __name__ == '__main__':
    for worksheet_tab in worksheet_list:
        worksheet_tab_title = worksheet_tab._title
        selected_worksheet = worksheet.worksheet(worksheet_tab_title)
        month_number = strptime(worksheet_tab_title, '%B').tm_mon
        start_date, end_date = get_start_date_and_end_date(year=2018, month=month_number)
        for single_date in daterange(start_date, end_date):
            msg = '{} {}'.format('Acquiring time spent for', worksheet_tab_title)
            logging_handler(msg)
            time_spent = get_time_spent(single_date)
            msg = '{} {}'.format('Acquiring time spent for', single_date)
            logging_handler(msg)
            time_spent_list[single_date] = time_spent
        first_sheet_number = 'B1'
        first_day_of_month = 1
        last_day_of_month = get_last_day(2018, month_number)
        last_sheet_number = '{}{}'.format('B', last_day_of_month+1)
        for x in range(first_day_of_month, last_day_of_month):
            current_day = date(2018, month_number, x)
            current_day_str = str(current_day.day)
            date_sheet_number = '{}{}'.format('A', x+1)
            hours_sheet_number = '{}{}'.format('B', x+1)
            msg = '{} {} {}'.format('Updating the spreadsheet for', worksheet_tab_title, current_day)
            logging_handler(msg)
            # Updating the day field
            selected_worksheet.update_acell(date_sheet_number, current_day_str)
            # Updating the hours field
            selected_worksheet.update_acell(hours_sheet_number, math.ceil(time_spent_list[current_day]/60))
        msg = '{} {}'.format('Finished updating', worksheet_tab_title)
        logging_handler(msg)
