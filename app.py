from collections import defaultdict
import calendar
import gspread
import math
from time import strptime
from oauth2client.service_account import ServiceAccountCredentials
from dbconnector import get_time_spent
from hebrew_day import get_hebrew_day
from datetime import timedelta, date
from logger import logging_handler

months_list = []
scope = ['https://spreadsheets.google.com/feeds']
credentials = ServiceAccountCredentials.from_json_keyfile_name('creds.json', scope)
gc = gspread.authorize(credentials)
worksheet = gc.open('generator_test')
worksheet_list = worksheet.worksheets()
time_spent_dict = {}
# This dictionary holds the number of the week and the current date for further usage calculations
day_and_week_dict = defaultdict(list)
print('s')


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


def calculate_weekly_usage(week, time_format='m'):
    weekly_usage = []
    # This function calculates the weekly usage
    for day in day_and_week_dict[week]:
        time_spent_this_day = time_spent_dict[day]
        weekly_usage.append(time_spent_this_day)
    weekly_usage_sum_minutes = math.ceil(sum(weekly_usage)/60)
    weekly_usage_sum_hours = math.ceil(sum(weekly_usage)/3600)
    del weekly_usage[:]
    if time_format == 'm':
        return weekly_usage_sum_minutes
    elif time_format == 'h':
        return weekly_usage_sum_hours


def get_the_week(current_day):
    # Checks the number of the week per entered day
    year, month, day = current_day.strftime('%Y %-m %-d').split(' ')
    week_number = date(int(year), int(month), int(day)+1).isocalendar()[1]
    msg = '{} {} {}'.format(current_day, 'is on week number', week_number)
    logging_handler(msg)
    return week_number

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
            time_spent_dict[single_date] = time_spent
        first_sheet_number = 'B1'
        first_day_of_month = 1
        last_day_of_month = get_last_day(2018, month_number)
        last_sheet_number = '{}{}'.format('B', last_day_of_month+1)
        for x in range(first_day_of_month, last_day_of_month):
            current_day = date(2018, month_number, x)
            what_week = get_the_week(current_day)
            current_day_str = str(current_day.day)
            # day_and_week_dict[what_week].append(current_day_str)
            day_and_week_dict[what_week].append(current_day)
            hebrew_date_name = '{}{}'.format('A', x+1)
            date_sheet_number = '{}{}'.format('B', x+1)
            hours_sheet_number = '{}{}'.format('C', x+1)
            hebrew_day = get_hebrew_day(current_day.strftime("%A"))
            msg = '{} {} {}'.format('Updating the spreadsheet for', worksheet_tab_title, current_day)
            logging_handler(msg)
            # Updating the Hebrew day field
            selected_worksheet.update_acell(hebrew_date_name, hebrew_day)
            # Updating the day field
            selected_worksheet.update_acell(date_sheet_number, current_day_str)
            # Updating the hours field
            selected_worksheet.update_acell(hours_sheet_number, math.ceil(time_spent_dict[current_day]/60))
        for week in day_and_week_dict:
            weekly_usage = calculate_weekly_usage(week, 'h')
            msg = '{} {}{} {} {}'.format('The weekly usage for the', week, 'th week is', weekly_usage, 'hours')
            logging_handler(msg)
        msg = '{} {}'.format('Finished updating', worksheet_tab_title)
        logging_handler(msg)
