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
weekly_usage_list = []
workday_usage_list = []
weekend_usage_list = []


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


def calculate_workdays_weekends_usage(week, workdays=True, time_format='m'):
    workday_usage = []
    weekend_usage = []
    # This function calculates the workdays usage
    for day in day_and_week_dict[week]:
        if day.weekday() == 3 or day.weekday() == 4 or day.weekday() == 5:
            time_spent_this_day = time_spent_dict[day]
            weekend_usage.append(time_spent_this_day)
        else:
            time_spent_this_day = time_spent_dict[day]
            workday_usage.append(time_spent_this_day)
    weekend_usage_sum_minutes = math.ceil(sum(weekend_usage) / 60)
    weekend_usage_sum_hours = math.ceil(sum(weekend_usage) / 3600)
    workday_usage_sum_minutes = math.ceil(sum(workday_usage) / 60)
    workday_usage_sum_hours = math.ceil(sum(workday_usage) / 3600)
    del workday_usage[:]
    del weekend_usage[:]
    if workdays:
        if time_format == 'm':
            return workday_usage_sum_minutes
        elif time_format == 'h':
            return workday_usage_sum_hours
    else:
        if time_format == 'm':
            return weekend_usage_sum_minutes
        elif time_format == 'h':
            return weekend_usage_sum_hours


def get_the_week(current_day):
    # Checks the number of the week per entered day
    year, month, day = current_day.strftime('%Y %-m %-d').split(' ')
    week_number = date(int(year), int(month), int(day)+1).isocalendar()[1]
    msg = '{} {} {}'.format(current_day, 'is on week number', week_number)
    logging_handler(msg)
    return week_number


def insert_weekly_calc_cells(cell_num):
    # Insert blank for the weekly calculation
    hebrew_date_name = '{}{}'.format('A', cell_num)
    date_sheet_number = '{}{}'.format('B', cell_num)
    hours_sheet_number = '{}{}'.format('C', cell_num)
    blank_val = ''
    selected_worksheet.update_acell(hebrew_date_name, blank_val)
    selected_worksheet.update_acell(date_sheet_number, blank_val)
    selected_worksheet.update_acell(hours_sheet_number, blank_val)


def get_saturdays_amt(year=2018, month=1):
    # Calculates amount of Saturdays per relevant month
    return len([1 for i in calendar.monthcalendar(year, month) if i[6] != 0])


if __name__ == '__main__':
    # Looping through every month per tab
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
        saturdays_amt = get_saturdays_amt(year = start_date.year, month = start_date.month)
        cells_number = len(range(first_day_of_month, last_day_of_month)) + saturdays_amt + 1
        days_number = len(range(first_day_of_month, last_day_of_month))
        cell_num = 1
        day_num = 0
        sum_cell_list = []
        # Looping through every cell for every day in month + a cell for every week for the weekly amount
        while cell_num < cells_number:
            day_num += 1
            current_day = date(2018, month_number, day_num)
            what_week = get_the_week(current_day)
            current_day_str = str(current_day.day)
            day_and_week_dict[what_week].append(current_day)
            hebrew_date_name = '{}{}'.format('A', cell_num+1)
            date_sheet_number = '{}{}'.format('B', cell_num+1)
            hours_sheet_number = '{}{}'.format('C', cell_num+1)
            hebrew_day = get_hebrew_day(current_day.strftime("%A"))
            msg = '{} {} {}'.format('Updating the spreadsheet for', worksheet_tab_title, current_day)
            logging_handler(msg)
            # Updating the Hebrew day field
            selected_worksheet.update_acell(hebrew_date_name, hebrew_day)
            # Updating the day field
            selected_worksheet.update_acell(date_sheet_number, current_day_str)
            # Updating the hours field
            monthly_hours, monthly_minutes = divmod(math.ceil(time_spent_dict[current_day]/60), 60)
            if not monthly_minutes == 0 and not monthly_hours == 0:
                if not monthly_minutes > 10:
                    monthly_minutes = str(monthly_minutes).zfill(2)
                displayed_sum = '{}:{}{}'.format(monthly_hours, monthly_minutes, 'h')
                selected_worksheet.update_acell(hours_sheet_number, displayed_sum)
            cell_num += 1
            if hebrew_day == 'ש':
                cell_num += 1
                insert_weekly_calc_cells(cell_num)
                sum_cell_list.append(cell_num)
        for week in day_and_week_dict:
            weekly_usage = calculate_weekly_usage(week, 'm')
            weekly_usage_list.append(weekly_usage)
            msg = '{} {}{} {} {}'.format('The weekly usage for the', week, 'th week is', weekly_usage, 'minutes')
            logging_handler(msg)
            # Weekends and Workdays
            workdays_usage = calculate_workdays_weekends_usage(week, time_format='m')
            workday_usage_list.append(workdays_usage)
            weekends_usage = calculate_workdays_weekends_usage(week, workdays=False, time_format='m')
            weekend_usage_list.append(weekends_usage)
        # Calculating the weekly amount
        weekly_usage_index = 0
        for summary_cell in sum_cell_list:
            summary_cell_title_location = '{}{}'.format('A', summary_cell)
            summary_cell_location = '{}{}'.format('C', summary_cell)
            workday_cell_location = '{}{}'.format('D', summary_cell)
            weekend_cell_location = '{}{}'.format('E', summary_cell)

            selected_worksheet.update_acell(summary_cell_title_location, 'Work Sum')

            weekly_usage_in_minutes = weekly_usage_list[weekly_usage_index]
            weekly_hours, weekly_minutes = divmod(weekly_usage_list[weekly_usage_index], 60)
            weekly_minutes = str(weekly_minutes).zfill(2)
            weekly_displayed_sum = '{}:{}{}'.format(weekly_hours, weekly_minutes, 'h')

            workday_usage_in_minutes = workday_usage_list[weekly_usage_index]
            workday_hours, workday_minutes = divmod(workday_usage_list[weekly_usage_index], 60)
            workday_minutes = str(workday_minutes).zfill(2)
            workday_displayed_sum = '{}:{}{}'.format(workday_hours, workday_minutes, 'h')

            weekend_usage_in_minutes = weekend_usage_list[weekly_usage_index]
            weekend_hours, weekend_minutes = divmod(weekend_usage_list[weekly_usage_index], 60)
            weekend_minutes = str(weekend_minutes).zfill(2)
            weekend_displayed_sum = '{}:{}{}'.format(weekend_hours, weekend_minutes, 'h')

            displayed_sum = '{} : {}'.format('סך הכל', weekly_displayed_sum)
            selected_worksheet.update_acell(summary_cell_location, displayed_sum)

            displayed_sum = '{} : {}'.format('אמצע שבוע', workday_displayed_sum)
            selected_worksheet.update_acell(workday_cell_location, displayed_sum)

            displayed_sum = '{} : {}'.format('סופ״ש', weekend_displayed_sum)
            selected_worksheet.update_acell(weekend_cell_location, displayed_sum)

            weekly_usage_index += 1
        msg = '{} {}'.format('Finished updating', worksheet_tab_title)
        logging_handler(msg)
