import os
import io
import requests
from docx import Document
import urllib.parse
import json
import os
import pandas as pd
from os.path import exists
import datetime
import calendar
import email_action as email


MONTHS = {'Jan': '01', 'Feb': '02', 'Mar': '03', 'Apr': '04', 'May': '05', 'Jun': '06', 'Jul': '07', 'Aug': '08',
          'Sep': '09', 'Oct': '10', 'Nov': '11', 'Dec': '12'}


'''************************ DATE PARSING FUNCTIONS *****************************'''


def date_eog_than_today(date):
    date_no_hours = str(date_as_datetime(date)).split()[0]
    today = str(datetime.datetime.today()).split()[0]
    if date_no_hours == today or date > datetime.datetime.today():
        return True
    return False


def date_in_csv(date):
    return str(int(date.year)) + '-' + str(int(date.month)) + '-' + str(int(date.day))


def check_msg_date_and_send(msg_date_str, system_name, recalibration_date, location, attachments):
    msg_date = str(date_as_datetime(msg_date_str)).split()[0]
    today = str(datetime.datetime.today()).split()[0]
    if msg_date == today:
        email.send_notification(system_name, recalibration_date, location, attachments)
        return True
    return False


def check_initial_msg(system):
    if is_nan(system['Initial msg']):
        return False
    else:
        return True


def compare_dates(date1, date2):
    if date1 is None or is_nan(date1):
        return 1
    if date2 is None or is_nan(date2):
        return -1
    idx = 0
    date1_list = date_as_list(date1)
    date2_list = date_as_list(date2)
    while idx < 3:
        if date1_list[idx] < date2_list[idx]:
            return 1
        elif date1_list[idx] > date2_list[idx]:
            return -1
        idx += 1
    return 0


def is_nan(num):
    if type(num) != float:
        return False
    return num != num


def add_months(source_date, months):
    month = source_date.month - 1 + months
    year = source_date.year + month // 12
    month = month % 12 + 1
    day = min(source_date.day, calendar.monthrange(year, month)[1])
    return datetime.datetime(year, month, day)


def parse_date(date):
    return_date = []
    splitted = date.split('T')
    y_m_d = splitted[0].split('-')
    for i in range(len(y_m_d)):
        return_date.append(y_m_d[i])
    hour = splitted[1].strip('Z').split(':')
    for i in range(len(hour)):
        return_date.append(hour[i])
    str_date = return_date[0]
    for time in return_date[1:]:
        str_date = str_date + '-' + time
    return str_date


def date_as_list(date):
    if type(date) == str:
        date_list = [time for time in date.split('-')]
    else:
        date_list = date
    return date_list


def date_as_datetime(date):
    if type(date) == list:
        qf5_date_list = [int(time) for time in date]
    elif type(date) == datetime.datetime:
        return date
    elif len(date.split('-')) > 1:
        qf5_date_list = [int(time) for time in date.split('-')]
    elif len(date.split('/')) > 1:
        qf5_date_list = [int(time) for time in date.split('/')]
        return datetime.datetime(qf5_date_list[2], qf5_date_list[1], qf5_date_list[0])

    return datetime.datetime(qf5_date_list[0], qf5_date_list[1], qf5_date_list[2])


def date_preformed(file_name):
    document = Document(file_name)
    for para in document.paragraphs:
        if para.text and para.text[0:4] == 'Date':
            return format_date_from_docx(para.text)


def format_date_from_docx(date_text):
    if len(date_text.split('.')) > 1:
        only_date = date_text.split(':')[1]
        date = [time for time in only_date.strip().split('.')]

    elif len(date_text.split()) == 3:
        date = date_text.split()[2].split('-')

    elif len(date_text.split('_')) > 1:
        only_date = date_text.split('_')[1]
        date = [time for time in only_date.strip().split('-')]

    else:
        only_date_num = date_text.split(':')[1].strip()
        date = only_date_num.split('-')

    if date[1] == '9' and len(date[2]) == 2:
        date[1] = 'Sep'
        date[2] = '20' + date[2]

    if date[1] == 'SEP':
        date[1] = 'Sep'

    date[1].lower().capitalize()

    to_ret = date[2] + '-' + MONTHS[date[1]] + '-' + date[0]
    return to_ret



