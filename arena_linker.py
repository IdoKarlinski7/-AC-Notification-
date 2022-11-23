import os
import io
import requests
import traceback as tb
from docx import Document
import urllib.parse
import json
import os
import pandas as pd
from os.path import exists
import datetime
import calendar
import update_utils as up
import email_action as email

USERNAME = 'ericm@fusmobile.com'
PASSWORD = 'PresharedKey1!'
WORKSPACE_ID = '898485829'
LOGIN_INFO = {'email': USERNAME, 'password': PASSWORD, 'workspaceId': WORKSPACE_ID}


'''************************ AUXILIARY FUNCTIONS *****************************'''


def download_file(access_token, file_name, file_guid):
    file_request = 'https://api.arenasolutions.com/v1/files/' + file_guid + '/content'
    response = requests.get(file_request, headers=access_token)
    if response.status_code == 200:
        f = open(file_name, 'x')
        f.close()
        file = open(file_name, 'wb')
        file.write(response.content)
        file.close()
    else:
        return None


def parse_title(title):
    split_title = title.split()
    sys_name = split_title[-1]
    sys_name_split = sys_name.split('_')
    sys_name = sys_name_split[-1]
    return sys_name


def qf05_file_name(system):
    return 'QF-00005 Acoustic Power Calibration Form ' + system + '.docx'


def qf24_file_name(system):
    return 'QF-00024 Acoustic Power Calibration Form ' + system + '.docx'


def log_file_name(system):
    return 'log files - ' + system


def update_calibration_date(system, access_token, sys_name):
    try:
        file_name = qf05_file_name(sys_name)
        if exists(file_name):
            os.remove(file_name)
        download_file(access_token, file_name, system['QF5 guid'])
        system['Date Preformed'] = up.date_preformed(file_name)
        return system
    except FileNotFoundError:
        exception = tb.format_exc()
        email.send_notification(sys_name, datetime.datetime.today(), error_msg=exception)
        return None


def update_qad(access_token):
    files_request = 'https://api.arenasolutions.com/v1/files?title=QAD-00007_Neurolyser XR Traceability_ver_1.3_clean'
    response = requests.get(files_request, headers=access_token)
    if response.status_code == 200:
        guid = response.json()['results'][0]['guid']
        name = 'QAD-00007_Neurolyser XR Traceability.docx'
        content_req = 'https://api.arenasolutions.com/v1/files/' + guid + '/content'
        resp = requests.get(content_req, headers=access_token)
        if resp.status_code == 200:
            if exists(name):
                os.remove(name)
            f = open(name, 'x')
            f.close()
            file = open(name, 'wb')
            file.write(resp.content)
            file.close()


def update_attachments(systems, access_token):
    for sys in systems:
        if 'QF24 guid' in systems[sys].keys():
            file_name = qf24_file_name(sys)
            if exists(file_name):
                os.remove(file_name)
            download_file(access_token, file_name, systems[sys]['QF24 guid'])
        '''if 'log guid' in systems[sys].keys():
            file_name = log_file_name(sys)
            if exists(file_name):
                os.remove(file_name)
            download_file(access_token, log_file_name(sys), systems[sys]['log guid'])'''


''' CALLED FROM DAILY SCRIPT'''


def sys_attachments(system, sys_name):
    attachments = [qf05_file_name(sys_name)]
    if not up.is_nan(system['QF24 guid']):
        name = qf24_file_name(sys_name)
        attachments.append(name)
    '''if not up.is_nan(system['log guid']):
        name = log_file_name(system)
        attachments.append(name)'''
    return attachments


def get_location(system_name):
    try:
        document = Document('QAD-00007_Neurolyser XR Traceability.docx')
        table = document.tables[1]
        data = []
        keys = None
        for i, row in enumerate(table.rows):
            text = (cell.text for cell in row.cells)
            if i == 0:
                keys = tuple(text)
                continue
            row_data = dict(zip(keys, text))
            data.append(row_data)
        for item in data:
            if item['System SN'].find(system_name) != -1:
                return item['Location ']
    except Exception:
        exception = tb.format_exc()
        email.send_notification(system_name, datetime.datetime.today(), error_msg=exception)
        return None


def get_latest(forms, sys_name, access_token):
    files_dict = {}
    final_name = qf05_file_name(sys_name)
    for i, form in enumerate(forms):
        name = str(i) + ' ' + final_name
        files_dict[name] = form
        download_file(access_token, name, form['guid'])

    file_names = [key for key in files_dict.keys()]
    if not file_names:
        return
    latest = file_names[0]
    if len(file_names) > 1:
        for file in file_names[1:]:
            latest_date = [time for time in up.date_preformed(latest).split('-')]
            date = [time for time in up.date_preformed(file).split('-')]
            if up.compare_dates(latest_date, date) == 1:
                os.remove(latest)
                latest = file
            else:
                os.remove(file)

    latest_file = files_dict[latest]
    if exists(final_name):
        latest_date = [time for time in up.date_preformed(latest).split('-')]
        existing_date = [time for time in up.date_preformed(final_name).split('-')]
        if up.compare_dates(existing_date, latest_date) == 1:
            os.remove(final_name)
            os.rename(latest, final_name)
        else:
            os.remove(latest)
    else:
        os.rename(latest, final_name)
    return latest_file


'''******************************* SCRIPT FUNCTIONS - IN SEQUENCE ******************************'''


def get_token_header():
    response = requests.post('https://api.arenasolutions.com/v1/login', json=LOGIN_INFO)
    if response.status_code == 200:
        j_response = response.json()
        return {'arena_session_id': j_response['arenaSessionId']}
    else:
        return None


def get_working_systems(access_header):
    systems = {}
    req = 'https://api.arenasolutions.com/v1/items?criteria='
    criteria= '[ [ { "attribute": "category.name", "value": "DHR Neurolyser XR System", "operator": "IS_EQUAL_TO" } ] ]'
    encoded_req = req + urllib.parse.quote(criteria)
    response = requests.get(encoded_req, headers=access_header)
    if response.status_code == 200:
        for result in response.json()['results']:
            if result['lifecyclePhase']['name'] == 'Design' and result['number'].startswith('NXR100-000'):
                systems[result['number']] = {'guid': result['guid']}
        return systems
    else:
        return None


def get_calibration_files(access_header, systems):
    req = 'https://api.arenasolutions.com/v1/items/'
    for sys in systems:
        files = systems[sys]['guid'] + '/files'
        files_request = req + files
        response = requests.get(files_request, headers=access_header)
        if response.status_code != 200:
            return sys
        qf5s = []
        qf24 = None

        for result in response.json()['results']:
            if result['file']['title'].find('Log files') != -1 or result['file']['title'].find('Log Files') != -1:
                systems[sys]['log guid'] = result['file']['guid']
                systems[sys]['log date'] = up.parse_date(result['file']['creationDateTime'])

            if result['file']['title'].find('QF-00005') != -1 and result['file']['format'] == 'docx':
                qf5s.append(result['file'])

            if result['file']['title'].find('QF-00024') != -1 and result['file']['format'] == 'docx':
                if not qf24 or up.compare_dates(qf24['creationDateTime'], result['file']['creationDateTime']) == 1:
                    qf24 = result['file']

        latest_qf5 = get_latest(qf5s, sys, access_header)
        if latest_qf5:
            systems[sys]['QF5 guid'] = latest_qf5['guid']
            systems[sys]['QF5 date'] = up.parse_date(latest_qf5['creationDateTime'])
        else:
            systems[sys]['QF5 guid'] = None

        if qf24:  # NOT EVERY SYSTEM HAS A QF24
            systems[sys]['QF24 guid'] = qf24['guid']
            systems[sys]['QF24 date'] = up.parse_date(qf24['creationDateTime'])
    return systems


def update_table(arena_systems, access_header):
    current_systems = pd.read_csv('Acoustic_Calibration_Master_Table.csv', index_col=0).to_dict()

    operating_systems = {sys: arena_systems[sys] for sys in arena_systems if arena_systems[sys]
                         ['QF5 guid']}
    for sys_number in operating_systems:
        if sys_number not in current_systems.keys() or up.compare_dates(current_systems[sys_number]['QF5 date'],
                                                                        operating_systems[sys_number]['QF5 date']) >= 0:
            current_systems[sys_number] = update_calibration_date(operating_systems[sys_number], access_header,
                                                                  sys_number)
            update_attachments(operating_systems, access_header)
            schedule_msg(current_systems[sys_number])

    df = pd.DataFrame(operating_systems)
    df.to_csv('Acoustic_Calibration_Master_Table.csv')


def schedule_msg(system):
    if len(system['Date Preformed'].split('-')) > 1:
        date_list = [time for time in system['Date Preformed'].split('-')]
        calibration_date = datetime.datetime(int(date_list[0]), int(date_list[1]), int(date_list[2]))
    else:
        date_list = [time for time in system['Date Preformed'].split('/')]
        calibration_date = datetime.datetime(int(date_list[2]), int(date_list[1]), int(date_list[0]))

    recalibration_date = up.add_months(calibration_date, 6)
    first_msg_date = up.add_months(recalibration_date, -1)
    second_msg_date = recalibration_date - datetime.timedelta(days=14)
    third_msg_date = recalibration_date - datetime.timedelta(days=7)
    system['Initial msg'] = None

    if up.date_eog_than_today(first_msg_date):
        system['month ahead msg'] = up.date_in_csv(first_msg_date)
    else:
        system['month ahead msg'] = 'Message Sent'

    if up.date_eog_than_today(second_msg_date):
        system['two weeks ahead msg'] = up.date_in_csv(second_msg_date)
    else:
        system['two weeks ahead msg'] = 'Message Sent'

    if up.date_eog_than_today(third_msg_date):
        system['week ahead msg'] = up.date_in_csv(third_msg_date)
    else:
        system['week ahead msg'] = 'Message Sent'

'''
if __name__ == '__main__':
    header = get_token_header()
    print('got header')
    syss = get_working_systems(header)
    print('got systems from arena')
    sys_w_files = get_calibration_files(header, syss)
    print('got calibration files')
    update_table(sys_w_files, header)
    print('table is updated')
'''
