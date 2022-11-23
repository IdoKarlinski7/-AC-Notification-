import arena_linker as al
import pandas as pd
import datetime
import email_action as email
import update_utils as up

SENT_MSG = 'Message Sent'


access_token = al.get_token_header()
if not access_token:
    email.send_notification('All systems', datetime.datetime.today(), error_msg='access token was not claimed')

sys_names = al.get_working_systems(access_token)
if not sys_names:
    email.send_notification('All systems', datetime.datetime.today(), error_msg=
                            'get Neurolyser systems from arena')


arena_systems = al.get_calibration_files(access_token, sys_names)
if type(arena_systems) == str:
    email.send_notification(arena_systems, datetime.datetime.today(), error_msg="get it's files from arena")
al.update_qad(access_token)
al.update_table(arena_systems, access_token)
systems_table = pd.read_csv('Acoustic_Calibration_Master_Table.csv', index_col=0).to_dict()

# CHECK FOR SYSTEMS WITH A NOTIFICATION DATE THAT IS TODAY
info_msg = ''
for system in systems_table:
    location = al.get_location(system)
    calibration_date = up.date_as_datetime(systems_table[system]['Date Preformed'])
    recalibration_date = up.add_months(calibration_date, 6)
    attachments = al.sys_attachments(systems_table[system], system)

    month_ahead_msg = systems_table[system]['month ahead msg']
    two_weeks_ahead_msg = systems_table[system]['two weeks ahead msg']
    week_ahead_msg = systems_table[system]['week ahead msg']
    time_since_due_date = recalibration_date - datetime.datetime.today() + datetime.timedelta(days=1)
    days_to_recal = time_since_due_date.days
    info_msg = info_msg + system + ' Next recalibration date is - ' + str(recalibration_date).split()[0] + '. in ' \
               + str(days_to_recal) + ' Days.\n\n'

    if month_ahead_msg != SENT_MSG and up.check_msg_date_and_send(month_ahead_msg, system, recalibration_date,
                                                                  location, attachments):
        systems_table[system]['month ahead msg'] = SENT_MSG
        if not up.check_initial_msg(systems_table[system]):
            systems_table[system]['Initial msg'] = SENT_MSG

    elif two_weeks_ahead_msg != SENT_MSG and up.check_msg_date_and_send(two_weeks_ahead_msg, system, recalibration_date
                                                                        , location, attachments):
        systems_table[system]['two weeks ahead msg'] = SENT_MSG
        if not up.check_initial_msg(systems_table[system]):
            systems_table[system]['Initial msg'] = SENT_MSG

    elif week_ahead_msg != SENT_MSG and up.check_msg_date_and_send(week_ahead_msg, system, recalibration_date
                                                                   , location, attachments):
        systems_table[system]['week ahead msg'] = SENT_MSG
        if not up.check_initial_msg(systems_table[system]):
            systems_table[system]['Initial msg'] = SENT_MSG

    elif up.is_nan(systems_table[system]['Initial msg']) and 7 > days_to_recal >= 0:
        systems_table[system]['Initial msg'] = SENT_MSG
        email.send_notification(system, recalibration_date, location, attachments)

    elif days_to_recal <= 0 and days_to_recal % 7 == 0:
        email.send_notification(system, recalibration_date, location, attachments)

info_msg = info_msg + '\n Daily Exec ran successfully.'

email.send_info_msg(info_msg)
df = pd.DataFrame(systems_table)
df.to_csv('Acoustic_Calibration_Master_Table.csv')





