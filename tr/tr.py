"""Main module."""
import json
import os
import calendar
import pygsheets
from clockifyclient.client import APISession
from clockifyclient.api import APIServer
from datetime import *
from dateutil import tz

#TODO move TIME_ZONE to main user attribute? to ClockifyClient
#TODO get the TimeZone from Google Spreadshhet settings?
TIME_ZONE = tz.tzoffset('MSK', 10800)  # ('Europe/Moscow') https://www.epochconverter.com/timezones

#TODO move to utils
def week_start_end_datetime(week_number, year=date.today().year):
    if week_number > 55 or week_number <= 0:
        raise Exception("недель то всего 54")
    else:
        start_date = date.fromisocalendar(year, week_number, 1)
        end_date = date.fromisocalendar(year, week_number, 7)
        start_datetime = datetime.combine(start_date, datetime.min.time(), tzinfo=TIME_ZONE)
        end_datetime = datetime.combine(end_date, datetime.max.time(), tzinfo=TIME_ZONE)
        return start_datetime, end_datetime

#TODO Rewrite program to remove this function
def comparison_with_none(first, second) -> 'Boolean':
    return first == second

#TODO move to utils
def format_timedelta(td) -> str:
    minutes, seconds = divmod(td.seconds + td.days * 86400, 60)
    hours, minutes = divmod(minutes, 60)
    return '{:d}:{:02d}'.format(hours, minutes)

#SETUP
CLOCKIFY_SETUP_FILE = os.path.abspath('../res/report_setup.json')

with open(CLOCKIFY_SETUP_FILE, 'r', encoding='utf-8') as task:
    setup_dict = json.load(task)

api_key = setup_dict['api_key']
url = "https://api.clockify.me/api/v1/"
api_session = APISession(APIServer(url), api_key)
WORKSPACE = [ws for ws in api_session.get_workspaces() if ws.name == setup_dict['workspace_name']][0]

CREDENTIALS_FILE = os.path.abspath('../res/' + setup_dict['google_credentials_file'])  # имя файла с закрытым ключом
REPORT_SPREADSHEET_ID = setup_dict['report_spreadsheet_id']
REPORT_SHEET_ID = setup_dict['report_sheet_id']
#END SETUP

print("================START===================")
client = pygsheets.authorize(service_file=CREDENTIALS_FILE)

# Open the spreadsheet and the first sheet.
sh = client.open_by_key(REPORT_SPREADSHEET_ID)
wks = sh.worksheet('id', REPORT_SHEET_ID)
#TODO Refactor line and file formatting
wks.update_value('D1', date.today().year)
PROJECT_COLUMN = 'A'
TASK_COLUMN = 'B'
week_hours_column = 'E'
week_money_column = 'F'
DATA_LINE = 4
model_cell = pygsheets.Cell("A1")
# Collect data from Clockify
weeks_in_RP = [9]
week = weeks_in_RP[0]
month = date.fromisocalendar(date.today().year, week, 1).month
wks.update_value('E1', calendar.month_name[month])
wks.update_value('E2', week)
projects_with_tasks = api_session.get_projects_with_tasks(workspace=WORKSPACE)
users = api_session.get_users(workspace=WORKSPACE)

cell = wks.cell('A1')
cell.set_text_format('bold', True)
start, end = week_start_end_datetime(week)
time_entries = []
for user in users:
    time_entries += api_session.get_time_entries(WORKSPACE, user, start, end)
curr_line = DATA_LINE




wks.unlink()
print("Total time entries: ", len(time_entries))


myranges=[]

for project in [*projects_with_tasks]:
    proj_line = curr_line
    wks.update_value(PROJECT_COLUMN + str(curr_line), project.name if project else "No Project")
    curr_line += 1
    proj_timedelta, proj_amount = timedelta(minutes=0), 0

    for task in projects_with_tasks[project] if project else [None]:
        wks.update_value(TASK_COLUMN + str(curr_line), task.name if task else "No Task")
        task_time_entries = [time_entry for time_entry in time_entries
                             if comparison_with_none(time_entry.project, project) and
                             comparison_with_none(time_entry.task, task)]
        elapsed_timedelta = sum([time_entry.end - time_entry.start for time_entry in task_time_entries], timedelta())

        if elapsed_timedelta > timedelta(minutes=1):
            #TODO measure time for that operation
            api_session.api.substitute_api_id_entities(task_time_entries, users, projects_with_tasks)
            elapsed_amount = sum([(time_entry.end - time_entry.start).seconds / 3600 *
                                  time_entry.project.get_hourly_rate(WORKSPACE, time_entry.user).amount
                                  for time_entry in task_time_entries])
        else:
            elapsed_amount = 0
        wks.update_value(week_hours_column + str(curr_line), format_timedelta(elapsed_timedelta))
        wks.update_value("F" + str(curr_line), elapsed_amount)
        proj_timedelta = proj_timedelta + elapsed_timedelta
        proj_amount += elapsed_amount
        curr_line += 1
#formatting
    project_range = wks.range('A1:D5')
    print(type(wks))
    assert type(project_range) == bool
    #project_range = wks.range(week_hours_column + str(proj_line) + ":" + week_money_column + str(proj_line))
    #TODO Refactor using Lambda or map or generator
    #make_bold = lambda cell: (cell.text_format['bold']:= True)
    #map(make_bold, project_range.cells)
    #for cell in project_range:
    #    cell.text_format['bold'] = True
    model_cell.set_text_format('bold', True)
    left_corner_cell = PROJECT_COLUMN + str(proj_line)
    right_corner_cell = week_money_column + str(proj_line)
    myrange = pygsheets.DataRange(
        left_corner_cell, right_corner_cell, worksheet=wks
    )
    myranges.append(myrange)



    wks.update_value(week_hours_column + str(proj_line), format_timedelta(proj_timedelta))
    wks.update_value(week_money_column + str(proj_line), elapsed_amount)

wks.link()
for my_range in myranges:
    my_range.apply_format(model_cell)
    my_range.update_borders(True, True, True, True, style='SOLID_THICK')
print("==================END=====================")
