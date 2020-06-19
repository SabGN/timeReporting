"""Main module."""
import cProfile
import calendar
import io
import pstats
from datetime import *
import pygsheets
from pygsheets import DataRange
from loguru import logger  # https://gspread.readthedocs.io/en/latest/user-guide.html

# from test_tr import *
from trsetup import *
from trutils import week_start_end_datetime, format_timedelta_hhmm


# TODO doc the func
def profile(fnc):
    """A decorator that used cProfile to profile a function"""

    def inner(*args, **kwargs):
        pr = cProfile.Profile()
        pr.enable()
        retval = fnc(*args, **kwargs)
        pr.disable()
        s = io.StringIO()
        sortby = 'cumulative'
        ps = pstats.Stats(pr, stream=s).sort_stats(sortby)
        ps.print_stats()
        logger.info(s.getvalue())
        return retval

    return inner


@profile
def main_work():
    logger.add("mylog.log", rotation="5 MB")
    logger.debug("================START===================")
    client = pygsheets.authorize(service_file=CREDENTIALS_FILE)

    # Open the spreadsheet and the first sheet.
    sh = client.open_by_key(REPORT_SPREADSHEET_ID)
    wks = sh.worksheet('id', REPORT_SHEET_ID)
    # TODO Refactor line and file formatting
    wks.update_value('D1', date.today().year)
    TASK_COLUMN = 'B'
    week_hours_column = 'E'
    # TODO move next column
    month_sunday = None

    # Collect data from Clockify
    projects_with_tasks = api_session.get_projects_with_tasks(workspace=WORKSPACE)
    users = api_session.get_users(workspace=WORKSPACE)

    # TODO change using yield
    time_entries = []
    curr_line = DATA_LINE - 1

    myranges = []
    report_dict = {}

    # loop for headers
    for project in [*projects_with_tasks]:
        curr_line += 1
        report_dict.update({project: []})
        report_dict[project] += [project.name if project else "No Project", '', '', '']
        proj_line = curr_line
        for task in projects_with_tasks[project] if project else [None]:
            curr_line += 1
            report_dict.update({(project, task): []})
            report_dict[(project, task)] += ['', task.name if task else "No Task", '', '']
        # proj_line = curr_line
        # TODO refactor 4 to some constant
        # myrange = DataRange(left_corner_cell.address, right_corner_cell.address, worksheet=wks)
        myrange = DataRange((proj_line, PROJECT_COLUMN), (curr_line, 4),
                            worksheet=wks).update_borders(True, True, True, True, style='SOLID_THICK')
        # myranges.append(myrange)
    week_column = 5 - 2
    # loop for weeks
    for week in weeks_in_RP:
        start, end = week_start_end_datetime(week)
        time_entries = []
        for user in users:
            time_entries += api_session.get_time_entries(WORKSPACE, user, start, end)
        week_column += 2
        # вернуть дату, соответствующую календарной дате ISO, указанной по году, неделе и дню
        month_monday = date.fromisocalendar(date.today().year, week, 1).month
        if month_sunday != month_monday:
            wks.update_value((1, week_column), calendar.month_name[month_monday])
        month_sunday = date.fromisocalendar(date.today().year, week, 7).month
        if month_monday != month_sunday:
            wks.update_value((1, week_column + 1), calendar.month_name[month_sunday])

        # TODO later month border
        wks.update_value((2, week_column), week)

        for project in [*projects_with_tasks]:
            proj_timedelta, proj_amount = timedelta(minutes=0), 0
            for task in projects_with_tasks[project] if project else [None]:
                task_time_entries = [time_entry for time_entry in time_entries
                                     if (time_entry.project == project) and (
                                         time_entry.task == task)]  # remove extra cond with project
                elapsed_timedelta = sum([time_entry.end - time_entry.start for time_entry in task_time_entries],
                                        timedelta())

                if elapsed_timedelta > timedelta(minutes=1):
                    # TODO measure time for that operation
                    task_time_entries = api_session.api.substitute_api_id_entities(task_time_entries, users,
                                                                                   projects_with_tasks)
                    # TODO don't forget about time_entries with 24h and more
                    elapsed_amount = sum([(time_entry.end - time_entry.start).seconds / 3600 *
                                          time_entry.user.get_hourly_rate(WORKSPACE, time_entry.user).amount
                                          for time_entry in task_time_entries])

                    # TODO keep in mind currency
                else:
                    elapsed_amount = 0

                report_dict[(project, task)] += [format_timedelta_hhmm(elapsed_timedelta), elapsed_amount]
                proj_timedelta += elapsed_timedelta
                proj_amount += elapsed_amount

            # TODO Refactor using Lambda or map or generator
            # model_cell.set_text_format('bold', True)

            report_dict[project] += [format_timedelta_hhmm(proj_timedelta), proj_amount]
            proj_none = [*projects_with_tasks][0]
            task_none = projects_with_tasks[[proj_none][0]][0]
            assert report_dict[proj_none] == report_dict[task_none]
            special_list = []
            for x in projects_with_tasks.keys():
                special_list += [x]
                special_list += [(x, t) for t in projects_with_tasks[x]]
            logger.info('speciallist: ', len(special_list))

    # Памятка
    # A4 = (1,4) = (DATA_LINE, PROJECT_COLUMN)
    # A6 = (1,6) = (proj_line, PROJECT_COLUMN)
    logger.info(curr_line)
    model_range = DataRange((DATA_LINE, PROJECT_COLUMN), (curr_line, week_column + 1), wks)
    logger.info(model_range.range)
    model_range.update_values([report_dict[x] for x in special_list])
    wks.sync()
    logger.warning("==================END=====================")


main_work()
