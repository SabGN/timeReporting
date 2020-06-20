import cProfile
import calendar
import io
import pstats
import gspread
from datetime import *
from gspread import Cell, Worksheet
from loguru import logger  # https://gspread.readthedocs.io/en/latest/user-guide.html
# from test_tr import *
from trsetup import *
from trutils import week_start_end_datetime, format_timedelta_hhmm
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
    client = gspread.service_account(filename=CREDENTIALS_FILE) #1
    sh = client.open_by_key(REPORT_SPREADSHEET_ID)# нет запроса
    wks = sh.worksheet("Clockify Projects SabGN")#нет запроса
    #cell_grisha = wks.acell('A1')
    # for x in range(10):
    #     datetime.date(x)

    cell_list = wks.range('A1:A5') #2
    cell_list[0].value = 'Пользователь'#нет запроса
    wks.format("A1:P23",
    {
        "backgroundColor": {
          "red": 13,
          "green": 53,
          "blue": 95
        },
        "horizontalAlignment": "LEFT",
        "borders": {
            "bottom": {
              "style": 'SOLID',
              "width": 1,
              "color": {
                  "red": 231,
                  "green": 55,
                  "blue": 67,
              }
            }
        },
        "textFormat": {
          "foregroundColor": {
            "red": 252,
            "green": 35,
            "blue": 10
          },
          "fontSize": 8,
          "bold": True
        }
    })#3
    wks.update_cells(cell_list)#4
    # print(cell_grisha)
main_work()

