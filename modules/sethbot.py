from modules.sprint_stuff import do_sprint_work
from modules.bug_stuff import do_bug_work
from openpyxl import Workbook, load_workbook

def week_to_column(week_num):
    return chr(66 + week_num)

def load_workbooks(sprint="./Sprint Status Elera Pay MC (5).xlsx", bugs="./Bugs (NEW).xlsx", backlog="./Stories (NEW).xlsx"):
    wb_sprint = load_workbook(sprint)
    wb_bugs = load_workbook(bugs)
    wb_export = load_workbook(backlog)
    return wb_sprint, wb_bugs, wb_export

def seth_bot(sprint="./SprintStatusNEW.xlsx", bugs="./Export (4).xlsx", backlog="./Export (1).xlsx", week=5, new_sprint=1):
    try:
        wb_sprint, wb_bugs, wb_backlog = load_workbooks(sprint, bugs, backlog)
        sprint_week = week_to_column(week)
        if new_sprint:
            do_sprint_work(wb_sprint.active, wb_backlog.active, sprint_week, wb_sprint.create_sheet('New Sprint'))
        else:
            do_sprint_work(wb_sprint.active, wb_backlog.active, sprint_week)
        # do_bug_work(wb_sprint['Defect Status'], wb_bugs.active, sprint_week)
        wb_sprint.save('Update.xlsx')
        return True
    except Exception as E:
        print(E)
        return False

if __name__ == "__main__":
    seth_bot()