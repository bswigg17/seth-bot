from openpyxl import Workbook, load_workbook
from openpyxl.styles import PatternFill, Border, Side, Alignment, Font


font = Font(name='Arial',
            size=11,
            bold=True, 
            )

fill_resolved = PatternFill(fill_type='solid',
                            start_color='0000B050',
                            end_color='0000B050')

fill_waiting = PatternFill(fill_type='solid',
                            start_color='A9D08E',
                            end_color='A9D08E')

fill_testing = PatternFill(fill_type='solid',
                            start_color='0000B0F0',
                            end_color='0000B0F0')

side = Side(border_style='thin')


border = Border(top=side,bottom=side, left=side, right=side)

# completed_ids = set()

def week_to_col(week_num):
    return chr(66 + int(week_num))

# def new_sprint(val):
#     if val: 
#         completed_ids = set()


def do_bug_work(sprint, bugs, week):
    # clean_ws(sprint)
    id_hash = make_hash(bugs)
    merge(sprint, id_hash)


# def load_books(book1_name, book2_name):
#     wb_export = load_workbook(book2_name)
#     wb_sprint = load_workbook(book1_name)
#     return wb_sprint, wb_export
    
# def load_sheets(wb_sprint, wb_export):
#     ws = wb_sprint['Defect Status']
#     ws_export = wb_export.active
#     return ws, ws_export

def format_value(cell, val):
    cell.font = font
    if val in {'Ready', 'In Progress', ''}:
        cell.fill = fill_waiting
        cell.value = 'Awaiting Resolution'
    elif val == 'Test':
        cell.fill = fill_testing
        cell.value = 'Testing'
    elif val in {'Done', 'Accepted'}:
        cell.fill = fill_resolved
        cell.value = 'RESOLVED'
    elif val == 'NEW':
        cell.value = val
    cell.border = border
    cell.alignment = Alignment(horizontal='center')

# def clean_ws(ws):
#     """Iterate over row D and delete rows with resolved"""
#     row_number = 1 
#     for status in ws['D']:
#         if status.value == 'RESOLVED':
#             completed_ids.add(ws[f'B{row_number}'].value)
#             ws.delete_rows(row_number)
#             continue
#         row_number += 1

#     """Move Over Row D to Row C"""
#     ws.move_range(f"D2:D{row_number}", cols=-1)


def make_hash(ws_export):
    """Make Set of ID's for quick lookup"""
    id_hash = {}
    i = 1
    for id_ in ws_export['B']:
        #Check for ID Row and None Row
        if id_.value == 'ID':
            i += 1
            continue
        # elif id_.value in completed_ids:
        #     i += 1
        #     continue
        #Check for Duplicate Rows
        # if id_.value in id_hash:
        #     ws.delete_rows(i)
        #     continue
        id_hash[id_.value] = {'Title': ws_export[f'A{i}'].value, 'Status': ws_export[f'D{i}'].value}
        i += 1
    return id_hash


def merge(ws, id_hash):
    i = 1
    for id_ in ws['B']:
        if id_.value == 'ID':
            i += 1
            continue
        elif id_.value in id_hash:
            format_value(ws[f'D{i}'], id_hash[id_.value]['Status'])
            ws[f'E{i}'] = id_hash[id_.value]['Points']
            del id_hash[id_.value]
            i += 1
        else:
            ws.delete_rows(i)

    # ws.delete_rows(i+1)
    if len(id_hash) > 0:
        for id_ in id_hash:
            data = id_hash[id_]
            ws[f'A{i}'] = data['Title']
            ws[f'B{i}'] = id_ 
            format_value(ws[f'C{i}'], 'NEW')
            format_value(ws[f'D{i}'], data['Status'])
            ws[f'E{i}'] = data['Points']
            i += 1

# if __name__ == "__main__":
#     bug_work("./SprintStatus.xlsx", "./Export-11.xlsx")