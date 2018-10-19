import pandas as pd
import datetime as dt
import openpyxl as xl
import sys
import os
from openpyxl.styles import PatternFill, Alignment, Font


'''
Returns the Sunday before <date> if <date> is not already Sunday
'''
def getSunday(date):
    ndate = date
    if type(date) is str :
        if len(date) == 10:
            ndate = dt.datetime(int(date[-4:]), int(date[0:2]), int(date[-7:-5]))
        else:
            ndate = None
    elif type(ndate) is dt.date:
        ndate = ndate.datetime()
    dow = ndate.weekday()
    if dow < 6:
        ndate = ndate - dt.timedelta(dow+1)
    return ndate


def create_headers(ws, title, dcmap, is_actual, max_name_len):
    ws.title = title
    ws.insert_rows(1, 2)
    curr_col = 1
    f = PatternFill(patternType=None, start_color="DCE6F1", end_color="DCE6F1", fill_type="solid")
    font = Font(bold=True)
    ws.cell(1,curr_col,'Row Labels')
    if is_actual:
        curr_col += 1
        ws.cell(1, curr_col, 'Rate')
    curr_col += 1
    r = Alignment(text_rotation=90, horizontal='center')
    for curr_date in dcmap.keys():
        c = ws.cell(1,curr_col,curr_date)
        c.fill = f
        c.alignment = r
        c.font = font
        curr_col += 1
    c = ws.cell(1, curr_col, "Total Hours")
    c.fill = f
    r = Alignment(wrap_text=True, horizontal='center')
    c.alignment = r
    c.font = font
    if is_actual:
        c = ws.cell(1, curr_col+1, "Total Cost")
        c.fill = f
        c.alignment = r
        c.font = font
        ws.freeze_panes = 'C2'
    else:
        ws.freeze_panes = 'B2'
    ws.column_dimensions['A'].width = max_name_len


def actuals_sheet(ws, actuals, date_col):
    row = 1
    start_col = 3
    actuals_gb = actuals.groupby(by=['Associate Name', 'Rate', 'Entry Date']).sum()
    name = None
    rate = -1.0
    for index, group in actuals_gb.iterrows():
        #print('{0} {1} {2} {3} {4}'.format(index[0], index[1],index[2], group['Total Hours'], group['Billable Amt. In USD']))
        if name != index[0] or rate != index[1]:
            row += 1
            ws.cell(row, 1, index[0])  # name
            ws.cell(row, 2, index[1]) #rate
        ws.cell(row,start_col + date_col[index[2].date()], group['Total Hours'])
        name = index[0]
        rate = index[1]


def forecast_sheet(ws, fcst, date_col):
    row = 1
    start_col = 2
    fcst_gb = fcst.groupby(by=['User Last Name', 'User First Name', 'Date']).sum()
    curr_name = ''
    for index, group in fcst_gb.iterrows():
        name = '{0}, {1}'.format(index[0],index[1])
        if name != curr_name:
            row += 1
            ws.cell(row, 1, name)  # name
        ws.cell(row,start_col + date_col[index[2].date()], group['Total Booking Hours'])
        curr_name = name


def create_date_seq(start, end):
    seq = dict()
    shift = 0
    for date in pd.date_range(start=start, end=end, freq='W'):
        seq[date.date()] = shift
        shift += 1
    return seq


def setColorToZeros(ws, pf):
    for row in ws.iter_rows(min_row=2, min_col=2):
        if row[0].value == 0.0:
            for col in row:
                col.fill = pf


def addFormulas(ws, shift, is_actual):
    start_col = 3 if is_actual else 2
    last_col = ws.cell(1, start_col+shift).column
    hours_formula = "=SUM({0}{1}:{2}{1})"
    cost_formula = "={0}{1}*B{1}"
    hours = start_col + shift + 1
    cost = hours + 1
    for r in ws.iter_rows(min_row=2):
        c = ws.cell(r[0].row, hours)
        c.set_explicit_value(hours_formula.format('C',r[0].row,last_col), data_type = 'f')
        if is_actual:
            h_col = c.column
            c = ws.cell(r[0].row,cost)
            c.set_explicit_value(cost_formula.format(h_col,r[0].row),data_type='f')


if __name__ == '__main__':
    if len(sys.argv) >= 3:
        fcstxl = sys.argv[1]
        actxl = sys.argv[2]
        if len(sys.argv) >= 4:
            out = sys.argv[3]
        else:
            out = "TEST.xlsx"
        if fcstxl.endswith('.xlsx') and actxl.endswith('.xlsx'):
            try:
                os.stat(fcstxl)
                os.stat(actxl)

                '''
                This function will process a spreadsheet where columns are:
                .__________________________________.
                |   Column data      |Column index |
                .__________________________________.
                |User Last Name      |     0       |
                |User First Name     |     1       |
                |Role                |     2       |
                |Project             |     3       |
                |Date                |     4       |
                |Actual Hours        |     5       |
                |Total Booking Hours |     6       |
                .__________________________________.

                Relevant data to be kept is consultant name (columns 0 and 1), date (columns 4) and booking hours (column 6).
                Data is consolidated by week. Weeks start on Sunday.
                First row with data is 9.
                '''
                fcst = pd.read_excel(fcstxl,header=6, converters={4: getSunday})
                fcst.drop(['Role', 'Project', 'Actual Hours'], 1)
                from_date = fcst['Date'].min().date()
                to_date = fcst['Date'].max().date()

                '''
                This function will process a spreadsheet where columns are:
                .__________________________________________________.
                |         Column data               | Column index |
                .__________________________________________________.
                | Client Reporting Unit            	|      0       |
                | Client Name	                    |      1       | 
                | Project Name	                    |      2       |
                | Client PO#	                    |      3       |
                | Project Manager	                |      4       |
                | Associate Name	                |      5       |
                | AssociateId	                    |      6       |
                | AssociateType	                    |      7       |
                | Period Start Date	                |      8       |
                | Entry Date	                    |      9       |
                | Task Name	                        |     10       |
                | Total Hours	                    |     11       |
                | Billable Hours	                |     12       |
                | Billing Rule Name	                |     13       |
                | Billable Amt. In Project Currency	|     14       |
                | Project Currency Name	            |     15       |
                | Billable Amt. In USD	            |     16       |
                | Timesheet is Submitted	        |     17       |
                | Timesheet is Approved           	|     18       |
                | Timesheet Workflow State	        |     19       |
                | JTRAX Invoiced Status/Ref #	    |     20       |
                | Invoice Through Date	            |     21       |
                | Service Delivery Location	        |     22       |
                | Entry Note	                    |     23       |
                | On Hold for Billing	            |     24       |
                | On Hold Reason                    |     25       |
                .__________________________________________________.
                Relevant data to be kept is associate name (column 5), entry date (column 9),  hours (column 11), amount in USD (column 16).
                Amount by itself is not kept, it is used to calculate rate (rate= <amount> / <hours>). Hours are consolidated by week/rate.
                Weeks start on Sunday.
                First row with data is 9. 
                '''
                actuals = pd.read_excel(actxl,header=7, converters={'Entry Date': getSunday})
                actuals.drop(['Client Reporting Unit', 'Client Name', 'Project Name', 'Client PO#', 'Project Manager',
                              'AssociateId', 'AssociateType', 'Period Start Date', 'Task Name', 'Billable Hours',
                              'Billing Rule Name', 'Billable Amt. In Project Currency', 'Project Currency Name',
                              'Timesheet is Submitted', 'Timesheet is Approved', 'Timesheet Workflow State',
                              'JTRAX Invoiced Status/Ref #', 'Invoice Through Date', 'Service Delivery Location',
                              'Entry Note', 'On Hold for Billing', 'On Hold Reason'], 1)
                dt = actuals['Entry Date'].min().date()
                if dt < from_date:
                    from_date = dt
                dt = actuals['Entry Date'].max().date()
                if dt > to_date:
                    to_date = dt
                max_name_len = actuals['Associate Name'].map(len).max()
                actuals['Rate'] = actuals['Billable Amt. In USD'] / actuals['Total Hours']

                wb = xl.Workbook()
                dcmap = create_date_seq(from_date, to_date)
                create_headers(wb.active, 'Actual', dcmap, True, max_name_len)
                create_headers(wb.create_sheet(), 'Forecast', dcmap, False, max_name_len)

                ws = wb['Actual']
                actuals_sheet(ws, actuals, dcmap)
                setColorToZeros(ws, PatternFill(patternType=None, start_color="FCD5B4", end_color="FCD5B4",
                                                fill_type="solid"))
                addFormulas(ws, max(dcmap.values()), True)

                ws = wb['Forecast']
                forecast_sheet(ws, fcst, dcmap)
                addFormulas(ws, max(dcmap.values()), False)

                wb.save(out)
                wb.close()
                print('Done!')

            except WindowsError:
                print("Error: file does not exists")
        else:
            print("Error: files are not in OOXML formal .xlsx")
    else:
        print("Error: Invocation should be:\n python JDAProjActuals.py <forecast workbook path> <actuals workbook path>")
