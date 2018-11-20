import time
from datetime import date, timedelta, datetime
import pandas as pd
import openpyxl as xl
import sys
import os
from openpyxl.styles import PatternFill, Alignment, Font

DF_ANAME = 'Associate Name'
DF_DATE  = 'Date'
DF_HOURS = 'Hours'
DF_COST  = 'Cost'
DF_RATE  = 'Rate'
DF_APPRV = 'Timesheet is Approved'
WS_THOURS = 'Total Hours'
WS_TCOST  = 'Total Cost'
WSN_ACTUALS = 'Actuals'
WSN_FCST = 'Forecast'
WSN_ACTFCST = 'Actuals+Forecast'
WSN_NOTAPPRV = 'Not Approved'

ACTUALS2DROP= ['Client Reporting Unit', 'Client Name', 'Project Name', 'Client PO#', 'Project Manager',
               'AssociateId', 'AssociateType', 'Period Start Date', 'Task Name', 'Billable Hours', 'Billing Rule Name',
               'Billable Amt. In Project Currency', 'Project Currency Name', 'Timesheet is Submitted', 'Timesheet Workflow State',
               'JTRAX Invoiced Status/Ref #', 'Invoice Through Date', 'Service Delivery Location', 'Entry Note',
               'On Hold for Billing', 'On Hold Reason']
FCST2DROP = ['Role', 'Project', 'Actual Hours','User Last Name', 'User First Name']

FORECAST_COLOR="CCFFCC"
TODAY_COLOR="00B050"
ZERO_COST_COLOR="FCD5B4"
DATES_COLOR="DCE6F1"
UNAPPRV_COLOR = "FF3300"

ZERO_FILLER = PatternFill(patternType=None, start_color=ZERO_COST_COLOR, end_color=ZERO_COST_COLOR,fill_type="solid")
FCST_FILLER = PatternFill(patternType=None, start_color=FORECAST_COLOR, end_color=FORECAST_COLOR,fill_type="solid")
TODAY_FILLER = PatternFill(patternType=None, start_color=TODAY_COLOR, end_color=TODAY_COLOR, fill_type="solid")
DATES_FILLER = PatternFill(patternType=None, start_color=DATES_COLOR, end_color=DATES_COLOR, fill_type="solid")
UNAPPRV_FILLER = PatternFill(patternType=None, start_color=UNAPPRV_COLOR, end_color=UNAPPRV_COLOR, fill_type="solid")


'''
Returns the Sunday before <date> if <date> is not already Sunday
'''
def getSunday(pdate):
    ndate = pdate
    if type(pdate) is str :
        if len(pdate) == 10:
            ndate = datetime(int(pdate[-4:]), int(pdate[0:2]), int(pdate[-7:-5]))
        else:
            ndate = None
    elif type(ndate) is date:
        ndate = datetime(ndate.year, ndate.month, ndate.day)
    dow = ndate.weekday()
    if dow < 6:
        ndate = ndate - timedelta(dow+1)
    return ndate


''' 
Build headers for the worksheet. Has a row for resource name, if it contains actuals a column
with resource rate, one column for each week from the first week in project up to the last week in
the forecast. There are two columns after weeks one to sum hours and one to sum cost
'''
def create_headers(ws, title, dcmap, is_actual, max_name_len):
    #Set worksheet name
    ws.title = title
    ws.insert_rows(1, 2)
    curr_col = 1
    font = Font(bold=True)
    ws.cell(1,curr_col,'Resources')
    if is_actual:
        curr_col += 1
        ws.cell(1, curr_col, DF_RATE)
    curr_col += 1
    r = Alignment(text_rotation=90, horizontal='center')
    for curr_date in dcmap.keys():
        c = ws.cell(1,curr_col,curr_date.date())
        c.fill = DATES_FILLER
        c.alignment = r
        c.font = font
        curr_col += 1
    c = ws.cell(1, curr_col, WS_THOURS)
    c.fill = DATES_FILLER
    r = Alignment(wrap_text=True, horizontal='center')
    c.alignment = r
    c.font = font
    ws.column_dimensions[c.column].width = 13
    if is_actual:
        c = ws.cell(1, curr_col+1, WS_TCOST)
        c.fill = DATES_FILLER
        c.alignment = r
        c.font = font
        ws.column_dimensions[c.column].width = 13
        ws.freeze_panes = 'C2'
    else:
        ws.freeze_panes = 'B2'
    ws.column_dimensions['A'].width = max_name_len


"""
Processing for actuals sheet loaded in actuals DataFrame dat will be grouped by Associate Name
and rate. Rate 0 is set to a different color
"""
def actuals_sheet(ws, actuals, date_col, apprv):
    row = 1
    start_col = 3
    actuals_gb = actuals.groupby(by=[DF_ANAME, DF_RATE, DF_DATE]).sum()
    name = None
    rate = -1.0
    for index, group in actuals_gb.iterrows():
        if name != index[0] or rate != index[1]:
            row += 1
            ws.cell(row, 1, index[0])  # name
            ws.cell(row, 2, index[1]) #rate
            if index[1] == 0:
                set_color(ws.iter_cols(min_row=row, max_row=row, min_col=2),ZERO_FILLER)
        if apprv and group[DF_APPRV] < 0:
            set_color(ws.iter_cols(min_row=row, max_row=row, min_col=2), UNAPPRV_FILLER)
        ws.cell(row,start_col + date_col[index[2]], group[DF_HOURS])
        if apprv and group[DF_APPRV] < 0:
            set_color(ws.iter_cols(min_row=row, max_row=row, min_col=2), UNAPPRV_FILLER)
        name = index[0]
        rate = index[1]


"""
Processings for forecast sheet loaded in fcst DataFrame
"""
def forecast_sheet(ws, fcst, date_col):
    row = 1
    start_col = 2
    fcst_gb = fcst.groupby(by=[DF_ANAME, DF_DATE]).sum()
    curr_name = ''
    for index, group in fcst_gb.iterrows():
        name = index[0]
        if name != curr_name:
            row += 1
            ws.cell(row, 1, name)  # name
        ws.cell(row,start_col + date_col[index[1]], group[DF_HOURS])
        curr_name = name


"""
Combining Actuals and Forecast data in one sheet. Actuals data fills the sheet up to last week. Since
last week to last week with forecast data will fill next weeks. Rate for forecast has to be set by 
spradsheet user. Special colors ha been set for rate value 0 and for forecast rows.
"""
def fcst_act_sheet(ws, fcst, actuals, date_col):
    fcst_date = getSunday(date.today())
    row = 1
    start_col = 3
    actuals_gb = actuals.groupby(by=[DF_ANAME, DF_RATE, DF_DATE]).sum()
    name = None
    rate = -1.0
    for index, group in actuals_gb.iterrows():
        if name != index[0] or rate != index[1]:
            row += 1
            if name != index[0]:
                if name is not None:
                    c = ws.cell(row, 1, name)
                    set_color(ws.iter_cols(min_row=row, max_row=row, min_col=1), FCST_FILLER)
                    res_fcst = fcst.loc[(fcst[DF_ANAME] == name) & (fcst[DF_DATE] >= fcst_date)]
                    res_fcst_gb = res_fcst.groupby(by=[DF_ANAME, DF_DATE]).sum()
                    for index_fcst, group_fcst in res_fcst_gb.iterrows():
                        ws.cell(row, start_col + date_col[index_fcst[1]], group_fcst[DF_HOURS])
                    row += 1
                name = index[0]
            rate = index[1]
            ws.cell(row, 1, name)
            ws.cell(row, 2, rate)
            if rate == 0:
                set_color(ws.iter_cols(min_row=row, max_row=row, min_col=2),ZERO_FILLER)
        ws.cell(row,start_col + date_col[index[2]], group[DF_HOURS])
        if group[DF_APPRV] < 0:
            set_color(ws.iter_cols(min_row=row, max_row=row, min_col=2), UNAPPRV_FILLER)
    res_fcst = fcst.loc[fcst[DF_DATE] >= fcst_date]
    res_fcst_gb=res_fcst[~res_fcst[DF_ANAME].isin(actuals[DF_ANAME])].groupby(by=[DF_ANAME, DF_DATE]).sum()
    name = None
    for index_fcst, group_fcst in res_fcst_gb.iterrows():
        if name != index_fcst[0]:
            name = index_fcst[0]
            row += 1
            c = ws.cell(row, 1, name)
            set_color(ws.iter_cols(min_row=row, max_row=row, min_col=1), FCST_FILLER)

        ws.cell(row, start_col + date_col[index_fcst[1]], group_fcst[DF_HOURS])

    col = date_col[fcst_date]+2
    set_color(ws.iter_cols(min_row=2, min_col=col, max_col=col),TODAY_FILLER)


"""
Returns a weekly date sequence
"""
def create_date_seq(start, end):
    seq = dict()
    shift = 0
    for date in pd.date_range(start=start, end=end, freq='W').to_pydatetime():
        seq[date] = shift
        shift += 1
    return seq


"""
Set a filller color for a set of cells
"""
def set_color(iter, fill):
    for sset in iter:
        for cell in sset:
            cell.fill = fill


"""
Add sum formula to all rows with data
"""
def addFormulas(ws, shift, is_actual):
    start_col = 3 if is_actual else 2
    last_col = ws.cell(1, start_col+shift).column
    hours_formula = "=SUM({0}{1}:{2}{1})"
    autosum_formula = "=SUM({0}2:{0}{1})"
    cost_formula = "={0}{1}*B{1}"
    hours = start_col + shift + 1
    cost = hours + 1
    for r in ws.iter_rows(min_row=2):
        c = ws.cell(r[0].row, hours)
        c.number_format = '#,##0.00'
        c.set_explicit_value(hours_formula.format('C',r[0].row,last_col), data_type = 'f')
        if is_actual:
            h_col = c.column
            c = ws.cell(r[0].row,cost)
            c.number_format = '#,##0.00'
            c.set_explicit_value(cost_formula.format(h_col,r[0].row),data_type='f')
    c = ws.cell(ws.max_row+1, hours)
    c.set_explicit_value(autosum_formula.format(c.column,c.row-1),data_type='f')
    c.number_format = '#,##0.00'
    if is_actual:
        c = ws.cell(ws.max_row, cost)
        c.set_explicit_value(autosum_formula.format(c.column,c.row-1),data_type='f')
        c.number_format = '#,##0.00'


start = time.time()
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
                #A new column is created as the concatenation of last and first names
                fcst[DF_ANAME] = fcst['User Last Name']+ ", "+ fcst['User First Name']
                #When there is no resource assign (ie: last name is NA) the Associate Name column is filled with the Role description
                fcst[DF_ANAME].fillna(fcst['Role'], inplace=True)
                #Housekeeping, all unnecessary columns are dropped
                fcst.drop(FCST2DROP, 1)
                #Renaming column
                fcst = fcst.rename(index=str,columns={"Total Booking Hours":DF_HOURS})

                from_date = fcst[DF_DATE].min().date()
                to_date = fcst[DF_DATE].max().date()
                max_name_len = fcst[DF_ANAME].map(len).max()

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
                actuals = pd.read_excel(actxl,header=7, converters={'Entry Date': getSunday,'Timesheet is Approved':lambda x: 0 if x == 'Y' else -1})
                actuals.drop(ACTUALS2DROP, 1)
                actuals = actuals.rename(index=str,columns={"Entry Date":DF_DATE, "Total Hours":DF_HOURS, "Billable Amt. In USD":DF_COST})
                dt = actuals[DF_DATE].min().date()
                if dt < from_date:
                    from_date = dt
                dt = actuals[DF_DATE].max().date()
                if dt > to_date:
                    to_date = dt
                tmp_len = actuals[DF_ANAME].map(len).max()
                if tmp_len > max_name_len:
                    max_name_len = tmp_len
                actuals[DF_RATE] = actuals[DF_COST] / actuals[DF_HOURS]

                wb = xl.Workbook()
                dcmap = create_date_seq(from_date, to_date)
                create_headers(wb.active, WSN_ACTUALS, dcmap, True, max_name_len)
                create_headers(wb.create_sheet(), WSN_FCST, dcmap, False, max_name_len)
                create_headers(wb.create_sheet(), WSN_ACTFCST, dcmap, True, max_name_len)


                actuals_sheet(wb[WSN_ACTUALS], actuals, dcmap, True)
                addFormulas(wb[WSN_ACTUALS], max(dcmap.values()), True)

                forecast_sheet(wb[WSN_FCST], fcst, dcmap)
                addFormulas(wb[WSN_FCST], max(dcmap.values()), False)

                fcst_act_sheet(wb[WSN_ACTFCST], fcst,actuals, dcmap)
                addFormulas(wb[WSN_ACTFCST], max(dcmap.values()), True)

                df_napprv = actuals[actuals[DF_APPRV]<0]
                if df_napprv.size > 0:
                    from_date = df_napprv[DF_DATE].min().date()
                    to_date = df_napprv[DF_DATE].max().date()
                    dcmap = create_date_seq(from_date,to_date)
                    create_headers(wb.create_sheet(), WSN_NOTAPPRV, dcmap, True, max_name_len)
                    actuals_sheet(wb[WSN_NOTAPPRV],df_napprv,dcmap,False)
                    addFormulas(wb[WSN_NOTAPPRV], max(dcmap.values()), True)
                wb.save(out)
                wb.close()
                print('Done!')

            except WindowsError:
                print("Error: file does not exists")
        else:
            print("Error: files are not in OOXML formal .xlsx")
    else:
        print("Error: Invocation should be:\n python JDAProjActuals.py <forecast workbook path> <actuals workbook path>")
    print("Executed in {0:.5f} seconds".format((time.time()-start)))
