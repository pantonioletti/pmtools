import sys
if sys.platform == 'win32':
    from waitress import serve

import falcon
import pandas as pd
from openpyxl.writer.excel import save_virtual_workbook
from werkzeug.wrappers import Request
import io
import UsingPandas as up


CONST_FCST_FILE="ForecastFile"
CONST_ACT_FILE="ActualsFile"
CONST_RATE_FILE="RatesFile"
FCST_DATE=4
ACT_DATE='Entry Date'
ACT_IS_APPRVD='Timesheet is Approved'
HTML_XL_CONTENT_TYPE='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'


class PMTools(object):

    """def on_get(self, req, resp):
        wb = upd.process(fcst, act)
        resp.set_header("Content-Disposition", "attachment; filename=\"workbook.xlsx\"")
        resp.context_type = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        resp.data = save_virtual_workbook(wb)
        #resp.stream = save_virtual_workbook(wb)
        resp.status = falcon.HTTP_200
        #resp.body='This site is still building'
    """
    def on_post(self, req, resp):
        wz = Request(req.env)

        files = wz.files
        if len(files) > 1:
            engine = None
            if files[CONST_FCST_FILE].filename is not None and files[CONST_FCST_FILE].filename.lower().endswith(".xls"):
                engine = 'xlrd'
            fcst = pd.read_excel(io.BytesIO(files[CONST_FCST_FILE].stream.read()),engine=engine, header=6, converters={FCST_DATE: up.get_sunday})
            engine = None
            if files[CONST_ACT_FILE].filename is not None and files[CONST_ACT_FILE].filename.lower().endswith(".xls"):
                engine = 'xlrd'
            actuals = pd.read_excel(io.BytesIO(files[CONST_ACT_FILE].stream.read()),engine=engine, header=7, converters={ACT_DATE: up.get_sunday,
                                                         ACT_IS_APPRVD: lambda x: 0 if x == 'Y' else -1})
            if files[CONST_RATE_FILE].filename is not None and files[CONST_RATE_FILE].filename.lower().endswith(".xls"):
                engine = 'xlrd'
            rates = pd.read_excel(io.BytesIO(files[CONST_RATE_FILE].stream.read()),engine=engine, header=0, converters={'Actual Billing Rate': up.parse_rate})
            fplusa = up.process(fcst,actuals,rates)

            resp.set_header("Content-Disposition", "attachment; filename=forecast_actuals.xlsx")
            resp.context_type = 'application/octet-stream;'
            resp.body = save_virtual_workbook(fplusa)
            resp.status = falcon.HTTP_200
        else:
            resp.status = falcon.HTTP_500


def main(app, pm):
    serve(app, host='127.0.0.1',port='9980')

app = falcon.API()
pm = PMTools()
app.add_route('/pmtools',pm)

if sys.platform == 'win32':
    if __name__== '__main__':
        main(app, pm)

