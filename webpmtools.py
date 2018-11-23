from waitress import serve
import falcon
import pandas as pd
from openpyxl.writer.excel import save_virtual_workbook
import UsingPandas as upd
from werkzeug.wrappers import Request
import io
import UsingPandas as up


fcst ="C:\\dev\\projects\\JDA\\internal\\PythonWorkshop\\data\\F CER.xlsx"
act = "C:\\dev\\projects\\JDA\\internal\\PythonWorkshop\\data\\A CER.xlsx"


class PMTools(object):

    def on_get(self, req, resp):
        wb = upd.process(fcst, act)
        resp.set_header("Content-Disposition", "attachment; filename=\"workbook.xlsx\"")
        resp.context_type = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        resp.data = save_virtual_workbook(wb)
        #resp.stream = save_virtual_workbook(wb)
        resp.status = falcon.HTTP_200
        #resp.body='This site is still building'

    def on_post(self, req, resp):
        wz = Request(req.env)

        #frm = wz.form.get('ActualsAndForecast',None)
        files = wz.files
        if len(files) > 1:
            fcst = pd.read_excel(io.BytesIO(files['ForecastFile'].stream.read()),engine='xlrd', header=6, converters={4: up.getSunday})
            actuals = pd.read_excel(io.BytesIO(files['ActualsFile'].stream.read()),engine='xlrd', header=7, converters={'Entry Date': up.getSunday,
                                                         'Timesheet is Approved': lambda x: 0 if x == 'Y' else -1})
            fplusa = up.process(fcst,actuals)
            resp.set_header("Content-Disposition", "attachment; filename=\"forecast_actuals.xlsx\"")
            resp.context_type = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
            resp.data = save_virtual_workbook(fplusa)
            resp.status = falcon.HTTP_200
        else:
            resp.status = falcon.HTTP_500

def main():
    app = falcon.API()
    pm = PMTools()
    app.add_route('/pmtools',pm)
    serve(app, host='127.0.0.1',port='9980')

if __name__== '__main__':
    main()