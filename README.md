# Project Manager Tools

Scripts to manipulate Excel spreadsheet containing actuals and forecast for resources in a project.

## Output spreadsheets

As output scripts write a workbook with three spreadsheets:
1.  **Actuals** spreadsheet contains time and cost by resource split by week
2. **Forecast** spreadsheet contains forecast time set in tenrox
3. **Actuals+Forecast** spreadsheet contains actuals and forecast time and cost in the same sheet helping PMM to to project an estimate to complete.
4. **Rates** spreadsheet contains rates assigned by default to resources to calculate their forcast cost
5. **Unapproved** spreadsheet contains data for time sheets submitted by resoources being pending for approval from project manager.

## Invocation

The way to use the script is as following:
`python UsingPandas.py [-h] --fcst FCST --act ACT [--out OUT] [--rates RATES]`

where:
- FCST is Forecast data spreadsheet (.xls or .xlsx file)
- ACT is Actuals data spreadsheet (.xls or .xlsx file)
- OUT is Output Excel file (.xlsx)
- RATES is Rates data spreadsheet (.xls or .xlsx file)


