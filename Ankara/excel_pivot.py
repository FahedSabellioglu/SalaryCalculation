import win32com.client
import os

def Pivot(dirc_path,name):
    Excel= win32com.client.gencache.EnsureDispatch('Excel.Application') # Excel = win32com.client.Dispatch('Excel.Application')
    win32c = win32com.client.constants
    results_path = os.path.join(dirc_path,name)
    Excel.Visible= False
    Excel.DisplayAlerts = False
    wb = Excel.Workbooks.Open(results_path)

    src_data = wb.Worksheets('Outcome')
    data = src_data.UsedRange
    nrows = data.Row + data.Rows.Count -1
    ncols = data.Column + 15

    src_cl1 = src_data.Cells(1,1)
    src_cl2 = src_data.Cells(nrows,ncols)
    pv_src = src_data.Range(src_cl1,src_cl2)

    wb.Sheets.Add (After=wb.Sheets(1))
    Sheet2 = wb.Worksheets(2)
    Sheet2.Name = 'Pivot Sheet'

    cl3=Sheet2.Cells(4,1)
    PivotTargetRange=  Sheet2.Range(cl3,cl3)
    PivotTableName = 'ReportPivotTable'


    PivotCache = wb.PivotCaches().Create(SourceType=win32c.xlDatabase, SourceData=pv_src, Version=win32c.xlPivotTableVersion14)

    PivotTable = PivotCache.CreatePivotTable(TableDestination=PivotTargetRange, TableName=PivotTableName, DefaultVersion=win32c.xlPivotTableVersion14)


    wb.Save()
    wb.Close()

    Excel.Visible= True
    Excel.DisplayAlerts = True
    Excel.Application.Quit()

    Excel = None
    wb = None
    src_data = None
