import win32com.client as win32
from datetime import datetime,timedelta

now=datetime.now()
yesterday=now+timedelta(days=-1)
sheet_all_name='01-{} (RSBCL)'.format(yesterday.strftime('%d %b'))
sheet_rkl='groupwise_rkl'

win32c = win32.constants

def pivot_table(wb: object, ws1: object, pt_ws: object, ws_name: str, pt_name: str, pt_rows: list, pt_cols: list,
                pt_filters: list, pt_fields: list, na_subtotals: list):
    """
    wb = workbook1 reference
    ws1 = worksheet1
    pt_ws = pivot table worksheet number
    ws_name = pivot table worksheet name
    pt_name = name given to pivot table
    pt_rows, pt_cols, pt_filters, pt_fields: values selected for filling the pivot tables
    na_subtotals = to remove subtotals from specific fields
    """

    # pivot table location
    pt_loc = len(pt_filters) + 5

    # grab the pivot table source data
    pc = wb.PivotCaches().Create(SourceType=win32c.xlDatabase, SourceData=ws1.UsedRange)

    # create the pivot table object
    pt=pc.CreatePivotTable(TableDestination=pt_ws.Range('B5'), TableName=pt_name)
    # pt=pc.CreatePivotTable(TableDestination=f'{ws_name}!R{pt_loc}C1',TableName=pt_name)

    # select the pivot table work sheet and location to create the pivot table
    pt_ws.Select()
    pt_ws.Cells(pt_loc, 1).Select()

    # Sets the rows, columns and filters of the pivot table
    for field_list, field_r in (
    (pt_filters, win32c.xlPageField), (pt_rows, win32c.xlRowField), (pt_cols, win32c.xlColumnField)):
        for i, value in enumerate(field_list):
            pt_ws.PivotTables(pt_name).PivotFields(value).Orientation = field_r
            pt_ws.PivotTables(pt_name).PivotFields(value).RepeatLabels = True
            if value in na_subtotals:
                pt_ws.PivotTables(pt_name).PivotFields(value).Subtotals = tuple(False for _ in range(12))
            pt_ws.PivotTables(pt_name).PivotFields(value).Position = i + 1

    # Sets the Values of the pivot table
    for field in pt_fields:
        pt_ws.PivotTables(pt_name).AddDataField(pt_ws.PivotTables(pt_name).PivotFields(field[0]), field[1],
                                                field[2]).NumberFormat = field[3]
    # for pt_field in pt.PivotFields():
    #     pt_field.RepeatLabels=True
    # Visiblity True or Valse
    pt.TableStyle2 = 'PivotStyleMedium2'
    pt_ws.PivotTables(pt_name).RowAxisLayout(win32c.xlTabularRow)
    pt_ws.PivotTables(pt_name).ShowValuesRow = True
    pt_ws.PivotTables(pt_name).ColumnGrand = True
    pt_ws.UsedRange.Columns.AutoFit()


def run_all_excel(excel,wb):
    try:
        pivot_data = [
            {
                'sheet_name':'Town Industry',
                'pt_rows':['HQ','Town'],
                'pt_cols':['BRAND NAME'],
                'pt_name':'Town Wise Pivot',
                'na_subtotal':[]
            },
            {
                'sheet_name':'Licensee Industry',
                'pt_rows':['HQ','Town','LICENSEE_NAME','Vend','Location','Group Name'],
                'pt_cols':['BRAND NAME'],
                'pt_name':'Licensee Wise Pivot',
                'na_subtotal':['LICENSEE_NAME','Vend','Location','Group Name']
            }
        ]
        for pdata in pivot_data:
            # set worksheet
            ws1 = wb.Sheets(sheet_all_name)
            # Setup and call pivot_table
            ws2_name = pdata['sheet_name']
            wb.Sheets.Add(After=ws1).Name = ws2_name
            ws2 = wb.Sheets(ws2_name)

            pt_name = pdata['pt_name']  # must be a string
            pt_rows = pdata['pt_rows']  # must be a list
            pt_cols = pdata['pt_cols']  # must be a list
            pt_filters = []  # must be a list
            # [0]: field name [1]: pivot table column name [3]: calculation method [4]: number format
            pt_fields=[['Cases','Sum of Cases',win32c.xlSum,'0.000']]

            pivot_table(wb, ws1, ws2, ws2_name, pt_name, pt_rows, pt_cols, pt_filters, pt_fields,pdata['na_subtotal'])
        # wb.Close(SaveChanges=1)
    except:
        excel.Application.Quit()
        raise Exception('Your excel file doesn\'t have required columns')
    
def run_rkl_excel(excel,wb,fields):
    try:
        pivot_data = [
            {
                'sheet_name':'Groupwise_RKL_Lifting(Licensee)',
                'pt_rows':['HQ','District','Category','Group_Name','LICENSEE_NAME','Vends'],
                'pt_cols':[],
                'pt_name':'RKL Groupwise Cases (With Licensee)',
            },
            {
                'sheet_name':'Groupwise_RKL_Lifting',
                'pt_rows':['HQ','District','Category','Group_Name'],
                'pt_cols':[],
                'pt_name':'RKL Groupwise Cases',
             }

        ]
        for pdata in pivot_data:
            # set worksheet
            ws1 = wb.Sheets(sheet_rkl)
            # Setup and call pivot_table
            ws2_name = pdata['sheet_name']
            wb.Sheets.Add(After=ws1).Name = ws2_name
            ws2 = wb.Sheets(ws2_name)

            pt_name = pdata['pt_name']  # must be a string
            pt_rows = pdata['pt_rows']  # must be a list
            pt_cols = pdata['pt_cols']  # must be a list
            pt_filters = []  # must be a list
            # [0]: field name [1]: pivot table column name [3]: calculation method [4]: number format
            # pt_fields=[['Cases','Sum of Cases',win32c.xlSum,'0.000']]
            pt_fields = fields

            na_subtotals=['LICENSEE_NAME','Vend']

            pivot_table(wb, ws1, ws2, ws2_name, pt_name, pt_rows, pt_cols, pt_filters, pt_fields,na_subtotals)
        # wb.Close(SaveChanges=1)
    except:
        excel.Application.Quit()
        raise Exception('Your excel file doesn\'t have required columns')