import xlrd
import xlwt
import from openpyxl import Workbook
wb_crsp = xlrd.open_workbook("CRSP.xlsx")
wb_compustat = xlrd.open_workbook("wb_compustat.xlsx")
wb_fisd = xlrd.open_workbook("FISD.xlsx")
crsp_table = wb_crsp.sheets()[0]
compustat_table = wb_compustat.sheets()[0]
fisd_table = wb_fisd.sheets()[0]
crsp_nrows = crsp_table.nrows
compustat_nrows = compustat_table.nrows
fisd_nrows = fisd_table.nrows
crsp_cusip = []
compustat_cusip = []
fisd_cusip = []
for i in range(crsp_nrows):
    if(i == 0):
        continue
    cusip_str = str(crsp_table.cell_value(i,2))
    if(cusip_str != ''):
        crsp_cusip.append(cusip_str)
for i in range(compustat_nrows):
    if(i == 0):
        continue
    cusip_str = str(compustat_table.cell_value(i,8))
    if(cusip_str != ''):
        compustat_cusip.append(cusip_str)
for i in range(fisd_nrows):
    if(i==0):
        continue
    cusip_str = str(fisd_table.cell_value(i,7))
    if(cusip_str != ''):
        fisd_cusip.append(cusip_str)
wb = Workbook()
sheet.title = 'crsp_to_fisd'
for i in 