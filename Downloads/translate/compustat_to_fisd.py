import xlrd
import xlwt
wb_compustat = xlrd.open_workbook("Compustat.xlsx")
wb_fisd = xlrd.open_workbook("FISD.xlsx")
compustat_table = wb_compustat.sheets()[0]
fisd_table = wb_fisd.sheets()[0]
compustat_nrows = compustat_table.nrows
fisd_nrows = fisd_table.nrows
compustat_cusip = []
fisd_cusip = []
correspond_qusip = []
for i in range(compustat_nrows):
    if(i == 0):
        continue
    cusip_str = str(compustat_table.cell_value(i,8))[0:6]
    if(cusip_str != ''):
        compustat_cusip.append(cusip_str)
print(compustat_cusip)
for i in range(fisd_nrows):
    if(i==0):
        continue
    cusip_str = str(fisd_table.cell_value(i,7))[0:6]
    if(cusip_str != ''):
        fisd_cusip.append(cusip_str)
print(fisd_cusip)
for i in compustat_cusip[0:20]:
    if(i[0:6] in fisd_cusip):
        correspond_qusip.append(i)
workbook = xlwt.Workbook()
worksheet = workbook.add_sheet('crsp_to_compustat')
for i in range(0,len(correspond_qusip)):
    if(correspond_qusip[i]!='00000000'):
        worksheet.write(i,0,correspond_qusip[i])
workbook.save('compustat_to_fisd.xls')