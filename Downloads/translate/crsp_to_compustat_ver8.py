import xlrd
import xlwt
crsp_qusip = []
compustat_quisp = []
correspond_qusip = []
wb = xlrd.open_workbook('CRSP.xlsx')
wb1 = xlrd.open_workbook('Compustat.xlsx')
crsp_table = wb.sheets()[0]
Compustat_table = wb1.sheets()[0]
crsp_nrows = crsp_table.nrows
compustat_nrows = Compustat_table.nrows
for i in range(crsp_nrows):
    if(i == 0):
        continue
    qusip_str = str(crsp_table.cell_value(i,2))
    if(len(qusip_str)<8):
        str1 = (8-len(qusip_str))*"0"
        qusip_str = str1 + qusip_str
    crsp_qusip.append(qusip_str)
print("step1")
for i in range(compustat_nrows):
    if(i==0):
        continue
    compustat_quisp.append(str(Compustat_table.cell_value(i,8))[0:8])
print("step2")
for i in crsp_qusip:
    if(i[0:8] in compustat_quisp):
        correspond_qusip.append(i[0:8])
print(correspond_qusip)
workbook = xlwt.Workbook()
worksheet = workbook.add_sheet('crsp_to_compustat_version8')
for i in range(0,len(correspond_qusip)):
    if(correspond_qusip[i]!='00000000'):
        worksheet.write(i,0,correspond_qusip[i])
workbook.save('crsp_to_compustat_version8.xls')