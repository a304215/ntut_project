import xlrd
from openpyxl import Workbook
dict_crsp = {}
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
    dict_crsp[qusip_str[0:6]] = '0'
print("step1")
for i in range(compustat_nrows):
    if(i==0):
        continue
    compustat_quisp.append(str(Compustat_table.cell_value(i,8)))
print("step2")
for i in compustat_quisp:
    if(i[0:6] in dict_crsp):
        correspond_qusip.append(i[0:8])
wb = Workbook()
sheet  = wb['Sheet']
for i in range(len(correspond_qusip)):
    row = 'A'+str(i+1)
    sheet[row] = correspond_qusip[i]
wb.save('crsp_to_compustat.xlsx')