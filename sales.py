import pyodbc
from collections import defaultdict
from openpyxl import Workbook

x = input('Masukkan Tanggal awal: ')
y = input('Masukkan Tanggal akhir: ')


cnxn_str = ("Driver={SQL Server Native Client 11.0};"
            "Server=192.168.10.250;"
            "Database=SBO_NSI_USD_LIVE;"
            "UID=sa;"
            "PWD=P@ssw0rd;")
            
cnxn = pyodbc.connect(cnxn_str)

cursor = cnxn.cursor()
cursor.execute("SELECT T0.[ItemCode], sum(T0.[CmpltQty]) as 'Receive Qty' FROM OWOR T0 LEFT JOIN WOR1 T1 ON T0.[DocEntry] = T1.[DocEntry] WHERE T0.[CmpltQty] >= 1 and T0.[UserSign] in (19,22) and T0.[Warehouse] = 'WHWIPMF1' and T0.[Status]" + 
"not in ('C') and T0.[PostDate] >= '%s' and T0.[PostDate] <= '%s' Group by T0.[ItemCode], T0.[CmpltQty]" % (x,y))


for i in cursor:
    print(i) 