import pyodbc
from collections import defaultdict
import pandas as pds

file =('C:/Users/NB22/Documents/python/STO 31 OKTOBER 2022.xlsx')
newData = pds.read_excel(file,sheet_name='SALDOAWAL.',engine='openpyxl')
data1 = newData.iloc[:, [4,5]]
print(data1)

# conn = ("Driver={SQL Server Native Client 11.0};"
#             "Server=192.168.10.250;"
#             "Database=SBO_NSI_USD_LIVE;"
#             "UID=sa;"
#             "PWD=P@ssw0rd;")

# con_nsap = pyodbc.connect(conn)

