from openpyxl import load_workbook

wb = load_workbook('C:/Users/NB22/Documents/python/10. Oktober.xlsx')

sheet = wb.active
# sheetsemua = []
sheetsemua = list(sheet['D'])
sheetsemua.extend(list(sheet['K']))

# Access all cells in column A
for data in sheetsemua:
    if data.value != None and data.value != 'PART No.' and data.value != 'PART NO' :
        setrip = data.value.rfind('-')
        print(data.value[0:setrip])