from pdfreader import convert_pdf_to_txt
import re
import xlsxwriter

#variables
result = {} #results are stored in this dictionary
temp = []
start, stop, adstart, adstop = 0,0,0,0

rawdata = convert_pdf_to_txt('Invoice 1.pdf')
data = rawdata.split('\n')

for i in  data:
    if i != '':
        temp.append(i)
data = temp
for i in data:
    inv = re.search('Invoice No : # (.+)$', i)
    if inv:
        print(inv.group(1))
        result['invoice'] = inv.group(1)
    ord = re.search('Order ID: (.+)$', i)
    if ord:
        print(ord.group(1))
        result['Order id'] = ord.group(1)


for i in range(0,len(data)):
    if re.search('Sold',data[i]):
        start = i
    if re.search('Order ID',data[i]):
        stop = i
    if re.search('Order Date:',data[i]):
        result['order date'] = data[i+1]
    if re.search('Invoice Date:',data[i]):
        result['invoice date'] = data[i+1]
    if re.search('Billing',data[i]):
        adstart = i
    if re.search('Shipping',data[i]):
        adstop = i

address,biladdress = '',''
for i in range(start+1,stop-1):
    address = address + data[i]
for i in range(adstart+1,adstop-1):
    biladdress = biladdress+data[i]

result['address'] = address
result['billing address'] = biladdress
result['company'] = data[stop-1]

for i in result.keys():
    print(i,':',result[i])

#writing to excel:
workbook = xlsxwriter.Workbook('output.xlsx')
worksheet = workbook.add_worksheet()
row = 0
col = 0
for i in result.keys():
    worksheet.write(row,col,i)
    col += 1
col = 0
for i in result.values():
    worksheet.write(row+1,col,i)
    col = col+1
workbook.close()
