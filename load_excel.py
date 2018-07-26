import openpyxl
import sys

wb = openpyxl.load_workbook('/Users/alaragebes/Desktop/solar_data/STATES/Louisiana.xlsx')
change = wb['Cross-Ref']
r = change.max_row

common = []
for i in range(2,r+1):
        common.append(change.cell(row=i, column=3).value)

for i in range(2,r+1):
        common.append(change.cell(row=i, column=5).value)


result = set()
for name in common:
    if name != None :
        result.add(name)

print result
print len(result)
