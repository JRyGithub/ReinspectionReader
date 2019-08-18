import pandas as pd
from pandas import ExcelWriter
from pandas import ExcelFile

sheet = pd.read_excel("/Users/joshua.ryland/Desktop/UNI/test.xlsx")

print("Column headings:")
print(sheet.columns)

reinspectDate =  sheet['Reinspect Date']
print(reinspectDate)

sheet['Reinspect Date'] = pd.to_datetime(sheet['Reinspect Date'])

start_date = '01-01-2018'
end_date = '01-01-2020'

mask = (sheet['Reinspect Date'] > start_date) & (sheet['Reinspect Date'] <= end_date)

sheet = sheet.loc[mask]
print("There are a total of "+ str(len(sheet)) +" overdue items for reinspection.")
print("They are between the dates 01/01/2018 and 01/01/2020.")
sheet
