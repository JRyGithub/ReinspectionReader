import pandas as pd
import sys
from pandas import ExcelWriter
from pandas import ExcelFile


sheet = pd.read_excel("K:/Jobs/T/Trust Management (TRU)/AMS_Properties/General Church Trust Board/Master Register V1 General Church Trust Board.xlsx")
#West Auckland Trust Services
#K:/Jobs/T/Trust Management (TRU)/AMS_Properties/West Auckland Trust Services Limited/Master V1 West Auckland Trust Services Limited.xlsx
#Trust Investment Property Fund
#K:/Jobs/T/Trust Management (TRU)/AMS_Properties/Trust Investment Property Fund/Master Register V1 Trust Investment Properry Fund.xlsx
#St Stephens & Queen Victoria Trust Board
#K:/Jobs/T/Trust Management (TRU)/AMS_Properties/St Stephens & Queen Victoria Trust Board/Master Register V1 St Stephens & Queen Victoria Trust Board.xlsx
#General Trust Board
#K:/Jobs/T/Trust Management (TRU)/AMS_Properties/General Church Trust Board/Master Register V1 General Church Trust Board.xlsx
#print("Column headings:")
#print(sheet.columns)

#reinspectDate =  sheet['Reinspect Date']
#print(reinspectDate)

sheet['Reinspect Date'] = pd.to_datetime(sheet['Reinspect Date'])

start_date = '01-01-2018'
end_date = '01-01-2020'

mask = (sheet['Reinspect Date'] > start_date) & (sheet['Reinspect Date'] <= end_date)


sheet = sheet.loc[mask]

print("There are a total of "+ str(len(sheet)) +" overdue items for reinspection.")
print("They are between the dates 01/01/2018 and 01/01/2020.")
#Print Data Sheet if there are values of worth.
if(len(sheet) == 0):
    sys.exit
sheet

#Look at a gui system