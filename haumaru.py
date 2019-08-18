import pandas as pd
from pandas import ExcelWriter
from pandas import ExcelFile
#useful pathnames K:/Jobs/H/Haumaru Housing/HAU-A-2001-Haumaru Housing/11-AMP/Haumaru App B Register AMP v9 190325.xlsx", 'App C - Asb Reg - Updated
#"K:/Jobs/C/Colliers International/COL-A-2011 AMP/Colliers Master Asbestos Register v2 190225.xlsx", "Master Register"
sheet = pd.read_excel("K:/Jobs/C/Colliers International/COL-A-2011 AMP/Colliers Master Asbestos Register v2 190225.xlsx", "Master Register")

#sheet.dropna(inplace = True)

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
#Colliers Only
cityList = sheet['City'].value_counts().to_dict()
for city, amount in cityList.items():
    print(city, ":", amount)
#End of Colliers Only
print("They are between the dates 01/01/2018 and 01/01/2020.")
sheet