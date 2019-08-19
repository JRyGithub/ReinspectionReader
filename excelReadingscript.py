import pandas as pd
from pandas import ExcelWriter
from pandas import ExcelFile
import PySimpleGUI as sg

layout = [[sg.Text('Filename:')],
            [sg.Input(), sg.FileBrowse(button_color=('white','navy'))],
            [sg.Submit(button_color=('white','navy')), sg.Cancel(button_color=('white','navy'))],
            [sg.Output(size=(80,20))]]
            
window = sg.Window("Reinspection Wizard",layout, default_element_size=(130,250))

while True:
    event,values = window.Read()
    if event is None or event == "Cancel":
        break
    elif event is event == "Submit":
        pathname = values[0]
        print(pathname)
        sheet = pd.read_excel(pathname)
        sheet['Reinspect Date'] = pd.to_datetime(sheet['Reinspect Date'])

        start_date = '01-01-2018'
        end_date = '01-01-2020'

        mask = (sheet['Reinspect Date'] > start_date) & (sheet['Reinspect Date'] <= end_date)

        sheet = sheet.loc[mask]
        if(len(sheet)==0):
            print("")
            print("There are no reinspections required at this time.")
            continue
        print("There are a total of "+ str(len(sheet)) +" overdue items for reinspection.")
        print("They are between the dates 01/01/2018 and 01/01/2020.")
        print("These are at the following properties:")
        print("")
        propertyDict = sheet['Property Name'].to_dict()
        for amount, properties in propertyDict.items():
            print(properties)
            print("")
window.Close()

# sheet = pd.read_excel("/Users/joshua.ryland/Desktop/UNI/test.xlsx")

# print("Column headings:")
# print(sheet.columns)

# reinspectDate =  sheet['Reinspect Date']
# print(reinspectDate)

# sheet['Reinspect Date'] = pd.to_datetime(sheet['Reinspect Date'])

# start_date = '01-01-2018'
# end_date = '01-01-2020'

# mask = (sheet['Reinspect Date'] > start_date) & (sheet['Reinspect Date'] <= end_date)

# sheet = sheet.loc[mask]
# print("There are a total of "+ str(len(sheet)) +" overdue items for reinspection.")
# print("They are between the dates 01/01/2018 and 01/01/2020.")
# sheet

#Open GUI Window

