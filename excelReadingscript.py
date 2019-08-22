import pandas as pd
from pandas import ExcelWriter
from pandas import ExcelFile
import PySimpleGUI as sg

layout = [[sg.Text('Filename:')],
            [sg.Input(), sg.FileBrowse(button_color=('white','navy'))],
            [sg.Submit(button_color=('white','navy')), sg.Cancel(button_color=('white','navy'))],
            [sg.Output(size=(100,50))]]
            
window = sg.Window("Reinspection Wizard",layout)

while True:
    event,values = window.Read()
    if event is None or event == "Cancel":
        break
    elif event is event == "Submit":
        pathname = values[0]
        print(pathname +"\n")
        #checks if pathname contains AMP, as that would mean a different register has been added which is on a different sheet
        if("AMP" in pathname):
            try:
                sheet = pd.read_excel(pathname, 'App C - Asb Reg - Updated')
            except:
                try:
                    sheet = pd.read_excel(pathname)
                except:
                    print("Pathname error.")
                    continue
        else:
            sheet = pd.read_excel(pathname)
        sheet['Reinspect Date'] = pd.to_datetime(sheet['Reinspect Date'])
        #dates to check between, can change as time goes, <TODO> obtain wanted dates via user input
        start_date = '01-01-2018'
        end_date = '01-01-2020'

        mask = (sheet['Reinspect Date'] > start_date) & (sheet['Reinspect Date'] <= end_date)

        sheet = sheet.loc[mask]
        #if no reinspections
        if(len(sheet)==0):
            print("")
            print("There are no reinspections required at this time.")
            continue
        print("There are a total of "+ str(len(sheet)) +" overdue items for reinspection.")
        print("They are between the dates 01/01/2018 and 01/01/2020.")
        print("These are at the following properties:")
        print("")
        #Runs through dictionary of properties printing out names requring reinsoections if property is already said it continues.
        propertyDict = sheet['Property Name'].to_dict()
        listOfProperties = []
        for amount, properties in propertyDict.items():
            if(properties in listOfProperties):
                continue
            else:
                print(properties)
                listOfProperties.append(properties)
                try:
                    sheetSample = sheet.loc[sheet['Property Name'] == properties]
                    print(sheetSample['Sample Type']+" located:"+ sheetSample['Location'])
                except:
                    sheetSample = sheet.loc[sheet['Property Name'] == properties]
                    sampleCategories = sheetSample['Sample Category'].to_dict()
                    for index, category in sampleCategories.items():
                        locale = sheetSample.loc[sheet['Sample Category'] == category]
                        print(category +" located: \n"+locale['Location of Sample'].to_string(index=False))
                #<TODO> Clean the above line up to maybe print each item sperately without number etc.
            print("")
            
window.Close()