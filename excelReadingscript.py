import pandas as pd
from pandas import ExcelWriter
from pandas import ExcelFile
import PySimpleGUI as sg

sg.SetOptions(
    background_color='#FFCD11',
    element_background_color='#FFD948',
    text_color='#5C4B0C',
    button_color=('white','#98834B')
    )
layout = [[sg.Text('Filename:')],
            [sg.Input(background_color='#FFD948'), sg.FileBrowse()],
            [sg.Submit(), sg.Cancel()],
            [sg.Output(size=(100,50),background_color='#FFD948')]]

            
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
        start_date = '01-01-2016'
        end_date = '06-01-2020'

        mask = (sheet['Reinspect Date'] > start_date) & (sheet['Reinspect Date'] <= end_date)

        sheet = sheet.loc[mask]
        #if no reinspections
        if(len(sheet)==0):
            print("")
            print("There are no reinspections required at this time.")
            continue
        print("There are a total of "+ str(len(sheet)) +" overdue items for reinspection.")
        print("They are between the dates "+start_date+" and "+end_date+".")
        print("These are at the following properties:")
        print("")
        #Runs through dictionary of properties printing out names requring reinsoections if property is already said it continues.
        try:
            propertyDict = sheet['Property Name'].to_dict()
        except:
            propertyDict = sheet['Address'].to_dict()
        listOfProperties = []
        for amount, properties in propertyDict.items():
            if(properties in listOfProperties):
                continue
            else:
                print(properties)
                listOfProperties.append(properties)
                try:
                    sheetSample = sheet.loc[sheet['Property Name'] == properties]
                    print(properties+" has "+ str(len(sheetSample))+" overdue samples.")
                    sampleTypes = sheetSample['Sample Type'].to_dict()
                    for index, category in sampleTypes.items():
                        locale = sheetSample.loc[sheet['Sample Type'] == category]
                        print(category +" located: \n"+locale['Location'].to_string(index=False))
                except:
                    sheetSample = sheet.loc[sheet['Property Name'] == properties]
                    sampleCategories = sheetSample['Sample Category'].to_dict()
                    for index, category in sampleCategories.items():
                        locale = sheetSample.loc[sheet['Sample Category'] == category]
                        try:
                            print(category +" located: \n"+locale['Location of Sample'].to_string(index=False))
                        except:
                            print('Type Error')
            print("")
            
window.Close()