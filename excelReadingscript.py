import pandas as pd
from pandas import ExcelWriter
from pandas import ExcelFile
import PySimpleGUI as sg

layout = [[sg.Text('Filename:')],
            [sg.Input(), sg.FileBrowse()],
            [sg.Submit(), sg.Cancel()],
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
        print("There are a total of "+ str(len(sheet)) +" overdue items for reinspection.")
        print("They are between the dates 01/01/2018 and 01/01/2020.")
window.Close()

# def ChatBot():
#     layout = [[(sg.Text('This is where standard out is being routed', size=[40, 1]))],
#               [sg.Output(size=(80, 20))],
#               [sg.Multiline(size=(70, 5), enter_submits=True),
#                sg.Button('SEND', button_color=(sg.YELLOWS[0], sg.BLUES[0])),
#                sg.Button('EXIT', button_color=(sg.YELLOWS[0], sg.GREENS[0]))]]

#     window = sg.Window('Chat Window', layout, default_element_size=(30, 2))

#     # ---===--- Loop taking in user input and using it to query HowDoI web oracle --- #
#     while True:
#         event, value = window.Read()
#         if event == 'SEND':
#             print(value)
#         else:
#             break

# ChatBot()
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

