import PySimpleGUI as sg
import pandas as pds
import numpy as np


def format_db(file_path, header_row, columns_list, output_name):
    # Import the excel into dataframe only specified cells and with custom names
    dbook = pds.read_excel(file_path, header=header_row,
                           names=['Date', 'Particulars', 'Vch Type', 'Vch No.', 'debit', 'credit'],
                           usecols=columns_list,
                           index_col=False)

    # xlsxwriter engine
    writer = pds.ExcelWriter(f'{output_name}.xlsx', engine='xlsxwriter')

    # Forward fill voucher types and voucher number
    dbook[['Vch Type', 'Vch No.']] = dbook[['Vch Type', 'Vch No.']].ffill()
    dbook['Vch Type - No'] = dbook['Vch Type'] + " - " + dbook['Vch No.']
    dbook = dbook.drop(['Vch Type', 'Vch No.'], axis=1)

    # Combine Debit and Credit Amounts and delete both columns
    dbook['Amount'] = dbook['debit'].fillna(0) - dbook['credit'].fillna(0)
    dbook = dbook.drop(['debit', 'credit'], axis=1)

    # Detect Entries and Narrations
    dbook['Entry Type'] = dbook['Amount'].apply(lambda x: 'narration' if x == 0 else "entry")

    # replace NaT with 0 to make it easier
    dbook['Date'] = dbook['Date'].fillna(0)

    # Create index column and fillna date with 0 for easy detection
    dbook['ind'] = dbook.index

    # create temp1 column to detect change in entries
    dbook['temp1'] = dbook.apply(lambda x: x['ind'] if x['Date'] != 0 else np.NaN, axis=1)

    # forward fill temp1 field to detect same entries
    dbook["temp1"] = dbook['temp1'].fillna(method="ffill")

    # assign incremental sl to each entry (assign entry number)
    i = dbook.temp1
    dbook['Sl. No.'] = i.ne(i.shift()).cumsum()

    # forward fill dates
    dbook['Date'] = dbook['Date'].apply(lambda x: x if x != 0 else np.NaN)
    dbook['Date'] = dbook['Date'].fillna(method="ffill")
    dbook['Date'] = dbook['Date'].dt.strftime('%d/%m/%Y')

    # delete temporary columns
    dbook = dbook.drop(['ind', 'temp1'], axis=1)

    # transfer narrations and sl.no to temporary framework df1
    df1 = dbook[dbook['Entry Type'] == "narration"]
    df1 = df1[["Sl. No.", "Particulars"]]
    df1.columns = ['Sl. No.', 'Narration']
    dbook.drop(dbook[dbook['Entry Type'] == 'narration'].index, inplace=True)

    # merge 2 dataframes
    dbook = pds.merge(dbook, df1, on="Sl. No.", how='left')
    dbook = dbook.drop('Entry Type', axis=1)

    # reorder columns
    dbook = dbook[['Sl. No.', 'Date', 'Particulars', 'Vch Type - No', 'Amount', 'Narration']]

    # export to excel without index
    dbook.to_excel(writer, sheet_name='Sheet1', index=False)

    # create workbook and sheet
    workbook = writer.book
    worksheet = writer.sheets['Sheet1']

    # create formats
    text_wrap = workbook.add_format({'text_wrap': True})
    border_format = workbook.add_format({'bottom': 1})
    drcr_format = workbook.add_format({'num_format': '[blue]#,##0.00_);[red](#,##0.00);-'})

    # set column formats
    worksheet.set_column('A:A', 8)
    worksheet.set_column('B:B', 12)
    worksheet.set_column('C:C', 30, text_wrap)
    worksheet.set_column('D:D', 15)
    worksheet.set_column('E:E', 15, drcr_format)
    worksheet.set_column('F:F', 30, text_wrap)

    # set conditional format for bottom border
    worksheet.conditional_format("A1:F1000000",
                                 {"type": "formula",
                                  "criteria": '=($A1<>$A2)',
                                  "format": border_format})

    # Final save
    writer.save()


# ----------------------------------------Layout Design---------------------------------------- #

sg.theme("DarkGrey6")

file_layout = [[sg.Text("Input File Location")],
               [sg.InputText(size=(60, 1)),
                sg.FileBrowse("Browse", key='-IN-', size=(5, 1), file_types=(("Excel Files", "*.xlsx"),))],
               [sg.Text("Output File Name")],
               [sg.InputText(size=(68, 1), key='-OUT-')]]

columns = ["A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V",
           "W", "X", "Y", "Z"]
detail_layout = [
    [sg.Text("Header Row (The row on which 'Date' is written)"), sg.InputText(key="-HEADER-", size=(27, 1))],
    [sg.Text("Date Column", size=(34, 1)), sg.Combo(columns, key="-DATE-", size=(25, 1))],
    [sg.Text("Particulars Column", size=(34, 1)), sg.Combo(columns, key="-PARTICULARS-", size=(25, 1))],
    [sg.Text("Vch Type Column", size=(34, 1)), sg.Combo(columns, key="-VCHTYPE-", size=(25, 1))],
    [sg.Text("Vch No Column", size=(34, 1)), sg.Combo(columns, key="-VCHNO-", size=(25, 1))],
    [sg.Text("Debit Column", size=(34, 1)), sg.Combo(columns, key="-DEBIT-", size=(25, 1))],
    [sg.Text("Credit Column", size=(34, 1)), sg.Combo(columns, key="-CREDIT-", size=(25, 1))]]

layout = [[sg.Frame("Input", file_layout, font="bold")],
          [sg.Frame("Additional Details", detail_layout, font="bold")],
          [sg.Button("Submit", key="-SUBMIT-", size=(61, 0))]]

window = sg.Window("Tally Daybook", layout)

while True:
    event, values = window.read()
    columns = [values['-DATE-'], values['-PARTICULARS-'], values['-VCHTYPE-'], values['-VCHNO-'], values['-DEBIT-'],
               values['-CREDIT-']]
    if event == sg.WIN_CLOSED or event is None:
        break
    if event == "-SUBMIT-":
        if values['-IN-'] == "":
            sg.Popup("Input file has to be selected. Select a xlsx file exported from Tally")
            continue
        if values['-OUT-'] == "":
            sg.Popup("Output file name has to be provided")
            continue
        if values['-HEADER-'] == '' or values['-HEADER-'] == '0':
            sg.Popup("Header Row Number has to be provided. It is the row on which 'date' appears. Row 5 by default.")
            continue
        if values['-HEADER-'] != '':
            try:
                int(values["-HEADER-"])
            except ValueError:
                sg.Popup("Row Header has to be a integer.")
                continue
        if "" in columns:
            sg.Popup("All Additional details regarding columns must be selected")
            continue
        if len(columns) != len(set(columns)):
            sg.Popup("There cannot be duplicate selections in additional details")
            continue
        format_db(values['-IN-'], int(values['-HEADER-']), ', '.join(columns), values['-OUT-'])
        sg.popup("Formatted copy stored in root directory")

window.close()
