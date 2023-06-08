import PySimpleGUI as sg
import pandas as pd
import xlwings as xw

# Read data from Excel file
wb = xw.Book('D:/Pycharm/PycharmProjects/Temperature.xlsx')
sht = wb.sheets('Sheet1')
data_headings = sht.range('A1').expand('right').value
data_values_disp = sht.range('A2').expand('table').value

# Create the GUI
sg.theme("LightGrey6")
layout_frame = [
    [sg.Text('Year', size=(4, 1)), sg.Combo(sorted(set([row[0] for row in data_values_disp])), key='Year', default_value="ALL", expand_x=True),
     sg.Text('Month', size=(8, 1)), sg.Combo(sorted(set([row[1] for row in data_values_disp])), key='Month', default_value="ALL", expand_x=True)],
    [sg.Text('Season', size=(8, 1)), sg.Combo(sorted(set([row[2] for row in data_values_disp])), key='Season', default_value="ALL", expand_x=True),
     sg.Text('Salinity', size=(8, 1)), sg.Input('29.19', key='Salinity', expand_x=True)],
    [sg.Text('Temperature', size=(9, 1)), sg.Input('4.0', key='Temperature')],
    [sg.Text('CHLFa', size=(5, 1)), sg.Input('4', key='CHLFa', expand_x=True)],
    [sg.Text('Area', size=(8, 1)), sg.Combo(sorted(set([row[6] for row in data_values_disp])), key='Area', default_value="ALL", expand_x=True)],
    [sg.Button('Search', key='Search', size=(8, 1))]
]

layout = [
    [sg.Frame("Loc Du Lieu", layout_frame, size=(400, 180))],
    [sg.Table(values=data_values_disp, headings=data_headings,
              background_color="#D1AAAA",
              num_rows=20,
              max_col_width=20,
              justification="Center",
              text_color="Black",
              header_background_color="#9A3D3D",
              header_font=("Arial", 11),
              key='_filestable',
              expand_x=True, expand_y=True)]
]

window = sg.Window("ADVANCED FILTER", layout, size=(480, 540))

while True:
    event, values = window.read()
    if event == sg.WINDOW_CLOSED or event == "Exit":
        break
    if event == "Search":
        year = values['Year']
        month = values["Month"]
        season = values["Season"]
        salinity = values["Salinity"]
        temperature = values["Temperature"]
        CHLFa = values["CHLFa"]
        area = values["Area"]

        # Filter the data based on selected criteria
        filtered_data = pd.DataFrame(data_values_disp, columns=data_headings)

        if year != "ALL":
            filtered_data = filtered_data[filtered_data["Year"] == year]

        if month != "ALL":
            filtered_data = filtered_data[filtered_data["Month"] == month]

        if season != "ALL":
            filtered_data = filtered_data[filtered_data["Season"] == season]

        # Update the table with filtered data
        filtered_values = filtered_data.values.tolist()
        window['_filestable'].update(values=filtered_values)

window.close()


