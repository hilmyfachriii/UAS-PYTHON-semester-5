import PySimpleGUI as sg
import pandas as pd
import os

sg.theme('DarkTeal')

EXCEL_FILE = 'Pendaftaran.xlsx'

# Check if the Excel file exists
if os.path.exists(EXCEL_FILE):
    df = pd.read_excel(EXCEL_FILE)
else:
    # If the file doesn't exist, initialize an empty DataFrame
    df = pd.DataFrame(columns=['Nama', 'Tlp', 'Alamat', 'Tgl Lahir', 'Jekel', 'Belajar', 'Menonton', 'Musik'])

layout = [
    [sg.Text('Masukan Data Kamu: ')],
    [sg.Text('Nama', size=(15, 1)), sg.InputText(key='Nama')],
    [sg.Text('No Telp', size=(15, 1)), sg.InputText(key='Tlp')],
    [sg.Text('Alamat', size=(15, 1)), sg.Multiline(key='Alamat')],
    [sg.Text('Tgl Lahir', size=(15, 1)), sg.InputText(key='Tgl Lahir'),
     sg.CalendarButton('Kalender', target='Tgl Lahir', format=('%d-%M-%Y'))],
    [sg.Text('Jenis Kelamin', size=(15, 1)), sg.Combo(['pria', 'wanita'], key='Jekel')],
    [sg.Text('Hobi', size=(15, 1)), sg.Checkbox('Belajar', key='Belajar'),
     sg.Checkbox('Menonton', key='Menonton'),
     sg.Checkbox('Musik', key='Musik')],
    [sg.Submit(), sg.Button('Clear'), sg.Exit()]
]

window = sg.Window('Form pendaftaran', layout)

def clear_input():
    for key in values:
        window[key].update('')
    return None

while True:
    event, values = window.read()
    if event == sg.WIN_CLOSED or event == 'EXIT':
        break
    if event == 'Clear':
        clear_input()
    if event == 'Submit':
        new_data = pd.DataFrame([values], columns=df.columns)
        df = pd.concat([df, new_data], ignore_index=True)
        df.to_excel(EXCEL_FILE, index=False)
        sg.popup('akhirnya jadi')
        clear_input()

window.close()