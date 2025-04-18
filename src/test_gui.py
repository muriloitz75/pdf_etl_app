import os
import sys
import PySimpleGUI as sg

def create_test_gui():
    sg.theme('Default1')  # Define o tema da interface
    
    layout = [
        [sg.Text('Teste da Interface Gráfica', size=(30, 1), font=('Arial', 16), justification='center')],
        [sg.Text('Este é um teste simples da interface gráfica.')],
        [sg.Button('OK', size=(15, 1))]
    ]

    window = sg.Window('Teste GUI', layout, resizable=True, finalize=True)

    while True:
        event, values = window.read()
        
        if event == sg.WIN_CLOSED or event == 'OK':
            break
            
    window.close()

if __name__ == '__main__':
    create_test_gui()
