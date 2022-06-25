import PySimpleGUI as sg
from visao_geral import gerar_visao_geral

layout = [[sg.Text('Pasta de relatórios:'), sg.Input(key='-DIR-', size=(40,1)), sg.FolderBrowse('Buscar')],
          [sg.Push(), sg.Button('Executar'), sg.Push()]]
window = sg.Window('Organizador de relatórios', layout)

while True:
    event, values = window.read()
    if event in [sg.WIN_CLOSED, 'Exit']:
        break
    elif event == 'Executar':
        try:
            dir, output_xlsx = values['-DIR-'], 'Visão geral.xlsx'
        except:
            sg.popup("Preencha todos os campos!")
        try:
            gerar_visao_geral(dir, output_xlsx)
            sg.popup("Visão geral gerada")
        except Exception as e:
            sg.popup("Um erro ocorreu!", e)

window.close()
