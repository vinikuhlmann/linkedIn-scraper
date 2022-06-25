import PySimpleGUI as sg
from linkedin_scraper import extrair_conexoes
from relatorio_individual import gerar_relatorio_individual

layout = [[sg.Text('Nome:', size=(7,1)), sg.Input(key='-NOME-', size=(40,1))],
          [sg.Text('Usuário:', size=(7,1)), sg.Input(key='-USUARIO-', size=(40,1))],
          [sg.Text('Senha:', size=(7,1)), sg.Input(key='-SENHA-', password_char='*', size=(40,1)), sg.Button('Mostrar', size=(10,1))],
          [sg.Push(), sg.Button('Executar', size=(15,1)), sg.Push()]]
window = sg.Window('Conexões no LinkedIn', layout)
password_is_hidden = True

while True:
    event, values = window.read()
    if event in [sg.WIN_CLOSED, 'Exit']:
        break
    elif event == 'Mostrar':
        if password_is_hidden:
            window['-SENHA-'].update(password_char='')
            password_is_hidden = False
        else:
            window['-SENHA-'].update(password_char='*')
            password_is_hidden = True
    elif event == 'Executar':
        nome, usuario, senha, perfis_busca_xlsx = values['-NOME-'], values['-USUARIO-'], values['-SENHA-'], "C:/Users/kuhlm/Desktop/LinkedIn Scraper 2.0/Perfis de busca.xlsx"
        #nome, usuario, senha, perfis_busca_xlsx = 'Lucas Wu', 'wu.lucas@protonmail.com', '8aBtOHqtjvYw?s9J', "C:/Users/kuhlm/Desktop/LinkedIn Scraper 2.0/Perfis de busca.xlsx"
        if nome == None or usuario == None or senha == None:
            sg.popup("Preencha todos os campos!")
        else:
            try:
                nome, conexoes_df = extrair_conexoes(nome, usuario, senha, perfis_busca_xlsx)
                gerar_relatorio_individual(nome, conexoes_df)
                sg.popup("Relatório individual gerado")
            except Exception as e:
                sg.popup("Um erro ocorreu!", e)

window.close()
