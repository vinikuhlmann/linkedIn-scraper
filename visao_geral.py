import pandas as pd
import xlsxwriter
from os import getcwd
from os import listdir

pd.options.mode.chained_assignment = None

def gerar_visao_geral(dir_relatorios, output_xlsx):

    nomes = listdir(dir_relatorios)
    # Recebe o nome dos relatórios e o caminho
    relatorios = [(path[:-5], dir_relatorios+"\\"+path) for path in listdir(dir_relatorios)]

    relatorio_cols = ['LinkedIn', 'Conhece?']
    perfis_busca_set = set()
    conexoes_lst = []

    for relatorio in relatorios:

        socio = relatorio[0]
        relatorio_xlsx = relatorio[1]
        relatorio_df = pd.read_excel(relatorio_xlsx, sheet_name='Questionário', skiprows=5, usecols=relatorio_cols)
        relatorio_df.dropna(inplace=True)
        conhece = relatorio_df['LinkedIn']

        conexoes_cols = ['Empresa', 'Nome', 'Cargo', 'LinkedIn', 'Nome da Conexão', 'Cargo da Conexão', 'LinkedIn da Conexão']
        conexoes_df = pd.read_excel(relatorio_xlsx, sheet_name='Dados', usecols=conexoes_cols)

        # Armazena os nomes de perfis de busca
        perfis_busca_set |= set(conexoes_df['Nome'].unique().tolist())

        # Marca os contatos conhecidos
        conexoes_df['Sócio'] = socio
        conexoes_df.loc[conexoes_df['LinkedIn'].isin(conhece), 'Conexão direta?'] = 'X'
        conexoes_df.loc[conexoes_df['LinkedIn da Conexão'].isin(conhece), 'Conhece?'] = 'X'
        conexoes_lst.append(conexoes_df)

    vertical1_df = pd.concat(conexoes_lst, ignore_index=True)

    vertical2_df_lst = []
    metadados_df_lst = []

    # Junta os perfis de busca e as conexões em uma única coluna
    for perfil_busca in vertical1_df['Nome'].unique():
        
        fatia_df = vertical1_df[vertical1_df['Nome'] == perfil_busca]

        perfil_busca_df = fatia_df[['Empresa', 'Nome', 'Cargo', 'LinkedIn', 'Conexão direta?', 'Sócio']]
        perfil_busca_df.drop_duplicates(inplace=True)
        perfil_busca_df.rename(columns={'Conexão direta?': 'Conhece?'}, inplace=True)

        conexao_df = fatia_df[['Nome da Conexão', 'Cargo da Conexão', 'LinkedIn da Conexão', 'Sócio', 'Conhece?']]
        conexao_df.rename(columns={'Nome da Conexão': 'Nome', 'Cargo da Conexão': 'Cargo', 'LinkedIn da Conexão': 'LinkedIn'}, inplace=True)

        fatia_vertical2_df = pd.concat([perfil_busca_df, conexao_df], ignore_index=True)
        fatia_vertical2_df.sort_values(['Empresa', 'Nome', 'Sócio'], inplace=True)
        vertical2_df_lst.append(fatia_vertical2_df)

        perfil_busca_metadados_df = fatia_df[['Empresa', 'Nome', 'Cargo', 'LinkedIn']]
        perfil_busca_metadados_df.drop_duplicates(inplace=True)
        
        conexao_df.drop(['Sócio', 'Conhece?'], axis=1, inplace=True)
        conexao_df.drop_duplicates(inplace=True)

        fatia_metadados_df = pd.concat([perfil_busca_metadados_df, conexao_df], ignore_index=True)
        metadados_df_lst.append(fatia_metadados_df)

    # Concatena os dfs
    vertical2_df = pd.concat(vertical2_df_lst, ignore_index=True).drop_duplicates()
    metadados_df = pd.concat(metadados_df_lst, ignore_index=True)

    vertical2_df.drop(['Empresa', 'Nome', 'Cargo'], axis=1, inplace=True)

    vertical2_df.reset_index(drop=True, inplace=True)
    horizontal_df = pd.pivot(vertical2_df, index='LinkedIn', columns='Sócio', values='Conhece?')
    horizontal_df.reset_index(inplace=True)
    horizontal_df.columns.name = None

    visao_geral_df = pd.merge(metadados_df, horizontal_df, how='left', on='LinkedIn')

    idx_empresas_duplicadas = visao_geral_df.dropna(subset='Empresa').drop_duplicates('Empresa')['Empresa'].index
    visao_geral_df.loc[~visao_geral_df.index.isin(idx_empresas_duplicadas) & ~visao_geral_df['Empresa'].isna(), 'Empresa'] = pd.NA

    idx_empresas = list(visao_geral_df.dropna(subset='Empresa')['Empresa'].index)
    visao_geral_df = visao_geral_df.T
    for idx in idx_empresas:
        visao_geral_df.insert(idx, '', pd.NA, allow_duplicates=True)
        for i in range(len(idx_empresas)):
            idx_empresas[i] += 1
    visao_geral_df = visao_geral_df.T
    visao_geral_df['Empresa'] = visao_geral_df['Empresa'].shift(-1)
    visao_geral_df.reset_index(drop=True, inplace=True)

    # Escreve o documento Excel
    with xlsxwriter.Workbook(output_xlsx) as workbook:

        startrow=5
        startcol=1

        # Formatações
        header_format = workbook.add_format({'bg_color': '#92D050', 'font_color': 'white'})
        empresa_format = workbook.add_format({'bold': True, 'underline': True})
        perfil_busca_format = workbook.add_format({'bold': True})
        perfil_busca_url_format = workbook.add_format({'bold': True, 'text_wrap': True})
        conexao_format = workbook.add_format({'indent': 1})
        conexao_url_format = workbook.add_format({'indent': 1, 'text_wrap': True})
        socio_format = workbook.add_format({'bg_color': '#00AF50', 'font_color': 'white'})

        # Criação da worksheet
        worksheet = workbook.add_worksheet()
        worksheet.hide_gridlines(2)
        worksheet.freeze_panes(6, 3)
        worksheet.set_column('A:A', width=2)
        worksheet.set_column('B:B', width=20)
        worksheet.set_column('C:E', width=40)

        # Escreve o cabeçalho
        for col, value in enumerate(visao_geral_df.columns.values):
            if col < 4:
                worksheet.write(startrow, col+1, value, header_format)
            else:
                worksheet.write(startrow, col+1, value, socio_format)
                worksheet.set_column(col+1, col+1, width=10)
        
        # Escreve as empresas
        for row, empresa in visao_geral_df.loc[~visao_geral_df['Empresa'].isnull(), 'Empresa'].iteritems():
            worksheet.write(startrow+row+1, startcol, empresa, empresa_format)
        
        # Escreve o resto dos valores
        for row, value in visao_geral_df.loc[~visao_geral_df['Nome'].isnull(), visao_geral_df.columns != 'Empresa'].iterrows():
            nome, cargo, linkedIn = value[0], value[1], value[2]
            if nome in perfis_busca_set:
                format = perfil_busca_format
                url_format = perfil_busca_url_format
                select_format = perfil_busca_format
            else:
                format = conexao_format
                url_format = conexao_url_format
                select_format = None
            worksheet.write(startrow+row+1, startcol+1, nome, format)
            worksheet.write(startrow+row+1, startcol+2, cargo, format)
            worksheet.write(startrow+row+1, startcol+3, linkedIn, url_format)
            for i in range(3, len(value)):
                if value[i] != value[i]:
                    continue
                worksheet.write(startrow+row+1, startcol+i+1, value[i], select_format)

        level_1_rows = visao_geral_df[~visao_geral_df['Empresa'].notna()].index
        for row in level_1_rows:
            worksheet.set_row(startrow+row+1, options={'level': 1})

        worksheet.protect()

if __name__ == '__main__':
    dir_relatorios = getcwd()+"/Relatorios"
    output_xlsx = getcwd()+"/Visão geral.xlsx"
    gerar_visao_geral(dir_relatorios, output_xlsx)