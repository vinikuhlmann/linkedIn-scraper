from ctypes import alignment
import pandas as pd
from pandas import ExcelWriter

pd.options.mode.chained_assignment = None

def juntar_nomes(df: pd.DataFrame) -> pd.DataFrame:
    
    df.sort_values(by=['Empresa', 'Nome', 'Nome da Conexão'], key=lambda col: col.str.lower(), inplace=True)
    df.reset_index(drop=True, inplace=True)

    # Cria uma lista com um df pra cada diretor
    df_lista = []
    
    for perfil_de_busca in df['Nome'].unique():
        
        fatia_df = df[df['Nome'] == perfil_de_busca]
        perfil_busca_df = fatia_df[['Empresa', 'Nome', 'Cargo', 'LinkedIn']]
        perfil_busca_df = perfil_busca_df.head(1)
        perfil_busca_df.reset_index(drop=True, inplace=True)

        conexao_df = fatia_df[['Nome da Conexão', 'Cargo da Conexão', 'LinkedIn da Conexão']]
        conexao_df.rename(columns={'Nome da Conexão': 'Nome', 'Cargo da Conexão': 'Cargo', 'LinkedIn da Conexão': 'LinkedIn'}, inplace=True)
        conexao_df.reset_index(drop=True, inplace=True)

        novo_df = pd.concat([perfil_busca_df, conexao_df], ignore_index=True)
        novo_df.reset_index(drop=True, inplace=True)
        df_lista.append(novo_df)

    # Concatena os dfs
    concat_df = pd.concat(df_lista)
    concat_df.reset_index(drop=True, inplace=True)
    
    return concat_df

def gerar_relatorio_individual(nome: str, conexoes_df: pd.DataFrame):

    relatorio_xlsx = nome+'.xlsx'

    relatorio_df = juntar_nomes(conexoes_df)
    relatorio_df.drop('Empresa', axis=1, inplace=True)
    relatorio_df.drop_duplicates(inplace=True)
    relatorio_df.reset_index(drop=True, inplace=True)
    relatorio_df['Conhece?'] = pd.NA

    # Formata e escreve os dados
    with ExcelWriter(relatorio_xlsx, engine='xlsxwriter') as writer:

        startrow=5
        startcol=1

        workbook = writer.book
        header_format = workbook.add_format({'bg_color': '#92D050', 'font_color': 'white'})
        perfil_busca_format = workbook.add_format({'bold': True, 'underline': True})
        conexao_format = workbook.add_format({'indent': 1})
        url_format = workbook.add_format({'text_wrap': True})
        workbook.default_url_format = url_format

        relatorio_df.to_excel(writer, sheet_name='Questionário', startrow=startrow+1, startcol=startcol, index=False, header=False)

        worksheet = writer.sheets['Questionário']

        for col, value in enumerate(relatorio_df.columns.values):
            worksheet.write(startrow, col+1, value, header_format)

        conexoes_rows = relatorio_df[relatorio_df['Nome'].isin(conexoes_df['Nome'])].index
        for row in conexoes_rows:
            worksheet.set_row(row=row+startrow+1, cell_format=perfil_busca_format)
        conexoes_rows = relatorio_df[~relatorio_df['Nome'].isin(conexoes_df['Nome'])].index
        for row in conexoes_rows:
            worksheet.set_row(row=row+startrow+1, cell_format=conexao_format)

        worksheet.set_column('A:A', width=2)
        worksheet.set_column('B:D', width=40)

        conexoes_df.to_excel(writer, sheet_name='Dados', index=False)

        worksheet = writer.sheets['Dados']
        worksheet.protect()