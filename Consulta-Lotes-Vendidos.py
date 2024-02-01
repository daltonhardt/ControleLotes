import pandas as pd
from datetime import datetime
import os.path
# from IPython.display import display

# ------------------  DEFINIÇÕES GERAIS ------------------

# ARQUIVO QUE CONTEM OS DADOS DOS PRODUTOS COM LOTE E VALIDADE
dir_dados_lotes = '..\\Lotes\\'
arq_dados_lotes = dir_dados_lotes + 'LOTES.xlsx'

# DEFINIÇÃO DA PASTA ONDE SERÁ MOVIDO O AQUIVO HISTÓRICO DE LOTES NO FINAL DA EXECUÇÃO
destino_lotes = dir_dados_lotes + 'Lotes-historicos'

# ARQUIVO QUE CONTEM OS DADOS DAS NFes DE VENDAS
dir_dados_xml = '..\\XML\\Dados\\'
arq_dados_xml = dir_dados_xml + 'DADOS-XML-SAIDA.dat'

# DEFINIÇÃO DA PASTA ONDE SERÁ MOVIDO O AQUIVO DE DADOS XML NO FINAL DA EXECUÇÃO
destino = dir_dados_xml + 'Saida-processada'

# ARQUIVO QUE CONTEM OS DADOS DE LOTES VENDIDOS
dir_dados_lotes_vendidos = '..\\Vendas\\'
arq_dados_lotes_vendidos = dir_dados_lotes_vendidos + 'LOTES-VENDIDOS.xlsx'

# DEFINIÇÃO DA PASTA ONDE SERÁ CRIADA UMA CÓPIA DO AQUIVO DE LOTES-VENDIDOS PARA HISTÓRICO
destino_lotes_vendidos = dir_dados_lotes_vendidos + 'Vendas-historicos'

# DIA DE HOJE NO FORMATO dd-mm-aaaa
today = datetime.now()
hoje = today.strftime('%d/%m/%Y')
hora = today.strftime('%H:%M:%S')

# DIAS DA SEMANA EM PORTUGUES
dias_semana = ['Segunda-feira', 'Terça-feira', 'Quarta-feira', 'Quinta-feira', 'Sexta-feira', 'Sábado', 'Domingo']
wkday = today.weekday()

# Para mostrar todas as linhas e colunas do Dataframe
pd.set_option('display.max_rows', None)
pd.set_option('display.max_columns', None)

pd.options.display.max_columns = None
pd.options.display.max_rows = None

print('------------------------------------------------------')
print('|                                                    |')
print('|      CONSULTA DE  L O T E S   V E N D I D O S      |')
print('|                                                    |')
print('------------------------------------------------------\n')
print('[INFO]', dias_semana[wkday], hoje, hora)

# |------------------------------------------|
# |     LENDO O ARQUIVO DE LOTES VENDIDOS    |
# |------------------------------------------|
if os.path.exists(arq_dados_lotes_vendidos):
    df_lotes = pd.read_excel(arq_dados_lotes_vendidos)
    # print('*** dtypes\n', df_lotes.dtypes)
    # print('\n[INFO] dataframe original com os LOTES')
    df_lotes['NF'] = df_lotes['NF'].values.astype(str)
    df_lotes['CNPJ/CPF'] = df_lotes['CNPJ/CPF'].values.astype(str)
    df_lotes['LOTE'] = df_lotes['LOTE'].values.astype(str)
    df_lotes['LOTE'] = df_lotes['LOTE'].str.strip()
    df_lotes = df_lotes[['LOTE', 'CODIGO', 'PRODUTO', 'QTDE', 'CNPJ/CPF', 'CLIENTE', 'NF', 'DT-SAIDA']]
    # print(df_lotes)

    while True:
        cod = input('\nQual código (?): ')
        if cod == '' or cod.lower() == 'fim':
            print('\n*** F I M ***\n')
            break
        elif cod == '?':
            df_produtos = df_lotes[['PRODUTO', 'CODIGO']]
            df_produtos = df_produtos.drop_duplicates()
            df_produtos = df_produtos.sort_values(by='PRODUTO')
            df_produtos = df_produtos.to_string(index=False)
            print('Códigos Vendidos:')
            print(df_produtos, '\n')
        else:
            if cod.lower().startswith('her'):
                cod = cod.upper()
            else:
                cod = 'HER' + cod

            df_item = df_lotes.loc[df_lotes['CODIGO'] == cod]
            if df_item.empty:
                print('*** Código não existe. Tente outro!')
                pass
            else:
                # print('dataframe com código que interessa:\n')
                # display(df_item)
                produto = df_item['PRODUTO'].values[0]
                lotes = df_item['LOTE'].to_list()
                # print('lotes =', lotes)
                lotes_unicos = sorted(set(lotes))
                # print('lotes unicos =', lotes_unicos)
                # for x in lotes_unicos:
                # print(x)

                lote = input('Escolha um lote:' + str(lotes_unicos) + ' : ')

                if lote in lotes_unicos:
                    # print('dataframe com os Lotes que interessam:\n')
                    df_item = df_item.loc[df_item['LOTE'] == lote]
                    df_item_agrupado = df_item.groupby(['LOTE', 'CODIGO', 'PRODUTO']).sum('QTDE')
                    df_item_agrupado = df_item_agrupado.rename({'QTDE': 'TOTAL VENDIDO'}, axis=1)
                    print('\n+--------------------------------------------------------------------------------+')
                    print(df_item_agrupado)
                    print('+--------------------------------------------------------------------------------+')
                    df_item = df_item[['QTDE', 'CNPJ/CPF', 'CLIENTE', 'NF', 'DT-SAIDA']]
                    df_item = df_item.rename({'DT-SAIDA': 'DATA'}, axis=1)
                    # df_item.style.set_properties(subset=['CLIENTE'], **{'text-align': 'left'})
                    # print('\nClientes que compraram o lote:', lote, 'do produto', cod, produto, '\n')
                    print('\nClientes que compraram desse lote:\n')
                    print(df_item.to_string(index=False))
                    print('\n+--------------------------------------------------------------------------------+')

                else:
                    print('*** LOTE não existe para este código. Tente outro!')
                    pass

else:
    print('\n[INFO] Arquivo', arq_dados_lotes_vendidos, 'não encontrado!')
