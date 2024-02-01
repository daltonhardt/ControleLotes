import pandas as pd
from datetime import datetime
import os.path
import matplotlib.pyplot as plt

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
pd.set_option('display.width', 1000)
pd.set_option('max_colwidth', 25)

print('----------------------------------------')
print('|                                      |')
print('|      CONSULTA DE   L O T E S         |')
print('|                                      |')
print('----------------------------------------\n')
print('[INFO]', dias_semana[wkday], hoje, hora)


def addLabels(x, y, alinhado, cor):
    for i in range(len(x)):
        plt.text(i, y[i], y[i], fontsize='large', ha=alinhado, bbox=dict(facecolor=cor, alpha=0.8))


# |---------------------------------|
# |     LENDO O ARQUIVO DE LOTES    |
# |---------------------------------|
if os.path.exists(arq_dados_lotes):
    df_lotes = pd.read_excel(arq_dados_lotes)
    # print('*** dtypes\n', df_lotes.dtypes)
    # print('\n[INFO] dataframe original com os LOTES')
    df_lotes['NF'] = df_lotes['NF'].values.astype(str)
    df_lotes['LOTE'] = df_lotes['LOTE'].values.astype(str)
    df_lotes = df_lotes[['CODIGO', 'PRODUTO', 'QTDE-REAL', 'LOTE', 'VALIDADE', 'DT-ULTIMA-SAIDA', 'SALDO']]

    df_lotes['VALIDADE'] = pd.to_datetime(df_lotes['VALIDADE'], dayfirst=True)  # dayfirst para não trocar dia/mês
    df_lotes['DT-ULTIMA-SAIDA'] = pd.to_datetime(df_lotes['DT-ULTIMA-SAIDA'], dayfirst=True)
    df_lotes['CODIGO'] = df_lotes['CODIGO'].str.strip()
    df_lotes['LOTE'] = df_lotes['LOTE'].str.strip()
    df_lotes['PRODUTO'] = df_lotes['PRODUTO'].str.strip()

    # print(df_lotes)

    while True:
        cod = input('\nQual código (?): ')

        if cod == '' or cod.lower() == 'fim':
            print('\n*** F I M ***')
            break
        elif cod == '?':
            df_produto = df_lotes[['PRODUTO', 'CODIGO']]
            df_produto = df_produto.drop_duplicates()
            df_produto = df_produto.sort_values(by='PRODUTO')
            df_produto = df_produto.to_string(index=False)
            print('Códigos existentes:')
            print(df_produto, '\n')
        else:
            if cod.lower().startswith('her'):
                cod = cod.upper()
            else:
                cod = 'HER' + cod

            df_item = df_lotes.loc[df_lotes['CODIGO'] == cod]
            # df_item = df_item.drop_duplicates()
            # df_item = df_item.sort_values(by='LOTE')
            df_item = df_item.sort_values(by='VALIDADE')
            # print('df_item:\n', df_item)

            if df_item.empty:
                print('*** Código não encontrado. Tente outro!')
                pass
            else:
                df_item_sem_indice = df_item.to_string(index=False)
                print('Lista encontrada:\n', df_item_sem_indice)

                lista_item = df_item.values.tolist()
                titulo = '\n' + cod + '\n' + str(lista_item[0][1] + '\n\n')

                # pegando as datas de primeira e ultima saida do produto
                # df_item['DT-ULTIMA-SAIDA'] = pd.to_datetime(df_item['DT-ULTIMA-SAIDA'], infer_datetime_format=True)

                primeira_saida = df_item['DT-ULTIMA-SAIDA'].min()
                if str(primeira_saida) != 'NaT':
                    data_primeira_saida = primeira_saida.strftime('%d-%m-%Y')
                else:
                    data_primeira_saida = ''

                ultima_saida = df_item['DT-ULTIMA-SAIDA'].max()
                if str(ultima_saida) != 'NaT':
                    data_ultima_saida = ultima_saida.strftime('%d-%m-%Y')
                else:
                    data_ultima_saida = ''

                df_item = df_item.groupby(['LOTE', 'VALIDADE'], sort=False).sum().reset_index()
                # df_item = df_item.groupby('LOTE', sort=False).sum().reset_index()
                # print('df_item_agrupado:\n', df_item)

                # selecionando os dados
                x = []
                y1 = []
                y2 = []
                for i in range(len(df_item)):
                    x.append('\n' + df_item['LOTE'].values[i] + '\n'
                             + str(df_item['VALIDADE'].dt.strftime('%m-%Y').values[i]))
                    y1.append(df_item['QTDE-REAL'].values[i])
                    y2.append(df_item['SALDO'].values[i])

                # print('x  = ', x)
                # print('y1 = ', y1)
                # print('y2 = ', y2)

                # plotando as barras (modo overlaping, uma na frente da outra)
                plt.bar(x, y1, hatch='//', fill=False)
                plt.bar(x, y2, color='b', width=0.7)

                # chamando a função para adicionar o label na barra de QTDE-REAL
                addLabels(x, y1, 'right', 'lightgray')

                # chamando a função para adicionar o label na barra de SALDO
                addLabels(x, y2, 'left', 'red')

                if data_ultima_saida == '':
                    plt.xlabel('\n[ LOTES / VALIDADE ]', fontweight='bold', fontsize=12)
                else:
                    plt.xlabel('\n[ LOTE / VALIDADE ]\nÚltima saída em: ' + data_ultima_saida, fontweight='bold', fontsize=12)

                plt.ylabel('[ Quantidade ]\n', fontweight='bold', fontsize=15)
                plt.title(titulo, fontweight='bold', fontsize='18')
                plt.tick_params(axis='x', labelsize=15)

                plt.tight_layout()
                plt.show()

else:
    print('\n[INFO] Arquivo', arq_dados_lotes, 'não encontrado!')
