import pandas as pd
from datetime import datetime
import os.path
import shutil
import sys
from openpyxl import load_workbook

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

# print('\n[INFO] Lotes, Validades e Saídas dos produtos')
# print('[INFO]', dias_semana[wkday], hoje, 'às', hora)

# Para mostrar todas as linhas e colunas do Dataframe
pd.set_option('display.max_rows', None)
pd.set_option('display.max_columns', None)
pd.set_option('display.width', 1000)
pd.set_option('max_colwidth', 25)

# |---------------------------------|
# |     LENDO O ARQUIVO DE LOTES    |
# |---------------------------------|
if os.path.exists(arq_dados_lotes):
    df_lotes = pd.read_excel(arq_dados_lotes)

    df_lotes['NF'] = df_lotes['NF'].values.astype(str)
    df_lotes['LOTE'] = df_lotes['LOTE'].values.astype(str)
    df_lotes['LOTE'] = df_lotes['LOTE'].str.strip()
    df_lotes['VALIDADE'] = pd.to_datetime(df_lotes['VALIDADE'], dayfirst=True)  # dayfirst para não trocar dia/mês
    df_lotes['DT-ULTIMA-SAIDA'] = pd.to_datetime(df_lotes['DT-ULTIMA-SAIDA'], dayfirst=True)
    df_lotes['CODIGO'] = df_lotes['CODIGO'].str.strip()
    df_lotes['PRODUTO'] = df_lotes['PRODUTO'].str.strip()

    df_lotes = df_lotes.sort_values(by='VALIDADE')
    # print('\n[INFO] dataframe LOTES ordenado por VALIDADE\n', df_lotes)

else:  # SE ARQUIVO NÃO EXISTE
    sys.exit('*** [INFO] Não foi encontrado o arquivo ' + arq_dados_lotes)

# TRANSFORMANDO O DATAFRAME EM UMA LISTA
lista_lotes = df_lotes.values.tolist()
# print('Lista original dos LOTES:\n', lista_lotes)

# |--------------------------------|
# |    LENDO OS DADOS DE VENDAS    |
# |--------------------------------|

dados_vendas = []
df_vendas = pd.DataFrame(dados_vendas)

if os.path.exists(arq_dados_xml):
    df_vendas = pd.read_csv(arq_dados_xml, sep=';')
    # print('\n[INFO] dataframe com os dados de VENDAS')
    df_vendas['NF'] = df_vendas['NF'].values.astype(str)
    df_vendas['CNPJ/CPF'] = df_vendas['CNPJ/CPF'].values.astype(str)
    # print(df_vendas)
else:
    print('\n[INFO] Não existe nenhum arquivo de vendas para processar. Tente mais tarde!\n')
    quit()

# TRANSFORMANDO O DATAFRAME EM UMA LISTA
lista_vendas = df_vendas.values.tolist()
# print('Lista com os dados de VENDAS', lista_vendas)

print('[INFO]', datetime.now().strftime('%H:%M:%S'), 'atualizando arquivo LOTES_VENDIDOS.xlsx')
lista_lote_cliente = []
lista_temp = []

# LOOP SOBRE A LISTA DE ITENS VENDIDOS
for i in range(len(lista_vendas)):
    data_vendido = lista_vendas[i][0]
    nf_vendido = lista_vendas[i][1]
    item_vendido = lista_vendas[i][2]
    produto_vendido = lista_vendas[i][3]
    qtd_item_vendido = int(lista_vendas[i][4])
    cnpjcpf_vendido = lista_vendas[i][5]
    cliente_vendido = lista_vendas[i][6]

    # print('VENDA', item_vendido, qtd_item_vendido)

    # LOOP SOBRE A LISTA DE LOTES
    for j in range(len(lista_lotes)):
        item_lote = lista_lotes[j][2]
        if item_vendido == item_lote:
            # print('linha da lista de itens vendidos', lista_vendas[i])
            lote_num = lista_lotes[j][6]
            lote_val = lista_lotes[j][7]
            lote_ultima_saida = lista_lotes[j][8]
            lote_qtd = lista_lotes[j][9]
            # print('      LOTE', lote_num, lote_val, lote_qtd)

            # VERIFICA SE TEM QUANTIDADE SUFICIENTE DENTRO DO LOTE
            if lote_qtd >= qtd_item_vendido:
                # print('      Tem quantidade suficiente')
                lista_lotes[j][8] = hoje  # Coluna=DT-ULTIMA-SAIDA
                lista_lotes[j][9] = lote_qtd - qtd_item_vendido  # Coluna=SALDO
                # print('      consumindo', qtd_item_vendido, 'do lote', lote_num)
                # print('      sobrou', lista_lotes[j][9], 'no lote', lote_num)
                # print('      adiciona item na lista_lote_cliente')
                lista_temp = [data_vendido, nf_vendido, cnpjcpf_vendido, cliente_vendido, item_vendido, produto_vendido,
                              qtd_item_vendido, lote_num, lote_val]
                lista_lote_cliente.append(lista_temp)
                # print('      ', lista_lote_cliente)
                break

            else:
                if lote_qtd > 0:
                    # print('      Lote insuficiente')
                    qtd_item_vendido = qtd_item_vendido - lote_qtd
                    lista_lotes[j][8] = hoje  # Coluna=DT-ULTIMA-SAIDA
                    lista_lotes[j][9] = 0  # Coluna=SALDO
                    # print('      consumindo', lote_qtd, 'do lote', lote_num)
                    # print('      sobrou', lista_lotes[j][9], 'no lote', lote_num)
                    # print('      adiciona item na lista_lote_cliente')
                    lista_temp = [data_vendido, nf_vendido, cnpjcpf_vendido, cliente_vendido, item_vendido,
                                  produto_vendido, lote_qtd, lote_num, lote_val]
                    lista_lote_cliente.append(lista_temp)
                    # print('      ', lista_lote_cliente)

            # print('final loop IF')

# print('\n[INFO] atualizando arquivos EXCEL...\n')

###########################################
# PROCESSANDO A LISTA DE LOTES ATUALIZADA #
###########################################
# print('Lista LOTES atualizada')
# print(lista_lotes)

# CONVERTENDO A LISTA EM UM DATAFRAME
df_lotes_new = pd.DataFrame(lista_lotes, columns=['DT-ENTRADA', 'NF', 'CODIGO', 'PRODUTO', 'QTDE', 'QTDE-REAL', 'LOTE',
                                                  'VALIDADE', 'DT-ULTIMA-SAIDA', 'SALDO'])

df_lotes_new['DT-ENTRADA'] = pd.to_datetime(df_lotes_new['DT-ENTRADA'], dayfirst=True)
df_lotes_new['VALIDADE'] = pd.to_datetime(df_lotes_new['VALIDADE'], dayfirst=True)
df_lotes_new['DT-ULTIMA-SAIDA'] = pd.to_datetime(df_lotes_new['DT-ULTIMA-SAIDA'], dayfirst=True)
df_lotes_new['LOTE'] = df_lotes_new['LOTE'].str.strip()
df_lotes_new['PRODUTO'] = df_lotes_new['PRODUTO'].str.strip()
# print('df_lotes_new:\n', df_lotes_new)

# SALVANDO OS DADOS DOS LOTES CONFORME DATAFRAME `df_lotes_new`
# PRIMEIRO RENOMEIA O ARQUIVO ATUAL DE LOTES E DEPOIS MOVE PARA A PASTA `Lotes-historicos'
fonte = arq_dados_lotes
basename = os.path.basename(fonte)
arq_excel_processado = os.path.splitext(basename)[0]
arq_excel_processado = arq_excel_processado + '-' + today.strftime('%Y_%m_%d_%H') + '.xlsx'
dest = os.path.join(destino_lotes, arq_excel_processado)
if not os.path.exists(dest):
    # print('[INFO] movendo o arquivo excel LOTES.xlsx para a pasta de histórico')
    shutil.move(fonte, dest)

# SALVA OS DADOS ATUALIZADOS DOS LOTES ***E CRIA UM NOVO*** ARQUIVO EXCEL
with pd.ExcelWriter(arq_dados_lotes) as writer:
    # print('[INFO] criando um novo arquivo excel LOTES.xlsx')
    df_lotes_new.to_excel(writer, sheet_name="Sheet1", index=False, header=True)

###############################################
#   PROCESSANDO A LISTA DE LOTES X CLIENTES   #
###############################################
# print('Lista final lotes X cliente')
# print(lista_lote_cliente)

# CONVERTENDO A LISTA EM UM DATAFRAME
df_lote_cliente = pd.DataFrame(lista_lote_cliente, columns=['DT-SAIDA', 'NF', 'CNPJ/CPF', 'CLIENTE', 'CODIGO',
                                                            'PRODUTO', 'QTDE', 'LOTE', 'VALIDADE'])

df_lote_cliente['DT-SAIDA'] = pd.to_datetime(df_lote_cliente['DT-SAIDA'], dayfirst=True)
df_lote_cliente['VALIDADE'] = pd.to_datetime(df_lote_cliente['VALIDADE'], dayfirst=True)
df_lote_cliente['LOTE'] = df_lote_cliente['LOTE'].str.strip()
df_lote_cliente['PRODUTO'] = df_lote_cliente['PRODUTO'].str.strip()
# print('df_lote_cliente:\n', df_lote_cliente)

# SALVANDO OS DADOS DOS LOTES X CLIENTE CONFORME DATAFRAME `df_lote_cliente`
if os.path.exists(arq_dados_lotes_vendidos):  # SE ARQUIVO EXISTE, FAZ UM APPEND COM OS NOVOS DADOS NO FINAL DO AQUIVO

    # PRIMEIRO FAZ UMA CÓPIA DO ARQUIVO ATUAL E ***MOVE*** PARA A PASTA `Vendas-historicos'
    # NÃO FAZ A CÓPIA SE JÁ HOUVER UM ARQUIVO NO HSTÓRICO QUE TENHA SIDO COPIADO DENTRO DA MESMA HORA
    # OU SEJA, SOMENTE CRIA UMA CÓPIA A CADA HORA CHEIA
    fonte = arq_dados_lotes_vendidos
    basename = os.path.basename(fonte)
    arq_excel_processado = os.path.splitext(basename)[0]
    # arq_excel_processado = arq_excel_processado + '-' + today.strftime('%Y_%m_%d_%H_%M_%S') + '.xlsx'
    arq_excel_processado = arq_excel_processado + '-' + today.strftime('%Y_%m_%d_%H') + '.xlsx'
    dest = os.path.join(destino_lotes_vendidos, arq_excel_processado)
    if not os.path.exists(dest):
        shutil.copy(fonte, dest)

    # print('[INFO] adicionando dados no arquivo existente LOTES-VENDIDOS.XLSX')
    wb = load_workbook(arq_dados_lotes_vendidos)
    page = wb.active
    item_df = df_lote_cliente.values.tolist()
    for item in item_df:
        page.append(item)
    wb.save(filename=arq_dados_lotes_vendidos)

else:  # SE ARQUIVO NÃO EXISTE, CRIA UM NOVO EXCEL COM OS DADOS
    with pd.ExcelWriter(arq_dados_lotes_vendidos) as writer:
        # print('[INFO] criando novo arquivo em Excel LOTES-VENDIDOS.XLSX')
        df_lote_cliente.to_excel(writer, sheet_name="Sheet1", index=False, header=True)

# MOVE O ARQUIVO COM OS DADOS XML PARA A PASTA `Saida-processada' E RENOMEIA O ARQUIVO
fonte = arq_dados_xml
basename = os.path.basename(fonte)
arq_processado = os.path.splitext(basename)[0]
arq_processado = arq_processado + '-' + today.strftime('%Y_%m_%d_%H_%M_%S') + '.old'
dest = os.path.join(destino, arq_processado)
shutil.move(fonte, dest)

# EXIBE MENSAGEM FINAL
print('[INFO]', datetime.now().strftime('%H:%M:%S'), 'itens atualizados com sucesso!\n')
