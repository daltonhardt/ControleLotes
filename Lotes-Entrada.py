from tkinter import *
import tkinter as tk
from tkcalendar import DateEntry
from idlelib.tooltip import Hovertip
import numpy as np
import pandas as pd
from datetime import datetime
import os.path
import shutil
import sys
from openpyxl import load_workbook


# --------------------  DEFINIÇÕES GERAIS  --------------------

# ARQUIVO QUE CONTEM OS DADOS DAS NFes DE ENTRADA
dir_dados_xml = '..\\XML\\Dados\\'
arq_dados_xml = dir_dados_xml + 'DADOS-XML-ENTRADA.dat'

# ARQUIVO QUE CONTEM OS DADOS DOS PRODUTOS COM LOTE E VALIDADE
arq_dados_lotes = '..\\Lotes\\LOTES.xlsx'

# DEFINIÇÃO DA PASTA ONDE SERÁ MOVIDO O AQUIVO DE DADOS XML NO FINAL DA EXECUÇÃO
destino = dir_dados_xml + 'Entrada-processada'

# DEFINIÇÃO DA PASTA ONDE SERÁ CRIADA UMA CÓPIA DO AQUIVO DE LOTES PARA HISTÓRICO
dir_lotes_historicos = '..\\Lotes\\Lotes-historicos'

# DEFINIÇÃO DO CÓDIGO DE PRODUTO QUE DEVE SER CONSIDERADO NA PESQUISA
# Herbia terá somente códigos HER----
# Tem um fornecedor que não consegue alterar o código e por isso os produtos 642 e 643 ainda devem aparecer
# item 642 = HER057
# item 643 = HER058
# >>>> SE PRECISAR ALTERAR A LINHA ABAIXO COM A LISTA DE CODIGOS QUE INTERESSAM
# lista_codigo_interesse = 'HER|642|643'  # para separar os  tipos usar barra vertical (pipe)
lista_codigo_interesse = 'HER'

# DIA DE HOJE NO FORMATO dd-mm-aaaa
today = datetime.now()
hoje = today.strftime('%d/%m/%Y')
hora = today.strftime('%H:%M:%S')
mes = today.strftime('%m')
ano = today.strftime('%Y')

# DIAS DA SEMANA EM PORTUGUES
dias_semana = ['Segunda-feira', 'Terça-feira', 'Quarta-feira', 'Quinta-feira', 'Sexta-feira', 'Sábado', 'Domingo']
wkday = today.weekday()

# Para mostrar todas as linhas e colunas do Dataframe
pd.set_option('display.max_rows', None)
pd.set_option('display.max_columns', None)
pd.set_option('display.width', 1000)
pd.set_option('max_colwidth', 25)

# MENSAGEM DE SAUDAÇÃO
print('----------------------------------------')
print('|                                      |')
print('|       ENTRADA DE NOVOS LOTES         |')
print('|                                      |')
print('----------------------------------------\n')
print('[INFO]', dias_semana[wkday], hoje)


# --------------  CRIANDO A LISTA DE NF E DE PRODUTOS --------------

if os.path.exists(arq_dados_xml):
    colunas = ['DT-ENTRADA', 'NF', 'CODIGO', 'PRODUTO', 'QTDE']
    df_xml = pd.DataFrame(columns=colunas)
    df_xml = df_xml.append(pd.read_csv(arq_dados_xml, sep=';', header=0))
    # print('dataframe xml\n', df_xml)
    df_xml['CODIGO'] = df_xml['CODIGO'].astype(str)
    df_xml = df_xml.loc[df_xml['CODIGO'].str.contains(lista_codigo_interesse)]  # seleciona só códigos que interessam
    # print('dataframe xml selecionado\n', df_xml)
    df_xml_sorted = df_xml.sort_values(by=['NF', 'CODIGO'], ascending=True)  # lista dos itens em ordem crescente
    # print('dataframe xml sorted\n', df_xml_sorted)
    df_itens = df_xml_sorted[['NF', 'CODIGO', 'PRODUTO', 'QTDE']]  # lista com os itens que interessam
    df_itens = df_itens.reset_index(drop=True)  # fazendo reset no index
    nro_itens = len(df_itens)
    if nro_itens > 0:
        print('[INFO] Nro de itens  :', nro_itens)
        print('[INFO] Lista de itens:\n', df_itens.to_string(index=False))
        lista_NF = df_xml_sorted['NF'].unique()
        # print('lista_NF unica', lista_NF)
        lista_PROD = df_xml_sorted['CODIGO'].to_list()
        # print('lista_PROD', lista_PROD)
        # for index, row in df_itens.iterrows():
        #    print(row['NF'], row['CODIGO'], row['PRODUTO'], row['QTDE'])
    else:
        sys.exit('[INFO] nenhum código de interesse dentro do arquivo ' + arq_dados_xml)
else:
    print('[INFO] Não existe arquivo com dados de XML para processar. Tente mais tarde!\n')
    quit()


def sair():
    botao_sair['relief'] = 'sunken'
    root.destroy()


def cadastrar():
    global botao_cadastrar
    botao_cadastrar['relief'] = 'sunken'
    entry_list = []
    # i = 0
    faltam_itens = False
    for entries in entradas:
        if str(entries.get()) == '':
            faltam_itens = True
        # print(str(i), entries.get())
        # i += 1
        mensagem.config(text='Preencha todos os itens antes de salvar!', bg='yellow', fg='black')
        botao_cadastrar['relief'] = 'raised'
    # i = 0
    if not faltam_itens:
        for entries in entradas:
            entry_list.append(str(entries.get()))
            # campo.config(text=entry_list)
            # print(str(i) + '...' + entradas[i].get())
            # i += 1

        # gravando dataframe com os valores de QTDE (CORRIGIDO), LOTE e VALIDADE
        df_entradas = pd.DataFrame(np.array(entry_list).reshape(nro_itens, 3), columns=['QTDE-REAL', 'LOTE',
                                                                                        'VALIDADE'])
        # concatenando os dataframes ao longo das colunas (axis=1)
        # tem que fazer reset do index do df_xml_sorted para manter a ordem das linhas do dataframe df_final
        df_final = pd.concat([df_xml_sorted.reset_index(drop=True), df_entradas], axis=1)

        # CONVERTENDO A QTDE_REAL PARA INTEIRO
        df_final['QTDE-REAL'] = df_final['QTDE-REAL'].astype(int)

        # CONVERTENDO A NF PARA OBJECT(TEXT)
        df_final['NF'] = df_final['NF'].astype(str)

        # ADICIONANDO A COLUNA COM A DATA DA ULTIMA SAIDA EM BRANCO PARA SER PREENCHIDA DEPOIS QUANDO VENDER O PRODUTO
        df_final['DT-ULTIMA-SAIDA'] = ''

        # ADICIONANDO A COLUNA SALDO=QTDE_REAL. DEPOIS O SALDO SERÁ ATUALIZADO QUANDO VENDER O PRODUTO
        df_final['SALDO'] = df_final['QTDE-REAL'].astype(int)

        print('[INFO] Dataframe final: \n', df_final.to_string(index=False))

        # CRIA ARQUIVO TIPO EXCEL COM OS DADOS DOS LOTES CONFORME DATAFRAME `df_final`
        if os.path.exists(arq_dados_lotes):  # SE ARQUIVO EXISTE, FAZ UM APPEND COM OS NOVOS DADOS NO FINAL DO AQUIVO

            # PRIMEIRO FAZ UMA CÓPIA DO ARQUIVO ATUAL E MOVE PARA A PASTA `Vendas-historicos'
            fonte = arq_dados_lotes
            basename = os.path.basename(fonte)
            arq_lotes_atual = os.path.splitext(basename)[0]
            # arq_lotes_atual = arq_lotes_atual + '-' + today.strftime('%Y_%m_%d_%H_%M_%S') + '.xlsx'
            arq_lotes_atual = arq_lotes_atual + '-' + today.strftime('%Y_%m_%d_%H') + '.xlsx'
            dest = os.path.join(dir_lotes_historicos, arq_lotes_atual)
            if not os.path.exists(dest):
                shutil.copy(fonte, dest)

            wb = load_workbook(arq_dados_lotes)
            page = wb.active
            item_df = df_final.values.tolist()
            for item in item_df:
                page.append(item)
            try:
                wb.save(filename=arq_dados_lotes)
            except OSError:
                print('[INFO] erro ao salvar o arquivo LOTES.XLSX')
                mensagem.config(text='Erro ao salvar o arquivo LOTES.XLSX', bg='red', fg='yellow')
                return
        else:  # SE ARQUIVO NÃO EXISTE, CRIA UM NOVO EXCEL COM OS DADOS
            # df_final.to_csv(arq_dados_lotes, index=False, header=True, sep=';')
            with pd.ExcelWriter(arq_dados_lotes) as writer:
                df_final.to_excel(writer, sheet_name="Sheet1", index=False, header=True)

        # MOVE O ARQUIVO COM OS DADOS XML PARA A PASTA `Entrada-processada' E RENOMEIA O ARQUIVO
        fonte = arq_dados_xml
        basename = os.path.basename(fonte)
        arq_processado = os.path.splitext(basename)[0]
        arq_processado = arq_processado + '-' + today.strftime('%Y_%m_%d_%H_%M_%S') + '.old'
        dest = os.path.join(destino, arq_processado)
        shutil.move(fonte, dest)

        # EXIBE MENSAGEM FINAL E DESABILITA O BOTÃO 'SALVAR' DEIXANDO ATIVO SOMENTE O BOTÃO 'SAIR'
        mensagem.config(text='Itens cadastrados com sucesso!', bg='green', fg='yellow')
        botao_cadastrar['state'] = 'disabled'

        print('[INFO] Itens cadastrados com sucesso!')


# -----------------------  TKINTER -----------------------
root = tk.Tk()

# DIMENSOES DO MONITOR
screen_width = root.winfo_screenwidth()
screen_height = root.winfo_screenheight()
# print('Screen dimensions: ', screen_width, 'x', screen_height)

# DIMENSOES DA JANELA
# w = int(screen_width/2)
# h = int(screen_height/2)
w = 1100
h = int(nro_itens*50 + 100)
janela = str(w) + 'x' + str(h)
# print(janela)

titulo = 'GRUPO HERBIA       -   E N T R A D A    D E    L O T E S   -       ' + hoje + '  (' + dias_semana[wkday] + ')'
root.title(titulo)
root.geometry(janela)
root.iconbitmap('')

Label(root, text='NF').grid(column=0, row=0, sticky=W)
Label(root, text='CODIGO').grid(column=1, row=0, sticky=W)
Label(root, text='PRODUTO').grid(column=2, row=0, sticky=W)
Label(root, text='QTDE').grid(column=3, row=0, sticky=W)
Label(root, text='LOTE').grid(column=4, row=0, sticky=W)
Label(root, text='VALIDADE').grid(column=5, row=0, sticky=W)

entradas = []
for y in range(nro_itens):
    for x in range(6):
        if x == 0:  # campo NF
            conteudo = StringVar()
            conteudo.set(df_itens.iloc[y][x])
            campo = tk.Entry(root, width=10, state=DISABLED, textvariable=conteudo)
        elif x == 1:  # campo CODIGO
            conteudo = StringVar()
            conteudo.set(df_itens.iloc[y][x])
            campo = tk.Entry(root, width=10, state=DISABLED, textvariable=conteudo)
        elif x == 2:  # campo PRODUTO
            conteudo = StringVar()
            conteudo.set(df_itens.iloc[y][x])
            campo = tk.Entry(root, width=50, state=DISABLED, textvariable=conteudo)
        elif x == 3:  # campo QTDE
            conteudo = StringVar()
            conteudo.set(df_itens.iloc[y][x])
            campo = tk.Entry(root, width=10, textvariable=conteudo)
            dica = 'NÃO altere a quantidade se for emitir uma NF de retorno.\n' \
                 'SOMENTE altere a quantidade se o fornecedor compensar a falta ou perda do item sem emitir uma NF.'
            myTip = Hovertip(campo, dica)
            entradas.append(campo)  # colocando a QTDE-REAL na lista de entradas

        elif x == 4:  # campo LOTE
            campo = tk.Entry(root, width=10, bd=3)
            entradas.append(campo)  # colocando LOTE na lista de entradas

        elif x == 5:  # campo VALIDADE
            campo = DateEntry(root, selectmode='day', locale='pt_BR', date_pattern='dd/MM/yyyy', year=int(ano)+2)
            entradas.append(campo)  # colocando VALIDADE na lista de entradas

        campo.grid(row=y + 1, column=x, pady=10, padx=5)
        # print(y, x)

botao_cadastrar = tk.Button(root, text='SALVAR', command=cadastrar, font='Arial 14',
                            relief='raised', bd=3, height=2, width=10)
botao_cadastrar.grid(row=nro_itens+2, column=5, pady=5)
botao_sair = tk.Button(root, text="Sair", command=sair, font='Arial 14',
                       relief='raised', bd=3, height=2, width=10)
botao_sair.grid(row=nro_itens+2, column=0, pady=5)

mensagem = tk.Label(root, font='Arial 14')
mensagem.grid(row=nro_itens+2, column=2, pady=5)

root.mainloop()
