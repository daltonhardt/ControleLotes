import xml.etree.ElementTree as ET
import pandas as pd
import os
import shutil
import sys
from datetime import datetime

# --------------------  DEFINIÇÕES GERAIS  --------------------
# PASTA DE ENTRADA DAS NFes DE COMPRAS
dir_notasXML = '..\\XML\\Entrada'

# ARQUIVO QUE VAI CONTER OS DADOS COPIADOS DAS NFes
arq_dados_xml = '..\\XML\\Dados\\DADOS-XML-ENTRADA.dat'

# COLUNAS QUE SERÃO GRAVADAS NO ARQUIVO
colunas_xml = ['DT-ENTRADA', 'NF', 'CODIGO', 'PRODUTO', 'QTDE']

# PASTA PARA ONDE SERÃO REMOVIDOS OS ARQUIVOS PROCESSADOS
dir_processado = '..\\XML\\Entrada\\Processado'

# ARQUIVO COM A TABELA DE CODIGOS DE:PARA (cod.Fornecedor para cod.Herbia)
arq_tabela_codigos = '..\\Lotes\\Tabela-Codigos.dat'

# DIA DE HOJE NO FORMATO dd/mm/aaaa
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

# CRIANDO UM DATAFRAME COM AS DEFINIÇÕES DA TABELA DE CÓDIGO DO FORNECEDOR PARA CÓDIGO HERBIA
if not os.path.isfile(arq_tabela_codigos):  # se o arquivo de codigos não existe
    sys.exit('[INFO] não encontrado o arquivo ' + arq_tabela_codigos)
else:  # se existe
    df_codigos = pd.read_csv(arq_tabela_codigos)
    lista_cod = df_codigos.values.tolist()  # transformando o dataframe numa lista


# -------------- PROCESSAMENTO DOS ARQUIVOS XML --------------

def processaXML():
    # print('[INFO] processando arquivos XML de Entrada')
    df = pd.DataFrame(columns=colunas_xml)

    # INICIA O LOOP DOS ARQUIVOS XML NA PASTA DE ENTRADA
    for arquivo in notas_entrada:

        # PEGANDO O ELEMENTO DO TOPO (root) QUE CONTÉM OS DEMAIS ELEMENTOS DA NOTA FISCAL
        root = ET.parse(dir_notasXML + '\\' + arquivo).getroot()
        # print(root.tag)

        # PEGA O NÚMERO DA NOTA FISCAL
        for child in root.iter('{http://www.portalfiscal.inf.br/nfe}ide'):
            for nro in child.iter('{http://www.portalfiscal.inf.br/nfe}nNF'):
                nroNF = nro.text
                # print('Processando NF:', nroNF)

        # FAZ O LOOP PARA PEGAR O SUB-ELEMENTO `det` DA NOTA
        for child in root.iter('{http://www.portalfiscal.inf.br/nfe}det'):
            # print(child.tag, child.attrib)
            item = child.get('nItem')
            # print('Item nr.: ', item)

            # fazendo loop para pegar o código do produto
            for cod in child.iter('{http://www.portalfiscal.inf.br/nfe}cProd'):
                # print(cod.text)
                codigo = cod.text
                # print('Código  : ', codigo)
                short_code = codigo.split('/')[0]
                # print('Short Code: ', short_code)
                for item in lista_cod:  # converte o cod. do Fornecedor para o cod. da Herbia
                    if str(item[0]) == str(short_code):
                        short_code = item[1]

            # fazendo loop para pegar a descrição do produto
            for prod in child.iter('{http://www.portalfiscal.inf.br/nfe}xProd'):
                # print(prod.text)
                produto = prod.text
                # print('Produto : ', produto)

            # fazendo o loop para pegar a quantidade do produto
            for value in child.iter('{http://www.portalfiscal.inf.br/nfe}qCom'):
                # print(value.text, '\n')
                qtde = int(float(value.text))
                # print('Qtde    : ', qtde, '\n')

            # CRIANDO A LINHA DE DADOS PARA SALVAR NO DATAFRAME
            new_data = {'DT-ENTRADA': hoje, 'NF': nroNF, 'CODIGO': short_code, 'PRODUTO': produto, 'QTDE': int(qtde)}

            # ADICIONANDO AO DATAFRAME
            df = df.append(new_data, ignore_index=True)

        # MOVENDO O ARQUIVO XML PARA `pasta_xml_entrada_processado`
        shutil.move(dir_notasXML + '\\' + arquivo, dir_processado)

    # SALVANDO DADOS DO XML DE ENTRADA NO ARQUIVO QUE DEPOIS SERÁ USADO PELO PROGRAMA `LOTES-ENTRADA.PY`
    if not os.path.isfile(arq_dados_xml):  # se o arquivo não existe
        df.to_csv(arq_dados_xml, header=colunas_xml, index=False, sep=';')
    else:  # se existe faz um `append`
        df.to_csv(arq_dados_xml, mode='a', header=False, index=False, sep=';')

    # print('[INFO] Dataframe final: \n', df.to_string(index=True))
    # print('----------------------------------')
    # print('[INFO]', datetime.now().strftime('%H:%M:%S'), 'arquivo XML processado com sucesso!')
    print('[INFO]', datetime.now().strftime('%H:%M:%S'), 'NF:', nroNF, 'processada!')


def main():
    # DEFININDO VARIAVEIS GLOBAIS
    global produto, short_code, qtde, nroNF, notas_entrada

    # LISTANDO ARQUIVOS XML PARA PROCESSAR SOMENTE OS DE CNPJ DIFERENTE DA HERBIA
    # notas_entrada = [f for f in sorted(os.listdir(dir_notasXML)) if (f.endswith('.xml') and
    # f.find('07732785000136') == -1)]

    # LISTANDO QUALQUER ARQUIVO XML PARA PROCESSAR
    notas_entrada = [f for f in sorted(os.listdir(dir_notasXML)) if f.endswith('.xml')]
    # print(notas_entrada)

    if len(notas_entrada) > 0:
        # CHAMA FUNÇÃO PARA PROCESSAR OS ARQUIVOS TIPO XML PARA EXTRAIR OS DADOS E CRIAR O DATAFRAME
        print('[INFO]', datetime.now().strftime('%H:%M:%S'), 'executando programa LeituraXML-entrada...')
        processaXML()

    else:
        print('\n-------------------------------------------------')
        print('Nenhum arquivo tipo XML de entrada foi encontrado')
        print('-------------------------------------------------')


if __name__ == '__main__':
    main()
