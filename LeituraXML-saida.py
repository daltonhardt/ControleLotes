import xml.etree.ElementTree as ET
import pandas as pd
import glob
import os
import shutil
import sys
from datetime import datetime

# --------------------  DEFINIÇÕES GERAIS  --------------------
# PASTA DE ENTRADA DAS NFes DE VENDAS
dir_notasXML = '..\\XML\\Saida\\'

# ARQUIVO QUE VAI CONTER OS DADOS COPIADOS DAS NFes
dir_dados_xml = '..\\XML\\Dados\\'
arq_dados_xml = dir_dados_xml + 'DADOS-XML-SAIDA.dat'

# COLUNAS QUE SERÃO GRAVADAS NO ARQUIVO
colunas_xml = ['DT-SAIDA', 'NF', 'CODIGO', 'PRODUTO', 'QTDE', 'CNPJ/CPF', 'CLIENTE']

# PASTA PARA ONDE SERÃO REMOVIDOS OS ARQUIVOS XML PROCESSADOS
dir_processado = '..\\XML\\Saida\\Processado\\'

# ARQUIVO FINAL QUE CONTEM OS DADOS DOS PRODUTOS COM LOTE E VALIDADE
arq_dados_vendas = '..\\Vendas\\LOTES-VENDIDOS.xlsx'

# DEFINIÇÃO DA PASTA ONDE SERÁ MOVIDO O AQUIVO DE DADOS XML NO FINAL DA EXECUÇÃO
destino = dir_dados_xml + 'Saida-processada'

# LISTA DE CFOP DAS NOTAS FISCAIS QUE DEVEM SER PROCESSADAS
# >>>> SE PRECISAR ALTERE A LINHA ABAIXO COM A LISTA DE CFOP QUE INTERESSA
lista_cfop_interessa = ['5102', '5108', '5117', '5202', '5901', '5910', '5915', '5917', '5949',
                        '6102', '6108', '6117', '6202', '6901', '6910', '6915', '6917', '6949']

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


# -------------- PROCESSAMENTO DOS ARQUIVOS XML --------------

def processaXML():
    # print('[INFO] processando arquivos XML de Saída')
    df = pd.DataFrame(columns=colunas_xml)

    # PROCESSA SOMENTE O PRIMEIRO ARQUIVO XML NA PASTA DE VENDAS
    # OS OUTROS (SE EXISTIR) VAI PROCESSAR NA PROXIMA INTERAÇÃO
    arquivo = notas_saida[0]

    # PEGANDO O ELEMENTO DO TOPO (root) QUE CONTÉM OS DEMAIS ELEMENTOS DA NOTA FISCAL
    root = ET.parse(dir_notasXML + '/' + arquivo).getroot()
    # print(root.tag)

    # FAZENDO LOOP NOS DEMAIS ELEMENTOS DA NOTA
    for child in root.iter('{http://www.portalfiscal.inf.br/nfe}ide'):

        # PEGANDO A NATUREZA DA NOTA
        for nat in child.iter('{http://www.portalfiscal.inf.br/nfe}natOp'):
            natureza = nat.text
            # print('natureza: ', natureza)

        # PEGANDO O NUMERO DA NOTA
        for nro in child.iter('{http://www.portalfiscal.inf.br/nfe}nNF'):
            nroNF = nro.text
            # print('nroNF: ', nroNF)

        # PEGANDO NO CNPJ/CPF DO CLIENTE
        for item in root.iter('{http://www.portalfiscal.inf.br/nfe}dest'):
            cnpj = ''
            cpf = ''
            for subitem in item.iter('{http://www.portalfiscal.inf.br/nfe}CNPJ'):
                cnpj = subitem.text
                # print('cnpj: ', cnpj)
            for subitem in item.iter('{http://www.portalfiscal.inf.br/nfe}CPF'):
                cpf = subitem.text
                # print('cpf: ', cpf)
            for subitem in item.iter('{http://www.portalfiscal.inf.br/nfe}xNome'):
                cliente = subitem.text
                # print('cliente: ', cliente)

        # PEGANDO OS ITENS (PRODUTOS) DA NOTA
        for item in root.iter('{http://www.portalfiscal.inf.br/nfe}det'):
            elemento = item.get('nItem')
            # print('Item nr.: ', elemento)

            # PEGANDO NUMERO DO CFOP
            for num in item.iter('{http://www.portalfiscal.inf.br/nfe}CFOP'):
                cfop = num.text
                # print('CFOP  : ', cfop)

            # PEGANDO O CODIGO DO PRODUTO
            for cod in item.iter('{http://www.portalfiscal.inf.br/nfe}cProd'):
                codigo = cod.text
                # print('Código  : ', codigo)
                short_code = codigo.split('/')[0]
                # print('Short Code: ', short_code)

            # PEGANDO A DESCRIÇÃO DO PRODUTO
            for prod in item.iter('{http://www.portalfiscal.inf.br/nfe}xProd'):
                produto = prod.text
                # print('Produto : ', produto)

            # PEGANDO A QUANTIDADE DO PRODUTO
            for value in item.iter('{http://www.portalfiscal.inf.br/nfe}qCom'):
                qtde = int(float(value.text))
                # print('Qtde    : ', qtde)

            # PROCESSANDO SOMENTE SE A NOTA FOR COM CFOP QUE INTERESSA
            if cfop in lista_cfop_interessa:
                # CRIANDO A LINHA DE DADOS PARA SALVAR NO DATAFRAME
                if cnpj != '':
                    cadastro = cnpj
                if cpf != '':
                    cadastro = cpf
                new_data = {'DT-SAIDA': hoje, 'NF': nroNF, 'CNPJ/CPF': cadastro, 'CLIENTE': cliente,
                            'CODIGO': short_code, 'PRODUTO': produto, 'QTDE': int(qtde)}
                # print('new_data ', new_data)

                # ADICIONANDO AO DATAFRAME
                df = df.append(new_data, ignore_index=True)

    # MOVENDO O ARQUIVO XML PARA A PASTA DE ARQUIVOS PROCESSADOS
    shutil.move(dir_notasXML + '\\' + arquivo, dir_processado)

    # SALVANDO DADOS DO XML DE SAIDA NO ARQUIVO QUE DEPOIS SERÁ USADO PELO PROGRAMA `Lotes-Saida.py`
    if not os.path.isfile(arq_dados_xml):  # se o arquivo não existe
        df.to_csv(arq_dados_xml, header=colunas_xml, index=False, sep=';')
    else:  # se existe faz um `append`
        df.to_csv(arq_dados_xml, mode='a', header=False, index=False, sep=';')

    # print('[INFO] Dataframe final: \n', df.to_string(index=True))
    # print('----------------------------------')
    # print('[INFO]', datetime.now().strftime('%H:%M:%S'), 'arquivo XML processado com sucesso!')
    print('[INFO]', datetime.now().strftime('%H:%M:%S'), '--> processando NF:', nroNF)


def main():
    # DEFININDO VARIAVEIS GLOBAIS
    global arquivo, produto, short_code, qtde, nroNF, notas_saida, cpf, cnpj, cadastro, cliente

    # LISTANDO ARQUIVOS XML PARA PROCESSAR SOMENTE OS DE CNPJ IGUAL AO DA HERBIA
    notas_saida = [f for f in sorted(os.listdir(dir_notasXML)) if (f.lower().endswith('.xml') and
                                                                   f.find('07732785000136') != -1)]

    # print('Encontrada(s)', len(notas_saida), 'NF de saida')
    if len(notas_saida) > 0:
        arquivo = notas_saida[0]
        if os.path.exists(dir_processado + '/' + arquivo):
            print('[INFO] *** Descartando arquivo XML já processado anteriormente:', arquivo)
            os.remove(dir_notasXML + '/' + arquivo)
        else:
            # CHAMA FUNÇÃO PARA PROCESSAR OS ARQUIVOS TIPO XML PARA EXTRAIR OS DADOS E CRIAR O DATAFRAME
            # print('[INFO]', datetime.now().strftime('%H:%M:%S'), 'executando programa LeituraXML-saida...')
            processaXML()

            print('[INFO]', datetime.now().strftime('%H:%M:%S'), 'atualizando arquivo LOTES.xlsx')
            os.system('python "..\\Programas\\Lotes-Saida.py"')

            # print('[INFO]', datetime.now().strftime('%H:%M:%S'), 'FIM arquivo LeituraXML_saida')

    else:
        print('\n-----------------------------------------------------')
        print('Nenhum arquivo tipo XML de saida da Herbia encontrado')
        print('-----------------------------------------------------')


if __name__ == '__main__':
    main()
