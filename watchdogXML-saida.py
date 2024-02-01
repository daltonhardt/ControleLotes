import time
import os
import glob
from datetime import datetime
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler

# --------------------  DEFINIÇÕES GERAIS  --------------------

# PASTA DE ENTRADA DAS NFes DE VENDAS
dir_notasXML = '..\\XML\\Saida\\'

# DIA DE HOJE NO FORMATO dd-mm-aaaa
today = datetime.now()
hoje = today.strftime('%d-%m-%Y')
hora = today.strftime('%H:%M:%S')

# DIAS DA SEMANA EM PORTUGUES
dias_semana = ['Segunda-feira', 'Terça-feira', 'Quarta-feira', 'Quinta-feira', 'Sexta-feira', 'Sábado', 'Domingo']
wkday = today.weekday()

# MENSAGEM DE SAUDAÇÃO
print('----------------------------------------------')
print('|                                            |')
print('|      MONITORAMENTO  XML  DE  SAIDA         |')
print('|                                            |')
print('----------------------------------------------\n')
print('[INFO]', dias_semana[wkday], hoje)


# -------------- PROCESSAMENTO DOS ARQUIVOS XML --------------

class Watcher:
    DIRECTORY_TO_WATCH = dir_notasXML

    print('[INFO]', datetime.now().strftime('%H:%M:%S'), 'início monitoramento em',
          DIRECTORY_TO_WATCH)

    def __init__(self):
        self.observer = Observer()

    def run(self):
        event_handler = Handler()
        self.observer.schedule(event_handler, self.DIRECTORY_TO_WATCH, recursive=False)
        self.observer.start()
        try:
            while True:
                time.sleep(5)
        except KeyboardInterrupt:
            self.observer.stop()
            print('\n[INFO]', datetime.now().strftime('%H:%M:%S'), 'processo finalizado!\n')

        self.observer.join()


class Handler(FileSystemEventHandler):

    @staticmethod
    def on_any_event(event, *args):
        if event.is_directory:
            return None

        elif event.event_type == 'created':
            time.sleep(5)  # espera 3 segundos para dar tempo de entrar mais arquivos no diretorio
            if event.src_path[-4:].lower() == '.xml':
                # print('[INFO]', datetime.now().strftime('%H:%M:%S'), 'entrou: %s' % event.src_path)

                # ELIMINA OS ARQUIVOS DE NOTAS FISCAIS CANCELADAS
                # LISTANDO ARQUIVOS XML DO TIPO CANCELADO E ELIMINANDO DO DIRETORIO. FINAL = 'can.xml'
                notas_canceladas = [f for f in sorted(os.listdir(dir_notasXML)) if (f.lower().endswith('can.xml') and
                                                                                    f.find('07732785000136') != -1)]

                # print('Encontrada(s)', len(notas_canceladas), 'NF cancelada(s)')

                if len(notas_canceladas) > 0:
                    for arq in notas_canceladas:
                        arq_nf = dir_notasXML + arq[:44] + '*.xml'

                        # Pega todos os arquivos com o mesmo número da NF
                        files = glob.glob(arq_nf)
    
                        # Loop sobre a lista de arquivos para remover
                        for file in files:
                            print('[INFO] *** Removendo arquivo da NFe CANCELADA' + file)
                            os.remove(file)

                # RODA O PROGRAMA QUE RETIRA OS DADOS NECESSÁRIOS DO ARQUIVO XML
                print('[INFO]', datetime.now().strftime('%H:%M:%S'), 'executando programa LeituraXML-saida...')
                os.system('python "..\\Programas\\LeituraXML-saida.py"')

        '''
        elif event.event_type == 'modified':
            if event.src_path[-3:] == 'xml':
                print(datetime.now().strftime('%d/%m/%Y %H:%M:%S'), '| Modificado arquivo: %s' % event.src_path)

        elif event.event_type == 'deleted':
            if event.src_path[-3:] == 'xml':
                print(datetime.now().strftime('%d/%m/%Y %H:%M:%S'), '| Saiu arquivo: %s' % event.src_path)

        elif event.event_type == 'moved':
            if event.src_path[-3:] == 'xml':
                print(datetime.now().strftime('%d/%m/%Y %H:%M:%S'), '| Movido arquivo: %s' % event.src_path)
        '''


if __name__ == '__main__':
    w = Watcher()
    w.run()
