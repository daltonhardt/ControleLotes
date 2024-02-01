### HERBIA - Sistema Controle De Lotes

1 A partir do ERP Bling os arquivos XML das notas fiscais de Entrada e de Saída devem ser salvos (de forma manual ou
automática) em pastas específicas.
→ NF de entrada: G:\Meu Drive\Controle de Lotes\XML\Entrada<br>
→ NF de saída: G:\Meu Drive\Controle de Lotes\XML\Saida

2 O sistema de Controle de Lotes possui programas que "monitoram" as pastas onde os arquivos XML são salvos.
→ Programa: watchdogXML-entrada.py<br>
→ Programa: watchdogXML-saida.py

3 Assim que um arquivo XML entrar, automaticamente são disparados outros programas que extraem os dados necessários.
→ Programa: LeituraXML-entrada.py<br>
→ Programa: LeituraXML-saida.py

4 Os dados extraídos do XML são armazenados em arquivos temporários e são usados posteriormente pelo programa de entrada de lotes e pelo programa de saída de lotes.
→ G:\Meu Drive\Controle de Lotes\XML\Dados\DADOS-XML-ENTRADA.dat<br>
→ G:\Meu Drive\Controle de Lotes\\XML\Dados\DADOS-XML-SAIDA.dat

5 O programa "Lotes-Entrada.py" deve ser executado manualmente ao longo do dia (após a entrada da NF) para registrar a entrada de um novo Lote e sua respectiva validade.
Os dados dos Lotes ficam armazenados num arquivo em formato Excel. 
→ Programa: Lotes-Entrada.py<br>
→ Arquivo: G:\Meu Drive\Controle de Lotes\Lotes\LOTES.xlsx
