https://drive.google.com/drive/folders/1wxF54KdwCc13NtMss5d1s3niU1tc3kHK

[HERBIA - Product Batch Control System-2.pdf](https://github.com/daltonhardt/ProductBatchControl/files/14377584/HERBIA.-.Product.Batch.Control.System-2.pdf)


### HERBIA - Product Batch Control System [EN-US]

#### 1 From ERP Bling, the XML files for Input and Output invoices must be saved (manually or automatically) in specific folders.<br>
→ Input Invoice: G:\Meu Drive\Controle de Lotes\XML\Entrada<br>
→ Output Invoice: G:\Meu Drive\Controle de Lotes\XML\Saida

#### 2 The Batch Control system has programs that "monitor" the folders where XML files are saved.<br>
→ Program: watchdogXML-entrada.py<br>
→ Program: watchdogXML-saida.py

#### 3 As soon as an XML file enters, other programs are automatically launched to extract the necessary data.<br>
→ Program: LeituraXML-entrada.py<br>
→ Program: LeituraXML-saida.py

#### 4 The data extracted from the XML file is stored in temporary files and is later used by the batch input program and the batch output program.<br>
→ G:\Meu Drive\Controle de Lotes\XML\Dados\DADOS-XML-ENTRADA.dat<br>
→ G:\Meu Drive\Controle de Lotes\\XML\Dados\DADOS-XML-SAIDA.dat

#### 5 The "Lotes-Entrada.py" program must be executed manually throughout the day (after the Invoice entry) to register the entry of a new Batch of product and its respective expiration date.<br>
Batch data is stored in an Excel file.<br>
→ Program: Lotes-Entrada.py<br>
→ File: G:\Meu Drive\Controle de Lotes\Lotes\LOTES.xlsx

#### 6 The program "Lotes-Saida.py"  runs automatically and updates the consumption of each Batch of product according to sales, listing the buyer's CNPJ/CPF, date of sale, product, quantity and Batch consumed.<br>
All data is stored in an Excel format file.<br>
→ Program: Lotes-Saida.py<br>
→ File: G:\Meu Drive\Controle de Lotes\Lotes\LOTES-VENDIDOS.xlsx




### HERBIA - Sistema Controle De Lotes [PT-BR]

#### 1 A partir do ERP Bling os arquivos XML das notas fiscais de Entrada e de Saída devem ser salvos (de forma manual ou automática) em pastas específicas.<br>
→ NF de entrada: G:\Meu Drive\Controle de Lotes\XML\Entrada<br>
→ NF de saída: G:\Meu Drive\Controle de Lotes\XML\Saida

#### 2 O sistema de Controle de Lotes possui programas que "monitoram" as pastas onde os arquivos XML são salvos.<br>
→ Programa: watchdogXML-entrada.py<br>
→ Programa: watchdogXML-saida.py

#### 3 Assim que um arquivo XML entrar, automaticamente são disparados outros programas que extraem os dados necessários.<br>
→ Programa: LeituraXML-entrada.py<br>
→ Programa: LeituraXML-saida.py

#### 4 Os dados extraídos do XML são armazenados em arquivos temporários e são usados posteriormente pelo programa de entrada de lotes e pelo programa de saída de lotes.<br>
→ G:\Meu Drive\Controle de Lotes\XML\Dados\DADOS-XML-ENTRADA.dat<br>
→ G:\Meu Drive\Controle de Lotes\\XML\Dados\DADOS-XML-SAIDA.dat

#### 5 O programa "Lotes-Entrada.py" deve ser executado manualmente ao longo do dia (após a entrada da NF) para registrar a entrada de um novo Lote e sua respectiva validade.<br>
Os dados dos Lotes ficam armazenados num arquivo em formato Excel.<br>
→ Programa: Lotes-Entrada.py<br>
→ Arquivo: G:\Meu Drive\Controle de Lotes\Lotes\LOTES.xlsx

#### 6 O programa "Lotes-Saida.py" roda de forma automática e faz a atualização do consumo de cada Lote de produto de acordo com as vendas, relacionando o CNPJ/CPF do comprador, a data da venda, o produto, a quantidade e o Lote consumido.<br>
Todos os dados ficam armazenados num arquivo em formato Excel.<br>
→ Programa: Lotes-Saida.py<br>
→ Arquivo: G:\Meu Drive\Controle de Lotes\Lotes\LOTES-VENDIDOS.xlsx
