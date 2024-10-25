# Scrapy-Pregao-Eletronico
Criação de um scrapy para consulta dos dados mais importantes para o pregoeiro durante sua interação com as empresas, acessando o site https://www.gov.br/compras/pt-br no perfil do pregoeiro, bucando o processo licitatório e extraindo as informações:  
-CNPJ da empresa participante  
-Se é Micro empresa ou empresa de pequeno porte  
-Nome da empresa  
-Numero do item no certame  
-Descrição resumida  
-Valor estimado  
-Quantidade de cada item  
-Valor ofertado pela empresa  

Após efetuar obter essas informações para todos os itens do certame, é criada uma planilha conforme exemplo: [Planilha Apoio Pregao 31_2023.xlsx] para o pregoeiro utilizar para controle das ações que foram executadas em para cada item ou empresa. Para auxiliar o pregoeiro, os itens acima do estimado, são destacados em vermelho na aba "Análise das Empresas"  

Tal ferramenta foi desenvovida devido a necessidade de todo pregoeiro criar uma planilha de controle, contendo os itens acima elencados, de ações tomadas em todo pregão eletrônico que participa, porém essa planilha acaba sendo feita manualmente por falta de ferramenta de extração de dados dentro do sistema de pregão eletrônico, que a dependender da quantidade de itens do processo, se torna algo extremamente lento e passivel de falhas. Esta ferramenta automatiza este processo, tirando o risco de erro humano na eaboração desta planiha, além de ser extremamente mais rápido.

# Instruções para utilizar o script:
-Tendo em vista que o sistema acessa o perfil do pregoeiro, é necessário que o usuário tenha cadastro no sistem compras.gov como Governo.  
-Faça o download dos arquivos [Criador de planilha.exe] e [geckodriver.exe]  
-É necessário o browser Mozilla Firefox na versão 119  
-Execute o arquivo [Criador de planilha.exe]  
-Preencha as informações solicitadas e aguarde a execução do programa.  

# Bibliotecas utilizadas na elaboração do código:
-Selenium  
-Openpyxl  
-PySimpleGUI  

# Linguagem utilizada:
-Python
