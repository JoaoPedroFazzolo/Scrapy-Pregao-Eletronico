
# Scrapy Prão Eletrônico

Criação de um secrapy para compilar dados importantes de um processo licitatório, como empresas participantes e seu CNPJ, itens, quantitativo, valor estimado, valor ofertado, tipo de empresa.

Tais dados são compilados em uma planilha para uso pelo pregoeiro, de forma a gerenciar as interações e atitudes efetuadas em cada item e empresa, de forma a permitir maior controle do processo.

Tal ferramenta foi desenvovida devido a necessidade de todo pregoeiro criar uma planilha de controle, contendo os itens acima elencados, de ações tomadas em todo pregão eletrônico que participa, porém essa planilha acaba sendo feita manualmente por falta de ferramenta de extração de dados dentro do sistema de pregão eletrônico, que a dependender da quantidade de itens do processo, se torna algo extremamente lento e passivel de falhas. Esta ferramenta automatiza este processo, tirando o risco de erro humano na eaboração desta planiha, além de ser extremamente mais rápido.


## Etiquetas

![Python](https://img.shields.io/badge/python-3670A0?style=for-the-badge&logo=python&logoColor=ffdd54)

![Selenium](https://img.shields.io/badge/-selenium-%43B02A?style=for-the-badge&logo=selenium&logoColor=white)
## Uso

-Tendo em vista que o sistema acessa o perfil do pregoeiro, é necessário que o usuário tenha cadastro no sistem compras.gov como Governo.

-Faça o download dos arquivos [Criador de planilha.exe] e [geckodriver.exe]

-É necessário o browser Mozilla Firefox na versão 119
-Execute o arquivo [Criador de planilha.exe]

-Preencha as informações solicitadas e aguarde a execução do programa.


## Stack utilizada


-Selenium
-Openpyxl
-PySimpleGUI
## Feedback

Se você tiver algum feedback, por favor nos deixe saber por meio de admfazzolo@gmail.com

