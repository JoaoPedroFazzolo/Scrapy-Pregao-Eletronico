import time
import re
import random
import openpyxl
import PySimpleGUI as sg
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from openpyxl import Workbook, load_workbook, writer
from openpyxl.styles import Alignment, NamedStyle, PatternFill
from openpyxl.worksheet.filters import FilterColumn, CustomFilter, CustomFilters, DateGroupItem, Filters


uasg = 120071
qntEmpresas = 6
numero = "90013/2024"
pregao = "Pregão Eletrônico " + str(uasg) + " - " + str(numero)

wb = openpyxl.Workbook()
wb.create_sheet('Empresas',0)
wbEmpresas = wb['Empresas']
#cabeçalho
wbEmpresas.append(['CNPJ', 'ME/EPP', 'NOME EMPRESA'])

numero1 = numero.replace('/', '_')
wb.save('Planilha Apoio Pregao '+ str(numero1)+'.xlsx')


def centralizando(w):
    alignment = openpyxl.styles.Alignment(horizontal='center', vertical='center')
    for row in w.iter_rows():
        for cell in row:
            cell.alignment = alignment

def tamanhoColunaComum(a):
    for col in a.columns:
        max_length = 0
        column = col[0].column_letter
        for cell in col:
            try:
                if len(str(cell.value)) > max_length:
                    max_length = len(cell.value)
            except:
                pass
        adjusted_width = (max_length + 2)
        a.column_dimensions[column].width = adjusted_width



options = Options()
options.page_load_strategy = 'normal'

navegador = webdriver.Chrome()
time.sleep(10)
navegador.get('https://cnetmobile.estaleiro.serpro.gov.br/comprasnet-area-trabalho-web/seguro/governo/area-trabalho')
urlAtual = navegador.current_url
time.sleep(10)
##########################         criando espera para o login do usuario de 5min       ############################
try:
    waitLogin = WebDriverWait(navegador, 900).until(
        EC.presence_of_element_located((By.XPATH, '//*[@id="buscaCompra"]'))
    )
except:
    navegador.execute_script("var e = alert('Sistema encerrando pois usuario não efetuou login, tente novamente e efetue o login.', '');document.body.setAttribute('loginscrapynãofeito', f)")
    time.sleep(10)
    
##########################         criando esperas para o código        ############################

def wait60(navegador,urlAtual):
    try:
        wait60 = WebDriverWait(navegador, 60) 
        wait60.until(EC.url_changes(urlAtual))
        time.sleep(3)
    except:
        navegador.execute_script("var e = alert('o tempo de carregamento da pagina passou do limite, reinicie o app.'), '');document.body.setAttribute('tempoEspirador', g)")


def randomWait(navegador):
    try:
        aleatorio = random.randint(5, 55)
        print(f'inicio da espera aleatoria de {aleatorio} segundos')
        time.sleep(aleatorio)

    except:
        navegador.execute_script("var e = alert('o tempo de carregamento da pagina passou do limite, reinicie o app.'), '');document.body.setAttribute('tempoEspirador', g)")

##########################         abrindo o pregao 14.133        ############################
wait60(navegador,urlAtual)
urlAtual = navegador.current_url
navegador.find_element(By.XPATH, '//*[@id="buscaCompra"]').send_keys(numero)
navegador.find_element(By.XPATH, '/html/body/app-root/div/app-area-governo/div/div[2]/div[2]/div/div[2]/button').click()
wait60(navegador,urlAtual)
for i in range(10):
    urlXPATH = str('/html/body/app-root/div/app-pesquisa-rapida/div/div[5]/p-dataview/div/div['+ str(i) +']/div/div/div[2]/span')
    pregoesPesquisa = navegador.find_elements(By.XPATH, urlXPATH )
    for e in pregoesPesquisa:
        a = 1
        if e.text == str(pregao):
            bottonXpath = str('/html/body/app-root/div/app-pesquisa-rapida/div/div[5]/p-dataview/div/div[2]/div['+ str(a) +']/div/div[5]/i')
            navegador.find_element(By.XPATH, bottonXpath).click()
            time.sleep(3)
            navegador.find_element(By.XPATH, '/html/body/app-root/div/app-pesquisa-rapida/div/p-dialog/div/div/div[2]/app-item-trabalho-detalhe/div[2]/div[2]/div/div[2]/app-redirect-sistemas/span/span').click()
            break
        else:
            a =+ 1 
            continue
        break



time.sleep(5)                                 
##########################         alterando janela        ############################
handles = navegador.window_handles
navegador.switch_to.window(handles[1])


##########################         iterando sobre os itens do pregão        ############################
#abrindo o grupo
time.sleep(5)
urlAtual = navegador.current_url
navegador.find_element(By.XPATH, '/html/body/app-root/div/div/div/app-cabecalho-selecao-fornecedores-governo/div[2]/app-selecao-fornecedores-governo/div/p-tabview/div/div[2]/p-tabpanel[1]/div/app-selecao-fornecedores-governo-itens/div[2]/p-dataview/div/div/div/app-card-item/div/div[3]/div[2]/app-botao-icone/span/button').click()
wait60(navegador,urlAtual)

#abrindo empresa
def abrirEmpresa(indice):
    time.sleep(5)
    urlAtual = navegador.current_url
    navegador.find_element(By.XPATH, '/html/body/app-root/div/div/div/app-cabecalho-selecao-fornecedores-governo/div[2]/app-selecao-fornecedores-governo-item/div/div/app-selecao-fornecedores-governo-propostas-item/div/div/div/p-dataview/div/div/div['+ str(indice) + ']/app-dados-proposta-item-em-selecao-fornecedores/div/div[3]/div[2]/div/app-botao-icone/span/button').click()
    wait60(navegador,urlAtual)



##########################        função para abrir os itens e retirar as informações necessárias        ############################
def informaçoesItens(indice):
    #descrição do item:
    numItem = navegador.find_element(By.XPATH, '/html/body/app-root/div/div/div/app-cabecalho-selecao-fornecedores-governo/div[2]/app-selecao-fornecedores-governo-propostas-itens-grupo/div/div[4]/app-selecao-fornecedores-governo-proposta-item/p-tabview/div/div[2]/p-tabpanel[1]/div/div/span/div/app-listagem-propostas-subitens-governo/div[2]/p-dataview/div/div/div[' + str(indice) + ']/app-card-proposta-subitem-em-selecao-fornecedores/div/div[1]/div/div/app-identificacao-e-fase-item/div[1]').text.split(' ')[0]
    #valor estimado:
    valorEstimado = navegador.find_element(By.XPATH, '/html/body/app-root/div/div/div/app-cabecalho-selecao-fornecedores-governo/div[2]/app-selecao-fornecedores-governo-propostas-itens-grupo/div/div[4]/app-selecao-fornecedores-governo-proposta-item/p-tabview/div/div[2]/p-tabpanel[1]/div/div/span/div/app-listagem-propostas-subitens-governo/div[2]/p-dataview/div/div/div[' + str(indice) + ']/app-card-proposta-subitem-em-selecao-fornecedores/div/div[2]/div/div/div[2]/div[2]').text
    qntItem = navegador.find_element(By.XPATH, '/html/body/app-root/div/div/div/app-cabecalho-selecao-fornecedores-governo/div[2]/app-selecao-fornecedores-governo-propostas-itens-grupo/div/div[4]/app-selecao-fornecedores-governo-proposta-item/p-tabview/div/div[2]/p-tabpanel[1]/div/div/span/div/app-listagem-propostas-subitens-governo/div[2]/p-dataview/div/div/div[' + str(indice) + ']/app-card-proposta-subitem-em-selecao-fornecedores/div/div[2]/div/div/div[2]/div[1]').text
    #valor ofertado:
    valorOfertado = navegador.find_element(By.XPATH, '/html/body/app-root/div/div/div/app-cabecalho-selecao-fornecedores-governo/div[2]/app-selecao-fornecedores-governo-propostas-itens-grupo/div/div[4]/app-selecao-fornecedores-governo-proposta-item/p-tabview/div/div[2]/p-tabpanel[1]/div/div/span/div/app-listagem-propostas-subitens-governo/div[2]/p-dataview/div/div/div[' + str(indice) + ']/app-card-proposta-subitem-em-selecao-fornecedores/div/div[3]/div/div/div/div[2]/div[1]').text
    return numItem, valorEstimado, qntItem, valorOfertado

#########################        função para abrir a aba empresas e retirar as informações de cada empresa resumida       ############################
def informaçoesEmpresas(indice):
    cnpjEmpresaCompleto = navegador.find_element(By.XPATH, '/html/body/app-root/div/div/div/app-cabecalho-selecao-fornecedores-governo/div[2]/app-selecao-fornecedores-governo-item/div/div/app-selecao-fornecedores-governo-propostas-item/div/div/div/p-dataview/div/div/div[' + str(indice) + ']/app-dados-proposta-item-em-selecao-fornecedores/div/div[1]/div/app-identificacao-e-situacao-participante-no-item/div/div[1]').text
    nomeEmpresa = navegador.find_element(By.XPATH, '/html/body/app-root/div/div/div/app-cabecalho-selecao-fornecedores-governo/div[2]/app-selecao-fornecedores-governo-item/div/div/app-selecao-fornecedores-governo-propostas-item/div/div/div/p-dataview/div/div/div[' + str(indice) + ']/app-dados-proposta-item-em-selecao-fornecedores/div/div[1]/div/app-identificacao-e-situacao-participante-no-item/div/div[2]/span').text        
    if len(cnpjEmpresaCompleto.split('\n')) == 1:
        cnpjEmpresa = cnpjEmpresaCompleto
        return cnpjEmpresa, nomeEmpresa, 'Não'
    else:
        cnpjEmpresa = cnpjEmpresaCompleto.split('\n')[0]
        return cnpjEmpresa, nomeEmpresa, 'Sim'

##########################        iterando sobre todos os itens        ############################

def limparNomeAba(nome):
    # Define os caracteres não permitidos
    caracteres_nao_permitidos = r'[\/\\\*\?\:\[\]]'
    # Remove os caracteres não permitidos usando expressão regular
    nome_limpo = re.sub(caracteres_nao_permitidos, '', nome)
    return nome_limpo


for i in range (1 , qntEmpresas + 1):
    urlAtual = navegador.current_url
    cnpjEmpresa, nomeEmpresa, meEPP = informaçoesEmpresas(i)
    wbEmpresas
    wbEmpresas.cell(row= i + 1, column=1, value=cnpjEmpresa)
    wbEmpresas.cell(row= i + 1, column=2, value=meEPP)
    wbEmpresas.cell(row= i + 1, column=3, value=nomeEmpresa)
    centralizando(wbEmpresas)
    wbEmpresas.auto_filter.ref = wbEmpresas.dimensions
    tamanhoColunaComum(wbEmpresas)
    wb.save('Planilha Apoio Pregao ' + str(numero1) + '.xlsx')
    abrirEmpresa(i)
    wait60(navegador,urlAtual)
    nomeEpresa1 = limparNomeAba(nomeEmpresa)
    wb.create_sheet(nomeEpresa1, i)
    wbItens = wb[nomeEpresa1]
    wbItens.append(['Item',	'Empresa', 'Qnt Solicitada', 'Valor Estimado', 'Valor ofertado pela empresa', 'Negociou Valor?', 'Especificação Técnica', 'Validade da Proposta', 'SICAF', 'Sanção / Ocorrência', 'CEIS', 'CNJ', 'TCU', 'Empresário Individual:Inscrição no Registro Público', 'MEI:Certificado da Condição de Microempreendedor Individual Verificar autenticidade', 'Sociedade Empresária ou Empresa Individual de Responsabilidade Limitada: Ato Constitutivo, Estatuto ou Contrato Social', 'Ato Constitutivo', 'Inscrição CNPJ', 'Regularidade Fiscal Fazenda Nacional', 'FGTS', 'CNDT', 'Inscrição Contribuintes Estadual -Excluído para ME/EPP', 'Regularidade Fazenda Estadual - Excluído para ME/EPP', 'Certidão Negativa de Falência', 'Balanço Patrimonial - Excluído para ME/EPP', 'Boa Situação Financeira', 'Habilitação Técnica'])
    xpathProximaPagina = navegador.find_element(By.XPATH, '/html/body/app-root/div/div/div/app-cabecalho-selecao-fornecedores-governo/div[2]/app-selecao-fornecedores-governo-propostas-itens-grupo/div/div[4]/app-selecao-fornecedores-governo-proposta-item/p-tabview/div/div[2]/p-tabpanel[1]/div/div/span/div/app-listagem-propostas-subitens-governo/div[2]/p-dataview/div/p-paginator/div/button[3]')
    wb.save('Planilha Apoio Pregao ' + str(numero1) + '.xlsx')
    while True:
        try:
            for j in range (1 , 11):
                wbItens = wb[nomeEpresa1]
                numItem, valorEstimado, qntItem, valorOfertado = informaçoesItens(j)
                wbItens.append([numItem, nomeEmpresa, qntItem, valorEstimado, valorOfertado])
                wb.save('Planilha Apoio Pregao ' + str(numero1) + '.xlsx')
            xpathProximaPagina = navegador.find_element(By.XPATH, '/html/body/app-root/div/div/div/app-cabecalho-selecao-fornecedores-governo/div[2]/app-selecao-fornecedores-governo-propostas-itens-grupo/div/div[4]/app-selecao-fornecedores-governo-proposta-item/p-tabview/div/div[2]/p-tabpanel[1]/div/div/span/div/app-listagem-propostas-subitens-governo/div[2]/p-dataview/div/p-paginator/div/button[3]')
            xpathProximaPagina.click()
            randomWait(navegador)
        except:
            ##########################       SAIR DOS ITENS PARA O GRUPO (BOTAO VOLTAR DENTRO DOS ITENS) OU PROXIMA PAGINA       ############################
            urlAtual = navegador.current_url
            navegador.find_element(By.XPATH, '/html/body/app-root/div/div/div/app-cabecalho-selecao-fornecedores-governo/div[2]/app-selecao-fornecedores-governo-propostas-itens-grupo/div/div[4]/app-selecao-fornecedores-governo-proposta-item/app-acoes-governo-na-proposta-item/div/button').click()
            wait60(navegador,urlAtual)
            break
##########################        FIM DO SCRAPY        ############################