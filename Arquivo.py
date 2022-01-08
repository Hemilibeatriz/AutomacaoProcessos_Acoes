'''
    Instalado pelo próprio PyCharm
        install package selenium
        install package pandas
    Instalado pelo Terminal
        pip install pywin32
    Instalado pelo Settings
        openpyxl
'''
#importar a data atual
from datetime import datetime

data = (datetime.today().strftime('%d-%m-%Y'))

from selenium import webdriver
from selenium.webdriver.common.keys import Keys

# para rodar o chrome em 2º plano
# from selenium.webdriver.chrome.options import Options
# chrome_options = Options()
# chrome_options.headless = True
# navegador = webdriver.Chrome(options=chrome_options)

# abrir um navegador
navegador = webdriver.Chrome('chromedriver')

navegador.get("https://www.google.com/")

# Passo 1: Pegar as cotações de interesse

#Drogasil
navegador.find_element_by_xpath(
    '/html/body/div[1]/div[3]/form/div[1]/div[1]/div[1]/div/div[2]/input').send_keys("ações raia drogasil")

navegador.find_element_by_xpath(
    '/html/body/div[1]/div[3]/form/div[1]/div[1]/div[3]/center/input[1]').send_keys(Keys.ENTER)

drogasil = navegador.find_element_by_xpath(
    '//*[@id="knowledge-finance-wholepage__entity-summary"]/div/g-card-section/div/g-card-section/div[2]/div[1]/span[1]/span/span[1]').text

#Magazine Luiza
navegador.get("https://www.google.com/")
navegador.find_element_by_xpath(
    '/html/body/div[1]/div[3]/form/div[1]/div[1]/div[1]/div/div[2]/input').send_keys("ações magalu")

navegador.find_element_by_xpath(
    '/html/body/div[1]/div[3]/form/div[1]/div[1]/div[3]/center/input[1]').send_keys(Keys.ENTER)

magazineluiza = navegador.find_element_by_xpath(
    '//*[@id="knowledge-finance-wholepage__entity-summary"]/div/g-card-section/div/g-card-section/div[2]/div[1]/span[1]/span/span[1]').text

#Amazon
navegador.get("https://www.google.com/")
navegador.find_element_by_xpath(
    '/html/body/div[1]/div[3]/form/div[1]/div[1]/div[1]/div/div[2]/input').send_keys("ações amazon")

navegador.find_element_by_xpath(
    '/html/body/div[1]/div[3]/form/div[1]/div[1]/div[3]/center/input[1]').send_keys(Keys.ENTER)

amazon = navegador.find_element_by_xpath(
    '//*[@id="knowledge-finance-wholepage__entity-summary"]/div/g-card-section/div/g-card-section/div[2]/div[1]/span[1]/span/span[1]').text

#Amaricanas
navegador.get("https://www.google.com/")
navegador.find_element_by_xpath(
    '/html/body/div[1]/div[3]/form/div[1]/div[1]/div[1]/div/div[2]/input').send_keys("ações americanas")

navegador.find_element_by_xpath(
    '/html/body/div[1]/div[3]/form/div[1]/div[1]/div[3]/center/input[1]').send_keys(Keys.ENTER)

americanas = navegador.find_element_by_xpath(
    '//*[@id="knowledge-finance-wholepage__entity-summary"]/div/g-card-section/div/g-card-section/div[2]/div[1]/span[1]/span/span[1]').text

#Netflix
navegador.get("https://www.google.com/")
navegador.find_element_by_xpath(
    '/html/body/div[1]/div[3]/form/div[1]/div[1]/div[1]/div/div[2]/input').send_keys("ações netflix")

navegador.find_element_by_xpath(
    '/html/body/div[1]/div[3]/form/div[1]/div[1]/div[3]/center/input[1]').send_keys(Keys.ENTER)

netflix = navegador.find_element_by_xpath(
    '//*[@id="knowledge-finance-wholepage__entity-summary"]/div/g-card-section/div/g-card-section/div[2]/div[1]/span[1]/span/span[1]').text

#Petrobrás
navegador.get("https://www.google.com/")
navegador.find_element_by_xpath(
    '/html/body/div[1]/div[3]/form/div[1]/div[1]/div[1]/div/div[2]/input').send_keys("ações petrobrás")

navegador.find_element_by_xpath(
    '/html/body/div[1]/div[3]/form/div[1]/div[1]/div[3]/center/input[1]').send_keys(Keys.ENTER)

petrobras = navegador.find_element_by_xpath(
    '//*[@id="knowledge-finance-wholepage__entity-summary"]/div/g-card-section/div/g-card-section/div[2]/div[1]/span[1]/span/span[1]').text

#Taesa
navegador.get("https://www.google.com/")
navegador.find_element_by_xpath(
    '/html/body/div[1]/div[3]/form/div[1]/div[1]/div[1]/div/div[2]/input').send_keys("ações taee4f")

navegador.find_element_by_xpath(
    '/html/body/div[1]/div[3]/form/div[1]/div[1]/div[3]/center/input[1]').send_keys(Keys.ENTER)

taesa = navegador.find_element_by_xpath(
    '//*[@id="knowledge-finance-wholepage__entity-summary"]/div/g-card-section/div/g-card-section/div[2]/div[1]/span[1]/span/span[1]').text

#Ouro
navegador.get("https://www.melhorcambio.com/ouro-hoje")

ouro = navegador.find_element_by_xpath('//*[@id="comercial"]').get_attribute("value")

#Dólar
navegador.get("https://www.google.com/")
navegador.find_element_by_xpath(
    '/html/body/div[1]/div[3]/form/div[1]/div[1]/div[1]/div/div[2]/input').send_keys("cotação dólar")

navegador.find_element_by_xpath(
    '/html/body/div[1]/div[3]/form/div[1]/div[1]/div[3]/center/input[1]').send_keys(Keys.ENTER)

dolaroriginal = navegador.find_element_by_xpath(
    '//*[@id="knowledge-currency__updatable-data-column"]/div[1]/div[2]/span[1]').get_attribute("data-value")
dolar = "{:.2f}".format(float(dolaroriginal))
dolar = dolar.replace(".", ",")

navegador.quit()

# Passo 2: Importar a Planilha
import pandas as pd

tabela = pd.read_excel("Planilha-Ações.xlsx")

# Passo 3: atualizar a Planilha => na data correta
# nas linhas onde na coluna "Empresa" = Empresa
tabela.loc[tabela["Empresa"] == "Drogasil", data] = (drogasil)
tabela.loc[tabela["Empresa"] == "Magazine Luiza", data] = (magazineluiza)
tabela.loc[tabela["Empresa"] == "Amazon", data] = (amazon)
tabela.loc[tabela["Empresa"] == "Americanas", data] = (americanas)
tabela.loc[tabela["Empresa"] == "Netflix", data] = (netflix)
tabela.loc[tabela["Empresa"] == "Petrobrás", data] = (petrobras)
tabela.loc[tabela["Empresa"] == "Taesa", data] = (taesa)
tabela.loc[tabela["Empresa"] == "Ouro", data] = (ouro)
tabela.loc[tabela["Empresa"] == "Dólar", data] = (dolar)

# Passo 4: Salvar a Planilha
tabela.to_excel("Planilha-Ações.xlsx", index=False)

# Passo 5: Enviar um email com o dado que foi Inserido na Planilha
import win32com.client as win32
outlook = win32.Dispatch('outlook.application')
mail = outlook.CreateItem(0)
mail.To = 'lobeatrizlobeatriz@gmail.com'
mail.Subject = 'Valores das Ações na data de Hoje'
mail.HTMLBody = f'''
<p>Prezada mamãe</p>

<p>Consta abaixo o valor das ações que a gente acompanha na data de hoje, dia {data}.</p>

<p>Drogasil:
{drogasil}

<p>Magazine Luiza:
{magazineluiza}

<p>Amazon:
{amazon}

<p>Americanas:
{americanas}

<p>Netflix:
{netflix}

<p>Petrobrás:
{petrobras}

<p>Taesa:
{taesa}

<p>Ouro:
{ouro}

<p>Dólar:
{dolar}

<p>Beijos da Programadora,</p>
<p>Mi</p>
'''

mail.Send()