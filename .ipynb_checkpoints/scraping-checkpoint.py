import sys
import os
import re
sys.path.append(os.path.dirname(os.getcwd()))
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.keys import Keys
import pandas as pd
from time import sleep
import chromedriver_autoinstaller as auto


def scraping():
    #define as configurações do browser
    chromedriver = auto.install()
    service = Service(chromedriver)
    options = Options()
    options.add_argument("start-maximized")
    options.add_experimental_option("prefs",{"download.default_directory":"C:\plansBot"})
    browser = webdriver.Chrome(service=service,options=options)
    browser.maximize_window()
    
    #pesquisa por notebooks na pagina
    browser.get("https://www.magazineluiza.com.br/")
    browser.find_element(By.ID,"input-search").send_keys("notebooks")
    browser.find_element(By.ID,"input-search").send_keys(Keys.ENTER)
    
    #cria um dicionário que será transformado em um dataframe
    notebooks = {"URL": [],
    "PRODUTO": [],
    "QTD_AVAL": []
    }
    
    #espera os itens aparecerem na tela
    while len(browser.find_elements(By.XPATH, '*//div[@data-testid="product-list"]//ul//li')) == 0:
        sleep(1)
        
    #itera na lista de produtos e salva no dicionario as informações solicitadas
    for item in browser.find_elements(By.XPATH, '*//div[@data-testid="product-list"]//ul//li'):
        link = item.find_element(By.TAG_NAME, "a").get_attribute("href")
        nome = item.find_element(By.TAG_NAME, "h2").text
        avaliacoes = item.find_element(By.TAG_NAME, "span").text
        try:
            qtd_avaliacoes = re.search("\((.*)\)", avaliacoes).group(1)
        except AttributeError:
            pass

        notebooks['URL'].append(link)
        notebooks['PRODUTO'].append(nome)
        notebooks['QTD_AVAL'].append(qtd_avaliacoes)


    #aplica os filtros necessários e divide a base em 2 dataframes
    df_produtos = pd.DataFrame(notebooks)
    df_produtos["QTD_AVAL"] = df_produtos["QTD_AVAL"].astype('int')
    df_filtrado = df_produtos[df_produtos['QTD_AVAL'] > 0]

    Melhores = df_produtos[df_produtos['QTD_AVAL'] >= 100]
    Piores = df_produtos[df_produtos['QTD_AVAL'] < 100]

    # salva os dataframes no mesmo arquivo
    with pd.ExcelWriter('output//Notebooks.xlsx') as writer:
        # Escrevendo df1 na primeira planilha chamada 'Planilha1'
        Melhores.to_excel(writer, sheet_name='Melhores', index=False)

        # Escrevendo df2 na segunda planilha chamada 'Planilha2'
        Piores.to_excel(writer, sheet_name='Piores', index=False)
        
    browser.quit()
    