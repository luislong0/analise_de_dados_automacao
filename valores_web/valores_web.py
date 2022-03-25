from selenium import webdriver #controlar o navegador
from selenium.webdriver.common.keys import Keys #usar o teclado
from selenium.webdriver.common.by import By #localizar os itens do navegador
import pandas as pd
import numpy
import openpyxl
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager


#navegador - chrome
#chromedriver

#criar navegador

#formula 1 - Mini erro
# navegador = webdriver.Chrome(executable_path="chromedriver.exe")

#formula 2 - Sem erro

navegador = webdriver.Chrome(service=Service(ChromeDriverManager().install()))


#entraria no google e pesquisaria por cotação do dolar
navegador.get("https://www.google.com/")

#pegaria a cotação do dolar

navegador.find_element(By.XPATH, '/html/body/div[1]/div[3]/form/div[1]/div[1]/div[1]/div/div[2]/input').send_keys("cotação dolar")

navegador.find_element(By.XPATH, '/html/body/div[1]/div[3]/form/div[1]/div[1]/div[1]/div/div[2]/input').send_keys(Keys.ENTER)

cotacao_dolar = navegador.find_element(By.XPATH, '//*[@id="knowledge-currency__updatable-data-column"]/div[1]/div[2]/span[1]').get_attribute('data-value')

print(cotacao_dolar)
#entraria no google e pesquisaria a cotação do euro

navegador.get("https://www.google.com/")

navegador.find_element(By.XPATH, '/html/body/div[1]/div[3]/form/div[1]/div[1]/div[1]/div/div[2]/input').send_keys("cotação euro")

navegador.find_element(By.XPATH, '/html/body/div[1]/div[3]/form/div[1]/div[1]/div[1]/div/div[2]/input').send_keys(Keys.ENTER)

cotacao_euro = navegador.find_element(By.XPATH, '//*[@id="knowledge-currency__updatable-data-column"]/div[1]/div[2]/span[1]').get_attribute('data-value')

print(cotacao_euro)

#pegaria a cotação do ouro

navegador.get("https://www.melhorcambio.com/ouro-hoje")

cotacao_ouro = navegador.find_element(By.XPATH, '//*[@id="comercial"]').get_attribute('value')

cotacao_ouro = cotacao_ouro.replace(",", ".")
#pegaria a cotação do ouro

print(cotacao_ouro)

navegador.quit()

# Importar a base de dados

tabela = pd.read_excel("Produtos.xlsx")
print(tabela)

#atualizar cotações

#cotação do dolar
# tabela.loc [linha,coluna]
tabela.loc[tabela["Moeda"] == "Dólar", "Cotação"] = float(cotacao_dolar)
print(tabela)

#cotação do euro
# tabela.loc [linha,coluna]
tabela.loc[tabela["Moeda"] == "Euro", "Cotação"] = float(cotacao_euro)
print(tabela)

#cotação do ouro
# tabela.loc [linha,coluna]
tabela.loc[tabela["Moeda"] == "Ouro", "Cotação"] = float(cotacao_ouro)


#preço de compra

tabela["Preço de Compra"] = tabela["Preço Original"] * tabela["Cotação"]

#preço de venda
tabela["Preço de Venda"] = tabela["Preço de Compra"] * tabela["Margem"]
print(tabela)

#exportar base de dados

tabela.to_excel("ProdutosNovo.xlsx", index=False)








