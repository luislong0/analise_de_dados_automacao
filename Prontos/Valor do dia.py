import pyautogui
import pyperclip
import time
import pandas as pd
import numpy
import openpyxl
import datetime


pyautogui.PAUSE = 2

#Pegando as datas
data = datetime.date.today()
#retroceder o dia
data_anterior = data - datetime.timedelta(days=1)

nova_data = data.strftime("%d/%m/%Y")
datastr = str(data)
data_anterior_str = str(data_anterior)

print(data)
print(nova_data)

#Lendo a tabela
tabela = pd.read_excel(r"C:\Users\luiso\Desktop\Planilhas_Marcio\marcio_vendas.xlsx", sheet_name="Fluxo")
print(tabela)

#Valor recebido data atual
#Filtrando a tabela
filtro = tabela["Data"] == datastr
tabela2 = tabela[filtro]
print(tabela2)

#Soma dos valores selecionados
valor_vendido = tabela2["Valor Total"].sum()
valor_recebido = tabela2["Quantidade Paga"].sum()

print (valor_vendido)
print (valor_recebido)

#Valor recebido data anterior
#Filtrando a tabela
filtro2 = tabela["Data"] == data_anterior_str
tabela3 = tabela[filtro2]
print (tabela3)

valor_vendido2 = tabela3["Valor Total"].sum()
valor_recebido2 = tabela3["Quantidade Paga"].sum()
#se nao vender nada no dia anterior, se atribui o valor "1" para fazer a porcentagem
if(valor_vendido2 == 0.0):
    valor_vendido2 = 1.0
if(valor_vendido2 == 'NaN'):
    porcentagem = 0

#Criando porcentagem de venda do dia
print (valor_vendido2)
print (valor_recebido2)
#calculo da porcentagem
porcentagem = ((valor_vendido - valor_vendido2)/valor_vendido2) * 100
print(porcentagem)

#Texto para ser enviado
texto = f"""
Ola!!!

O total de vendas hoje em palmital foi de: R$ {valor_vendido:,.2f}
A quantidade recebida hoje foi de : R$ {valor_recebido:,.2f}
Porcentagem de venda do dia em comparação do dia anterior: {porcentagem:,.1f}%
abs: dia: {nova_data}

by: Luis Otavio Longo Pereira
"""
pyperclip.copy(texto)

#Automação
pyautogui.press("win")
pyautogui.write("Opera")
pyautogui.press("enter")

pyautogui.click(x=21, y=161)
time.sleep(5)
pyautogui.click(x=173, y=114)

pyautogui.write("mp cosme")
pyautogui.hotkey("enter")
pyautogui.hotkey("ctrl", "v")
# pyautogui.hotkey("enter")