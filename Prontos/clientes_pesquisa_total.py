import pyautogui
import pyperclip
import time
import pandas as pd
import numpy
import openpyxl
import datetime

pyautogui.PAUSE = 2

# #Pegando as datas
data = datetime.date.today()
nova_data = data.strftime("%d/%m/%Y")

#Lendo a tabela
tabela = pd.read_excel(r"C:\Users\luiso\Desktop\Planilhas_Marcio\marcio_vendas.xlsx", sheet_name="Fluxo")
print(tabela)


pessoa = str(input('Digite o nome da Pessoa: '))

#Filtrando as compras da cliente
filtro = tabela["Nome"] == pessoa
tabela2 = tabela[filtro]

print(tabela2)

# Pegando o Nome da cliente
filtronome = tabela2["Nome"] == pessoa
tabelanome = tabela2[filtronome]
nome_exibir = (tabelanome.iloc[0]["Nome"])
print(nome_exibir)

#Pegando o valor total de compras do cliente
valortotal = tabela2["Valor Total"].sum()
print(valortotal)

#Texto para ser enviado
texto = f"""
Ola!!!
O total de compras de {nome_exibir} foi de: R$ {valortotal:,.2f}
abs: dia da consulta: {nova_data}

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