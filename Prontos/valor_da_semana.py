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
tabela = pd.read_excel(r"C:\Users\luiso\Desktop\Planilhas_Marcio\marcio_vendas.xlsx", sheet_name=1)
print(tabela)

#variaveis
id = int(input('Digite o numero da semana: '))
mes = str(input("Digite o mês: "))
id2 = int(input('Digite o numero da outra semana: '))
mes2 = str(input("Digite o mês: "))

#Filtrando Valores da semana digitada

#filtrando o mes
order = tabela["Mes"] == mes
df = tabela[order]
#filtrando a semana
filtro = df["Semana"] == id
tabela2 = df[filtro]

print(tabela2)

#Filtrando Valores da segunda semana digitada
#filtrando o mes
order2 = tabela["Mes"] == mes2
df2 = tabela[order2]
#filtrando a semana
filtro2 = df2["Semana"] == id2
tabela3 = df2[filtro2]

print(tabela3)

#pegar o valor o mes
filtromes = tabela2["Mes"] == mes
tabelames = tabela2[filtromes]
#Fazendo com que a variavel vira uma string
tabelames = (tabelames.iloc[0]["Mes"])
print(tabelames)

#pegar o valor o mes 2
filtromes2 = tabela3["Mes"] == mes2
tabelames2 = tabela3[filtromes2]
#Fazendo com que a variavel vira uma string
tabelames2 = (tabelames2.iloc[0]["Mes"])
print(tabelames2)

#Os valores vendidos (usei o max pois so tem um valor no filtro e ele vira um numero float no final)
valor_vendido = tabela2["Total"].max()
valor_vendido2 = tabela3["Total"].max()

print(valor_vendido)
print(valor_vendido2)

#Criando porcentagem de venda das semanas selecionadas
porcentagem = ((valor_vendido - valor_vendido2)/valor_vendido2) * 100
print(porcentagem)

#Texto para ser enviado
texto = f"""
Ola!!!

O total de vendas da semana {id} de {tabelames} foi de: R$ {valor_vendido:,.2f}
O total de vendas da semana {id2} de {tabelames2} foi de: R$ {valor_vendido2:,.2f}
Porcentagem de venda das duas semanas foram: {porcentagem:,.1f}%
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