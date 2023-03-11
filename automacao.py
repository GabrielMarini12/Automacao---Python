import pandas as pd

#baixei o relatório
df = pd.read_excel("C:\\Users\\Usuario\Desktop\\PROJECTS\\ZAGONELL_PROJECTS\\EmitidosDezembro2022.xls")
print(df)

#tratando o relatório
df["Percentual de comissão"] = pd.to_numeric(df["Percentual de comissão"], errors="coerce")
print(df.info())

print(df)

#excluindo as colunas totalmente vazias
df = df.dropna(how="all", axis=1)
print(df)

print(df.info())

#caclculando a média de comissão
mediacomissao = df["Percentual de comissão"].sum()
mediacomissao = mediacomissao / 179
print(mediacomissao)

#contando os produtos
produtos = df["Célula"].value_counts()
print(produtos)

#contando os produtos em %
#produtos1 = df["Célula"].value_counts(normalize=True).map("{:.1%}".format)
#print(produtos1)

import pyautogui
import time
import pyperclip

pyautogui.PAUSE = 1

#abrir uma janela no google
pyautogui.press("winleft")
pyautogui.write("Chrome")
pyautogui.press("enter")
pyperclip.copy("https://www.icloud.com/")
pyautogui.hotkey("ctrl", "v")
pyautogui.press("enter")

time.sleep(5)

#entrando no icloud
pyautogui.click(x=670, y=675)

time.sleep(5)

#colocando o login
pyautogui.click(x=631, y=484)
pyperclip.copy("gabrielmarini12@icloud.com")
pyautogui.hotkey("ctrl", "v")
pyautogui.press("enter")

time.sleep(4)

#colocando a senha
pyperclip.copy("Drugm12*")
pyautogui.hotkey("ctrl", "v")
pyautogui.press("enter")

time.sleep(15)

#enviando o e-mail
pyautogui.click(x=1149, y=196)

time.sleep(8)

pyperclip.copy("gabrielmarini12@icloud.com")
pyautogui.hotkey("ctrl", "v")
pyautogui.hotkey("tab")
pyautogui.hotkey("tab")
pyautogui.hotkey("tab")
pyperclip.copy("Relatório Mensal - Vendas")
pyautogui.hotkey("ctrl", "v")
pyautogui.hotkey("tab")

texto = f"""
Prezados,

Segue relatório de vendas.
Média de comissão: {mediacomissao:.2f}
Produtos vendidos: {produtos}
    
Qualquer dúvida estou à disposição.
Att.,
Gabriel Marini

"""

pyperclip.copy(texto)
pyautogui.hotkey("ctrl", "v")

time.sleep(3)

pyautogui.click(x=1045, y=209)