import pyautogui
import pyperclip
from time import sleep
import pandas as pd
# .clik -> clicar
# .write -> escrever
# .press -> pressionar
# .hotkey -> atalhos
pyautogui.PAUSE = 1

# 1 - ENTRAR NO GOOGLE, ENTRAR NO DRIVE, FAZER DOWLOAD DO ARQUIVO
pyautogui.press("win")
pyautogui.write("chrome")
pyautogui.press("Enter")

#pyperclip.copy("https://drive.google.com/drive/folders/149xknr9JvrlEnhNWO49zPcw0PW5icxga")
#pyautogui.hotkey("ctrl", "v")

pyautogui.write("https://drive.google.com/drive/folders/149xknr9JvrlEnhNWO49zPcw0PW5icxga")
pyautogui.press("Enter")
sleep(5)
pyautogui.click(x=378, y=292, clicks=2)
sleep(2)
pyautogui.click(x=460, y=291)
pyautogui.click(x=1068, y=188)
pyautogui.click(x=850, y=618)

# 2 - LER ARQUIVO, SOMAR QUANTIDADE E FATURAMENTO
tabela = pd.read_excel(r"C:\Users\PICHAU\Downloads\Vendas - Dez.xlsx") #Colocar r antes do caminho do arquivo

print(tabela)

quantidade = tabela["Quantidade"].sum() #.sum() = soma
faturamento = tabela["Valor Final"].sum()

print(f"Foram vendido {quantidade} itens")
print(f"O faturamento foi de {faturamento}")

# 3 - ENVIAR EMAIL

pyautogui.press("win")
pyautogui.write("chrome")
pyautogui.press("Enter")

pyautogui.write("https://mail.google.com/mail/u/0/?tab=rm&ogbl#inbox")
pyautogui.press("Enter")
sleep(5)

pyautogui.click(x=95, y=196)
pyautogui.write("luizfilypeoliveiradesiqueira@gmail.com")
pyautogui.hotkey("tab")
pyautogui.hotkey("tab")
pyautogui.write("Relatorio de Vendas")
pyautogui.hotkey("tab")
texto = f"""Bom Dia
Segue a analise da quantidade de produtos e do faturamento do dia de ontem
Quantidade de Produtos: {quantidade}
Faturamento XX/XX/XXXX : {faturamento}

Ass. Luiz Oliveira"""

pyautogui.write(texto)
pyautogui.hotkey("ctrl", "enter")


