import pyautogui
import time
import pandas as pd

#====================================================================

# configuração
pyautogui.FAILSAFE = True # Caso precise parar é colocar o mouse para a esquerda
pyautogui.PAUSE = 1

#====================================================================

# Abrir menu do Windows 
pyautogui.press("win")

#====================================================================

# Abrir programa 
print("1. Abrindo programa")
pyautogui.write("google chrome")
time.sleep(1)
pyautogui.press("enter")
time.sleep(1)
pyautogui.press("Ctrl") # atalho para maximizar janela

#====================================================================

# Entrar no site 
time.sleep(2)
print("2. Entrando no site")
pyautogui.write("https://drive.google.com/drive/folders/14")
time.sleep(1) 
pyautogui.press("enter")

#====================================================================

# Abrir diretorio 
time.sleep(3)
print("3.  Abrir diretorio")
time.sleep(4)
pyautogui.click(x=383, y=252, clicks =2)

#====================================================================

# baixar arquivo 
time.sleep(4)
print("4. Baixar arquivo")
pyautogui.click(x=388, y=343)
pyautogui.click(x=1132, y=157)
pyautogui.click(x=907, y=560)

#====================================================================

# Entrar no arquivo 
time.sleep(8)
tabela = pd.read_excel("Vendas.xlsx", engine='openpyxl')
faturamento = tabela["Valor Final"].sum() # para "somar"
print(tabela)
print()
print("Faturamento total: R$", faturamento)

#====================================================================

# Abrir whatsapp 
print("5. Abrir whatsapp")
pyautogui.press("F9")
time.sleep(2) 
pyautogui.write("https://web.whatsapp.com/")
pyautogui.press("enter")
time.sleep(25)

#====================================================================

# Enviar mensagem
print("6. Enviar mensagem")
pyautogui.click(x=231, y=245)
pyautogui.write = ("Meu Numero")
pyautogui.click(x=267, y=362) # Selecionar chat
pyautogui.click(x=627, y=959) # Escrever
pyautogui.write = ("Meu faturamento: R$", faturamento)






