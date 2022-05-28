from selenium import webdriver

from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By

navegador= webdriver.Chrome(executable_path= r"./chromedriver.exe")
navegador.get("https://www.google.com.br/")
navegador.find_element(By.XPATH,'/html/body/div[1]/div[3]/form/div[1]/div[1]/div[1]/div/div[2]/input').send_keys("Cotaçao Dolar", Keys.ENTER)
cotaçao_dolar= navegador.find_element(By.XPATH, '//*[@id="knowledge-currency__updatable-data-column"]/div[1]/div[2]/span[1]').get_attribute("data-value")
print(cotaçao_dolar)
navegador.get("https://www.google.com.br/")
navegador.find_element(By.XPATH, '/html/body/div[1]/div[3]/form/div[1]/div[1]/div[1]/div/div[2]/input').send_keys("Cotaçao Euro", Keys.ENTER)
cotaçao_euro= navegador.find_element(By.XPATH,'//*[@id="knowledge-currency__updatable-data-column"]/div[1]/div[2]/span[1]').get_attribute("data-value")
print(cotaçao_euro)
navegador.get("https://www.melhorcambio.com/ouro-hoje")
cotaçao_ouro= navegador.find_element(By.XPATH, '//*[@id="comercial"]').get_attribute("value")
cotaçao_ouro= cotaçao_ouro.replace(",",".")
print(cotaçao_ouro)
navegador.quit()


import pyautogui
import pyperclip
import time

import pandas as pd
tabela= pd.read_excel(r"C:\Users\lukes\Downloads\Produtos.xlsx")
tabela.loc[tabela["Moeda"]=="Dólar", "Cotação"]= float(cotaçao_dolar)
tabela.loc[tabela["Moeda"]=="Euro", "Cotação"]= float(cotaçao_euro)
tabela.loc[6,"Cotação"]= float(cotaçao_ouro)

tabela["Preço de Compra"]= tabela["Preço Original"]* tabela["Cotação"]
tabela["Preço de Venda"]= tabela["Preço de Compra"]* tabela["Margem"]

tabela.to_excel(r"C:\Users\lukes\Downloads\Produtos Novos.xlsx", index= False)

time.sleep(2)
pyautogui.PAUSE=1

pyautogui.press("Win")
pyautogui.write("Chrome")
pyautogui.press("Enter")
pyperclip.copy("https://mail.google.com/mail/u/0/#inbox")
pyautogui.hotkey("ctrl", "v")
pyautogui.press("Enter")

time.sleep(4)

pyautogui.click(x=39, y=203)
pyautogui.write("emailqualquer@gmail.com")
pyautogui.press("tab", presses=2)
pyautogui.write("Relatorio")
pyautogui.press("tab")


cotaçao_dolar = round(float(cotaçao_dolar),2)
cotaçao_euro = round(float(cotaçao_euro),2)
cotaçao_ouro = round(float(cotaçao_ouro),2)

texto= f"""
BOM DIA, SENHORES.
AQUI ESTA A COTAÇAO DO DOLAR, EURO E OURO EM TEMPO REAL:
Dólar:{cotaçao_dolar}
Euro:{cotaçao_euro}
Ouro:{cotaçao_ouro}
Segue em anexo a atualização da tabela.
Abs, Lucas.
"""
pyperclip.copy(texto)
pyautogui.hotkey("ctrl","v")
pyautogui.click(x=961, y=698)
pyautogui.write("Produtos Novos")
pyautogui.press("PgDn")
pyautogui.press("Enter")
time.sleep(1)
pyautogui.hotkey("ctrl","Enter")
time.sleep(1)
pyautogui.click(x=1337, y=11)