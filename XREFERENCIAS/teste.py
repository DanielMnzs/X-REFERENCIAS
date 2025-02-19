import pyautogui
import pyscreeze
import PIL
import time 
import pandas as pd
caminho_rua = r"//depo0903/gpa$/PAC/Daniel Menezes/Python/XREFERENCIAS/1R.png"
caminho_excel = r"//depo0903/gpa$/PAC/Daniel Menezes/Python/XREFERENCIAS/xreferencias.xlsx"
caminho_pacotes = r"//depo0903/gpa$/PAC/Daniel Menezes/Python/XREFERENCIAS/PACOTES.PNG"
ler = pd.read_excel(caminho_excel)
tabela = pd.DataFrame(ler)
print(tabela)


x = 0
y = 0


while True:
    pyautogui.scroll(-10)