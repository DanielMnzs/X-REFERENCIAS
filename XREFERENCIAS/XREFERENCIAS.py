import pyautogui
import time 
import pyscreeze
import PIL
import pandas as pd 


caminho_rua = r"//depo0903/gpa$/PAC/Daniel Menezes/Python/XREFERENCIAS/1R.png"
caminho_excel = r"//depo0903/gpa$/PAC/Daniel Menezes/Python/XREFERENCIAS/xreferencias.xlsx"
caminho_pacotes = r"//depo0903/gpa$/PAC/Daniel Menezes/Python/XREFERENCIAS/PACOTES.PNG"
ler = pd.read_excel(caminho_excel)
tabela = pd.DataFrame(ler)
print(tabela)


x = 0
y = 0

while True and x < len (tabela):

    plu = tabela.iloc[x, y + 1]
    print(plu)
    rua = tabela.iloc[x, y + 11]
    print(rua)
    pyautogui.doubleClick(205, 195)
    time.sleep(1)
    pyautogui.write(str(plu),interval=0.111)
    time.sleep(1)
    pyautogui.leftClick(1087, 196)
    time.sleep(1)
    pyautogui.leftClick(33, 299)
    time.sleep(1)
    pyautogui.leftClick(30, 698)
    time.sleep(1.5)

    for i in range (True):
        try:
            pacotes = pyautogui.locateOnScreen(caminho_pacotes,confidence=0.585)
            if pacotes:
                print("IMAGEM ENCONTRADA")
                time.sleep(1)
                pyautogui.click(pacotes)
                break
            
            
        except:
            print("IMAGEM NÃO ENCONTRADA")
            time.sleep(1)
            print("TENTANDO NOVAMENTE")
            time.sleep(0.5)
            time.sleep(1)
    time.sleep(2)
    pyautogui.leftClick(42, 697)
    time.sleep(1)


    print("TENTANDO NOVAMENTE")
       
    for i in range(50):
        try:    
            localizacao = pyautogui.locateOnScreen(caminho_rua,confidence=0.699)
            if localizacao:
                print("ENCONTRADO")
                time.sleep(1)
                pyautogui.doubleClick(localizacao)
                pyautogui.hotkey("backspace")
                pyautogui.write(rua,interval=0.111)
                break
            
            
        except pyautogui.ImageNotFoundException:
            print("NÃO ENCONTRADO")
            time.sleep(1)
            pyautogui.scroll(-100)
            pyautogui.leftClick(1176, 400)
    pyautogui.leftClick(1237, 694)
    time.sleep(2)
    pyautogui.leftClick(21, 145)
    time.sleep(2)
    x = x + 1 

