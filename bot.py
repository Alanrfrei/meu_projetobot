import pandas
import pyautogui
import time

tabela = pandas.read_excel('Lista_Produtos.xlsx')
print (tabela)



pyautogui.PAUSE = 0.5

#passo 1 - entra no sistema 
pyautogui.press('win')

#passo 2 - entra no site
pyautogui.write('chorome')
pyautogui.press('enter')
  
#link site
pyautogui.position(x=620, y=101)

pyautogui.click(x=620, y=101)


pyautogui.write('https://drive.google.com/drive/u/0/folders/1R3fVLO99yb4b10wu5HM6_WEbulwz9-Ub')

pyautogui.press('enter')

time.sleep(2.5)

pyautogui.click(x=389, y=416)

pyautogui.press('enter')

time.sleep(2)

pyautogui.click(x=74, y=327)

for linha in tabela.index:
    codigo = str(tabela.loc[linha,'Código'])
    pyautogui.write(codigo)
    pyautogui.press('tab')

    descricao = str(tabela.loc[linha,'Descrição'])
    pyautogui.write(descricao)
    
    pyautogui.press('tab')

    und = str(tabela.loc[linha,'Unidade'])
    pyautogui.write(und)
    pyautogui.press('tab')

    preco = str(tabela.loc[linha,'Preço Custo'])
    pyautogui.write(preco)
    pyautogui.press('tab')  
     
    preco_venda = str (tabela.loc[linha,'Preço Venda'])
    pyautogui.write(preco_venda)
    pyautogui.press('tab') 

    preco_pra = str(tabela.loc[linha,'Preço Aprazo'])
    pyautogui.write(preco_pra)
    pyautogui.press('tab')  

    preco_ata = str(tabela.loc[linha,'Preço Atacado'])
    pyautogui.write(preco_ata)
    pyautogui.press('tab')  

    estoque = str(tabela.loc[linha,'Estoque'])
    pyautogui.write(estoque)
    pyautogui.press('tab')  

    pyautogui.press('enter')
    pyautogui.press('enter')

    pyautogui.hotkey('shift'+ 'tab')

   




