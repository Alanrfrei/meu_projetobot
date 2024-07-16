import pyautogui
import time
#pyautogui-click - clica o mouse
#pyautogui.write - escreve texto
#pyautogui.press - aperta uma clica
#pyautogui.hotkey - combinar teclar (ctrl + c)
#pyautogui.scroll - rola a tela

#pausa entre as abas 
pyautogui.PAUSE = 0.5

#passo 1 - entra no sistema 
pyautogui.press('win')

#passo 2 - entra no site
pyautogui.write('chorome')
pyautogui.press('enter')

#link site
pyautogui.position(x=1286, y=188)
pyautogui.press('enter')

#clica onde quiser 
pyautogui.click(x=1286, y=188)

#login
pyautogui.hotkey('ctrl','a')

pyautogui.write('alanrodrigues@gmail.com')

pyautogui.press('tab')
pyautogui.write('1234567')
pyautogui.click(x=664, y= 577)
time.sleep(2)

#banco de dados 
import pandas
tabela = pandas.read_excel('Lista_Produtos.xlsx')       #tabela de produtos 
 
print (tabela)

#cadastra produtos
#loop
for linha in tabela.index:  

    pyautogui.click(x=518, y=302)

    #codigo
    codigo = str(tabela.loc[linha,'codigo'])
    pyautogui.write(codigo)
    pyautogui.press('tab')

    #marca
    marca = str (tabela.loc[linha,'marca'])
    pyautogui.write(marca)
    pyautogui.press('tab')

    #tipo
    tipo = str (tabela.loc[linha, 'tipo'])  
    pyautogui.write(tipo)
    pyautogui.press('tab')

    #categoria
    categoria = str (tabela.loc[linha,'categoria'])
    pyautogui.write(categoria)
    pyautogui.press('tab')

    #preco_unitario
    preço = str (tabela.loc[linha,'preco_unitario'])
    pyautogui.write(preço)
    pyautogui.press('tab')

    #custo
    custo = str (tabela.loc[linha, 'custo'])
    pyautogui.write(custo)
    pyautogui.press('tab')
    #obs
    obs= str (tabela.loc[linha, 'obs'])
    if obs != 'nam':
        pyautogui.write(obs) 

    pyautogui.press('tab')
    pyautogui.press('enter')    

    #rolagem da tela (- e para baixo e + mais para cima)
    pyautogui.scroll(1000)
    

