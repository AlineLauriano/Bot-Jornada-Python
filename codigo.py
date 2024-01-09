#AUTOMAÇÃO DE TAREFAS
# Bibliotecas usadas:
    # pyautogui
    # pandas
    # numpy
    # openpyxl

###  Passo a passo do projeto ###

import pyautogui
import time
import pandas as pd
import numpy
import openpyxl

# PRINCIPAIS COMANDOS DO PROJETO
    # clicar -> pyautogui.click
    # escrever -> pyautogui.write
    # apertar uma tecla -> pyautogui.press
    # apertar -> pyautogui.hotkey()
    # rolar a tela -> pyautogui.scroll

######################################################################

# Passo 1 - entrar no sistema da empresa
    # https://dlp.hashtagtreinamentos.com/python/intensivao/login

pyautogui.PAUSE = 1 # defini que a cada comando irá esperar 1 segundo

# aperta a tecla do windows
pyautogui.press("win")

# digita o nome do programa chrome
pyautogui.write("chrome")

#pressiona a tecla enter
pyautogui.press("enter")

pyautogui.click(x=721, y=64)

#digitar o link
link = 'https://dlp.hashtagtreinamentos.com/python/intensivao/login'
pyautogui.write(link)

#pressiona a tecla enter
pyautogui.press("enter")

time.sleep(2)

# Passo 2 - Fazer Login

# Seleciona a caixa de email
pyautogui.click(x=631, y=410)

# digitar o e-mail
pyautogui.write("aline_rgt@hotmail.com")

# pressionando a tecla tab
pyautogui.press("tab")

# digitar a senha 
pyautogui.write("123")

# clicar no botão de login
pyautogui.click(x=686, y=567)
time.sleep(2)

# Passo 3 - Importar a base de dados
# o arquivo que está sendo importado deve estar na mesma pasta, ou você tem que passar o caminho inteiro c://..
tabela = pd.read_csv('produtos.csv')
print(tabela)

# Passo 4 - Importar um produto

for linha in tabela.index:
    #selecionando o primeiro campo do formulário
    pyautogui.click(x=484, y=291)

    #codigo
    codigo = tabela.loc[linha, "codigo"]
    pyautogui.write(codigo)
    pyautogui.press("tab")

    #marca
    marca = tabela.loc[linha,"marca"]
    pyautogui.write(marca)
    pyautogui.press("tab")

    #tipo
    tipo = tabela.loc[linha,"tipo"]
    pyautogui.write(tipo)
    pyautogui.press("tab")

    #categoria
    categoria = (str(tabela.loc[linha,"categoria"]))
    pyautogui.write(categoria)
    pyautogui.press("tab")

    #preço unitário
    preco_unitario = (str(tabela.loc[linha,"preco_unitario"]))
    pyautogui.write(preco_unitario)
    pyautogui.press("tab")

    #custo
    custo = (str(tabela.loc[linha,"custo"]))
    pyautogui.write(custo)
    pyautogui.press("tab")

    # OBS
    obs = tabela.loc[linha, "obs"]
    

    if not pd.isna(obs):
        pyautogui.write(obs)

    #enviar o produto
    pyautogui.press("tab")
    pyautogui.press("enter")

    #rolando a tela para cima
    pyautogui.scroll(500)

# Repetir o passao 4, até acabar a base de dados

