"""Bot que pega produtos numa planilha excel e 
cadastra eles numa plataforma""" 
import openpyxl
import pyperclip
import pyautogui
from time import sleep

#Entrar na planilha
workbook = openpyxl.load_workbook('produtos_ficticios.xlsx')
sheet_produtos = workbook['Produtos']

#Copiar informações de um campo e colar no seu campo correspondente
for linha in sheet_produtos.iter_rows(min_row=2):
   #Nome do produto
    nome_produto = linha[0].value
    pyperclip.copy(nome_produto)
    pyautogui.click(868,195, duration=1)
    pyautogui.hotkey('ctrl','v')
    
    #descricao
    descricao = linha[1].value
    pyperclip.copy(descricao)
    pyautogui.click(871,278, duration=1)
    pyautogui.hotkey('ctrl','v')

    #categoria
    categoria = linha[2].value
    pyperclip.copy(categoria)
    pyautogui.click(877,411, duration=1)
    pyautogui.hotkey('ctrl','v')

    #codigo_produto
    codigo_produto = linha[3].value
    pyperclip.copy(codigo_produto)
    pyautogui.click(874,497, duration=1)
    pyautogui.hotkey('ctrl','v')

    #peso
    peso = linha[4].value
    pyperclip.copy(peso)
    pyautogui.click(872,580, duration=1)
    pyautogui.hotkey('ctrl','v')

    #dimensoes
    dimensoes = linha[5].value
    pyperclip.copy(dimensoes)
    pyautogui.click(875,672, duration=1)
    pyautogui.hotkey('ctrl','v')

    #CLica no botão próximo
    pyautogui.click(873,717, duration=1)
    sleep(3)

    #Preço
    preco = linha[6].value
    pyperclip.copy(preco)
    pyautogui.click(898,219, duration=1)
    pyautogui.hotkey('ctrl','v')

    #quantidade_em_estoque
    quantidade_em_estoque = linha[7].value
    pyperclip.copy(quantidade_em_estoque)
    pyautogui.click(890,303, duration=1)
    pyautogui.hotkey('ctrl','v')

    #data_de_validade
    data_de_validade = linha[8].value
    pyperclip.copy(data_de_validade)
    pyautogui.click(904,394, duration=1)
    pyautogui.hotkey('ctrl','v')

    #cor
    cor = linha[9].value
    pyperclip.copy(cor)
    pyautogui.click(913,471, duration=1)
    pyautogui.hotkey('ctrl','v')

    #tamanho
    tamanho = linha[10].value
    pyautogui.click(894,566, duration=1)

    if tamanho == 'Pequeno':
        pyautogui.click(900,592, duration=1)
    elif tamanho == 'Médio':
        pyautogui.click(900,614, duration=1)
    else:
        pyautogui.click(900,636, duration=1)

    #material
    material = linha[11].value
    pyperclip.copy(material)
    pyautogui.click(903,647, duration=1)
    pyautogui.hotkey('ctrl','v')

    #Clica no botão próximo
    pyautogui.click(857,706, duration=1)
    sleep(3)

    #fabricante
    fabricante = linha[12].value
    pyperclip.copy(fabricante)
    pyautogui.click(908,238, duration=1)
    pyautogui.hotkey('ctrl','v')

    #Pais de origem
    pais_origem = linha[13].value
    pyperclip.copy(pais_origem)
    pyautogui.click(889,331, duration=1)
    pyautogui.hotkey('ctrl','v')

    #observações
    observacoes = linha[14].value
    pyperclip.copy(observacoes)
    pyautogui.click(880,416, duration=1)
    pyautogui.hotkey('ctrl','v')

    #codigo de barras
    codigo_de_barras = linha[15].value
    pyperclip.copy(codigo_de_barras)
    pyautogui.click(850,539, duration=1)
    pyautogui.hotkey('ctrl','v')

    #localização no armazem
    localizacao_armazem= linha[16].value
    pyperclip.copy( localizacao_armazem)
    pyautogui.click(864,638, duration=1)
    pyautogui.hotkey('ctrl','v')

#Botão Concluir
    pyautogui.click(868,684, duration=1)
#Botão OK
    pyautogui.click(1265,189, duration=1)
    sleep(2)
#Botão adicionar novo
    pyautogui.click(1076,454, duration=1)
    sleep(2)

#Após isso ele continua o processo de cadastro





"""Bot assistente que faça minhas funções de 
lançamento contábil de uma empresa, ele pega
a planilha do excel e faz o lançamento dos dados"""
