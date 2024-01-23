    #Programa de automação
    #Para que o programa funcione, basta certificar de que a aba do chrome esteja em tela cheia ou apenas que não esteja minimizado.
    #É necessária a extensão "excel viewer" para abrir o doc "produtos_ficticios.xlsx"

#Entrar na planilia
import openpyxl as op
import pyperclip as pc
import pyautogui as py
import time

workbook = op.load_workbook('produtos_ficticios.xlsx')
sheet_produtos = workbook['Produtos']

# Abrir a página/executar app de catálogo de produtos

py.PAUSE = 1

py.press("win")
py.write("chrome")
py.press("Enter")
link = "https://cadastro-produtos-devaprender.netlify.app/"
py.write(link)
py.press("enter")

time.sleep(4)

# Copiar as informações do xl e colar na área correspondente

for linha in sheet_produtos.iter_rows(min_row=2):
    nome_produto = linha[0].value
    pc.copy(nome_produto)
    py.press("tab")
    #pode usar o "py.click(x=x y=y)" indicando onde clicar com as coordenadas fornecidas pelo doc "auxiliar.py"
    py.hotkey("ctrl", "v")

    descricao = linha[1].value
    pc.copy(descricao)
    py.press("tab")
    py.hotkey("ctrl", "v")

    categoria = linha[2].value
    pc.copy(categoria)
    py.press("tab")
    py.hotkey("ctrl", "v")

    codigo_produto = linha[3].value
    pc.copy(codigo_produto)
    py.press("tab")
    py.hotkey("ctrl", "v")
    
    peso = linha[4].value
    pc.copy(peso)
    py.press("tab")
    py.hotkey("ctrl", "v")

    dimensoes = linha[5].value
    pc.copy(dimensoes)
    py.press("tab")
    py.hotkey("ctrl", "v")

    # Mudança para a pg 2

    py.press("tab")
    py.press("enter")
    time.sleep(3)

    preco = linha[6].value
    pc.copy(preco)
    py.press("tab")
    py.hotkey("ctrl", "v")

    quantidade_estoque = linha[7].value
    pc.copy(quantidade_estoque)
    py.press("tab")
    py.hotkey("ctrl", "v")

    validade = linha[8].value
    pc.copy(validade)
    py.press("tab")
    py.hotkey("ctrl", "v")

    cor = linha[9].value
    pc.copy(cor)
    py.press("tab")
    py.hotkey("ctrl", "v")

    tamanho = linha[10].value
    pc.copy(tamanho)
    py.press("tab")
    py.press("enter")

    # Condicional direcionada á caixa que indica se o produto é pequeno, médio ou grande
    if tamanho == 'Pequeno':
        py.press("enter")
    elif tamanho == "médio":
        py.hotkey("down", "enter")
    else:
        py.hotkey("down", "down", "enter")

    material = linha[11].value
    pc.copy(material)
    py.press("tab")
    py.hotkey("ctrl", "v")

    py.press("tab")
    py.press("enter")
    time.sleep(3)

    fabricante = linha[12].value
    pc.copy(fabricante)
    py.press("tab")
    py.hotkey("ctrl", "v")

    pais_origem = linha[13].value
    pc.copy(pais_origem)
    py.press("tab")
    py.hotkey("ctrl", "v")

    obs = linha[14].value
    pc.copy(obs)
    py.press("tab")
    py.hotkey("ctrl", "v")

    codigo_barras = linha[15].value
    pc.copy(codigo_barras)
    py.press("tab")
    py.hotkey("ctrl", "v")

    local_armazem = linha[16].value
    pc.copy(local_armazem)
    py.press("tab")
    py.hotkey("ctrl", "v")

    # Fim das informações de um produto e retorno para a pg inicial para registrar o proximo produto

    py.press("tab")
    py.press("enter")
    py.press("enter")
    time.sleep(0.5)
    py.press("tab")
    py.press("enter")
    time.sleep(3)