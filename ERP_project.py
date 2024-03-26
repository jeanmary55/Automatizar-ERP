
import pyautogui
import time
import pyperclip
import pandas as pd

pyautogui.PAUSE = 0.6
pyautogui.FAILSAFE = True
pyautogui.press('win')
time.sleep(1)
pyautogui.write('Fakturama')
time.sleep(1)
pyautogui.press('enter')

time.sleep(32)

def encontra_imagem(imagem) :
    while not pyautogui.locateOnScreen(imagem, grayscale=True, confidence=0.8) :
        time.sleep(1)
    encontrou = pyautogui.locateOnScreen(imagem, grayscale=True, confidence=0.8)
    return encontrou

def direita(posicoes_imagem):
    seu_click =  posicoes_imagem[0] + posicoes_imagem[2], posicoes_imagem[1] + posicoes_imagem[3]/2
    return seu_click


tabela = pd.read_excel('Produtos.xlsx')



#clicar  no  new
for linha in tabela.index :
    id = tabela.loc[linha, 'ID']
    nome = tabela.loc[linha, 'Nome']
    categoria = tabela.loc[linha, 'Categoria']
    gtin = tabela.loc[linha, 'GTIN']
    suplier = tabela.loc[linha, 'Supplier']
    descricao = tabela.loc[linha, 'Descrição']
    preco = tabela.loc[linha, 'Preço']
    custo = tabela.loc[linha, 'Custo']
    estoque =tabela.loc[linha, 'Estoque']
    imagem = tabela.loc[linha, 'Imagem']



    encontrou = encontra_imagem('new.png')
    pyautogui.click(pyautogui.center(encontrou))

    # clicar  no  new product
    encontrou = encontra_imagem('new_product.png')
    pyautogui.click(pyautogui.center(encontrou))
    time.sleep(2)

    # preencher  todos os campos
    encontrou =encontra_imagem('item.png')
    pyautogui.click(direita(encontrou))
    pyautogui.write(str(id))

    encontrou =encontra_imagem('name.png')
    pyautogui.click(direita(encontrou))
    pyautogui.write(str(nome))


    encontrou =encontra_imagem('categori.png')
    pyautogui.click(direita(encontrou))
    pyautogui.write(str(categoria))

    encontrou =encontra_imagem('gtin.png')
    pyautogui.click(direita(encontrou))
    pyautogui.write(str(gtin))

    encontrou =encontra_imagem('code.png')
    pyautogui.click(direita(encontrou))
    pyautogui.write(str(suplier))

    encontrou =encontra_imagem('descricao.png')
    pyautogui.click(direita(encontrou))
    pyautogui.write(str(descricao))
    
    encontrou =encontra_imagem('price.png')
    pyautogui.click(direita(encontrou))

    preco_texto = f'{preco:.2f}'.replace('.', ",")
    pyautogui.write(str(preco_texto))

    encontrou =encontra_imagem('price2.png')
    pyautogui.click(direita(encontrou))
    preco_texto = f'{custo:.2f}'.replace('.', ",")
    pyautogui.write(str(custo))


    encontrou =encontra_imagem('select.png')
    pyautogui.click(pyautogui.center(encontrou))
    pyautogui.write(rf'C:\Users\usuario\OneDrive\Documents\RPA_ERP\Imagens Produtos\{str(imagem)}')
    pyautogui.press('enter')
    time.sleep(2)


    encontrou =encontra_imagem('save.png')
    pyautogui.click(pyautogui.center(encontrou))
    time.sleep(3)











# esperar para carregar  a  imagem


# clicar no  new

