from selenium import webdriver
from bs4 import BeautifulSoup
from selenium.webdriver.chrome.options import Options
from time import sleep
import pandas as pd
from selenium.webdriver.common.keys import Keys
import os
import PySimpleGUI as sg
import openpyxl
#Opções do navegador
opcoes = Options()
opcoes.add_argument('window-size=1280,720')

#Criando o navegador

def IniciandoKabum():
    driver.get('https://www.kabum.com.br/')
    sleep(2)

    input_place = driver.find_element_by_id('input-busca')
    input_place.send_keys(produtodesejado)
    input_place.submit()
    sleep(2)



#Pegando o html da página
listaprodutos = []

def PegarProdutosKabum():
    f = True
    site = BeautifulSoup(driver.page_source,'html.parser')
    produtos = site.findAll('div',attrs={'class':'sc-GEbAx odTaK productCard'})
    for produto in produtos:
        titulo = produto.find('h2',attrs={'class':'sc-kHOZwM brabbc sc-fHeRUh jwXwUJ nameCard'})
        preco = produto.find('span', attrs={'class':'sc-iNGGcK fTkZBN priceCard'})
        link = produto.find('a',attrs={'class':'sc-eGRUor bTTpdH'})
        if preco:
            listaprodutos.append([titulo.text,preco.text,("https://www.kabum.com.br" + link['href'])])
        else:
            f = False

    return f

    sleep(1)


def VarrerSiteKabum():
    IniciandoKabum()
    f = True
    while f:
        f = PegarProdutosKabum()
        nextpage = driver.find_element_by_class_name('nextLink')
        try:
            nextpage.click()
        except:
            break
        else:
            sleep(1)


def Tela():
    sg.theme('Reddit')
    layout = [
        [sg.Text('Insira o produto: '),sg.Input('',key='produto')],
        [sg.Button('Procurar',key='go')]
    ]
    return sg.Window('BotDeProdutos',layout=layout,element_justification='center',finalize=True)

def ChamarTela():
    janela = Tela()
    while True:
        event, values = janela.read()
        if event == sg.WINDOW_CLOSED:
            break
        if event == 'go':
            global produtodesejado
            produtodesejado = values['produto']
            break



def Tela2():
    sg.theme('Reddit')
    linha = [
        [sg.Text('Deseja Filtrar item pelo preço? '),sg.Radio('Sim','escolha',key='sim'),sg.Radio('Não','escolha',key='nao')],
        [sg.Button('Enviar',key='enviar')]
    ]

    layout = [
        [sg.Frame('',layout=linha,key='desejo')]
    ]

    return sg.Window('O que deseja fazer?',layout=layout,finalize=True,element_justification='center')

def ChamarTela2():
    janela = Tela2()
    while True:
        event , values = janela.read()
        if event == sg.WINDOW_CLOSED:
            break
        elif event == 'enviar' and values['nao'] == True:
            break
        elif event == 'enviar' and values['sim'] == True:
            janela.extend_layout(janela['desejo'],[
                [sg.Text('Digite o preço máximo: '),sg.Input('',key='precodesejado')],
                [sg.Button('Filtrar',key='filtrar')]
            ])
        if event == 'filtrar':
            global precodesejado
            precodesejado = values['precodesejado']
            planilha = openpyxl.load_workbook(f'Planilha{produtodesejado}\{produtodesejado}.xlsx')

            pagina = planilha['Sheet1']

            listacerta = []
            contador = 0

            for rows in pagina.iter_rows(min_col=1, max_col=3, min_row=2):
                listacerta.append([rows[contador].value, rows[contador + 1].value, rows[contador + 2].value])

            lista = []
            for i in range(0, len(listacerta)):
                if listacerta[i][1] == 'Indisponível':
                    listacerta[i][1] = '0'
                    valor = float(
                        listacerta[i][1].replace('R$', '').replace('\xa0', '').replace('.', '').replace(',', '.'))
                else:
                    valor = float(
                        listacerta[i][1].replace('R$', '').replace('\xa0', '').replace('.', '').replace(',', '.'))
                    if valor <= float(precodesejado):
                        valor = str(valor)
                        listacerta[i][1] = "R$" + valor
                        lista.append(listacerta[i])
            v2 = pd.DataFrame(lista, columns=['Nome', 'Preço', 'Link'])

            if not os.path.isdir(f'Planilha{produtodesejado}'):
                os.mkdir(f'Planilha{produtodesejado}')
            v2.to_excel(f'Planilha{produtodesejado}\{produtodesejado}até{precodesejado}.xlsx', index=False)
            break


ChamarTela()
driver = webdriver.Chrome(options=opcoes)
VarrerSiteKabum()
driver.close()


if not os.path.isdir(f'Planilha{produtodesejado}'):
    os.mkdir(f'Planilha{produtodesejado}')
v = pd.DataFrame(listaprodutos,columns=["Nome","Preço","Link"])
v.to_excel(f'Planilha{produtodesejado}\{produtodesejado}.xlsx',index=False)

ChamarTela2()


