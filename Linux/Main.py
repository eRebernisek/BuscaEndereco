import os
import time
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
import xlsxwriter
import pandas as pd


# Configura o Selenium no ambiente Linux para utilizar o Firefox
firefox = webdriver.Firefox(executable_path='./geckodriver')
linhaExcel = 0


def lerExcel():
    global linhaExcel

    pesquisa = pd.read_excel('Pesquisa.xlsx')
    shape = pesquisa.shape

    resultados = list()

    for i in range(0, shape[0]):
        linhaExcel = i
        endereco = pesquisa.iloc[i, 0]
        cep = str(pesquisa.iloc[i, 1])
        tipo = str(pesquisa.iloc[i, 2])

        result = {}

        validaNome = validaItem(endereco, 'Nome')
        validaCep = validaItem(cep, 'CEP')

        if tipo == 'nan':
            if validaNome[0]:
                pesquisaCorreios(endereco)
                pass
            elif validaCep[0]:
                pesquisaCorreios(cep)
                pass
            else:
                geraExcel(validaCep[1])
                pass
        else:
            if tipo.lower() == 'cep':
                if validaCep[0]:
                    pesquisaCorreios(cep)
                    pass
                else:
                    geraExcel(validaCep[1])
                    pass
            elif tipo.lower() == 'nome':
                if validaNome[0]:
                    pesquisaCorreios(endereco)
                    pass
                else:
                    geraExcel(validaNome[1])
                    pass
            else:
                geraExcel({'Nome': 'Campo Busca Invalido'})
                pass


def pesquisaCorreios(busca):
    # Abre o site da Amazon no Firefox
    firefox.get('https://buscacepinter.correios.com.br/app/endereco/index.php')
    # firefox.maximize_window()
    time.sleep(1)
    # Realiza busca
    firefox.find_element_by_id("endereco").send_keys(str(busca))
    firefox.find_element_by_id("endereco").send_keys(Keys.RETURN)

    # Aguarda aparecer a mensagem de retorno da busca
    loading = True
    while loading:
        mensagem = firefox.find_element_by_id("mensagem-resultado").text
        loading = len(mensagem) < 1

    lerCorreios()


def lerCorreios():
    listaResult = firefox.find_element_by_id("resultado-DNEC")
    listBody = listaResult.find_element_by_tag_name('tbody')
    itens = listBody.find_elements_by_tag_name('tr')

    resultados = []

    for item in itens:
        result = item.find_elements_by_tag_name('td')
        nome = result[0].text
        bairro = result[1].text
        uf = result[2].text
        cep = result[3].text

        resultado = {'Nome': nome, 'Bairro': bairro, 'UF': uf, 'CEP': cep}
        resultados.append(resultado)

    geraExcel(resultados)


def geraExcel(resultados):
    global linhaExcel

    msg = len(resultados)

    indexExcel = []

    for i in range(1, len(resultados)+1):
        indexExcel.append(i)

    # print('Resultados:'+str(msg))
    # print(indexExcel)

    df1 = pd.DataFrame(resultados, index=indexExcel, columns=[
                       'Nome', 'Bairro', 'UF', 'CEP'])
    df1.to_excel('RetornoLinha'+str(linhaExcel+1)+'.xlsx', index=False)


def validaItem(item, tipo):
    if pd.isna(item) or item == 'nan':
        return [False, {'Nome': 'Campo '+tipo+' não informado'}]
    else:
        if tipo.lower() == 'cep':
            try:
                item = item.replace('-', '')

                if len(item) <= 8:
                    cep = int(item)
                    return [True, '']
                else:
                    return [False, {'Nome': 'Campo CEP maior do que permitido'}]
            except Exception as e:
                return [False, {'Nome': 'Campo CEP inválido'}]
        else:
            return [True, '']


if __name__ == '__main__':
    os.system('clear')
    try:
        lerExcel()
    except Exception as e:
        # os.system('clear')
        print(str(e))
    firefox.close()
    quit()
