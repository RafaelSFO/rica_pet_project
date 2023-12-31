# %%
# Teste API
from selenium.webdriver.common.by import By
from webdriver_manager.microsoft import EdgeChromiumDriverManager
from selenium.webdriver.edge.service import Service as EdgeService
from selenium import webdriver
from bs4 import BeautifulSoup as soup
from datetime import datetime
import pandas as pd
import time
import requests
import pandas as pd
import credentials

print('Iniciando Robô')
inicio = datetime.now().strftime("%d-%m-%Y %H-%M-%S")
# %%%
## Encontrando as Separações
api_key = credentials.api_key_thapet
url_search = 'https://api.tiny.com.br/api2/separacao.pesquisa.php'
print('Leitura da quantidade de notas a serem movidas')
def coleta_separacoes(situacao):
    # Configurando os cabeçalhos da requisição com a chave de API
    #Fazendo a requisição GET
    separacoes_iniciais = pd.DataFrame()
    count = 1
    num_pag = 1000000
    try:
        while count <= num_pag:
            parametros = {
            'token': f'{api_key}',
            'situacao': situacao,
            'pagina': count
            }
            response = requests.get(url_search, params=parametros)
            data = response.json()
            num_pag = data['retorno']['numero_paginas']
            situacao_mercadorias = pd.DataFrame().from_dict(data['retorno']['separacoes'])
            separacoes_iniciais = pd.concat([separacoes_iniciais, situacao_mercadorias])
            count += 1
        return separacoes_iniciais
    except KeyError:
        print('Sem dados em separação')
        return separacoes_iniciais
try:
    mudar_status = coleta_separacoes(1)
    mudar_status.reset_index(drop=True, inplace=True)
    lista_separacoes = mudar_status['id'].astype('int').to_list()
    def muda_status(ids):
        list_dict_id = []
        alt_sep  ='https://api.tiny.com.br/api2/separacao.alterar.situacao.php'
        for id in ids:
            list_dict_id.append({'token': f'{api_key}', 'situacao': 2, 'idSeparacao': id})

        for param in list_dict_id:
            requests.get(alt_sep, params=param)
    print('Mudando o status das notas')
    muda_status(lista_separacoes)
except:
    print('Sem notas para mover')
print('Leitura das notas em separação')
separacoes_iniciais = coleta_separacoes(2)
separacoes_iniciais.reset_index(drop=True, inplace=True)

status_separacao = 'https://api.tiny.com.br/api2/separacao.obter.php'
def qtd_dos_clicks(id_separacao, i):
    try:
        status_param = {
                    'token': f'{api_key}',
                    'idSeparacao': id_separacao
                }
        response = requests.get(status_separacao, params=status_param)
        num_clicks = int(float(response.json()['retorno']['separacao']['itens'][i]['quantidade']))
        return num_clicks
    except:
        pass
#path_inicial = r'C:\Users\rafae\Documents\EMPREGO\RicaPet\RPA Thapet\Separações Iniciais\execucao {}.xlsx'.format(inicio)
path_inicial = r'C:\Users\snt\Documents\RPA\Repositório\rica_pet_project\Separações Iniciais\execucao_thapet {}.xlsx'.format(inicio)
separacoes_iniciais.to_excel(path_inicial, index=False)
# %%
login = credentials.login_thapet
senha = credentials.senha_thapet
# %%
# Abrindo uma instância do Google Chrome
print('Inicia a impressão')
def inicia_chrome():
    try:
        global browser
        #service = Service()
        #options = webdriver.ChromeOptions()
        browser = webdriver.Edge(service=EdgeService(EdgeChromiumDriverManager().install()))
        #browser = webdriver.Chrome(service=service, options=options)
        browser.implicitly_wait(30)
        browser.maximize_window()
        # Acessando o site
        browser.get('https://erp.tiny.com.br/login/')
        time.sleep(5)
        browser.find_element(By.NAME, 'username').send_keys(login) #coloca login
        browser.find_element(By.NAME, 'senha').send_keys(senha) #coloca senha
        time.sleep(3)
        browser.find_element(By.XPATH, '//*[@id="root-pagina-login"]/div/div/div/div[1]/div[1]/div[5]/button').click()  #clica para entrar
        time.sleep(5)
        web = soup(browser.page_source, 'html.parser')
        error_message = 'Este usuário já está logado em outra máquina' 
        if error_message in web.text: #verifica se o usuário está logado
            browser.find_element(By.XPATH, '//*[@id="bs-modal-ui-popup"]/div/div/div/div[3]/button[1]').click() #confirma o acesso na nova aba
            time.sleep(5)
            browser.get('https://erp.tiny.com.br/separacao') #link das separações
        else:
            pass
            time.sleep(3)
            browser.get('https://erp.tiny.com.br/separacao') #link das separações
        time.sleep(3)
        try:
            browser.find_element(By.XPATH, '//*[@id="page-wrapper"]/div[2]/div[1]/div[3]/ul/li[6]/a').click()
        except:
            pass
        time.sleep(2)
        browser.find_element(By.XPATH, '//*[@id="opc-sit-S"]').click() #clica em separadas
        browser.find_element(By.XPATH, '//*[@id="page-wrapper"]/div[2]/div[1]/div[1]/div/button[1]').click() #embala
        try: 
            element = browser.find_element(By.XPATH, "//input[@name='acao-checkout' and @value='V']")
            browser.execute_script("arguments[0].click();", element)
            ## https://morioh.com/a/48c7e73de145/perform-actions-using-javascript-in-python-selenium-webdriver
        except:
            browser.find_element(By.XPATH, '//*[@id="bs-modal"]/div/div/div/div[2]/div/div[1]').click()

        browser.find_element(By.XPATH, '//*[@id="bs-modal"]/div/div/div/div[3]/button[1]').click() # clica para avançar
    except:
        browser.close()
        print('erro, reiniciar o robô')
inicia_chrome()

for i in range(len(separacoes_iniciais['id'])):
    try:
        time.sleep(3)
        browser.find_element(By.ID, 'ui_popup_prompt_input').send_keys(separacoes_iniciais.loc[i, 'numero']) #adiciona número da nota
        time.sleep(2)
        browser.find_element(By.XPATH, '//*[@id="bs-modal-ui-popup"]/div/div/div/div[3]/button[1]').click() #clica para embalar
        time.sleep(5)
        valida_pedido = soup(browser.page_source, 'html.parser') #em casos de pedidos que já começaram a ser embalados
        if "embalar mesmo assim" in valida_pedido.text:
            browser.find_element(By.XPATH, '//*[@id="bs-modal-ui-popup"]/div/div/div/div[3]/button[1]').click()
        time.sleep(3)
        qtd_produtos = '//*[@id="checkout-lote-lista-separacoes"]/tr[2]/td/table/tbody/tr/td[5]/button[2]'
        lista_clicks = []
        if len(browser.find_elements(By.XPATH, qtd_produtos)) >= 2:
            lista_produtos = browser.find_elements(By.XPATH, qtd_produtos)
            for j in range(len(lista_produtos)):
                lista_clicks.append(qtd_dos_clicks(separacoes_iniciais.loc[i, 'id'], j))
                time.sleep(5)
                for k in range(lista_clicks[j]):
                    time.sleep(7)
                    lista_produtos[j].click()
        else:
            time.sleep(5)
            num_clicks = qtd_dos_clicks(separacoes_iniciais.loc[i, 'id'], 0)
            for i in range(num_clicks):
                browser.find_element(By.XPATH, qtd_produtos).click()
        time.sleep(6)

        try:
            browser.find_element(By.XPATH, '//*[@id="acoes-checkout-lote-individualmente"]/div/div/button[1]').click() #clica para adicionar mais um item (Continuar)
        except:
            browser.find_element(By.XPATH, '//*[@id="acoes-checkout-lote-individualmente"]/div/div/button[2]').click()
             # finaliza a lista de separado
        time.sleep(55)
    except:
        browser.close() # fecha o chrome em casos de erro
        time.sleep(10)
        inicia_chrome() # reinicia o processo e tenta novamente
try:
    browser.close() # se tudo deu certo, fecha o chrome
except:
    pass # se deu algum erro no meio da execução e não tiver mais chrome aberto, ele vai dar erro
# Validação

print('Inicia validação')
validacao_separacoes = separacoes_iniciais[['id', 'numero']]
validacao_separacoes['id'] = validacao_separacoes['id'].astype('str')
dict_list = []
for value in validacao_separacoes['id'].astype('int').to_list():
    dict_list.append({'token': f'{api_key}', 'idSeparacao': value})
#Fazendo a requisição GET
data_list = []
def requests_in_batch(api_batch):
    for entry in api_batch:
        try:
            response = requests.get(status_separacao, params=entry)
            my_data = response.json()['retorno']['separacao']
            data_list.append(my_data)
            time.sleep(5)
        except:
            print('Erro na chamada da API')
requests_in_batch(dict_list)

situacao_final = pd.DataFrame().from_dict(data_list)
situacao_final['id'] = situacao_final['id'].astype('str')

final_validacao = validacao_separacoes.merge(situacao_final[['id', 'situacao']], how='left', on='id')
final_validacao.loc[final_validacao['situacao'] == '3', 'Status'] = 'Finalizado com sucesso'
final_validacao.loc[final_validacao['situacao'] != '3', 'Status'] = 'Finalizado sem sucesso'
path = r'C:\Users\snt\Documents\RPA\Repositório\rica_pet_project\Execução Final\execucao_thapet {}.xlsx'.format(inicio)
final_validacao.to_excel(path, index=False)
print("Finalizado com Sucesso")