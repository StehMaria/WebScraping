# 

import time
from selenium.webdriver.common.by import By
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.keys import Keys # Necessário para enviar caracteres especiais no comando .send_keys()
from webdriver_manager.chrome import ChromeDriverManager
from pprint import pprint

def WebScraping(software,data_inicial,data_final):

    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()))

    # Definir site
    Site = "https://nvd.nist.gov/vuln/search"

    # Abrir o site
    driver.get(Site)


    ### Funções
        
    def  getListFromElements(Caminho):    
        '''
        Retorna uma lista com o texto dos elementos encontrados no caminho(xpath) 
        
        '''
        Elemento = driver.find_elements(By.XPATH , Caminho)
        Lista = []
        for item in Elemento:
            Lista.append(item.text)
        return Lista

    def isLastPage():  
        '''
        Idendifica existência de página adicional e segue para próxima caso exista
        
        '''
        Navegador = driver.find_elements(By.XPATH, '//*[@id="row"]/div/div/nav/ul/li/a')
        Saida = True
        time.sleep(1) 
        for item in Navegador:
            if item.text == '>':
                print('Abrindo próxima página...')
                item.click()
                Saida = False
                break
            else:
                Saida = True
                
        return Saida   

    def itensSeveros(Severidades,SeveridadeCritica=8):
        itens = 0
        for severidade in Severidades:
            teste = (severidade >= SeveridadeCritica)
            if teste:
                itens += 1
        return itens
        

    def colorize(text,color):
        '''
        muda a cor do texto 'text' na cor 'color'
        '''
        text = color + text + colorReset
        return text


    print('Declarando varáveis')    
    colorReset = "\033[0;0m"
    colorRed   = "\033[1;31m"
    colorGrenn = "\033[0;32m"
    colorBlue  = "\033[1;34m"
    
    
    nomeUsuario = 'Teste'
    emailCadastrado = 'alerta_cve@teste.com'

    SeveridadeCritica = 8
    Severidades = [1,2,3,4,8,2.5,9,10]
    lstCVE = []
    lstCVSS_v20 = []
    lstCVSS_v31 = []
    lstDescr = []
    lstPubli = []
    sistema = []
    Links = []
    Referencias = []
    Configs = []

    print('Definindo xpaths')
    XpathCVE = '//*/tbody/tr/th/strong/a'
    XpathCVSS_v20 = '//*/tbody/tr/td[2]/span[2]'
    XpathCVSS_v31 = '//*/tbody/tr/td[2]/span[1]'
    XpathDescr = '//*/tbody/tr/td[1]/p'
    XpathPubli = '//*/table/tbody/tr/td[1]/span'
    xpathReferencia = '//*/table/tbody/tr/td/div/div[1]/div/div[1]/table/tbody/tr/td[1]/a'
    xpathConfig = '//*[@id="config-div-1"]/table/tbody/tr/td[1]/b'

    print('Ativando busca avançada')
    AdvavancedSearch = driver.find_element(value='SearchTypeAdvanced')
    AdvavancedSearch.click()

    print('Encontrando elementos na página')
    # Keyword Search //*[@id="Keywords"]
    Keyword = driver.find_element(value="Keywords")
    Keyword.clear()

    # Encontrar os elementos Data inicial e Data final
    StartDate = driver.find_element(value='published-start-date')
    StartDate.clear()

    EndDate = driver.find_element(value='published-end-date')
    EndDate.clear()

    print('Inserindo configurações de busca')
    Keyword.send_keys(software)
    StartDate.send_keys(data_inicial)
    EndDate.send_keys(data_final)

    print('Iniciando busca por vulnerabilidades em '+ colorize(software,colorGrenn))
    # Clicar no botão "Search"
    SubmitSearch = driver.find_element(value='vuln-search-submit')
    SubmitSearch.click()
    time.sleep(2)

    LastPage = False        
    while not(LastPage):
        lstCVE += getListFromElements(XpathCVE)
        lstCVSS_v20 += getListFromElements(XpathCVSS_v20)
        lstCVSS_v31 += getListFromElements(XpathCVSS_v31)
        lstDescr += getListFromElements(XpathDescr)
        lstPubli += getListFromElements(XpathPubli)
        LastPage = isLastPage()
        
    if len(lstCVE) == 0:
        print('Na busca por ' + colorize(software,colorBlue) + ', nenhuma vulnerabilidade foi encontrada')
    elif len(lstCVE) == 1:
        print(len('>>> Uma nova vulnerabilidade encontrada'))
    else:
        print('>>>', len(lstCVE), 'novas vulnerabilidades foram encontradas')

    
            
    for i in lstCVE:
        Links.append('https://nvd.nist.gov/vuln/detail/'+i)
        sistema.append(software)
        
    contador = len(Links)
    print('Buscando informações de referencia e configuração para cada vulnerabilidade...')
    for link in Links:
        driver.get(link)
        
        lstConfig = getListFromElements(xpathConfig)
        
        lstReferencia = getListFromElements(xpathReferencia)
        
        print(contador,'|',end=' ')
        Referencia = []    
        for r in range(len(lstReferencia)):
            Referencia.append(lstReferencia[r])
        Referencias.append(Referencia)
        
        Config = []    
        for l in range(len(lstConfig)):
            Config.append(lstConfig[l])
        Configs.append(Config)
        
        contador = contador - 1
        

    if len(lstCVE) > 0 and itensSeveros(Severidades, SeveridadeCritica):

        print("\n\n "+colorize("!!!ATENÇÃO!!!",colorRed) , itensSeveros(Severidades, SeveridadeCritica),  "vulnerabilidade(s) com severidade considerada crítica encontrada!\n\n")
        print('já estou enviando relatório para o e-mail: '+ colorize(emailCadastrado,colorBlue) + ' cadastrado\n')

    else:
        print("\nNenuma vulnerabilidade com severidade considerada crítica foi encontrada!\n") 
        
        
    print('Adicionando informações ao dicionário')
    Tabela={}
    Tabela.update({'CVE': lstCVE})
    Tabela.update({'CVSS_v20': lstCVSS_v20})
    Tabela.update({'CVSS_v31': lstCVSS_v31})
    Tabela.update({'Descrição': lstDescr})
    Tabela.update({'Publicação': lstPubli})    
    Tabela.update({'Link':Links})
    Tabela.update({'Sistema':sistema})
    Tabela.update({'Referencias':Referencias})
    Tabela.update({'Configuração':Configs})

    print("Aqui está, tudo o que encontrei até agora:\n")
    pprint(Tabela)
    print('Encerrando o navegador.')
    time.sleep(5)
    driver.quit()
    return Tabela