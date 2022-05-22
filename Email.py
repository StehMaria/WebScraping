import win32com.client as win32 # importar biblioteca para criar integraçãao entre o Python e o Outlook

import os,sys #auxilia na validação de Pathway para diretorio

import re #auxilia na validação de email

from datetime import date #auxilia no trabalho de datas

import pythoncom



def diretorio():

    path = os.getcwd()
    file_path = path + "\Planilha.xlsx"
    return file_path

def validar_nome(nome):                     #define a função validar_nome
    
    if " " in nome:                         #Se o nome não possuir espaçamento, considera como inválido
        [p_nome,s_nome]= nome.split()       #divide a variável nome em p_nome (primeiro nome) e s_nome (sobrenome)
        p = len(p_nome)                     #variável p recebe a quantidade de caracteres do primeiro nome
        s = len(s_nome)                     #variável s recebe a quantidade de caracteres do sobrenome
        if (p > 1 and s > 1):               #se p e s possuirem mais que uma letra, valida o nome
            nome_certo = True               #variável recebe True por cumprir os requisitos, caso contrário recebe False
        else:
            nome_certo = False
    else:
        nome_certo = False
    
    return nome_certo                       #retorna nome_certo

def validar_email(destinatario):              #define a função validar_email
    
    email_certo = re.search(r'[a-zA-Z0-9_-]+@[a-zA-Z0-9]+.[a-zA-Z]{1,3}$', destinatario)
    print(str(email_certo))
    
    if email_certo:
        return True
    else:
        return False

def Enviar_Email(lista, val, nome, destinatario, anexo_endereco):# valor menor que 8
    
    outlook = win32.Dispatch('outlook.application',pythoncom.CoInitialize()) # Cria a integração com o Outlook
    email = outlook.CreateItem(0) #Gera um email
    email.To = destinatario # define o endereço de email de destino
    email.Subject = "Assunto: email automático Python" # define o assunto do email
    
    #print(lista[1])
    
    #Valores transmitidos para a mensagem de texto. O valor de cada variável virá da planilha de vulnerabilidade
    Software = lista
    CVE = lista
    Severity = lista
    Occurrence = lista
    Link = lista
    #print(Software)
    #print(CVE)
    #print(Severity)
    #print(Occurrence)
    #print(Link)

    # como enviar anexos

    anexo = anexo_endereco   #"C://Users/aluno/WebScraping/teste.xlsx"
    email.Attachments.Add(anexo)   #anexa ao email o arquivo solicitado
    
    
    #Data de atualização do registro de vulnerabilidade
    
    from datetime import datetime
    atualizacao = datetime.now()
    atualizacao = atualizacao.strftime("%d/%m/%Y %H:%M")    
    
    # Corpo do texto, template
    
    
    email.HTMLBody = f""" 
    
    
    
    <p>Prezado, {nome}. Segue atualização sobre vulnerabilidades acima com severidade acima de 8 pontos.</p>
    
    <h1> Vulnerabilidades Encontradas </h1>
    
    
    <!--
        HIERARQUIA DE TABELAS (simples)
        TABLE = tabela
            TABLE ROW = linha de tabela
                TABLE HEADER = cabeçalho de tabela
                TABLE DATA = dado de tabela
                
    -->
    <table border = "1" border-collapse = "0" class = "dataframe>
        <thead> <!-- Cabeçalho -->
        
            <tr style = "text-align: center"> 
                <th>Software</th>
                <th>CVE</th>
                <th>Severity</th>
                <th>NVD Published Date</th>
                <th>Link para o respectivo CVE</th>
            </tr>
        </thead>
        <tbody>    
        
        """
    valores = 0
    for i in range(val):
        email.HTMLBody = email.HTMLBody + f"""
            <tr style = "text-align: right"> <!-- 1º linha -->
                <td>{Software[6+(valores)]}</td>
                <td>{CVE[0+(valores)]}</td>
                <td>{Severity[2+(valores)]}</td>
                <td>{Occurrence[4+(valores)]}</td>
                <td>{Link[5+(valores)]}</td> 
            </tr>
            
        </tbody>
        """
        valores=valores+9
    email.HTMLBody = email.HTMLBody + f"""
    </table>
            
    <p>As vulnerabilidades foram atualizadas em: {atualizacao} </p>

    <p>atenciosamente, </p>


    """ 
    #Corpo do email. O comando HTMLBody ao invés de criar apenas um texto, gera um template de email utilizando linguagem HTML.
    #Com isso, permite várias configurações de páginas como acrescentar imagens, arquivos e outros itens
    # Para criar um parágrafo em HTML, a linha precisa iniciar como <p>, e finalizar com </p>
    # o comando (f""" texto de email""") torna dinâmico o texto do email, possibilitando acrescentar variáveis calculadas ao longo do código
    print(email.HTMLBody)
    email.display()
    #email.Send() #envia o e-mail   (viniciusbarbosabrasil@gmail.com)


