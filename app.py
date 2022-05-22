from flask import Flask, render_template, request
import datetime
import WebScraping
from Planilha import planilha, organizar, valores
from Email import validar_nome, validar_email, diretorio, Enviar_Email
import Drive

app = Flask(__name__)

@app.route('/', methods=['GET','POST'])
def web():
    if request.method == 'POST':
        software = request.form['software']
        data_inicial = request.form['datainicial']
        data_final = request.form['datafinal']
        data_inicial = datetime.datetime.strptime(request.form['datainicial'], "%Y-%m-%d").date().strftime("%d%m%Y")
        data_final = datetime.datetime.strptime(request.form['datafinal'], "%Y-%m-%d").date().strftime("%d%m%Y")
        print (software)
        print (data_inicial)
        print (data_final)
        tabela = WebScraping.WebScraping(software,data_inicial,data_final)
        plan = planilha(tabela)
        organizar(plan)
        lista, qnt_v = valores(tabela)
        print(lista)
        print(qnt_v)
        nome_certo = False
        email_certo = False
        nome = request.form['nome']
        email = request.form['email']
        print(nome)
        print(email)
        if (nome != 'nenhum') and (email != 'nenhum@email'):
            nome_certo = validar_nome(nome)
            email_certo = validar_email(email)
            anexo_endereco = diretorio()
            Enviar_Email(lista,qnt_v,nome,email,anexo_endereco)
        else:
            print('Usuario não quer receber emails')
        
    return render_template('paginainicial.html')

@app.route('/planilha', methods=['GET','POST'])
def planilha_enviar():
    if request.method == 'POST':
        subject = request.form['subject']
        print(subject)
        if (subject == 'drive'):
            Drive.drive()
        else:
            print('Usuario não quer enviar')
    return render_template('planilha.html')

@app.route('/powerbi')
def power_bi():
    
    return render_template('powerbi.html')

if __name__ == '__main__':
    app.run()