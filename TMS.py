from selenium import webdriver
from selenium.webdriver.common.by import By
import time
import csv
import pandas as pd
from datetime import datetime
from bs4 import BeautifulSoup
import os
import smtplib
from email.message import EmailMessage


nome_csv = "D:\\Soft\\TMS_V8\\maquina.csv"
nome_txt = "D:\\Soft\\TMS_V8\\maquina.txt"
caminho_email_sucesso = "D:\\Soft\\TMS_V8\\Emails.txt"
caminho_email_erro = "D:\\Soft\\TMS_V8\\Emails_CPD.txt"

def enviar_email(corpo, tipo):#Função para enviar E-mail

    enderco_email = 'teste@gmail.com'
    senha = '123456#'

    #Seleciona o que vai ser escrito no assunto
    if tipo == 1: 
        with open(caminho_email_sucesso, "r") as emails:
            email_enviar = emails.readlines()
        assunto = "Exportado com Sucesso"
    
    else:
        #Caso de erro na exportação
        with open(caminho_email_erro, "r") as emails:
            email_enviar = emails.readlines()
        assunto = "Erro na hora exportar dados do TMS"

    data = datetime.now().strftime("%d/%m/%Y, %H:%M:%S")

    msg = EmailMessage()
    msg['Subject'] = f"{assunto} {data}" #Assunto do E-mail
    msg["From"] = "teste@gmail.com" #Correspondente
    msg["To"] = email_enviar #Pra quem vai enviar
    msg.set_content(f"{txt}") #Corpo do E-mail

    with smtplib.SMTP_SSL('serverTeste.net', 156) as smtp: #Abre o servidor executa as funções abaixo e fecha a conexão
        smtp.login(enderco_email, senha)#Faz login no Servidor
        smtp.send_message(msg)#Envia o E-mail

try:
    #Abrir o Navegador
    navegador = webdriver.Edge()
    #Caminho do TMS
    navegador.get("**")
    #navegador.get("**") Para testes
    time.sleep(1)
    #Coleta dados das maquinas
    navegador.find_element_by_xpath('/html/body/center/table[3]/tbody/tr[2]/td[2]/table/tbody/tr[1]/td/a/b').click()
    time.sleep(1)
    #Selciona todas as maquinas
    navegador.find_element_by_xpath('/html/body/center/form/table/tbody/tr[2]/td/input').click()
    time.sleep(1)
    #Enter
    navegador.find_element_by_xpath('/html/body/center/form/table/tbody/tr[3]/td/input').click()
    time.sleep(20) #Tempo de espera para coletar os dados da máquina
    #Seleciona o elemento tabela da pagina
    tabela = navegador.find_element_by_xpath('/html/body/center/table')
    #Pega os dados do HTML
    htmlConect = tabela.get_attribute("outerHTML")
    #Gera o HTML
    suop = BeautifulSoup(htmlConect, "html.parser")
    #Selciona o elemento "table" do HTML gerado
    maquinas = suop.find(name="table")
    #Lê o HTML
    df = pd.read_html(str(maquinas))[1]
    #Grava as informações em arquivo CSV
    df.to_csv(nome_csv, encoding="UTF-8", sep=";", index=False)

    #Lê o arquivo CSV
    with open(nome_csv, newline='') as csvfile:#Abre o arquivo CSV executa as funções abaixo e fecha o arquivo
        reader = csv.DictReader(csvfile)
        for row in reader:
            #Grava as linhas no txt
            #É gravado no txt para melhor vizualização no email
            with open(nome_txt, "a") as txt:
                txt.write(row['0;1'] + "\n")

    #Abre o arquivo txt
    with open(nome_txt, "r") as txt: #Abre o arquivo TXT executa as funções abaixo e fecha o arquivo
        txt = txt.read()

    #Voltar ao menu
    navegador.find_element_by_xpath('/html/body/center/nobr[1]/a/img').click()
    time.sleep(1)
    #Exportar para arquivos CSV
    navegador.find_element_by_xpath('/html/body/center/table[5]/tbody/tr[2]/td[3]/table/tbody/tr/td/a/b').click()
    time.sleep(1)
    #Pega o mês atual
    mes = datetime.today().strftime("%Y/%m")
    #Selecionar o mês
    dia = navegador.find_element(By.XPATH, f"//*[text()= '{mes}']")
    dia.click()
    time.sleep(1)
    #Gerar o Relatório
    navegador.find_element_by_xpath('/html/body/center/form/table[2]/tbody/tr[3]/td/input[2]').click()
    #Fecha o Navegador
    time.sleep(1)

    #Envia E-mail quando tudo ocorre tudo certo com a relação das maquinas
    enviar_email(txt, 1)

    #Remove os arquivos
    os.remove(nome_csv)
    os.remove(nome_txt)
    
    #Fecha o navegador e o Prompt de Comando
    navegador.quit()

except:
    #Mensagem que é enviada
    mensagem = "Erro na hora de exporta os Dados do TMS"

    #Envia E-mail quando tudo ocorre um erro 
    enviar_email(mensagem, 0)

    #Remove os arquivos
    os.remove(nome_csv)
    os.remove(nome_txt)

    #Fecha o navegador e o Prompt de Comando
    navegador.quit()