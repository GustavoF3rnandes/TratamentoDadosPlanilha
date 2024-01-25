import pandas as pd
import smtplib
import email.message
import os
from dotenv import load_dotenv

load_dotenv()
#Importar a base de dados
tabela_vendas = pd.read_excel("Vendas.xlsx")

#Visualizar a base de dados 
pd.set_option("display.max_columns", None)
print(tabela_vendas)
print('-' * 50)

#Faturamento por loja
faturamento = tabela_vendas[['ID Loja', 'Valor Final']].groupby('ID Loja').sum()
print(faturamento)
print('-' * 50)

#Quantidade de produtos vendidos por loja
quantidade  = tabela_vendas[['ID Loja', 'Quantidade']].groupby('ID Loja').sum()
print(quantidade)
print('-' * 50)

#Ticket medio por produto em cada loja
ticket_medio = (faturamento["Valor Final"] / quantidade["Quantidade"]).to_frame()
print(ticket_medio)

#Enviar um email com relatório no GMail
def enviar_email():
    corpo_email = f'''
                    <p>Prezados,</p>
                    
                    <p>Segue o Relatorio de vendas por cada loja.</p>
                    
                    <p>Faturamento:</p>
                    {faturamento.to_html()}
                    
                    <p>Quantidade Vendida:</p>
                    {quantidade.to_html()}
                    
                    <p>Ticket Medio dos Produtos em cada Loda</p>
                    {ticket_medio.to_html()}
                    
                    <p>Qualquer duvida, estou a disposição.</p>
                    
                    <p>att., </p>
                    <p>Gustavo Morais</p>
                    ''' 
    msg = email.message.Message()
    msg['Subject'] = "Relatorio de Vendas por Loja"
    msg['From'] = os.getenv("EMAIL")
    msg['To'] = os.getenv("EMAILTO")
    password = os.getenv("PASSWORD")
    msg.add_header('Content-Type', 'text/html')
    msg.set_payload(corpo_email)
    s = smtplib.SMTP('smtp.gmail.com: 587')
    s.starttls()
    s.login(msg["From"], password)
    s.sendmail(msg['From'], [msg['to']], msg.as_string().encode('utf-8'))
    
resposta = input("Deseja enviar o relatorio para o email? (S/N)").lower()
if resposta == "s":
    enviar_email()
    print("Email Enviado!")
    
elif resposta == "n":
    print("Finalizando Programa.")
    
else:
    print("Opção Invalida. Finalizando Programa.")
