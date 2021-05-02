# Programa para envio do Texto Di�rio no Email 

from selenium import webdriver
from selenium.webdriver.common.keys import Keys

# Para abrir o navegador
nav = webdriver.Chrome()

# Para rodar em segundo plano
# from selenium.webdriver.chrome.options import Options
# chrome_options = Options()
# chrome_options.headless = True # also works
# nav = webdriver.Chrome(options=chrome_options)

# acessar jw
nav.get('https://wol.jw.org/pt/wol/h/r5/lp-t')

# pegar informações do texto diário
titulo_texto = nav.find_element_by_xpath('//*[@id="dailyText"]/div[2]/div[3]/header').text
texto_do_dia = nav.find_element_by_xpath('//*[@id="p3"]/em[1]').text
versiculo = nav.find_element_by_xpath('//*[@id="p3"]/a/em').text
texto_diario = nav.find_element_by_xpath('//*[@id="p4"]').text

texto = list((titulo_texto, texto_do_dia, versiculo, texto_diario))

# sair do navegador
nav.quit()

# envio do email
import win32com.client as win32 

# integração python com outlook
outlook = win32.Dispatch('outlook.application')

#criar email
email = outlook.CreateItem(0) # cria item no outlook

# configurar as informações do email
# email.To = email_destino.loc[0, 'email'] # destino

email.Subject = texto[0] # assunto
email.HTMLBody = f"""
<body style="background-image: linear-gradient(#000272, #341677, #a32f80, #ff6363);">
    <h2 style="color: #ff6464; font-size: 50px; text-align: center; font-family: Verdana, Geneva, Tahoma, sans-serif;"><strong>{texto[0]}</strong></h2>
    <br>
    <p style="color: #eb4eb7; font-family: Verdana, Geneva, Tahoma, sans-serif;"><em>{texto[1]}</em> - 
        <b style="color:#b377d6; font-family: Verdana, Geneva, Tahoma, sans-serif;">{texto[2]}</b></p>
    <hr>
    <p style="color: #51d2d6; font-family: Verdana, Geneva, Tahoma, sans-serif;">{texto[3]}</p>
</body>
""" # corpo do email'

# anexo no email
# anexo = "caminho do anexo"
# email.Attachments.Add(anexo)

# email destino
import pandas as pd

email_destino = pd.read_csv('email.csv') # 'email@destino.com'

conjunto_emails = list() # cria lista para juntar emails do arquivo csv 

# itera os itens no arquivo csv para colocar na lista
for index, row in email_destino.iterrows():
    conjunto_emails.append(row['email'])

# transformar a lista de emails em uma string
conjunto_emails_string = "; ".join(conjunto_emails)

# enviar emails
email.To = conjunto_emails_string
email.Send()
print(f'\n\n Emails enviados para {conjunto_emails_string}\n\n')

# for index, row in email_destino.iterrows():
#     email.To = row['email']
#     # Enviar email
#     email.Send()
#     print(f'Email Enviado para {row}')