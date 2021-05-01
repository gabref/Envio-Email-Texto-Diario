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

# envio do email
import win32com.client as win32 

# integração python com outlook
outlook = win32.Dispatch('outlook.application')

#criar email
email = outlook.CreateItem(0) # cria item no outlook

# email destino
email_destino = 'email@destino.com'

# configurar as informações do email
email.To = email_destino # destino
email.Subject = texto['titulo_texto'] # assunto
email.HTMLBody = f"""
<h2>{texto[0]}</h2>
<br>
<p><em>{texto[1]}</em> - <b>{texto[2]}</b></p>
<hr>
<p>{texto[3]}</p>
""" # corpo do email

# anexo no email
# anexo = "caminho do anexo"
# email.Attachments.Add(anexo)

# Enviar email
email.Send()
print('Email Enviado')

# sair do navegador
# nav.quit()