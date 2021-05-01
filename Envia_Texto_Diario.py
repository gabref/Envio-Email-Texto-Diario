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
# nav.find_element_by_xpath('//*[@id="dailyText"]/div[2]/div[3]/header')

# sair do navegador
nav.quit()