from selenium.common.exceptions import *
from selenium.webdriver.support import expected_conditions as CondicaoExperada
from selenium.webdriver.support.ui import WebDriverWait
from selenium import webdriver
from selenium.webdriver.chrome.service import Service as ChromeService
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from time import sleep
import openpyxl
import smtplib
from email.message import EmailMessage

def iniciar_driver():
    chrome_options = Options()

    arguments = ['--lang=en-US', '--window-size=1920,1080',
                 '--incognito', '--disable-gpu', '--no-sandbox', '--headless', '--disable-dev-shm-usage']

    for argument in arguments:
        chrome_options.add_argument(argument)
    chrome_options.headless = True
    chrome_options.add_experimental_option('prefs', {
        'download.prompt_for_download': False,
        'profile.default_content_setting_values.notifications': 2,
        'profile.default_content_setting_values.automatic_downloads': 1

    })
    driver = webdriver.Chrome(service=ChromeService(
        ChromeDriverManager().install()), options=chrome_options)

    wait = WebDriverWait(
        driver,
        10,
        poll_frequency=1,
        ignored_exceptions=[
            NoSuchElementException,
            ElementNotVisibleException,
            ElementNotSelectableException,
        ]
    )
    return driver, wait


driver, wait = iniciar_driver()

# Configuração da planilha
workbook = openpyxl.Workbook()
del workbook['Sheet']
workbook.create_sheet('Celulares')
sheet_atual = workbook['Celulares']
sheet_atual.append(['Descrição','Valor'])    

# Configuração do email
## Configurações de login
EMAIL_ADDRESS = '????????'
EMAIL_PASSWORD = '????????'

## Criar um email
mail = EmailMessage()
mail['Subject'] = 'Seu relatório de preços'
mensagem = '''
Baixe seu relatório de preços agora!
'''
mail['From'] = EMAIL_ADDRESS
mail.add_header('Content-Type', 'text/html')
mail.set_payload(mensagem.encode('utf-8'))


# Início
driver.get('https://telefonesimportados.netlify.app')
emailTo = input('Digite o email para o qual o relatório deve ser enviado: ')
print(f'O relatório será enviado para o email: {emailTo} ...')
pagina = 2
while True:
    sleep(1)

    # Encontrar título dos produtos
    produtos = wait.until(CondicaoExperada.visibility_of_all_elements_located(
        (By.XPATH, "//div[@class='single-shop-product']/h2/a")))

    # Encontrar preços dos produtos
    precos = wait.until(CondicaoExperada.visibility_of_all_elements_located(
        (By.XPATH, "//div[@class='product-carousel-price']//ins")))
    sleep(1)

    # Gravar em planilha    
    print(f'################### Página: {pagina - 1} ###################')
    for produto, preco in zip(produtos, precos):
        print(produto.text, preco.text)
        sheet_atual.append([produto.text, preco.text])    

    # Buscar próxima página
    try:        
        botao_proxima_pagina = driver.find_element(By.LINK_TEXT, str(pagina))        
        pagina = pagina +1
        botao_proxima_pagina.click()
        print('Indo para a próxima página')
        sleep(1)        
    except:
        print('Chegamos a última página')
        workbook.save('Produtos.xlsx')
        break

try:
    print('Enviando e-mail')    
    sleep(1)
    mail['To'] = emailTo
    mail.add_attachment(maintype='application',subtype='octet-stream', filename='Produtos.xlsx')
    with smtplib.SMTP_SSL('smtp.gmail.com', 465) as email:
        email.login(EMAIL_ADDRESS, EMAIL_PASSWORD)
        email.send_message(mail)
    print('E-mail enviado com sucesso...')
    driver.get('https://telefonesimportados.netlify.app')
except:
    print('Erro ao enviar o e-mail')



