from selenium import webdriver # Biblioteca para automação de navegadores
import time # Para criar delays no código
from selenium.webdriver.common.by import By # Para localizar elementos na página
from selenium.webdriver.common.keys import Keys # Para enviar teclas para os elementos
from selenium.webdriver.support.ui import WebDriverWait # Para esperar elementos carregarem na página
from selenium.webdriver.support import expected_conditions as EC # Para condições esperadas na página
import win32com.client # Para manipulação do Excel e envio de e-mails via Outlook

# Nesse código, utilizei o geckodriver, que é o driver do Firefox, para automatizar o download, você pode utilizar outros drivers para utilizar outros navegadores, como o ChromeDriver para o Google Chrome.
driver = webdriver.Firefox()
# URL do site do formulário
url = ""
driver.get(url) # Abre o site do formulário
time.sleep(5)

# Ao entrar no site, precisamos fazer o login na conta microsoft, para isso, vamos localizar os campos de e-mail e senha e enviar as informações necessárias.
# Como ao entrar no site o campo de email já está selecionado, vamos enviar o e-mail diretamente.
driver.switch_to.active_element.send_keys("") # Envia o e-mail da conta microsoft
driver.switch_to.active_element.send_keys(Keys.ENTER)
time.sleep(2)
driver.switch_to.active_element.send_keys("") # Envia a senha da conta microsoft
driver.switch_to.active_element.send_keys(Keys.ENTER)
time.sleep(1)
driver.switch_to.active_element.send_keys(Keys.ENTER) # Pressiona Enter para entrar na conta
time.sleep(4)
dropdown_button = driver.find_element(By.ID, "ExcelDropdownMenu") # Procura o botão de mais opções no site do forms
dropdown_button.click() # Clica no botão de mais opções
time.sleep(1)
download_button = driver.find_element(By.XPATH, "//span[text()='Download a copy']") # Procura o botão de download do Excel
download_button.click() # Clica no botão de download do Excel
time.sleep(1)
driver.quit()# Fechar o navegador
print("Download concluído.")

excel = win32com.client.Dispatch("Excel.Application")
excel.visible = True  # Deixa o Excel visível ao usuário
wb = excel.Workbooks.Open(r"") # Caminho para o arquivo base onde serão inseridos os dados do download, essa incrementação é feita pelo VBA
excel.Application.Run("Planilha7.organizar") # "Planilha7" é o nome da planilha que contém a macro, "organizar" é o nome dessa macro
# O código continua pelo excel, onde há um botão que executa o código python "EnvioSemanalPgestores.py"