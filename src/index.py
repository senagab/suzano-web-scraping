from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from openpyxl import load_workbook
from selenium.webdriver.chrome.service import Service
import time

# Caminho do arquivo Excel
caminho_planilha = r"C:\Users\senag\Documents\0 - Projects\Banco\banco.xlsx"

# Carregar a planilha
wb = load_workbook(caminho_planilha)
sheet = wb.active

# Inicializar o WebDriver (Certifique-se de ter o ChromeDriver instalado)
# driver = webdriver.Chrome()
# driver = webdriver.Chrome(executable_path=r"C:\Users\senag\Documents\0 - Projects\Freelance\Suzano Web Scraping\chromedriver.exe")


service = Service(r"C:\Users\senag\Documents\0 - Projects\Freelance\Suzano Web Scraping\chromedriver.exe")
driver = webdriver.Chrome(service=service)

driver.get("https://contra-ten.vercel.app/")

time.sleep(3)  # Tempo para garantir que a página carregue

# Iterar sobre as linhas da planilha, pulando o cabeçalho
for row in sheet.iter_rows(min_row=2, values_only=True):
    nome, email, telefone, mensagem = row[:4]  # Pegando os 4 primeiros valores
    
    # Preencher os campos do formulário
    driver.find_element(By.ID, "nome").send_keys(nome)
    driver.find_element(By.ID, "email").send_keys(email)
    driver.find_element(By.ID, "telefone").send_keys(telefone)
    driver.find_element(By.ID, "mensagem").send_keys(mensagem)
    
    # Pressionar Enter para enviar o formulário (caso necessário)
    driver.find_element(By.ID, "mensagem").send_keys(Keys.RETURN)
    
    time.sleep(2)  # Pequeno delay entre envios

print("Formulários preenchidos com sucesso!")
input("Pressione Enter para fechar o navegador...")
driver.quit()
