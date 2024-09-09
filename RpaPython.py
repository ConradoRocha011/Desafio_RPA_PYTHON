# Importação das bibliotecas necessárias
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import pandas as pd
import time
import os

# Definindo variáveis
site = 'http://www.rpachallenge.com/'
file = 'challenge (1).xlsx'
log_file = r"C:\Users\Conrado\Documents\log.txt"
screenshot_file = r"C:\Users\Conrado\Documents\screenshot_final.png"
dt_registros = pd.read_excel(file)

# Função para escrever no arquivo de log
def escrever_log(mensagem):
    with open(log_file, 'a') as f:
        f.write(mensagem + '\n')

# Criação ou limpeza do arquivo de log
with open(log_file, 'w') as f:
    f.write("Log de execução\n")
    f.write("=================\n")

# Exibindo a tabela de registros no console
print("Tabela de registros:")
print(dt_registros)

# Abrindo o navegador Google Chrome na URL do RPA Challenge
escrever_log("Abrindo navegador Google Chrome na URL do RPA Challenge")
print("Abrindo navegador Google Chrome na URL do RPA Challenge")
service = Service(r"C:\Users\Conrado\Desafio_RPA_PYTHON\chromedriver.exe")
driver = webdriver.Chrome(service=service)
driver.get(site)

escrever_log("Iniciando preenchimento dos formulários")
print("Iniciando preenchimento dos formulários")

# Usando espera explícita para aguardar que o botão "Start" esteja clicável
wait = WebDriverWait(driver, 10)
btn_start = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[@class='waves-effect col s12 m12 l12 btn-large uiColorButton']")))
btn_start.click()

# Processando cada registro da tabela e atualizando o log e o arquivo Excel com o status
for i, r in dt_registros.iterrows():
    try:
        # Salvando dados em variáveis
        first_name = r['First Name']
        last_name = r['Last Name']
        company_name = r['Company Name']
        role_in_company = r['Role in Company']
        address = r['Address']
        email = r['Email']
        phone_number = r['Phone Number']

        escrever_log(f"Rodada {i}: Iniciando lançamento dos dados para {first_name} {last_name}")
        print(f'Rodada:', i)

        # Preenchendo o campo "First Name"
        textbox = wait.until(EC.visibility_of_element_located((By.XPATH, "//input[@ng-reflect-name='labelFirstName']")))
        textbox.clear()
        textbox.send_keys(first_name)
        escrever_log(f"Rodada {i}: Nome '{first_name}' lançado no campo 'First Name'")

        # Preenchendo o campo "Last Name"
        textbox = wait.until(EC.visibility_of_element_located((By.XPATH, "//input[@ng-reflect-name='labelLastName']")))
        textbox.clear()
        textbox.send_keys(last_name)
        escrever_log(f"Rodada {i}: Sobrenome '{last_name}' lançado no campo 'Last Name'")

        # Preenchendo o campo "Company Name"
        textbox = wait.until(EC.visibility_of_element_located((By.XPATH, "//input[@ng-reflect-name='labelCompanyName']")))
        textbox.clear()
        textbox.send_keys(company_name)
        escrever_log(f"Rodada {i}: Empresa '{company_name}' lançada no campo 'Company Name'")

        # Preenchendo o campo "Role in Company"
        textbox = wait.until(EC.visibility_of_element_located((By.XPATH, "//input[@ng-reflect-name='labelRole']")))
        textbox.clear()
        textbox.send_keys(role_in_company)
        escrever_log(f"Rodada {i}: Cargo '{role_in_company}' lançado no campo 'Role in Company'")

        # Preenchendo o campo "Address"
        textbox = wait.until(EC.visibility_of_element_located((By.XPATH, "//input[@ng-reflect-name='labelAddress']")))
        textbox.clear()
        textbox.send_keys(address)
        escrever_log(f"Rodada {i}: Endereço '{address}' lançado no campo 'Address'")

        # Preenchendo o campo "Email"
        textbox = wait.until(EC.visibility_of_element_located((By.XPATH, "//input[@ng-reflect-name='labelEmail']")))
        textbox.clear()
        textbox.send_keys(email)
        escrever_log(f"Rodada {i}: Email '{email}' lançado no campo 'Email'")

        # Preenchendo o campo "Phone Number"
        textbox = wait.until(EC.visibility_of_element_located((By.XPATH, "//input[@ng-reflect-name='labelPhone']")))
        textbox.clear()
        textbox.send_keys(phone_number)
        escrever_log(f"Rodada {i}: Telefone '{phone_number}' lançado no campo 'Phone Number'")

        # Clicando no botão "Submit"
        botao_submit = wait.until(EC.element_to_be_clickable((By.XPATH, "//input[@value='Submit']")))
        botao_submit.click()

        # Log de sucesso
        escrever_log(f"Rodada {i}: Dados enviados com sucesso para {first_name} {last_name}")
        dt_registros.at[i, 'Status'] = 'OK'  # Adiciona status de sucesso na tabela
        escrever_log(f"Rodada {i}: Status 'OK' marcado na coluna 'Status'")

    except Exception as e:
        # Log de erro
        escrever_log(f"Rodada {i}: Erro ao inserir dados para {first_name} {last_name} - {str(e)}")
        dt_registros.at[i, 'Status'] = 'Erro'  # Adiciona status de erro na tabela
        escrever_log(f"Rodada {i}: Status 'Erro' marcado na coluna 'Status'")

# Salvando o arquivo Excel atualizado com o status na coluna "Status"
dt_registros.to_excel(file, index=False)

# Aguardando 5 segundos
time.sleep(5)

# Tirando uma captura de tela e salvando
driver.save_screenshot(screenshot_file)
escrever_log(f"Captura de tela tirada e salva em {screenshot_file}")

# Fechando o navegador
driver.quit()

escrever_log("Fim do processo")
print("Fim do processo")
