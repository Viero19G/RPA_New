from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.action_chains import ActionChains
import time
import pandas as pd
import os

# Configurar o serviço do ChromeDriver automaticamente
service = Service(ChromeDriverManager().install())

# Inicializar o navegador Chrome
driver = webdriver.Chrome(service=service)

# Abrir uma página da web
driver.get("https://pncp.gov.br/app/editais?q=Engenharia&status=recebendo_proposta&pagina=1")

time.sleep(10)
# Localizar o campo de entrada pelo 'id'
search_box = driver.find_element(By.ID, "keyword")
search_box.clear()
# Preencher o campo com o texto desejado
search_box.send_keys("Obra", Keys.ENTER)

# # Simular Enter para executar a pesquisa
# search_box.send_keys(Keys.RETURN)


# # Botão pesquisar por Xpath
# search_button = driver.find_element(By.XPATH, '//*[@id="main-content"]/pncp-list/pncp-results-panel/pncp-tab-set/div/pncp-top-panel/div/div/div[3]/div[2]/button')
# driver.execute_script("arguments[0].scrollIntoView(true);", search_button)
# search_button.click()
time.sleep(5)
#Botão com número final será utilizado para o tamanho do for ser exato
last_btn =  driver.find_element(By.XPATH,'//*[@id="main-content"]/pncp-list/pncp-results-panel/pncp-tab-set/div/div/div/pncp-pagination/nav/ul/li[10]/button')
buttons = int(last_btn.text) -1 # menos um pois a primeira chamada do for não precisa ser considerada

    # Encontre todos os itens de edital
editais = driver.find_elements(By.XPATH, '/html/body/app-root/div/main/div/div/div[2]/div/pncp-list/pncp-results-panel/pncp-tab-set/div/pncp-tab[1]/div/div[2]/div/div[2]/pncp-items-list/div')


# Lista para armazenar os dados extraídos
resultados = []

# Percorra cada edital para capturar informações
for edital in editais:
    item = driver.find_element(By.XPATH, '//*[@id="main-content"]/pncp-list/pncp-results-panel/pncp-tab-set/div/pncp-tab[1]/div/div[2]/div/div[2]/pncp-items-list/div/div[1]/a')
    
    texto = item.text

    # Separar as linhas do texto
    linhas = texto.split("\n")

    # Extrair as informações específicas
    dados = {
        "Edital": linhas[0].replace("Edital nº", "").strip(),
        "ID Contratação PNCP": linhas[1].replace("Id contratação PNCP:", "").strip(),
        "Modalidade": linhas[2].replace("Modalidade da Contratação:", "").strip(),
        "Última Atualização": linhas[3].replace("Última Atualização:", "").strip(),
        "Órgão": linhas[4].replace("Órgão:", "").strip(),
        "Local": linhas[5].replace("Local:", "").strip(),
        "Objeto": linhas[6].replace("Objeto:", "").strip(),
    }

    # Adicionar os dados extraídos à lista de resultados
    resultados.append(dados)
print(dados)
breakpoint()


for _ in range(buttons):
    iframe = driver.find_element(By.XPATH, '//button[@data-next-page="data-next-page"]')
    ActionChains(driver)\
        .scroll_to_element(iframe)\
        .perform()
    time.sleep(5)
    clicar = driver.find_element(By.XPATH, '//button[@data-next-page="data-next-page"]')
    clicar.click()
    
    # Encontre todos os itens de edital
    editais = driver.find_elements(By.CLASS_NAME, "br-item")

    for edital in editais:
        item = driver.find_element(By.XPATH, '//*[@id="main-content"]/pncp-list/pncp-results-panel/pncp-tab-set/div/pncp-tab[1]/div/div[2]/div/div[2]/pncp-items-list/div/div[1]/a')
        
        texto = item.text

        # Separar as linhas do texto
        linhas = texto.split("\n")

        # Extrair as informações específicas
        dados = {
            "Edital": linhas[0].replace("Edital nº", "").strip(),
            "ID Contratação PNCP": linhas[1].replace("Id contratação PNCP:", "").strip(),
            "Modalidade": linhas[2].replace("Modalidade da Contratação:", "").strip(),
            "Última Atualização": linhas[3].replace("Última Atualização:", "").strip(),
            "Órgão": linhas[4].replace("Órgão:", "").strip(),
            "Local": linhas[5].replace("Local:", "").strip(),
            "Objeto": linhas[6].replace("Objeto:", "").strip(),
        }

        # Adicionar os dados extraídos à lista de resultados
        resultados.append(dados)

# Salvar os resultados em um DataFrame do pandas
df = pd.DataFrame(resultados)

# Defina o caminho da pasta
diretorio_destino = r"C:\Users\gabri\Documents\projetos_python\RPA\Results" 
if not os.path.exists(diretorio_destino):
    os.makedirs(diretorio_destino)  # Cria a pasta caso não exista

# nome do arquivo
nome_arquivo = "editais_resultados.xlsx"

# Caminho completo para salvar o arquivo
caminho_completo = os.path.join(diretorio_destino, nome_arquivo)

# Salve o arquivo Excel na pasta definida
df.to_excel(caminho_completo, index=False, engine='openpyxl')
print(f"Os resultados foram salvos em {caminho_completo}.")

    
breakpoint()
# Fechar o navegador
driver.quit()
