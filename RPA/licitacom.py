from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
import time
import pandas as pd
import os

# Configurar o serviço do ChromeDriver automaticamente
service = Service(ChromeDriverManager().install())

# Inicializar o navegador Chrome
driver = webdriver.Chrome(service=service)
try:
    # Abrir uma página da web
    driver.get("https://pncp.gov.br/app/editais?q=Engenharia&status=recebendo_proposta&pagina=1")

    time.sleep(10)
    # Localizar o campo de entrada pelo 'id'
    search_box = driver.find_element(By.ID, "keyword")
    search_box.clear()
    # Preencher o campo com o texto desejado
    search_box.send_keys("Obras")

    # Simular Enter para executar a pesquisa
    search_box.send_keys(Keys.RETURN)


    # Botão pesquisar por Xpath
    search_button = driver.find_element(By.XPATH, "//button[@aria-label='Buscar']")
    search_button.click()

    # Localizar todos os botões da lista
    buttons = driver.find_elements(By.CSS_SELECTOR, "li button.page")

    # Extrair os valores numéricos de cada botão
    button_values = [int(button.text.strip()) for button in buttons]

    # Encontrar o valor mais alto
    highest_value = max(button_values)

    # Localizar o botão com o valor mais alto e clicar nele
    # Lista para armazenar os dados extraídos
    resultados = []
    for button in buttons:
        button.click()
        time.sleep(5)
        
        # Encontre todos os itens de edital
        editais = driver.find_elements(By.CLASS_NAME, "br-item")

        

        # Percorra cada edital para capturar informações
        for edital in editais:
            titulo = edital.find_element(By.TAG_NAME, "strong").text
            id_contratacao = edital.find_element(By.XPATH, ".//span[contains(., 'Id contratação PNCP')]").text
            modalidade = edital.find_element(By.XPATH, ".//div[contains(., 'Modalidade da Contratação')]").text
            ultima_atualizacao = edital.find_element(By.XPATH, ".//div[contains(., 'Última Atualização')]").text
            orgao = edital.find_element(By.XPATH, ".//div[contains(., 'Órgão')]").text
            local = edital.find_element(By.XPATH, ".//div[contains(., 'Local')]").text
            objeto = edital.find_element(By.XPATH, ".//span[contains(., 'Objeto')]").text

            # Adiciona os dados em um dicionário
            resultados.append({
                "Título": titulo,
                "ID Contratação PNCP": id_contratacao,
                "Modalidade": modalidade,
                "Última Atualização": ultima_atualizacao,
                "Órgão": orgao,
                "Local": local,
                "Objeto": objeto,
            })

    # Salvar os resultados em um DataFrame do pandas
    df = pd.DataFrame(resultados)
    
    # Defina o caminho da pasta
    diretorio_destino = r"C:\Users\gabri\Documents\projetos_python\RPA\excel_results" 
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
except Exception as e:
    print(f"Erro: {e}")