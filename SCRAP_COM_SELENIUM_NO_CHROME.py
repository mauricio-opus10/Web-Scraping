import time
import openpyxl
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By

# Caminho para o ChromeDriver
driver_path = r"endereco_do_driver_no_seu_computador\chromedriver.exe"
service = Service(driver_path)

# Função para realizar o scroll


def scroll_page(driver):
    scroll_pause_time = 3  # Tempo para a página carregar após cada scroll
    last_height = driver.execute_script("return document.body.scrollHeight")

    while True:
        # Rolar até o final da página
        driver.execute_script(
            "window.scrollTo(0, document.body.scrollHeight);")
        time.sleep(scroll_pause_time)

        # Calcular nova altura da página
        new_height = driver.execute_script("return document.body.scrollHeight")
        if new_height == last_height:
            break  # Se a altura não mudou, terminamos o scroll
        last_height = new_height

# Função principal


def scrape_all_pages():
    print("Iniciando o driver...")
    driver = webdriver.Chrome(service=service)

    # Criar o arquivo Excel
    excel_file = "imoveis.xlsx"
    workbook = openpyxl.Workbook()
    sheet = workbook.active
    sheet.title = "Imóveis"
    sheet.append(["Título", "Localização", "Tamanho",
                 "Quartos", "Vagas", "Preço", "URL"])

    try:
        # URL inicial
        base_url = "endereco_da_url_do_site"
        driver.get(base_url)
        time.sleep(5)  # Espera inicial para carregar a página

        page_number = 1
        while True:
            print(f"Acessando a página {page_number}...")

            # Realizar scroll na página para carregar todos os imóveis
            print("Realizando scroll na página...")
            scroll_page(driver)
            time.sleep(2)  # Garantir que os elementos sejam carregados

            # Encontrar os imóveis na página atual
            imoveis = driver.find_elements(By.CLASS_NAME, "building-content")
            print(
                f"Encontrados {len(imoveis)} imóveis na página {page_number}.")

            # Processar os imóveis
            for imovel in imoveis:
                try:
                    titulo = imovel.find_element(
                        By.CLASS_NAME, "building-content--title").text
                except:
                    titulo = "N/A"

                try:
                    localizacao = imovel.find_elements(
                        By.CLASS_NAME, "block-title")[0].text
                except:
                    localizacao = "N/A"

                try:
                    tamanho = imovel.find_elements(
                        By.CLASS_NAME, "block-title")[1].text
                except:
                    tamanho = "N/A"

                try:
                    quartos = imovel.find_elements(
                        By.CLASS_NAME, "block-title")[2].text
                except:
                    quartos = "N/A"

                try:
                    vagas = imovel.find_elements(
                        By.CLASS_NAME, "block-title")[3].text
                except:
                    vagas = "N/A"

                try:
                    preco = imovel.find_element(By.CLASS_NAME, "price").text
                except:
                    preco = "N/A"

                try:
                    url = imovel.find_element(
                        By.XPATH, ".//ancestor::a").get_attribute("href")
                except:
                    url = "N/A"

                print("Título:", titulo)
                print("Localização:", localizacao)
                print("Tamanho:", tamanho)
                print("Quartos:", quartos)
                print("Vagas:", vagas)
                print("Preço:", preco)
                print("URL:", url)
                print("-" * 50)

                # Adicionar os dados na planilha Excel
                sheet.append([titulo, localizacao, tamanho,
                             quartos, vagas, preco, url])

            # Verificar se existe o botão "Próximo"
            try:
                next_button = driver.find_element(By.LINK_TEXT, "PRÓXIMO")
                next_button.click()
                page_number += 1
                time.sleep(5)  # Espera para carregar a próxima página
            except:
                print("Não há mais páginas. Finalizando a raspagem.")
                break

    finally:
        # Salvar o arquivo Excel
        workbook.save(excel_file)
        print(f"Dados salvos com sucesso no arquivo: {excel_file}")

        # Encerrar o driver
        print("Encerrando o driver...")
        driver.quit()


# Executar a função
scrape_all_pages()
