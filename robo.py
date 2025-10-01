import pandas as pd
import time, re
from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.
import Keys
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager

# Lê o Excel com os números de processo
df = pd.read_excel("C:/Users/rafae/OneDrive/Desktop/ESCRITÓRIO/robo/processos.xlsx")
resultados = []

def separar_numero_processo(numero_completo):
    numero_limpo = re.sub(r"[.-]", "", numero_completo)
    if len(numero_limpo) != 20:
        raise ValueError(f"Número de processo inválido: {numero_completo}")
    parte1 = numero_limpo[:13]
    parte3 = numero_limpo[16:]
    return parte1, parte3

# Configura o navegador
options = Options()
options.add_argument("--start-maximized")
driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)

for processo in df['numero_processo']:
    print(f"processo: {processo}")
    driver.get("https://esaj.tjsp.jus.br/cpopg/search.do?conversationId=&cbPesquisa=NUMPROC")
    time.sleep(2)

    try:
        parte1, parte3 = separar_numero_processo(processo)

        # Preenche os campos
        driver.find_element(By.ID, "numeroDigitoAnoUnificado").send_keys(parte1)
        driver.find_element(By.ID, "foroNumeroUnificado").send_keys(parte3)

        # Clica em "Pesquisar"
        driver.find_element(By.ID, "botaoConsultarProcessos").click()
        time.sleep(5)

        # Extrai dados principais
        classe = driver.find_element(By.ID, "classeProcesso").text
        assunto = driver.find_element(By.ID, "assuntoProcesso").text
        foro = driver.find_element(By.ID, "foroProcesso").text
        vara = driver.find_element(By.ID, "varaProcesso").text
        juiz = driver.find_element(By.ID, "juizProcesso").text

        # Extrai partes
        html_partes = driver.find_element(By.ID, "tablePartesPrincipais").get_attribute("outerHTML")
        soup_partes = BeautifulSoup(html_partes, "html.parser")

        requerentes = []
        devedores = []

        for linha in soup_partes.find_all("tr"):
            texto = linha.get_text(separator=" ", strip=True).upper()
            if any(palavra in texto for palavra in ["REQUERENTE", "REQTE", "EXEQUENTE", "PARTE ATIVA"]):
                requerentes.append(texto)
            elif any(palavra in texto for palavra in ["REQUERIDO", "EXECUTADO", "DEVEDOR", "PARTE PASSIVA"]):
                devedores.append(texto)

        # Extrai movimentações
        html_movs = driver.find_element(By.ID, "tabelaTodasMovimentacoes").get_attribute("outerHTML")
        soup_movs = BeautifulSoup(html_movs, "html.parser")

        lista_movs = []
        for linha in soup_movs.find_all("tr"):
            texto = linha.get_text(separator=" ", strip=True)
            lista_movs.append(texto)

        movimentacoes_formatadas = "\n".join(lista_movs)

    except Exception as e:
        print(f"Erro ao consultar {processo}: {e}")
        classe = assunto = movimentacoes_formatadas = "Erro ou não encontrado"
        requerentes = devedores = []
        foro = vara = juiz = "Erro"

    resultados.append({
        "numero_processo": processo,
        "classe": classe,
        "assunto": assunto,
        "foro": foro,
        "vara": vara,
        "juiz": juiz,
        "requerente": ", ".join(requerentes),
        "devedor": ", ".join(devedores),
        "movimentacoes": movimentacoes_formatadas
    })

driver.quit()

# Salva os resultados em um novo Excel
df_resultado = pd.DataFrame(resultados)
df_resultado.to_excel("resultados_processos_selenium.xlsx", index=False)
