import pandas as pd
import time, re
from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

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

# Extrai dados do processo
def extrair_texto_por_id(driver, id_elemento, timeout=10):
    try:
        WebDriverWait(driver, timeout).until(
            EC.presence_of_element_located((By.ID, id_elemento))
        )
        elemento = driver.find_element(By.ID, id_elemento)
        return elemento.get_attribute("innerText").strip()
    except:
        return None
    
# Configura o navegador
options = Options()
options.add_argument("--start-maximized")
driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)

# Abre a página de login
driver.get("https://esaj.tjsp.jus.br/cpopg/abrirConsultaDeRequisitorios.do")

# Pausa para login manual
input("⏸️ Faça o login manualmente e pressione ENTER para continuar...")

for processo in df['numero_processo']:
    print(f"processo: {processo}")
    driver.get("https://esaj.tjsp.jus.br/cpopg/abrirConsultaDeRequisitorios.do")
    time.sleep(2)

    try:
        parte1, parte3 = separar_numero_processo(processo)

        # Preenche os campos
        driver.find_element(By.ID, "numeroDigitoAnoUnificado").send_keys(parte1)
        driver.find_element(By.ID, "foroNumeroUnificado").send_keys(parte3)

        # Clica em consultar
        driver.find_element(By.ID, "botaoConsultarProcessos").click()
        time.sleep(5)

        # chamada da funcao e passagem de parametros
        classe     = extrair_texto_por_id(driver, "classeProcesso")
        assunto    = extrair_texto_por_id(driver, "assuntoProcesso")
        foro       = extrair_texto_por_id(driver, "foroProcesso")
        vara       = extrair_texto_por_id(driver, "varaProcesso")
        juiz       = extrair_texto_por_id(driver, "juizProcesso")
        dataHora   = extrair_texto_por_id(driver, "dataHoraDistribuicaoProcesso")
        controle   = extrair_texto_por_id(driver, "numeroControleProcesso")
        area       = extrair_texto_por_id(driver, "areaProcesso")
        valorAcao  = extrair_texto_por_id(driver, "valorAcaoProcesso")
        peticoes   = extrair_texto_por_id(driver, "processoSemDiversas")
        incidentes = extrair_texto_por_id(driver, "processoSemIncidentes")
        apensos    = extrair_texto_por_id(driver, "dadosApensosNaoDisponiveis")
        audiencia  = extrair_texto_por_id(driver, "processoSemAudiencias")
        

        # Extrai partes envolvidas
        html_partes = driver.find_element(By.ID, "tablePartesPrincipais").get_attribute("outerHTML")
        soup_partes = BeautifulSoup(html_partes, "html.parser")

        requerentes = []
        devedores = []

        #percorre todas as linha de soup na tag tr e depois faz uma condição na qual palavra no texto for igual a alguma STR dento da lista ele faz um append em requerente ou devedores
        for linha in soup_partes.find_all("tr"):#tr é uma tag HTML 
            texto = linha.get_text(separator=" ", strip=True).upper()
            if any(palavra in texto for palavra in ["REQUERENTE", "REQTE", "EXEQUENTE", "PARTE ATIVA"]):
                requerentes.append(texto)
            elif any(palavra in texto for palavra in ["REQUERIDO", "EXECUTADO", "DEVEDOR", "PARTE PASSIVA"]):
                devedores.append(texto)

        # Extrai todas as movimentações
        html_movs = driver.find_element(By.ID, "tabelaTodasMovimentacoes").get_attribute("outerHTML")
        soup_movs = BeautifulSoup(html_movs, "html.parser")

        lista_movs = []
        for linha in soup_movs.find_all("tr"):
            texto = linha.get_text(separator=" ", strip=True) #na hora de extrair o text das tags garante que em tags diferentes o texto nao fique grudado
            lista_movs.append(texto)

        movimentacoes_formatadas = "\n".join(lista_movs)

        # Clica no link para abrir a pasta
        WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.ID, "linkPasta"))).click()

        # Espera o botão de selecionar tudo aparecer e clica
        WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.ID, "selecionarButton"))).click()

        # Espera o botão de salvar aparecer e clica para baixar o PDF
        WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.ID, "salvarButton"))).click()


    except Exception as e:
        print(f"Erro ao consultar {processo}: {e}")
        classe = assunto = movimentacoes_formatadas = "Erro ou não encontrado"
        requerentes = devedores = []
        foro = vara = juiz = "Erro"

    resultados.append({
        "numero_processo": processo,
        "Classe": classe,
        "Assunto": assunto,
        "Foro": foro,
        "Vara": vara,
        "Juiz": juiz,
        "Distribuição": dataHora,
        "Controle": controle,
        "Area": area,
        "ValorAcao": valorAcao,
        "Requerente": ", ".join(requerentes),
        "Devedor": ", ".join(devedores),
        "Movimentacoes": movimentacoes_formatadas,
        "Petições diversas": peticoes,
        "Incidentes, ações incidentais, recursos e execuções de sentenças": incidentes,
        "Apensos, Entranhados e Unificados": apensos,
        "Audiências": audiencia
    })

    

driver.quit()

# Salva os resultados em um novo Excel
df_resultado = pd.DataFrame(resultados)
df_resultado.to_excel("resultados_processos_selenium.xlsx", index=False)
