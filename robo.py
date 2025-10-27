import pandas as pd
import time, re, os
from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains

# Lê o Excel com os números de processo
df = pd.read_excel("C:/Users/rafae/OneDrive/Desktop/robo/processos.xlsx")
resultados = []
pasta_download = r"G:\Drives compartilhados\Tecnologia\PDFs - Esaj TJSP"

def separar_numero_processo(numero_completo):
    numero_limpo = re.sub(r"[.-]", "", numero_completo)
    if len(numero_limpo) != 20:
        raise ValueError(f"Número de processo inválido: {numero_completo}")
    parte1 = numero_limpo[:13]
    parte3 = numero_limpo[16:]
    return parte1, parte3

# Extrai dados do processo que são colocados via JS
def extrair_texto_por_id(driver, id_elemento, timeout=10):
    try:
        elemento = WebDriverWait(driver, timeout).until(
            EC.presence_of_element_located((By.ID, id_elemento))
        )
        texto = elemento.text.strip()
        if texto:
            return texto
        return elemento.get_attribute("innerText").strip()
    except:
        return None

def aguardar_download(pasta, timeout=30):
    tempo_inicial = time.time()
    arquivos_antes = set(os.listdir(pasta))
    while time.time() - tempo_inicial < timeout:
        arquivos_agora = set(os.listdir(pasta))
        novos = arquivos_agora - arquivos_antes
        for arquivo in novos:
            if arquivo.endswith(".pdf"):
                print(f"PDF baixado: {arquivo}")
                return arquivo
        time.sleep(1)
    print("Tempo limite atingido. Nenhum PDF encontrado.")
    return None

# Configura o navegador
options = Options()
options.add_argument("--start-maximized")
prefs = {
    "download.default_directory": pasta_download,
    "download.prompt_for_download": False,
    "download.directory_upgrade": True,
    "safebrowsing.enabled": True
}
options.add_experimental_option("prefs", prefs)

driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)

# Abre a página de login
driver.get("https://esaj.tjsp.jus.br/sajcas/login?service=https%3A%2F%2Fesaj.tjsp.jus.br%2Fcpopg%2FabrirConsultaDeRequisitorios.do")

# Pausa para login manual
input("Faça o login manualmente e pressione ENTER para continuar...")

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
        html_movs = driver.find_element(By.ID, "tabelaUltimasMovimentacoes").get_attribute("outerHTML")
        soup_movs = BeautifulSoup(html_movs, "html.parser")

        lista_movs = []
        for linha in soup_movs.find_all("tr"):
            texto = linha.get_text(separator=" ", strip=True) #na hora de extrair o text das tags garante que em tags diferentes o texto nao fique grudado
            lista_movs.append(texto)

        movimentacoes_formatadas = "\n".join(lista_movs)

        # Clica no link para abrir a pasta
        WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.ID, "linkPasta"))).click()

        time.sleep(2)  # pequena pausa para garantir que a aba abriu
        abas = driver.window_handles
        driver.switch_to.window(abas[-1])


        # Espera o botão de selecionar tudo aparecer e clica
        WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.ID, "selecionarButton"))).click()
        # Espera o botão de salvar aparecer e clica para baixar o PDF
        WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.ID, "salvarButton"))).click()

        WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.ID, "opcao1")))
        driver.find_element(By.ID, "opcao1").click()

        botao_continuar = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "botaoContinuar")))
        ActionChains(driver).move_to_element(botao_continuar).pause(1).click().perform()
        WebDriverWait(driver, 30).until(EC.invisibility_of_element_located((By.ID, "msgAguarde")))

        WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.ID, "btnDownloadDocumento"))).click()

        

        pdf = aguardar_download(pasta_download) 
        caminho_pdf = os.path.join(pasta_download, pdf) if pdf else "Não baixado"


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
        "Distribuicao": dataHora,
        "Controle": controle,
        "Area": area,
        "ValorAcao": valorAcao,
        "Requerente": ", ".join(requerentes),
        "Devedor": ", ".join(devedores),
        "Movimentacoes": movimentacoes_formatadas,
        "Petições diversas": peticoes,
        "Incidentes, acoes incidentais, recursos e execucoes de sentencas": incidentes,
        "Apensos, Entranhados e Unificados": apensos,
        "Audiencias": audiencia,
        "PDF": f'=HYPERLINK("{caminho_pdf}", "{caminho_pdf}")'
    })

    

driver.quit()

# Salva os resultados em um novo Excel
df_resultado = pd.DataFrame(resultados)

nome = "resultados_processos"
extensao = ".xlsx"
contador = 1

while True:
    nome_arquivo = f"{nome}_{contador}{extensao}"
    if not os.path.exists(nome_arquivo):
        break
    contador += 1

df_resultado.to_excel(nome_arquivo, index=False)