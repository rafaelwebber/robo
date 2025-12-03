import os
import re
import pdfplumber
import pytesseract
from pdf2image import convert_from_path
import openpyxl
import unicodedata
import pandas as pd

# >>> Ajuste aqui para o caminho da pasta bin do Poppler <<<
POPPLER_PATH = r"C:\Users\rafae\Downloads\Release-25.11.0-0\poppler-25.11.0\Library\bin"
pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR"


# Regex para CPF (aceita diferentes formatos - mais flexível)
# Aceita: 123.456.789-10, 12345678910, 123 456 789 10, 123.456.789/10, etc.
padrao_cpf = re.compile(r'(\d{3}[\.\s\/-]?\d{3}[\.\s\/-]?\d{3}[\.\s\/-]?\d{2})')

# Regex para encontrar campo CPF (aceita diferentes variações: CPF, cpf, C.P.F, c.p.f, etc.)
padrao_campo_cpf = re.compile(r'c\.?\s*p\.?\s*f\.?', re.IGNORECASE)

def normalizar_nome(nome):
    """Normaliza o nome para comparação (remove espaços extras, converte para minúsculas, remove acentos)"""
    if not nome:
        return ""
    # Remove acentos
    nome_sem_acentos = ''.join(
        c for c in unicodedata.normalize('NFD', str(nome))
        if unicodedata.category(c) != 'Mn'
    )
    # Remove espaços extras e converte para minúsculas
    return re.sub(r'\s+', ' ', nome_sem_acentos.strip().lower())

def encontrar_cpf_proximo_ao_nome(texto, nome_requerente, debug=False):
    """Procura pelo nome do requerente no texto e retorna o CPF mais próximo"""
    nome_normalizado = normalizar_nome(nome_requerente)
    if not nome_normalizado:
        if debug:
            print("    [DEBUG] Nome do requerente vazio ou inválido")
        return None
    
    # Dividir texto em linhas
    linhas = texto.split('\n')
    
    # Primeiro, verificar se há CPFs no texto
    todos_cpfs = padrao_cpf.findall(texto)
    if debug:
        print(f"    [DEBUG] Total de CPFs encontrados no texto: {len(todos_cpfs)}")
        if todos_cpfs:
            print(f"    [DEBUG] CPFs encontrados: {todos_cpfs}")
    
    # Procurar linha(s) com o nome do requerente
    indices_nome = []
    palavras_nome = nome_normalizado.split()
    
    # Busca mais flexível: aceita nomes com 1 palavra também, mas prioriza nomes com mais palavras
    for i, linha in enumerate(linhas):
        linha_normalizada = normalizar_nome(linha)
        # Verifica quantas palavras do nome estão na linha
        palavras_encontradas = sum(1 for palavra in palavras_nome if palavra in linha_normalizada)
        
        # Se encontrou pelo menos 1 palavra (para nomes curtos) ou 2+ palavras (para nomes longos)
        if len(palavras_nome) == 1:
            if palavras_encontradas >= 1:
                indices_nome.append(i)
        else:
            if palavras_encontradas >= 2:
                indices_nome.append(i)
    
    if debug:
        print(f"    [DEBUG] Nome normalizado: '{nome_normalizado}'")
        print(f"    [DEBUG] Linhas com nome encontradas: {indices_nome}")
        if indices_nome:
            print(f"    [DEBUG] Exemplo de linha com nome: '{linhas[indices_nome[0]]}'")
    
    # Se não encontrou o nome, mas há CPFs, retornar o primeiro (fallback)
    if not indices_nome:
        if debug:
            print("    [DEBUG] Nome não encontrado no texto, mas tentando retornar primeiro CPF")
        return todos_cpfs[0] if todos_cpfs else None
    
    # Usar o primeiro índice encontrado como referência
    indice_nome = indices_nome[0]
    
    # Procurar CPF nas linhas próximas ao nome (até 15 linhas antes e depois)
    inicio_busca = max(0, indice_nome - 15)
    fim_busca = min(len(linhas), indice_nome + 16)
    
    # Estratégia 1: Procurar por padrão "CPF" seguido de números na região próxima
    for i in range(inicio_busca, fim_busca):
        linha = linhas[i]
        match_campo = padrao_campo_cpf.search(linha)
        if match_campo:
            # Se encontrou o campo CPF, procurar o número na mesma linha ou próximas
            texto_proximo = ' '.join(linhas[max(0, i-2):min(len(linhas), i+4)])
            cpfs = padrao_cpf.findall(texto_proximo)
            if cpfs:
                if debug:
                    print(f"    [DEBUG] CPF encontrado próximo ao campo 'CPF' na linha {i}: {cpfs[0]}")
                return cpfs[0]
    
    # Estratégia 2: Procurar CPF após o nome (nas linhas seguintes)
    for i in range(indice_nome, fim_busca):
        linha = linhas[i]
        cpfs = padrao_cpf.findall(linha)
        if cpfs:
            if debug:
                print(f"    [DEBUG] CPF encontrado na linha {i} (após o nome): {cpfs[0]}")
            return cpfs[0]
    
    # Estratégia 3: Procurar CPF em um bloco de texto próximo ao nome
    texto_proximo = ' '.join(linhas[inicio_busca:fim_busca])
    cpfs = padrao_cpf.findall(texto_proximo)
    if cpfs:
        if debug:
            print(f"    [DEBUG] CPF encontrado no bloco próximo ao nome: {cpfs[0]}")
        return cpfs[0]
    
    # Fallback: procurar qualquer CPF na página (último recurso)
    if todos_cpfs:
        if debug:
            print(f"    [DEBUG] Retornando primeiro CPF encontrado (fallback): {todos_cpfs[0]}")
        return todos_cpfs[0]
    
    if debug:
        print("    [DEBUG] Nenhum CPF encontrado")
    return None

def extrair_cpf_anexoII(caminho_pdf, nome_requerente, debug=False):
    pagina_anexo = None
    texto_pagina = ""

    # 1. Procurar página do Anexo II
    with pdfplumber.open(caminho_pdf) as pdf:
        for i, pagina in enumerate(pdf.pages):
            texto = pagina.extract_text()
            if texto and ("Anexo II" in texto or "ANEXO II" in texto or "anexo ii" in texto):
                pagina_anexo = i
                texto_pagina = texto
                # Tentar também extrair tabelas
                tabelas = pagina.extract_tables()
                if tabelas:
                    for tabela in tabelas:
                        for linha in tabela:
                            if linha:
                                texto_pagina += "\n" + " ".join([str(cell) if cell else "" for cell in linha])
                break

    if pagina_anexo is None:
        if debug:
            print("    [DEBUG] Página 'Anexo II' não encontrada no PDF")
        return None

    if debug:
        print(f"    [DEBUG] Página Anexo II encontrada: página {pagina_anexo + 1}")
        print(f"    [DEBUG] Tamanho do texto extraído: {len(texto_pagina)} caracteres")

    # 2. Tentar extrair CPF associado ao nome do requerente do texto extraído
    cpf = encontrar_cpf_proximo_ao_nome(texto_pagina, nome_requerente, debug=debug)
    if cpf:
        return cpf

    # 3. Se não achou, converter só a página do Anexo II em imagem e usar OCR
    if debug:
        print("    [DEBUG] Tentando OCR na página do Anexo II...")
    try:
        paginas = convert_from_path(
            caminho_pdf,
            first_page=pagina_anexo+1,
            last_page=pagina_anexo+1,
            poppler_path=POPPLER_PATH,
            dpi=300  # Maior resolução para melhor OCR
        )
        texto_ocr = pytesseract.image_to_string(paginas[0], lang="por")
        
        if debug:
            print(f"    [DEBUG] Tamanho do texto OCR: {len(texto_ocr)} caracteres")
        
        # Procurar CPF associado ao nome no texto do OCR
        cpf = encontrar_cpf_proximo_ao_nome(texto_ocr, nome_requerente, debug=debug)
        if cpf:
            return cpf
        
        # Fallback: retornar primeiro CPF encontrado se não conseguir associar ao nome
        cpfs = padrao_cpf.findall(texto_ocr)
        if debug and cpfs:
            print(f"    [DEBUG] CPFs encontrados no OCR (fallback): {cpfs}")
        return cpfs[0] if cpfs else None
    except Exception as e:
        if debug:
            print(f"    [DEBUG] Erro no OCR: {e}")
        return None


def atualizar_excel_iterativo(caminho_excel, pasta_pdfs):
    # Tentar abrir o arquivo Excel
    try:
        wb = openpyxl.load_workbook(caminho_excel)
    except PermissionError:
        print(f"\n[ERRO] O arquivo Excel está aberto ou sendo usado por outro programa!")
        print(f"Por favor, feche o arquivo '{caminho_excel}' e tente novamente.")
        return
    except FileNotFoundError:
        print(f"\n[ERRO] Arquivo Excel não encontrado: '{caminho_excel}'")
        return
    except Exception as e:
        print(f"\n[ERRO] Erro ao abrir o arquivo Excel: {e}")
        return
    
    ws = wb.active

    # Encontrar colunas existentes
    cabecalhos = [cell.value for cell in ws[1]]
    
    # Encontrar coluna do número do processo (geralmente primeira coluna)
    col_numero_processo = 1
    
    # Encontrar coluna "Requerente"
    col_requerente = None
    for idx, cabecalho in enumerate(cabecalhos, start=1):
        if cabecalho and "requerente" in str(cabecalho).lower() and "cpf" not in str(cabecalho).lower():
            col_requerente = idx
            break
    
    if col_requerente is None:
        print("[ERRO] Coluna 'Requerente' não encontrada no Excel!")
        return

    # Criar coluna "Requerente CPF" ao lado da coluna "Requerente"
    if "Requerente CPF" not in cabecalhos and "requerente_CPF" not in [str(c).lower() for c in cabecalhos]:
        # Inserir coluna após a coluna "Requerente"
        ws.insert_cols(col_requerente + 1)
        ws.cell(row=1, column=col_requerente + 1, value="Requerente CPF")
        col_cpf = col_requerente + 1
        # Ajustar índices das colunas seguintes se necessário
    else:
        # Encontrar coluna existente
        for idx, cabecalho in enumerate(cabecalhos, start=1):
            if cabecalho and ("requerente" in str(cabecalho).lower() and "cpf" in str(cabecalho).lower()):
                col_cpf = idx
                break
        else:
            col_cpf = col_requerente + 1

    # Iterar linhas
    for row in range(2, ws.max_row+1):
        numero_processo = str(ws.cell(row=row, column=col_numero_processo).value).strip()
        nome_requerente = str(ws.cell(row=row, column=col_requerente).value).strip() if ws.cell(row=row, column=col_requerente).value else ""
        
        caminho_pdf = os.path.join(pasta_pdfs, f"{numero_processo}.pdf")

        if os.path.exists(caminho_pdf):
            print(f"\n[...] Processo {numero_processo}: Processando... (Requerente: {nome_requerente})")
            # Ativar debug apenas para os primeiros processos ou quando não encontrar CPF
            debug = (row <= 3)  # Debug para os 3 primeiros processos
            cpf = extrair_cpf_anexoII(caminho_pdf, nome_requerente, debug=debug)
            if cpf:
                ws.cell(row=row, column=col_cpf, value=cpf)
                print(f"[OK] Processo {numero_processo}: CPF encontrado {cpf}")
            else:
                # Se não encontrou, tentar novamente com debug ativado
                if not debug:
                    print(f"[!] Processo {numero_processo}: CPF não encontrado, tentando com debug...")
                    cpf = extrair_cpf_anexoII(caminho_pdf, nome_requerente, debug=True)
                    if cpf:
                        ws.cell(row=row, column=col_cpf, value=cpf)
                        print(f"[OK] Processo {numero_processo}: CPF encontrado {cpf} (com debug)")
                    else:
                        ws.cell(row=row, column=col_cpf, value="CPF não encontrado")
                        print(f"[!] Processo {numero_processo}: CPF não encontrado")
                else:
                    ws.cell(row=row, column=col_cpf, value="CPF não encontrado")
                    print(f"[!] Processo {numero_processo}: CPF não encontrado")
        else:
            ws.cell(row=row, column=col_cpf, value="PDF não encontrado")
            print(f"[X] Processo {numero_processo}: PDF não encontrado")

        # Salvar imediatamente após cada linha
        try:
            wb.save(caminho_excel)
        except PermissionError:
            print(f"\n[ERRO] Não foi possível salvar o arquivo Excel!")
            print(f"O arquivo '{caminho_excel}' pode estar aberto. Feche-o e execute o script novamente.")
            wb.close()
            return
        except Exception as e:
            print(f"\n[ERRO] Erro ao salvar o arquivo Excel: {e}")
            wb.close()
            return
    
    # Fechar o arquivo ao final
    wb.close()
    print("\n[CONCLUÍDO] Processamento finalizado com sucesso!")


# Exemplo de uso
pasta_pdfs = r"G:\Drives compartilhados\Tecnologia\PDFs - Esaj TJSP"
caminho_excel = r"C:\Users\rafae\OneDrive\Desktop\robo\resultados_processo_final.xlsx"

atualizar_excel_iterativo(caminho_excel, pasta_pdfs)
