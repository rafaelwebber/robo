import os
import re
import pdfplumber
import pytesseract
from pdf2image import convert_from_path
import unicodedata

# >>> Ajuste aqui para o caminho da pasta bin do Poppler <<<
POPPLER_PATH = r"C:\Users\rafae\Downloads\Release-25.11.0-0\poppler-25.11.0\Library\bin"
pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR"

# Regex para CPF (aceita diferentes formatos - mais flexível)
padrao_cpf = re.compile(r'(\d{3}[\.\s\/-]?\d{3}[\.\s\/-]?\d{3}[\.\s\/-]?\d{2})')
padrao_campo_cpf = re.compile(r'c\.?\s*p\.?\s*f\.?', re.IGNORECASE)

def normalizar_nome(nome):
    """Normaliza o nome para comparação"""
    if not nome:
        return ""
    nome_sem_acentos = ''.join(
        c for c in unicodedata.normalize('NFD', str(nome))
        if unicodedata.category(c) != 'Mn'
    )
    return re.sub(r'\s+', ' ', nome_sem_acentos.strip().lower())

def testar_pdf(numero_processo, pasta_pdfs, nome_requerente=None):
    """Testa a extração de CPF de um PDF específico"""
    caminho_pdf = os.path.join(pasta_pdfs, f"{numero_processo}.pdf")
    
    if not os.path.exists(caminho_pdf):
        print(f"[ERRO] PDF não encontrado: {caminho_pdf}")
        return
    
    print(f"=" * 80)
    print(f"TESTANDO PDF: {numero_processo}")
    print(f"Caminho: {caminho_pdf}")
    print(f"=" * 80)
    
    # 1. Procurar página do Anexo II
    print("\n[1] Procurando página 'Anexo II'...")
    pagina_anexo = None
    texto_pagina = ""
    
    with pdfplumber.open(caminho_pdf) as pdf:
        print(f"    Total de páginas: {len(pdf.pages)}")
        for i, pagina in enumerate(pdf.pages):
            texto = pagina.extract_text()
            if texto:
                # Verificar variações de "Anexo II"
                if "Anexo II" in texto or "ANEXO II" in texto or "anexo ii" in texto or "ANEXO 2" in texto or "Anexo 2" in texto:
                    pagina_anexo = i
                    texto_pagina = texto
                    print(f"    [OK] Pagina Anexo II encontrada na pagina {i+1}")
                    break
                # Mostrar primeiras palavras de cada página para debug
                if i < 3:
                    palavras_inicio = texto[:100].replace('\n', ' ')
                    print(f"    Página {i+1} (início): {palavras_inicio}...")
    
    if pagina_anexo is None:
        print("    [ERRO] Pagina 'Anexo II' NAO encontrada!")
        print("\n[DEBUG] Tentando buscar em todas as páginas...")
        with pdfplumber.open(caminho_pdf) as pdf:
            for i, pagina in enumerate(pdf.pages):
                texto = pagina.extract_text()
                if texto:
                    # Procurar por palavras-chave relacionadas
                    texto_lower = texto.lower()
                    if "anexo" in texto_lower or "requerente" in texto_lower:
                        print(f"\n    Página {i+1} contém 'anexo' ou 'requerente'")
                        print(f"    Primeiras 200 caracteres: {texto[:200]}")
        return None
    
    # 2. Extrair tabelas também
    print("\n[2] Extraindo texto e tabelas da página do Anexo II...")
    with pdfplumber.open(caminho_pdf) as pdf:
        pagina = pdf.pages[pagina_anexo]
        tabelas = pagina.extract_tables()
        if tabelas:
            print(f"    [OK] {len(tabelas)} tabela(s) encontrada(s)")
            for idx, tabela in enumerate(tabelas):
                print(f"\n    Tabela {idx+1}:")
                for linha in tabela[:5]:  # Mostrar primeiras 5 linhas
                    if linha:
                        linha_texto = " | ".join([str(cell) if cell else "" for cell in linha])
                        print(f"      {linha_texto}")
                        texto_pagina += "\n" + linha_texto
        else:
            print("    Nenhuma tabela encontrada")
    
    # 3. Mostrar texto extraído
    print(f"\n[3] Texto extraído da página do Anexo II ({len(texto_pagina)} caracteres):")
    print("-" * 80)
    # Mostrar primeiras 1000 caracteres
    print(texto_pagina[:1000])
    if len(texto_pagina) > 1000:
        print(f"\n... (mais {len(texto_pagina) - 1000} caracteres)")
    print("-" * 80)
    
    # 4. Procurar CPFs no texto
    print("\n[4] Procurando CPFs no texto...")
    cpfs_encontrados = padrao_cpf.findall(texto_pagina)
    print(f"    Total de CPFs encontrados: {len(cpfs_encontrados)}")
    if cpfs_encontrados:
        for idx, cpf in enumerate(cpfs_encontrados, 1):
            print(f"    CPF {idx}: {cpf}")
    else:
        print("    [ERRO] Nenhum CPF encontrado no texto extraido!")
    
    # 5. Procurar campo "CPF" no texto
    print("\n[5] Procurando campo 'CPF' no texto...")
    linhas = texto_pagina.split('\n')
    for i, linha in enumerate(linhas):
        if padrao_campo_cpf.search(linha):
            print(f"    [OK] Campo 'CPF' encontrado na linha {i+1}: {linha[:100]}")
            # Procurar CPF próximo
            contexto = ' '.join(linhas[max(0, i-2):min(len(linhas), i+3)])
            cpfs_proximos = padrao_cpf.findall(contexto)
            if cpfs_proximos:
                print(f"    CPFs próximos ao campo: {cpfs_proximos}")
    
    # 6. Se não encontrou CPF, tentar OCR
    if not cpfs_encontrados:
        print("\n[6] Nenhum CPF encontrado no texto. Tentando OCR...")
        try:
            paginas = convert_from_path(
                caminho_pdf,
                first_page=pagina_anexo+1,
                last_page=pagina_anexo+1,
                poppler_path=POPPLER_PATH,
                dpi=300
            )
            texto_ocr = pytesseract.image_to_string(paginas[0], lang="por")
            print(f"    Texto OCR extraído ({len(texto_ocr)} caracteres):")
            print("-" * 80)
            print(texto_ocr[:1000])
            if len(texto_ocr) > 1000:
                print(f"\n... (mais {len(texto_ocr) - 1000} caracteres)")
            print("-" * 80)
            
            cpfs_ocr = padrao_cpf.findall(texto_ocr)
            print(f"\n    CPFs encontrados no OCR: {len(cpfs_ocr)}")
            if cpfs_ocr:
                for idx, cpf in enumerate(cpfs_ocr, 1):
                    print(f"    CPF {idx}: {cpf}")
        except Exception as e:
            print(f"    [ERRO] Erro no OCR: {e}")
    
    # 7. Se nome do requerente foi fornecido, procurar próximo ao nome
    if nome_requerente:
        print(f"\n[7] Procurando CPF próximo ao nome do requerente: '{nome_requerente}'...")
        nome_normalizado = normalizar_nome(nome_requerente)
        linhas = texto_pagina.split('\n')
        
        # Procurar linha com o nome
        indices_nome = []
        palavras_nome = nome_normalizado.split()
        for i, linha in enumerate(linhas):
            linha_normalizada = normalizar_nome(linha)
            palavras_encontradas = sum(1 for palavra in palavras_nome if palavra in linha_normalizada)
            if len(palavras_nome) == 1:
                if palavras_encontradas >= 1:
                    indices_nome.append(i)
            else:
                if palavras_encontradas >= 2:
                    indices_nome.append(i)
        
        if indices_nome:
            print(f"    [OK] Nome encontrado nas linhas: {indices_nome}")
            for idx in indices_nome[:3]:  # Mostrar primeiras 3 ocorrências
                print(f"    Linha {idx+1}: {linhas[idx][:100]}")
            
            # Procurar CPF próximo
            indice_nome = indices_nome[0]
            inicio = max(0, indice_nome - 10)
            fim = min(len(linhas), indice_nome + 11)
            contexto = ' '.join(linhas[inicio:fim])
            cpfs_proximos = padrao_cpf.findall(contexto)
            if cpfs_proximos:
                print(f"    [OK] CPFs proximos ao nome: {cpfs_proximos}")
            else:
                print("    [ERRO] Nenhum CPF encontrado proximo ao nome")
        else:
            print(f"    [ERRO] Nome '{nome_requerente}' nao encontrado no texto")
    
    print("\n" + "=" * 80)
    return cpfs_encontrados[0] if cpfs_encontrados else None

# Teste
if __name__ == "__main__":
    numero_processo = "0443679-76.2019.8.26.0500"
    pasta_pdfs = r"G:\Drives compartilhados\Tecnologia\PDFs - Esaj TJSP"
    
    # Nome do requerente encontrado no PDF
    nome_requerente = "Maurício Ferreira Leite"  # Nome encontrado no PDF
    
    resultado = testar_pdf(numero_processo, pasta_pdfs, nome_requerente)
    
    if resultado:
        print(f"\n[RESULTADO] CPF encontrado: {resultado}")
    else:
        print(f"\n[RESULTADO] Nenhum CPF encontrado")

