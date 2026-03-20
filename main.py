import re
import os
import time
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler
from datetime import datetime
import pdfplumber
import openpyxl

pdf_folder = r'C:\Users\05454745235\Projetos\leitor-pdf\pdf'

excel_file = openpyxl.Workbook()
excel_worksheet = excel_file.active

# CABEÇALHOS CORRETOS
excel_worksheet.append([
    "Nome", "CPF", "Curso", "Matrícula",
    "Semestre", "Disciplina", "Professor", 
    "Carga Horária", "Média Final", "Faltas", "Situação"
])

os.makedirs(pdf_folder, exist_ok=True)
data_atual = datetime.now().strftime("%d-%m-%Y")

def separar_coluna(texto, expected_count, is_componente=False):
    """Separa as células fundidas verticalmente baseando-se no número esperado de quebras"""
    texto = str(texto or '').strip()
    
    if expected_count <= 1:
        return [texto.replace('\n', ' ').strip()] if not is_componente else [texto]
        
    blocos = re.split(r'\n{2,}', texto)
    blocos_limpos = [b.strip() for b in blocos if b.strip()]
    
    if len(blocos_limpos) == expected_count:
        return blocos_limpos if is_componente else [b.replace('\n', ' ') for b in blocos_limpos]
        
    blocos_simples = [b.strip() for b in texto.split('\n') if b.strip()]
    
    if len(blocos_simples) == expected_count:
        return blocos_simples
        
    while len(blocos_limpos) < expected_count:
        blocos_limpos.append("-")
        
    return blocos_limpos[:expected_count]

def processar_linha_ancorada(linha):
    """Processa a linha encontrando a Carga Horária como âncora central"""
    linha_limpa = [str(c).strip() for c in linha if c is not None and str(c).strip() != '']
    
    if len(linha_limpa) < 5:
        return []
        
    ord_str = linha_limpa[0]
    ords = [x for x in re.split(r'\n+', ord_str) if x.strip().isdigit()]
    num_discs = len(ords)
    
    if num_discs == 0:
        return []
        
    # Encontra a "Âncora" -> Carga Horária (ex: "75h")
    idx_carga = -1
    for i, celula in enumerate(linha_limpa):
        if re.search(r'\d+\s*[hH]', celula):
            idx_carga = i
            break
            
    if idx_carga == -1:
        return [] 
        
    # Mapeamento do que vem antes e logo após a âncora
    semestre_str = linha_limpa[1]
    carga_str = linha_limpa[idx_carga]
    media_str = linha_limpa[idx_carga + 1] if idx_carga + 1 < len(linha_limpa) else "-"
    
    # --- NOVO: Extração Inteligente de Faltas e Situação ---
    # Pega TUDO que sobrou depois da Média, independente de como o PDF fundiu
    bloco_final = " ".join(linha_limpa[idx_carga + 2:])
    bloco_final_limpo = re.sub(r'\s+', ' ', bloco_final).strip()
    
    # 1. Pega as Situações procurando os padrões de aprovação/reprovação
    padrao_sit = r'(Aprovado|Reprovado\s*por\s*M.dia|Reprovado\s*por\s*Falta|Reprovado|Cursando|Trancado|Dispensado)'
    situacoes_encontradas = [s.strip().title() for s in re.findall(padrao_sit, bloco_final_limpo, flags=re.IGNORECASE)]
    
    # 2. Remove as palavras de situação do texto e pega os números que sobraram (as Faltas)
    sobra_faltas = re.sub(padrao_sit, ' ', bloco_final_limpo, flags=re.IGNORECASE)
    faltas_encontradas = re.findall(r'\b\d+\b', sobra_faltas)
    
    # 3. Garante que temos um item para cada disciplina processada na linha
    while len(faltas_encontradas) < num_discs: faltas_encontradas.append("-")
    while len(situacoes_encontradas) < num_discs: situacoes_encontradas.append("-")
    
    faltas = faltas_encontradas[:num_discs]
    situacoes = situacoes_encontradas[:num_discs]
    # -------------------------------------------------------
    
    # Disciplina e Professor
    meio = linha_limpa[2:idx_carga]
    if len(meio) >= 2:
        disciplina_str = " ".join(meio[:-1])
        professor_str = meio[-1]
    elif len(meio) == 1:
        disciplina_str = meio[0]
        professor_str = "-"
    else:
        disciplina_str, professor_str = "-", "-"

    # Separação
    semestres = separar_coluna(semestre_str, num_discs)
    cargas = separar_coluna(carga_str, num_discs)
    medias = separar_coluna(media_str, num_discs)
    disciplinas = separar_coluna(disciplina_str, num_discs, is_componente=True)
    professores = separar_coluna(professor_str, num_discs, is_componente=True)
    
    discs_desmembradas = []
    for i in range(num_discs):
        disc_atual = disciplinas[i].replace('\n', ' ') if i < len(disciplinas) else "-"
        prof_atual = professores[i].replace('\n', ' ') if i < len(professores) else "-"
        sem_atual = semestres[i] if i < len(semestres) else "-"
        carga_atual = cargas[i] if i < len(cargas) else "-"
        media_atual = medias[i] if i < len(medias) else "-"
        
        discs_desmembradas.append([
            ords[i], sem_atual, disc_atual, prof_atual,
            carga_atual, media_atual, faltas[i], situacoes[i]
        ])
        
    return discs_desmembradas

def extract_info_and_write_to_excel(pdf_file_path):
    global excel_worksheet

    text_content = ""
    todas_as_tabelas = []

    with pdfplumber.open(pdf_file_path) as pdf:
        for page in pdf.pages:
            text_content += page.extract_text() + "\n"
            tabelas_pagina = page.extract_tables()
            for tabela in tabelas_pagina:
                todas_as_tabelas.extend(tabela)

    if "Histórico Escolar" not in text_content:
        print(f"Ignorado (Não é histórico escolar válido): {pdf_file_path}")
        return

    print(f"\nProcessando histórico: {pdf_file_path}")

    nome = re.search(r'Aluno\(a\):\s*(.+?)(?:\s*CPF|\n)', text_content)
    cpf = re.search(r'CPF:\s*([\d\.\-]+)', text_content)
    curso = re.search(r'Curso:\s*(.+?)(?:\s*Matrícula|\n)', text_content)
    matricula = re.search(r'Matrícula:\s*(\d+)', text_content)

    nome_str = nome.group(1).strip() if nome else "Não encontrado"
    cpf_str = cpf.group(1).strip() if cpf else "Não encontrado"
    curso_str = curso.group(1).strip() if curso else "Não encontrado"
    matricula_str = matricula.group(1).strip() if matricula else "Não encontrado"

    disciplinas_encontradas = 0

    for linha in todas_as_tabelas:
        if not linha:
            continue
            
        ord_str = str(linha[0] or '').strip()
        if not ord_str or not ord_str[0].isdigit():
            continue
            
        linhas_processadas = processar_linha_ancorada(linha)
        
        for item in linhas_processadas:
            excel_worksheet.append([
                nome_str, cpf_str, curso_str, matricula_str,
                item[1], item[2], item[3], item[4], item[5], item[6], item[7]
            ])
            disciplinas_encontradas += 1

    caminho_salvamento = rf"{pdf_folder}\..\relatorio_{data_atual}.xlsx"
    excel_file.save(caminho_salvamento)
    print(f"Salvo com sucesso! {disciplinas_encontradas} disciplinas processadas perfeitamente alinhadas.")

class PDFHandler(FileSystemEventHandler):
    def on_created(self, event):
        if event.is_directory:
            return

        if event.src_path.endswith('.pdf'):
            print(f"Novo PDF detectado: {event.src_path}")
            time.sleep(1)
            extract_info_and_write_to_excel(event.src_path)

event_handler = PDFHandler()
observer = Observer()
observer.schedule(event_handler, path=pdf_folder, recursive=False)
observer.start()

print(f"Monitorando a pasta: {pdf_folder}...")
print("Pressione Ctrl+C para encerrar.\n")

try:
    while True:
        time.sleep(1)
except KeyboardInterrupt:
    observer.stop()

observer.join()