import re
import os
import sys
import time
import pdfplumber
import openpyxl
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from threading import Thread

class LeitorPDFApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Extrator de Históricos Escolares")
        self.root.geometry("600x450")
        self.root.configure(padx=20, pady=20)

        # Variáveis de controle
        self.pasta_selecionada = tk.StringVar()
        self.status_var = tk.StringVar(value="Aguardando seleção de pasta...")

        self.setup_ui()

    def setup_ui(self):
        # Título
        tk.Label(self.root, text="Conversor de PDF para Excel", font=("Arial", 14, "bold")).pack(pady=(0, 20))

        # Seleção de Pasta
        frame_pasta = tk.Frame(self.root)
        frame_pasta.pack(fill="x", pady=10)
        
        tk.Entry(frame_pasta, textvariable=self.pasta_selecionada, state="readonly", width=50).pack(side="left", padx=(0, 10), expand=True, fill="x")
        tk.Button(frame_pasta, text="Selecionar Pasta", command=self.selecionar_pasta).pack(side="right")

        # Log de Atividades
        tk.Label(self.root, text="Log de Processamento:").pack(anchor="w")
        self.log_text = tk.Text(self.root, height=10, state="disabled", font=("Consolas", 9))
        self.log_text.pack(fill="both", expand=True, pady=5)

        # Barra de Progresso
        self.progress = ttk.Progressbar(self.root, orient="horizontal", mode="determinate")
        self.progress.pack(fill="x", pady=10)

        # Botão Iniciar
        self.btn_iniciar = tk.Button(self.root, text="GERAR RELATÓRIOS", bg="#4CAF50", fg="white", 
                                     font=("Arial", 10, "bold"), height=2, command=self.iniciar_processamento)
        self.btn_iniciar.pack(fill="x", pady=10)

        # Status
        tk.Label(self.root, textvariable=self.status_var, fg="gray").pack()

    def log(self, mensagem):
        self.log_text.config(state="normal")
        self.log_text.insert("end", f"[{time.strftime('%H:%M:%S')}] {mensagem}\n")
        self.log_text.see("end")
        self.log_text.config(state="disabled")
        self.root.update_idletasks()

    def selecionar_pasta(self):
        pasta = filedialog.askdirectory()
        if pasta:
            self.pasta_selecionada.set(pasta)
            self.log(f"Pasta selecionada: {pasta}")

    def iniciar_processamento(self):
        pasta = self.pasta_selecionada.get()
        if not pasta:
            messagebox.showwarning("Aviso", "Por favor, selecione uma pasta primeiro!")
            return

        # Rodar em uma thread separada para não travar a janela
        self.btn_iniciar.config(state="disabled")
        Thread(target=self.processar_arquivos, args=(pasta,), daemon=True).start()

    def processar_arquivos(self, pasta):
        arquivos = [f for f in os.listdir(pasta) if f.lower().endswith('.pdf')]
        if not arquivos:
            self.log("Nenhum PDF encontrado na pasta.")
            self.btn_iniciar.config(state="normal")
            return

        total = len(arquivos)
        self.progress["maximum"] = total
        sucessos = 0

        for i, nome_arq in enumerate(arquivos):
            caminho = os.path.join(pasta, nome_arq)
            self.status_var.set(f"Processando {i+1} de {total}...")
            
            resultado = self.extrair_dados(caminho, pasta)
            
            if resultado:
                self.log(f"Sucesso: {nome_arq}")
                sucessos += 1
            else:
                self.log(f"Erro ou Formato Inválido: {nome_arq}")
            
            self.progress["value"] = i + 1
            self.root.update_idletasks()

        self.status_var.set("Concluído!")
        self.btn_iniciar.config(state="normal")
        messagebox.showinfo("Finalizado", f"Processamento concluído!\n\nSucessos: {sucessos}\nTotal: {total}")

    # --- Lógica de Extração (Sua lógica original adaptada) ---

    def separar_coluna(self, texto, expected_count, is_componente=False):
        texto = str(texto or '').strip()
        if expected_count <= 1:
            return [texto.replace('\n', ' ').strip()] if not is_componente else [texto]
        blocos = re.split(r'\n{2,}', texto)
        blocos_limpos = [b.strip() for b in blocos if b.strip()]
        if len(blocos_limpos) == expected_count:
            return blocos_limpos if is_componente else [b.replace('\n', ' ') for b in blocos_limpos]
        simples = [b.strip() for b in texto.split('\n') if b.strip()]
        if len(simples) == expected_count:
            return simples
        while len(blocos_limpos) < expected_count:
            blocos_limpos.append("-")
        return blocos_limpos[:expected_count]

    def processar_linha_ancorada(self, linha):
        linha_limpa = [str(c).strip() for c in linha if c is not None and str(c).strip() != '']
        if len(linha_limpa) < 5 or "Componente" in str(linha[0]): return []
        
        ords = [x for x in re.split(r'\n+', linha_limpa[0]) if x.strip().isdigit()]
        num_discs = len(ords)
        if num_discs == 0: return []
        
        idx_carga = -1
        for i, celula in enumerate(linha_limpa):
            if re.search(r'\d+\s*[hH]', celula):
                idx_carga = i
                break
        if idx_carga == -1: return [] 

        semestres = self.separar_coluna(linha_limpa[1], num_discs)
        cargas = self.separar_coluna(linha_limpa[idx_carga], num_discs)
        media_raw = linha_limpa[idx_carga + 1] if idx_carga + 1 < len(linha_limpa) else "-"
        medias = self.separar_coluna(media_raw, num_discs)
        
        bloco_restante = " ".join(linha_limpa[idx_carga + 2:])
        padrao_sit = r'(Aprovado|Reprovado|Cursando|Trancado|Dispensado|Disp\.)'
        situacoes = [s.title() for s in re.findall(padrao_sit, bloco_restante, flags=re.IGNORECASE)]
        faltas = re.findall(r'\b\d+\b', re.sub(padrao_sit, ' ', bloco_restante, flags=re.IGNORECASE))
        
        while len(situacoes) < num_discs: situacoes.append("-")
        while len(faltas) < num_discs: faltas.append("-")

        meio = linha_limpa[2:idx_carga]
        disc_str, prof_str = (" ".join(meio[:-1]), meio[-1]) if len(meio) >= 2 else (meio[0] if meio else "-", "-")
        disciplinas = self.separar_coluna(disc_str, num_discs, is_componente=True)
        professores = self.separar_coluna(prof_str, num_discs, is_componente=True)

        return [[ords[i], semestres[i], disciplinas[i], professores[i], cargas[i], medias[i], faltas[i], situacoes[i]] for i in range(num_discs)]

    def calcular_medias_semestrais(self, worksheet, nome_aluno):
        agrupamento = {}
        for row in worksheet.iter_rows(min_row=2, values_only=True):
            sem_raw = str(row[4] or '').strip()
            nota_raw = str(row[8] or '').strip().replace(',', '.')
            match_sem = re.search(r'(\d{4}[\./]\d|\d+)', sem_raw)
            if match_sem:
                sem_id = match_sem.group(0).replace('/', '.')
                try:
                    nota = float(nota_raw)
                    if sem_id not in agrupamento: agrupamento[sem_id] = []
                    agrupamento[sem_id].append(nota)
                except: continue

        worksheet.append([])
        worksheet.append([f"APLICANTE: {nome_aluno}"])
        worksheet.append(["Semestre", "Soma Notas", "Qtd. Disciplinas Validadas", "Média Período"])
        
        periodos = sorted(agrupamento.keys())[:4]
        medias_finais = []
        for sem in periodos:
            if agrupamento[sem]:
                m = sum(agrupamento[sem]) / len(agrupamento[sem])
                medias_finais.append(m)
                worksheet.append([f"Período {sem}", sum(agrupamento[sem]), len(agrupamento[sem]), round(m, 2)])
        
        if medias_finais:
            worksheet.append([])
            worksheet.append(["MÉDIA GERAL (DOS 4 SEMESTRES):", "", "", round(sum(medias_finais)/len(medias_finais), 2)])

    def extrair_dados(self, caminho, pasta_destino):
        try:
            with pdfplumber.open(caminho) as pdf:
                text = "".join([p.extract_text() or "" for p in pdf.pages])
                tabelas = [tab for p in pdf.pages for tab in (p.extract_tables() or [])]

            if "Histórico" not in text: return False

            n_match = re.search(r'(?:Aluno\(a\):|Nome:)\s*(.*?)(?=\s*Nascimento:|\n|\r|$)', text, re.IGNORECASE)
            nome_aluno = n_match.group(1).split('\n')[0].strip() if n_match else "Desconhecido"
            
            cpf = (re.search(r'CPF:\s*([\d\.\-]+)', text) or re.search(r'(\d{3}\.\d{3}\.\d{3}-\d{2})', text))
            mat = re.search(r'Matrícula:\s*(\d+)', text)
            cur = re.search(r'Curso:\s*(.+)', text)

            c_val = cpf.group(1) if cpf else "N/A"
            m_val = mat.group(1) if mat else "N/A"
            cu_val = cur.group(1).split('Matrícula')[0].strip() if cur else "N/A"
            
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.append(["Nome", "CPF", "Curso", "Matrícula", "Semestre", "Disciplina", "Professor", "Carga Horária", "Média Final", "Faltas", "Situação"])

            count = 0
            for tab in tabelas:
                for lin in tab:
                    if lin and len(lin) > 0 and str(lin[0]).strip().replace('\n', '').isdigit():
                        for item in self.processar_linha_ancorada(lin):
                            ws.append([nome_aluno, c_val, cu_val, m_val] + item[1:])
                            count += 1

            if count > 0:
                self.calcular_medias_semestrais(ws, nome_aluno)
                
                # Limpa o nome do aluno para o arquivo
                nome_limpo = "".join(x for x in nome_aluno if x.isalnum() or x==' ').replace(' ', '_')
                
                # NOVO: Define o nome do arquivo incluindo a matrícula (m_val)
                nome_arquivo_final = f"relatorio_{m_val}_{nome_limpo}.xlsx"
                
                # Salva com o novo nome
                wb.save(os.path.join(pasta_destino, nome_arquivo_final))
                return True
            return False
        except Exception as e:
            print(f"Erro interno: {e}")
            return False

if __name__ == "__main__":
    root = tk.Tk()
    app = LeitorPDFApp(root)
    root.mainloop()