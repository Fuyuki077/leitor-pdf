import re
import os
import time
import logging
import pdfplumber
import openpyxl
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from threading import Thread

# Configurações de Cores e Estilo
COLOR_BG = "#F5F7F8"          # Cinza muito claro (fundo)
COLOR_CARD = "#FFFFFF"        # Branco (áreas de conteúdo)
COLOR_PRIMARY = "#2E7D32"      # Verde (botão principal)
COLOR_ACCENT = "#1976D2"       # Azul (seleção de pasta)
COLOR_TEXT = "#333333"         # Cinza escuro (texto)

class LeitorPDFApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Extrator de Históricos UERR")
        self.root.geometry("700x600")
        self.root.configure(bg=COLOR_BG)

        self.pasta_selecionada = tk.StringVar()
        self.status_var = tk.StringVar(value="Pronto para iniciar")
        
        self.setup_styles()
        self.setup_ui()

    def setup_styles(self):
        style = ttk.Style()
        style.theme_use('clam')
        
        # Estilização da Progressbar
        style.configure("TProgressbar", thickness=10, troughcolor=COLOR_BG, background=COLOR_PRIMARY)
        
    def setup_ui(self):
        # Container Principal com Padding
        main_container = tk.Frame(self.root, bg=COLOR_BG, padx=30, pady=30)
        main_container.pack(fill="both", expand=True)

        # --- CABEÇALHO ---
        header_frame = tk.Frame(main_container, bg=COLOR_BG)
        header_frame.pack(fill="x", pady=(0, 20))
        
        tk.Label(
            header_frame,
            text="Conversor de Histórico UERR",
            font=("Segoe UI", 18, "bold"),
            bg=COLOR_BG,
            fg=COLOR_PRIMARY
        ).pack(anchor="w")
        
        tk.Label(
            header_frame,
            text="Selecione a pasta contendo os arquivos PDF para extração automática.",
            font=("Segoe UI", 10),
            bg=COLOR_BG,
            fg="#333333"
        ).pack(anchor="w")

        # --- CARD DE SELEÇÃO ---
        selection_card = tk.Frame(main_container, bg=COLOR_CARD, padx=15, pady=15, highlightthickness=1, highlightbackground="#DDDDDD")
        selection_card.pack(fill="x", pady=10)

        tk.Label(selection_card, text="Pasta de Origem:", font=("Segoe UI", 9, "bold"), bg=COLOR_CARD).pack(anchor="w", pady=(0, 5))
        
        path_frame = tk.Frame(selection_card, bg=COLOR_CARD)
        path_frame.pack(fill="x")

        self.entry_path = tk.Entry(
            path_frame,
            textvariable=self.pasta_selecionada,
            font=("Segoe UI", 10),
            state="readonly",
            relief="flat",
            bg="#F0F0F0"
        )
        self.entry_path.pack(side="left", fill="x", expand=True, ipady=8, padx=(0, 10))

        tk.Button(
            path_frame,
            text="Selecione...",
            font=("Segoe UI", 9),
            bg=COLOR_ACCENT,
            fg="white",
            relief="flat",
            activebackground="#1565C0",
            activeforeground="white",
            cursor="hand2",
            command=self.selecionar_pasta,
            width=12
        ).pack(side="right", ipady=4)

        # --- ÁREA DE LOG ---
        log_frame = tk.Frame(main_container, bg=COLOR_BG)
        log_frame.pack(fill="both", expand=True, pady=15)
        
        tk.Label(log_frame, text="Log de Atividades:", font=("Segoe UI", 9, "bold"), bg=COLOR_BG).pack(anchor="w")
        
        self.log_text = tk.Text(
            log_frame, 
            height=10, 
            font=("Consolas", 9),
            relief="flat",
            bg="#2C3E50",  # Fundo escuro para o log
            fg="#ECF0F1",
            padx=10,
            pady=10,
            state="disabled"
        )
        self.log_text.pack(fill="both", expand=True, pady=5)

        # --- RODAPÉ (Progresso e Ação) ---
        self.progress = ttk.Progressbar(main_container, orient="horizontal", mode="determinate", style="TProgressbar",)
        self.progress.pack(fill="x", pady=(10, 5))
        
        self.lbl_status = tk.Label(main_container, textvariable=self.status_var, font=("Segoe UI", 8), bg=COLOR_BG, fg="#888888")
        self.lbl_status.pack(anchor="e")

        self.btn_iniciar = tk.Button(
            main_container,
            text="GERAR RELATÓRIOS EXCEL",
            bg=COLOR_PRIMARY,
            fg="white",
            font=("Segoe UI", 11, "bold"),
            relief="flat",
            activebackground="#1B5E20",
            activeforeground="white",
            cursor="hand2",
            height=2,
            command=self.iniciar_processamento
        )
        self.btn_iniciar.pack(fill="x", pady=(15, 0))

    # --- LÓGICA MANTIDA E MELHORADA ---

    def log(self, mensagem):
        self.log_text.config(state="normal")
        self.log_text.insert("end", f"› {time.strftime('%H:%M:%S')} | {mensagem}\n")
        self.log_text.see("end")
        self.log_text.config(state="disabled")

    def selecionar_pasta(self):
        pasta = filedialog.askdirectory()
        if pasta:
            self.pasta_selecionada.set(pasta)
            self.log(f"Pasta selecionada: {pasta}")

    def iniciar_processamento(self):
        pasta = self.pasta_selecionada.get()
        if not pasta:
            messagebox.showwarning("Aviso", "Por favor, selecione uma pasta primeiro.")
            return

        self.btn_iniciar.config(state="disabled", bg="#999999")
        self.status_var.set("Iniciando processamento...")
        Thread(target=self.processar_arquivos, args=(pasta,), daemon=True).start()

    # (Mantenha as funções de processamento como limpar_texto, normalizar_linha, etc., iguais à última versão corrigida)
    
    def processar_arquivos(self, pasta):
        arquivos = [f for f in os.listdir(pasta) if f.lower().endswith(".pdf")]
        if not arquivos:
            self.log("Nenhum arquivo PDF encontrado na pasta.")
            self.root.after(0, lambda: self.btn_iniciar.config(state="normal", bg=COLOR_PRIMARY))
            return

        self.progress["maximum"] = len(arquivos)
        sucessos = 0
        
        for i, nome_arq in enumerate(arquivos):
            caminho = os.path.join(pasta, nome_arq)
            self.root.after(0, lambda v=i + 1: self.status_var.set(f"Processando arquivo {v} de {len(arquivos)}"))
            
            if self.extrair_dados(caminho, pasta):
                self.log(f"Sucesso: {nome_arq}")
                sucessos += 1
            else:
                self.log(f"Falha: {nome_arq}")
                
            self.root.after(0, lambda v=i + 1: self.progress.configure(value=v))

        self.status_var.set("Processamento concluído")
        self.root.after(0, lambda: self.btn_iniciar.config(state="normal", bg=COLOR_PRIMARY))
        self.root.after(0, lambda: messagebox.showinfo("Concluído", f"Processamento finalizado!\nArquivos com sucesso: {sucessos}/{len(arquivos)}"))

    # ... (restante das funções: limpar_texto, eh_cabecalho, inicio_de_registro, mesclar_continuacao, 
    # separar_faltas_situacao, normalizar_linha, extrair_dados, calcular_medias_semestrais, quebrar_linha_colapsada)

    def limpar_texto(self, texto):
        if texto is None:
            return ""
        texto = str(texto)
        texto = re.sub(r'[\x00-\x08\x0b\x0c\x0e-\x1f]', '', texto)
        texto = texto.replace("\r", "\n")
        return texto.strip()

    def eh_cabecalho(self, celulas):
        texto = " ".join(celulas).lower()
        return (
            "ord" in texto
            and "semestre" in texto
            and "disciplina" in texto
            and "professor" in texto
        )

    def inicio_de_registro(self, celulas):
        if not celulas:
            return False
        texto = " ".join(celulas)
        return bool(re.match(r"^\d+\s+20\d{2}\.\d", texto))

    def mesclar_continuacao(self, atual, complemento):
        limite = max(len(atual), len(complemento))
        while len(atual) < limite:
            atual.append("")
        while len(complemento) < limite:
            complemento.append("")

        for i, valor in enumerate(complemento):
            valor = self.limpar_texto(valor)
            if not valor:
                continue
            if atual[i]:
                if valor not in atual[i]:
                    atual[i] = f"{atual[i]}\n{valor}"
            else:
                atual[i] = valor
        return atual

    def separar_faltas_situacao(self, texto):
        texto = self.limpar_texto(texto)
        if not texto:
            return "", ""
        normal = re.sub(r"\s+", " ", texto).strip()
        if "Aproveitada" in normal and "Reprovado" not in normal and "Aprovado" not in normal:
            return "", "Disciplina Aproveitada"
        if "Reprovado" in normal or "Aprovado" in normal:
            m = re.search(r"\d+", normal)
            faltas = m.group(0) if m else ""
            situacao = re.sub(r"\d+", "", normal).strip()
            situacao = re.sub(r"\s+", " ", situacao).strip()
            return faltas, situacao
        m = re.match(r"^(\d+)\s*(.*)$", normal)
        if m:
            return m.group(1), m.group(2).strip()
        return "", normal

    def normalizar_linha(self, celulas):
        # Limpa as células e remove vazios
        celulas = [self.limpar_texto(c) for c in celulas if self.limpar_texto(c) != ""]
        
        if not celulas:
            return [""] * 8  # Aumentado para 8 para alinhar com o Regex

        # Se a linha estiver colapsada em uma única string
        if len(celulas) == 1:
            tentativa = self.quebrar_linha_colapsada(celulas[0])
            if tentativa: 
                return tentativa  # Retorna os 8 grupos do Regex

        # Se houver mais colunas (PDFs com tabelas bem definidas), 
        # garantimos um mínimo de 8 colunas sem cortar o excesso
        while len(celulas) < 8:
            celulas.append("")
            
        return celulas # Removido o [:7] para não perder a última coluna

    def extrair_dados(self, caminho, pasta_destino):
        try:
            dados_aluno = {"nome": "N/A", "cpf": "N/A", "mat": "0000", "curso": "N/A"}
            linhas_brutas = []
            registro_atual = None
            fim_alcancado = False

            with pdfplumber.open(caminho) as pdf:
                for page in pdf.pages:
                    if fim_alcancado: break
                    texto = page.extract_text()
                    if texto:
                        if dados_aluno["nome"] == "N/A":
                            n = re.search(r"Aluno\(a\):\s*(.*?)(Nascimento|CPF|$)", texto, re.DOTALL)
                            if n: dados_aluno["nome"] = self.limpar_texto(n.group(1))
                            c = re.search(r"CPF:\s*([\d\.\-]+)", texto)
                            if c: dados_aluno["cpf"] = c.group(1).strip()
                            m = re.search(r"Matrícula:\s*(\d+)", texto)
                            if m: dados_aluno["mat"] = m.group(1).strip()
                            cur = re.search(r"Curso:\s*(.*?)(Matrícula|Grade|$)", texto, re.DOTALL)
                            if cur: dados_aluno["curso"] = self.limpar_texto(cur.group(1))

                    tabelas = page.extract_tables() or []
                    for tabela in tabelas:
                        if fim_alcancado: break
                        for linha in tabela:
                            celulas = [self.limpar_texto(c) for c in linha if c is not None]
                            if not celulas: continue
                            
                            # TRAVA DE SEGURANÇA: Atividades Complementares
                            if any("ATIVIDADES COMPLEMENTARES" in str(c).upper() for c in celulas):
                                fim_alcancado = True
                                break

                            if self.eh_cabecalho(celulas): continue

                            if self.inicio_de_registro(celulas):
                                if registro_atual is not None:
                                    linhas_brutas.append(registro_atual)
                                registro_atual = self.normalizar_linha(celulas)
                            else:
                                texto_l = " ".join(celulas)
                                if re.match(r"^\d+\s+20\d{2}\.\d", texto_l):
                                    if registro_atual is not None:
                                        linhas_brutas.append(registro_atual)
                                    registro_atual = self.normalizar_linha(celulas)
                                elif registro_atual is not None:
                                    registro_atual = self.mesclar_continuacao(registro_atual, self.normalizar_linha(celulas))

            if registro_atual is not None:
                linhas_brutas.append(registro_atual)

            if not linhas_brutas: return False

            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "Historico"
            ws.append(["Aluno", "CPF", "Curso", "Matrícula", "Ord", "Semestre", "Disciplina", "Professor", "CH", "Média", "Faltas", "Situação"])

            for d in linhas_brutas:
                # Verificamos se o dado já veio separado (pelo Regex) ou se precisa dividir
                # d[6] = Faltas, d[7] = Situação
                if len(d) >= 8 and d[7].strip() != "":
                    faltas = d[6]
                    situacao = d[7]
                else:
                    # Se d[7] estiver vazio, tentamos extrair ambos de d[6]
                    faltas, situacao = self.separar_faltas_situacao(d[6])

                ws.append([
                    dados_aluno["nome"], 
                    dados_aluno["cpf"], 
                    dados_aluno["curso"], 
                    dados_aluno["mat"], 
                    d[0], # Ord
                    d[1], # Semestre
                    d[2], # Disciplina
                    d[3], # Professor
                    d[4], # CH
                    d[5], # Média
                    faltas, 
                    situacao
                ])
            # CHAMADA DO NOVO MÉTODO DE MÉDIAS
            self.calcular_medias_semestrais(ws, dados_aluno["nome"])

            safe_nome = re.sub(r"[^\w]", "", dados_aluno["nome"])[:15] or "Aluno"
            nome_saida = f"Relatorio_{dados_aluno['mat']}_{safe_nome}.xlsx"
            wb.save(os.path.join(pasta_destino, nome_saida))
            return True

        except Exception as e:
            print(f"Erro interno no arquivo {caminho}: {e}")
            return False

    def calcular_medias_semestrais(self, worksheet, nome_aluno):
        agrupamento = {}
        # min_row=2 para pular o cabeçalho
        for row in worksheet.iter_rows(min_row=2, values_only=True):
            # Se a linha estiver vazia ou for o início da nova tabela, para a leitura
            if not row[0] or "APLICANTE" in str(row[0]): break
            
            # Ajustado para os índices reais da sua planilha: Semestre(5) e Média(9)
            sem_raw = str(row[5] or '').strip()
            nota_raw = str(row[9] or '').strip().replace(',', '.')
            
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
        
        # Pega os 4 primeiros períodos encontrados e sorteia
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

    def quebrar_linha_colapsada(self, linha_texto):
        linha_texto = re.sub(r'\s+', ' ', linha_texto).strip()
        padrao = re.match(r'^(\d+)\s+(20\d{2}\.\d)\s+(.+?)\s+(.+?)\s+(\d+h)\s+([\d,]+)\s+(\d+)\s+(.+)$', linha_texto)
        if padrao: return list(padrao.groups())
        return None

    def processar_arquivos(self, pasta):
        arquivos = [f for f in os.listdir(pasta) if f.lower().endswith(".pdf")]
        if not arquivos:
            self.log("Nenhum PDF encontrado.")
            self.root.after(0, lambda: self.btn_iniciar.config(state="normal"))
            return

        self.progress["maximum"] = len(arquivos)
        sucessos = 0
        for i, nome_arq in enumerate(arquivos):
            caminho = os.path.join(pasta, nome_arq)
            self.root.after(0, lambda v=i + 1: self.status_var.set(f"Processando {v}..."))
            if self.extrair_dados(caminho, pasta):
                self.log(f"OK: {nome_arq}")
                sucessos += 1
            else:
                self.log(f"ERRO: {nome_arq}")
            self.root.after(0, lambda v=i + 1: self.progress.configure(value=v))

        self.root.after(0, lambda: self.btn_iniciar.config(state="normal"))
        self.root.after(0, lambda: messagebox.showinfo("Fim", f"Processados: {sucessos}\nTotal: {len(arquivos)}"))


if __name__ == "__main__":
    root = tk.Tk()
    app = LeitorPDFApp(root)
    root.mainloop()