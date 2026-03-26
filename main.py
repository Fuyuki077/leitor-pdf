import re
import os
import time
import pdfplumber as pdfp
import openpyxl
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from threading import Thread

# ─── Paleta de Cores ───────────────────────────────────────────────────────────
COLOR_BG       = "#FFFFFF"
COLOR_CARD     = "#FFFFFF"
COLOR_BORDER   = "#D1D5DB"
COLOR_TEXT     = "#000000"
COLOR_MUTED    = "#D1D5DB"
COLOR_PRIMARY  = "#000000"
COLOR_ACCENT   = "#000000"
COLOR_SELECTED = "#292B35"
COLOR_PROGRESS = "#49CE45"


class LeitorPDFApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Extrator de Históricos UERR")
        self.root.geometry("700x600")
        self.root.configure(bg=COLOR_BG)

        self.pasta_selecionada = tk.StringVar()
        self.status_var        = tk.StringVar(value="Pronto para iniciar")

        self._setup_styles()
        self._setup_ui()

    # ─── Configuração de Estilos ───────────────────────────────────────────────

    def _setup_styles(self):
        style = ttk.Style()
        style.theme_use("clam")
        style.configure(
            "TProgressbar",
            thickness=10,
            troughcolor=COLOR_BG,
            background=COLOR_PROGRESS,
        )

    # ─── Construção da Interface ───────────────────────────────────────────────

    def _setup_ui(self):
        main = tk.Frame(self.root, bg=COLOR_BG)
        main.pack(fill="both", expand=True, padx=20, pady=20)

        # Cabeçalho
        hdr = tk.Frame(main, bg=COLOR_BG)
        hdr.pack(fill="x", pady=(0, 20))
        tk.Label(
            hdr,
            text="Conversor de Histórico UERR",
            font=("Segoe UI", 18, "bold"),
            bg=COLOR_BG,
            fg=COLOR_PRIMARY,
        ).pack(anchor="w")
        tk.Label(
            hdr,
            text="Selecione a pasta contendo os arquivos PDF para extração automática.",
            font=("Segoe UI", 10),
            bg=COLOR_BG,
            fg="#242424",
        ).pack(anchor="w")

        # Seleção de pasta
        sel = tk.Frame(main, bg=COLOR_CARD, padx=15, pady=15)
        sel.pack(fill="x", pady=10)
        tk.Label(
            sel,
            text="Pasta de Origem:",
            font=("Segoe UI", 9, "bold"),
            bg=COLOR_CARD,
        ).pack(anchor="w", pady=(0, 5))

        path_row = tk.Frame(sel, bg=COLOR_CARD)
        path_row.pack(fill="x")
        tk.Entry(
            path_row,
            textvariable=self.pasta_selecionada,
            font=("Segoe UI", 10),
            state="readonly",
            relief="flat",
            bg="#F0F0F0",
        ).pack(side="left", fill="x", expand=True, ipady=8, padx=(0, 10))
        tk.Button(
            path_row,
            text="Selecione...",
            font=("Segoe UI", 9),
            bg=COLOR_ACCENT,
            fg="white",
            relief="flat",
            activebackground=COLOR_SELECTED,
            activeforeground="white",
            cursor="hand2",
            command=self._selecionar_pasta,
            width=12,
        ).pack(side="right", ipady=4)

        # Log
        log_frame = tk.Frame(main, bg=COLOR_BG)
        log_frame.pack(fill="both", expand=True, pady=15)
        tk.Label(
            log_frame,
            text="Log de Atividades:",
            font=("Segoe UI", 9, "bold"),
            bg=COLOR_BG,
        ).pack(anchor="w")
        self.log_text = tk.Text(
            log_frame,
            height=10,
            font=("Consolas", 9),
            relief="flat",
            bg="#121920",
            fg="#ECF0F1",
            padx=10,
            pady=10,
            state="disabled",
        )
        self.log_text.pack(fill="both", expand=True, pady=5)

        # Barra de progresso e status
        self.progress = ttk.Progressbar(
            main, orient="horizontal", mode="determinate", style="TProgressbar"
        )
        self.progress.pack(fill="x", pady=(10, 5))
        tk.Label(
            main,
            textvariable=self.status_var,
            font=("Segoe UI", 8),
            bg=COLOR_BG,
            fg="#888888",
        ).pack(anchor="e")

        # Botão principal
        self.btn_iniciar = tk.Button(
            main,
            text="GERAR RELATÓRIOS EXCEL",
            bg=COLOR_PRIMARY,
            fg="white",
            font=("Segoe UI", 11, "bold"),
            relief="flat",
            activebackground=COLOR_SELECTED,
            activeforeground="white",
            cursor="hand2",
            height=2,
            command=self._iniciar_processamento,
        )
        self.btn_iniciar.pack(fill="x", pady=(15, 0))

    # ─── Helpers de UI (thread-safe) ──────────────────────────────────────────

    def _log(self, mensagem: str):
        """Adiciona uma linha ao log. Pode ser chamado de qualquer thread."""
        def _append():
            self.log_text.config(state="normal")
            self.log_text.insert("end", f"› {time.strftime('%H:%M:%S')} | {mensagem}\n")
            self.log_text.see("end")
            self.log_text.config(state="disabled")

        self.root.after(0, _append)

    def _set_status(self, texto: str):
        self.root.after(0, lambda: self.status_var.set(texto))

    def _set_progress(self, valor: int):
        self.root.after(0, lambda: self.progress.configure(value=valor))

    def _set_btn(self, ativo: bool):
        cor = COLOR_PRIMARY if ativo else "#999999"
        estado = "normal" if ativo else "disabled"
        self.root.after(0, lambda: self.btn_iniciar.config(state=estado, bg=cor))

    # ─── Ações da Interface ────────────────────────────────────────────────────

    def _selecionar_pasta(self):
        pasta = filedialog.askdirectory()
        if pasta:
            self.pasta_selecionada.set(pasta)
            self._log(f"Pasta selecionada: {pasta}")

    def _iniciar_processamento(self):
        pasta = self.pasta_selecionada.get()
        if not pasta:
            messagebox.showwarning("Aviso", "Por favor, selecione uma pasta primeiro.")
            return
        self._set_btn(False)
        self._set_status("Iniciando processamento...")
        Thread(target=self._processar_arquivos, args=(pasta,), daemon=True).start()

    # ─── Processamento Principal ───────────────────────────────────────────────

    def _processar_arquivos(self, pasta: str):
        arquivos = [f for f in os.listdir(pasta) if f.lower().endswith(".pdf")]

        if not arquivos:
            self._log("Nenhum arquivo PDF encontrado na pasta.")
            self._set_btn(True)
            return

        self.root.after(0, lambda: self.progress.config(maximum=len(arquivos)))
        sucessos = 0

        for i, nome_arq in enumerate(arquivos):
            caminho = os.path.join(pasta, nome_arq)
            self._set_status(f"Processando {i + 1}/{len(arquivos)}: {nome_arq}")

            if self._extrair_dados(caminho, pasta):
                self._log(f"OK: {nome_arq}")
                sucessos += 1
            else:
                self._log(f"ERRO: {nome_arq}")

            self._set_progress(i + 1)

        total = len(arquivos)
        self._set_status("Processamento concluído")
        self._set_btn(True)
        self.root.after(
            0,
            lambda: messagebox.showinfo(
                "Concluído",
                f"Processamento finalizado!\nSucessos: {sucessos}/{total}",
            ),
        )

    # ─── Funções Auxiliares de Parsing ────────────────────────────────────────

    def _limpar_texto(self, texto) -> str:
        """Remove caracteres de controle e normaliza espaços."""
        if texto is None:
            return ""
        texto = str(texto)
        texto = re.sub(r"[\x00-\x08\x0b\x0c\x0e-\x1f]", "", texto)
        texto = texto.replace("\r", "\n")
        return texto.strip()

    def _eh_cabecalho(self, celulas: list) -> bool:
        """Retorna True se a linha parece ser o cabeçalho da tabela.
        Exige ao menos 2 das 4 palavras-chave para tolerar cabeçalhos
        parcialmente colapsados.
        """
        texto = " ".join(celulas).lower()
        palavras_chave = ["ord", "semestre", "disciplina", "professor"]
        return sum(p in texto for p in palavras_chave) >= 2

    def _inicio_de_registro(self, celulas: list) -> bool:
        """Retorna True se a linha inicia com número de ordem + ano/semestre."""
        if not celulas:
            return False
        texto = " ".join(celulas)
        return bool(re.match(r"^\d+\s+20\d{2}\.\d", texto))

    def _mesclar_continuacao(self, atual: list, complemento: list) -> list:
        """Mescla uma linha de continuação (quebra de página) no registro atual."""
        limite = max(len(atual), len(complemento))
        while len(atual) < limite:
            atual.append("")
        while len(complemento) < limite:
            complemento.append("")

        for i, valor in enumerate(complemento):
            valor = self._limpar_texto(valor)
            if not valor:
                continue
            if atual[i]:
                if valor not in atual[i]:
                    atual[i] = f"{atual[i]}\n{valor}"
            else:
                atual[i] = valor

        return atual

    def _separar_faltas_situacao(self, texto: str):
        """Separa o campo colapsado de faltas+situação em dois valores distintos.
        Retorna (faltas, situacao) como strings.
        """
        texto = self._limpar_texto(texto)
        if not texto:
            return "", ""

        normal = re.sub(r"\s+", " ", texto).strip()

        if "Aproveitada" in normal and "Reprovado" not in normal and "Aprovado" not in normal:
            return "", "Disciplina Aproveitada"

        if "Reprovado" in normal or "Aprovado" in normal:
            m = re.search(r"\d+", normal)
            faltas   = m.group(0) if m else ""
            situacao = re.sub(r"\d+", "", normal).strip()
            situacao = re.sub(r"\s+", " ", situacao).strip()
            return faltas, situacao

        m = re.match(r"^(\d+)\s*(.*)$", normal)
        if m:
            return m.group(1), m.group(2).strip()

        return "", normal

    def _normalizar_linha(self, celulas: list) -> list:
        """Garante sempre uma lista de exatamente 8 colunas.
        Tenta quebrar linhas colapsadas em uma única string.
        """
        celulas = [self._limpar_texto(c) for c in celulas if self._limpar_texto(c) != ""]

        if not celulas:
            return [""] * 8

        if len(celulas) == 1:
            tentativa = self._quebrar_linha_colapsada(celulas[0])
            if tentativa:
                return tentativa  # já tem 8 grupos

        while len(celulas) < 8:
            celulas.append("")

        # Garante tamanho fixo — evita IndexError ao acessar d[0]..d[7]
        return celulas[:8]

    def _quebrar_linha_colapsada(self, linha_texto: str):
        """Tenta extrair 8 campos de uma linha completamente colapsada num único texto."""
        linha_texto = re.sub(r"\s+", " ", linha_texto).strip()
        padrao = re.match(
            r"^(\d+)\s+(20\d{2}\.\d)\s+(.+?)\s+(.+?)\s+(\d+h)\s+([\d,]+)\s+(\d+)\s+(.+)$",
            linha_texto,
        )
        if padrao:
            return list(padrao.groups())
        return None

    # ─── Extração de Dados do PDF ──────────────────────────────────────────────

    def _extrair_dados(self, caminho: str, pasta_destino: str) -> bool:
        """Extrai dados acadêmicos de um PDF e salva um .xlsx na pasta de destino.
        Retorna True em caso de sucesso, False em caso de erro.
        """
        try:
            dados_aluno   = {"nome": "N/A", "cpf": "N/A", "mat": "0000", "curso": "N/A"}
            linhas_brutas = []
            registro_atual = None
            fim_alcancado  = False

            with pdfp.open(caminho) as pdf:
                for page in pdf.pages:
                    if fim_alcancado:
                        break

                    texto_completo = page.extract_text()
                    if not texto_completo:
                        continue

                    # Extração de metadados (apenas na primeira ocorrência)
                    if dados_aluno["nome"] == "N/A":
                        n = re.search(
                            r"Aluno\(a\):\s*(.*?)(Nascimento|CPF|$)",
                            texto_completo,
                            re.DOTALL,
                        )
                        if n:
                            dados_aluno["nome"] = self._limpar_texto(n.group(1))

                        c = re.search(r"CPF:\s*([\d.\-]+)", texto_completo)
                        if c:
                            dados_aluno["cpf"] = c.group(1).strip()

                        m = re.search(r"Matrícula:\s*(\d+)", texto_completo)
                        if m:
                            dados_aluno["mat"] = m.group(1).strip()

                        cur = re.search(
                            r"Curso:\s*(.*?)(Matrícula|Grade|$)",
                            texto_completo,
                            re.DOTALL,
                        )
                        if cur:
                            dados_aluno["curso"] = self._limpar_texto(cur.group(1))

                    # Processamento de tabelas
                    tabelas = page.extract_tables()
                    for tabela in tabelas:
                        for linha in tabela:
                            celulas = [self._limpar_texto(c) for c in linha if c is not None]
                            conteudo_real = [c for c in celulas if c != ""]
                            if not conteudo_real:
                                continue

                            texto_linha = " ".join(conteudo_real).upper()

                            # Trava de segurança — fim do histórico acadêmico
                            if any(
                                x in texto_linha
                                for x in ("ATIVIDADES COMPLEMENTARES", "CARGA HORÁRIA TOTAL")
                            ):
                                fim_alcancado = True
                                break

                            if self._eh_cabecalho(celulas):
                                continue

                            # Lógica de continuidade multipáginas
                            if self._inicio_de_registro(celulas):
                                if registro_atual is not None:
                                    linhas_brutas.append(registro_atual)
                                registro_atual = self._normalizar_linha(celulas)

                            elif registro_atual is not None:
                                # Ignora ruídos comuns de rodapé/cabeçalho de página
                                if any(
                                    x in texto_linha
                                    for x in ("PÁGINA", "UERR", "7 DE SETEMBRO", "EMISSÃO")
                                ):
                                    continue
                                registro_atual = self._mesclar_continuacao(
                                    registro_atual, self._normalizar_linha(celulas)
                                )

            # Adiciona o último registro pendente
            if registro_atual is not None:
                linhas_brutas.append(registro_atual)

            if not linhas_brutas:
                self._log(f"Nenhum dado encontrado em: {os.path.basename(caminho)}")
                return False

            # ── Geração do Excel ──────────────────────────────────────────────
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "Historico"
            ws.append([
                "Aluno", "CPF", "Curso", "Matrícula",
                "Ord", "Semestre", "Disciplina", "Professor",
                "CH", "Média", "Faltas", "Situação",
            ])

            for d in linhas_brutas:
                # Determina faltas e situação
                if len(d) >= 8 and d[7].strip():
                    faltas, situacao = d[6], d[7]
                else:
                    faltas, situacao = self._separar_faltas_situacao(d[6])

                try:
                    val_ord    = int(d[0]) if str(d[0]).isdigit() else d[0]
                    media_str  = str(d[5]).replace(",", ".").strip()
                    val_media  = float(media_str) if media_str and media_str != "None" else 0.0
                    val_faltas = int(faltas) if str(faltas).isdigit() else 0
                except Exception:
                    val_ord, val_media, val_faltas = d[0], 0.0, 0

                ws.append([
                    dados_aluno["nome"],
                    dados_aluno["cpf"],
                    dados_aluno["curso"],
                    dados_aluno["mat"],
                    val_ord,
                    d[1],   # semestre
                    d[2],   # disciplina
                    d[3],   # professor
                    d[4],   # CH
                    val_media,
                    val_faltas,
                    situacao,
                ])

            self._calcular_medias_semestrais(ws, dados_aluno["nome"])

            nome_saida = f"Relatorio_{dados_aluno['mat']}_{{{dados_aluno['nome']}}}.xlsx"
            wb.save(os.path.join(pasta_destino, nome_saida))
            return True

        except Exception as e:
            import traceback
            self._log(f"Erro crítico em {os.path.basename(caminho)}: {e}")
            self._log(traceback.format_exc())
            return False

    # ─── Cálculo de Médias Semestrais ─────────────────────────────────────────

    def _calcular_medias_semestrais(self, worksheet, nome_aluno: str):
        """Adiciona uma seção de resumo ao final da planilha com médias.
        
        Limitado aos 4 primeiros semestres cronológicos encontrados.
        """
        agrupamento: dict[str, list[float]] = {}

        for row in worksheet.iter_rows(min_row=2, values_only=True):
            # Interrompe se a linha estiver vazia ou já estiver no resumo antigo
            if not row or row[0] is None or "APLICANTE" in str(row[0]):
                break

            sem_raw  = str(row[5] or "").strip()
            # O índice 9 corresponde à coluna "Média" no seu código
            nota_raw = str(row[9] if row[9] is not None else "").strip().replace(",", ".")

            match_sem = re.search(r"(\d{4}[./]\d|\d+)", sem_raw)
            if not match_sem:
                continue

            sem_id = match_sem.group(0).replace("/", ".")
            try:
                nota = float(nota_raw)
                agrupamento.setdefault(sem_id, []).append(nota)
            except ValueError:
                continue

        worksheet.append([])
        worksheet.append([f"APLICANTE: {nome_aluno}"])
        worksheet.append(["Semestre", "Soma Notas", "Qtd. Disciplinas", "Média Período"])

        # Ordena os períodos e seleciona apenas os 4 primeiros
        periodos_ordenados = sorted(agrupamento.keys())
        periodos_limitados = periodos_ordenados[:4]  # <-- O limite de 4 semestres
        
        medias_finais = []

        for sem in periodos_limitados:
            notas = agrupamento[sem]
            if notas:
                media = sum(notas) / len(notas)
                medias_finais.append(media)
                worksheet.append([
                    f"Período {sem}",
                    round(sum(notas), 2),
                    len(notas),
                    round(media, 2),
                ])

        if medias_finais:
            worksheet.append([])
            worksheet.append([
                "MÉDIA GERAL (4 SEMESTRES):",
                "",
                "",
                round(sum(medias_finais) / len(medias_finais), 2),
            ])

# ─── Ponto de Entrada ──────────────────────────────────────────────────────────

if __name__ == "__main__":
    root = tk.Tk()
    app  = LeitorPDFApp(root)
    root.mainloop()