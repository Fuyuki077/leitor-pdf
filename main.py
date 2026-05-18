import os
import time
import logging
import traceback
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from threading import Thread, Event

from pdf_parser import PDFTranscriptParser, ExtractionError
from excel_writer import ExcelReportGenerator, FiltroSemestres

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
        self.root.geometry("700x760")
        self.root.configure(bg=COLOR_BG)

        self.pasta_selecionada     = tk.StringVar()
        self.status_var            = tk.StringVar(value="Pronto para iniciar")
        self.modo_semestres        = tk.StringVar(value="todos")
        self.qtd_semestres         = tk.IntVar(value=4)
        self.semestres_especificos = tk.StringVar()

        self._processando = False
        self._cancelar    = Event()
        self._logger      = self._setup_logging()

        self.root.protocol("WM_DELETE_WINDOW", self._ao_fechar)
        self._setup_styles()
        self._setup_ui()

    # ─── Logging em Arquivo ───────────────────────────────────────────────────

    def _setup_logging(self) -> logging.Logger:
        log_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "logs")
        os.makedirs(log_dir, exist_ok=True)
        log_file = os.path.join(log_dir, f"sessao_{time.strftime('%Y%m%d_%H%M%S')}.log")
        logger = logging.getLogger("leitor_pdf")
        logger.setLevel(logging.INFO)
        handler = logging.FileHandler(log_file, encoding="utf-8")
        handler.setFormatter(logging.Formatter("%(asctime)s | %(message)s", datefmt="%H:%M:%S"))
        logger.addHandler(handler)
        return logger

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

        # Configuração de Semestres no Resumo
        sem_card = tk.Frame(main, bg=COLOR_CARD, padx=15, pady=12)
        sem_card.pack(fill="x", pady=(0, 10))
        tk.Label(
            sem_card,
            text="Semestres no Resumo de Médias:",
            font=("Segoe UI", 9, "bold"),
            bg=COLOR_CARD,
        ).pack(anchor="w", pady=(0, 8))

        tk.Radiobutton(
            sem_card,
            text="Todos os semestres",
            variable=self.modo_semestres,
            value="todos",
            bg=COLOR_CARD,
            font=("Segoe UI", 9),
            activebackground=COLOR_CARD,
            command=self._atualizar_opcoes_semestres,
        ).pack(anchor="w")

        row_n = tk.Frame(sem_card, bg=COLOR_CARD)
        row_n.pack(anchor="w", pady=3)
        tk.Radiobutton(
            row_n,
            text="Primeiros",
            variable=self.modo_semestres,
            value="primeiros_n",
            bg=COLOR_CARD,
            font=("Segoe UI", 9),
            activebackground=COLOR_CARD,
            command=self._atualizar_opcoes_semestres,
        ).pack(side="left")
        self.spin_qtd = tk.Spinbox(
            row_n,
            from_=1,
            to=20,
            textvariable=self.qtd_semestres,
            width=4,
            font=("Segoe UI", 9),
            state="disabled",
            relief="flat",
            bg="#F0F0F0",
        )
        self.spin_qtd.pack(side="left", padx=6)
        tk.Label(
            row_n, text="semestres", font=("Segoe UI", 9), bg=COLOR_CARD
        ).pack(side="left")

        row_esp = tk.Frame(sem_card, bg=COLOR_CARD)
        row_esp.pack(anchor="w", pady=3)
        tk.Radiobutton(
            row_esp,
            text="Semestres específicos:",
            variable=self.modo_semestres,
            value="especificos",
            bg=COLOR_CARD,
            font=("Segoe UI", 9),
            activebackground=COLOR_CARD,
            command=self._atualizar_opcoes_semestres,
        ).pack(side="left")
        self.entry_semestres = tk.Entry(
            row_esp,
            textvariable=self.semestres_especificos,
            font=("Segoe UI", 9),
            width=22,
            state="disabled",
            relief="flat",
            bg="#F0F0F0",
            disabledbackground="#E8E8E8",
        )
        self.entry_semestres.pack(side="left", padx=6)
        tk.Label(
            row_esp,
            text="ex: 2022.1, 2023.2",
            font=("Segoe UI", 8),
            fg="#888888",
            bg=COLOR_CARD,
        ).pack(side="left")

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
            height=7,
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

    def _log(self, mensagem: str) -> None:
        self._logger.info(mensagem)

        def _append():
            self.log_text.config(state="normal")
            self.log_text.insert("end", f"› {time.strftime('%H:%M:%S')} | {mensagem}\n")
            self.log_text.see("end")
            self.log_text.config(state="disabled")

        self.root.after(0, _append)

    def _set_status(self, texto: str) -> None:
        self.root.after(0, lambda: self.status_var.set(texto))

    def _set_progress(self, valor: int) -> None:
        self.root.after(0, lambda: self.progress.configure(value=valor))

    def _set_btn(self, ativo: bool) -> None:
        cor = COLOR_PRIMARY if ativo else "#999999"
        estado = "normal" if ativo else "disabled"
        self.root.after(0, lambda: self.btn_iniciar.config(state=estado, bg=cor))

    # ─── Shutdown Gracioso ────────────────────────────────────────────────────

    def _ao_fechar(self) -> None:
        if self._processando:
            if messagebox.askokcancel(
                "Fechar",
                "Processamento em andamento.\nDeseja cancelar e fechar o programa?",
            ):
                self._cancelar.set()
                self._aguardar_encerramento()
        else:
            self.root.destroy()

    def _aguardar_encerramento(self) -> None:
        if self._processando:
            self.root.after(100, self._aguardar_encerramento)
        else:
            self.root.destroy()

    # ─── Ações da Interface ────────────────────────────────────────────────────

    def _selecionar_pasta(self) -> None:
        pasta = filedialog.askdirectory()
        if pasta:
            self.pasta_selecionada.set(pasta)
            self._log(f"Pasta selecionada: {pasta}")

    def _atualizar_opcoes_semestres(self) -> None:
        modo = self.modo_semestres.get()
        self.spin_qtd.config(state="normal" if modo == "primeiros_n" else "disabled")
        self.entry_semestres.config(state="normal" if modo == "especificos" else "disabled")

    def _get_filtro(self) -> FiltroSemestres:
        modo = self.modo_semestres.get()
        especificos = []
        if modo == "especificos":
            texto = self.semestres_especificos.get()
            especificos = [s.strip() for s in texto.split(",") if s.strip()]
        return FiltroSemestres(
            modo=modo,
            quantidade=self.qtd_semestres.get(),
            especificos=especificos,
        )

    def _iniciar_processamento(self) -> None:
        pasta = self.pasta_selecionada.get()
        if not pasta:
            messagebox.showwarning("Aviso", "Por favor, selecione uma pasta primeiro.")
            return
        self._processando = True
        self._cancelar.clear()
        self._set_btn(False)
        self._set_status("Iniciando processamento...")
        Thread(target=self._processar_arquivos, args=(pasta,), daemon=True).start()

    # ─── Processamento Principal ───────────────────────────────────────────────

    def _processar_arquivos(self, pasta: str) -> None:
        arquivos = [f for f in os.listdir(pasta) if f.lower().endswith(".pdf")]

        if not arquivos:
            self._log("Nenhum arquivo PDF encontrado na pasta.")
            self._processando = False
            self._set_btn(True)
            return

        self.root.after(0, lambda: self.progress.config(maximum=len(arquivos)))

        parser   = PDFTranscriptParser()
        writer   = ExcelReportGenerator()
        filtro   = self._get_filtro()
        sucessos = 0

        for i, nome_arq in enumerate(arquivos):
            if self._cancelar.is_set():
                self._log("Processamento cancelado pelo usuário.")
                break

            caminho = os.path.join(pasta, nome_arq)
            self._set_status(f"Processando {i + 1}/{len(arquivos)}: {nome_arq}")

            try:
                dados_aluno, linhas_brutas = parser.extrair(caminho)
                if not linhas_brutas:
                    self._log(f"AVISO: nenhum dado encontrado em {nome_arq}")
                else:
                    writer.gerar(dados_aluno, linhas_brutas, pasta, filtro)
                    self._log(f"OK: {nome_arq}")
                    sucessos += 1
            except ExtractionError as e:
                self._log(f"ERRO ao ler PDF '{nome_arq}': {e}")
            except PermissionError as e:
                self._log(f"ERRO de permissão em '{nome_arq}': {e}")
            except Exception as e:
                self._log(f"ERRO inesperado em '{nome_arq}': {e}")
                self._log(traceback.format_exc())

            self._set_progress(i + 1)

        self._processando = False
        total = len(arquivos)

        if self._cancelar.is_set():
            self._set_status("Cancelado")
        else:
            self._set_status("Processamento concluído")
            self.root.after(
                0,
                lambda: messagebox.showinfo(
                    "Concluído",
                    f"Processamento finalizado!\nSucessos: {sucessos}/{total}",
                ),
            )

        self._set_btn(True)


# ─── Ponto de Entrada ──────────────────────────────────────────────────────────

if __name__ == "__main__":
    root = tk.Tk()
    app  = LeitorPDFApp(root)
    root.mainloop()
