import re
import pdfplumber as pdfp

# Prefixos de título acadêmico usados para separar disciplina de professor
# em linhas colapsadas por quebra de página.
_TITULOS_PROFESSOR = re.compile(
    r"(?<!\w)(Dr\.|Dra\.|Msc\.|MSc\.|M\.Sc\.|Prof\.|Profa\.|Esp\.|Me\.|Ma\.)",
    re.IGNORECASE,
)


class ExtractionError(Exception):
    """Falha durante a extração de dados do PDF."""


class PDFTranscriptParser:
    """Extrai dados de histórico acadêmico UERR de um arquivo PDF."""

    def extrair(self, caminho: str) -> tuple[dict, list]:
        """Retorna (dados_aluno, linhas_brutas). Lança ExtractionError em caso de falha."""
        try:
            return self._processar(caminho)
        except ExtractionError:
            raise
        except FileNotFoundError as e:
            raise ExtractionError(f"Arquivo não encontrado: {e}") from e
        except PermissionError as e:
            raise ExtractionError(f"Sem permissão para ler o arquivo: {e}") from e
        except Exception as e:
            raise ExtractionError(f"Falha ao processar PDF: {e}") from e

    def _processar(self, caminho: str) -> tuple[dict, list]:
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

                if dados_aluno["nome"] == "N/A":
                    self._extrair_metadados(texto_completo, dados_aluno)

                tabelas = page.extract_tables()
                for tabela in tabelas:
                    for linha in tabela:
                        celulas = [self._limpar_texto(c) for c in linha if c is not None]
                        conteudo_real = [c for c in celulas if c != ""]
                        if not conteudo_real:
                            continue

                        texto_linha = " ".join(conteudo_real).upper()

                        if any(x in texto_linha for x in ("ATIVIDADES COMPLEMENTARES", "CARGA HORÁRIA TOTAL")):
                            fim_alcancado = True
                            break

                        if self._eh_cabecalho(celulas):
                            continue

                        if self._inicio_de_registro(celulas):
                            if registro_atual is not None:
                                linhas_brutas.append(registro_atual)
                            registro_atual = self._normalizar_linha(celulas)
                        elif registro_atual is not None:
                            if any(x in texto_linha for x in ("PÁGINA", "UERR", "7 DE SETEMBRO", "EMISSÃO")):
                                continue
                            registro_atual = self._mesclar_continuacao(
                                registro_atual, self._normalizar_continuacao(linha)
                            )

        if registro_atual is not None:
            linhas_brutas.append(registro_atual)

        return dados_aluno, linhas_brutas

    def _extrair_metadados(self, texto: str, dados: dict) -> None:
        n = re.search(r"Aluno\(a\):\s*(.*?)(Nascimento|CPF|$)", texto, re.DOTALL)
        if n:
            dados["nome"] = self._limpar_texto(n.group(1))

        c = re.search(r"CPF:\s*([\d.\-]+)", texto)
        if c:
            dados["cpf"] = c.group(1).strip()

        m = re.search(r"Matrícula:\s*(\d+)", texto)
        if m:
            dados["mat"] = m.group(1).strip()

        cur = re.search(r"Curso:\s*(.*?)(Matrícula|Grade|$)", texto, re.DOTALL)
        if cur:
            dados["curso"] = self._limpar_texto(cur.group(1))

    def _limpar_texto(self, texto) -> str:
        if texto is None:
            return ""
        texto = str(texto)
        texto = re.sub(r"[\x00-\x08\x0b\x0c\x0e-\x1f]", "", texto)
        texto = texto.replace("\r", "\n")
        return texto.strip()

    def _eh_cabecalho(self, celulas: list) -> bool:
        texto = " ".join(celulas).lower()
        palavras_chave = ["ord", "semestre", "disciplina", "professor"]
        return sum(p in texto for p in palavras_chave) >= 2

    def _inicio_de_registro(self, celulas: list) -> bool:
        if not celulas:
            return False
        texto = " ".join(celulas)
        return bool(re.match(r"^\d+\s+20\d{2}\.\d", texto))

    def _mesclar_continuacao(self, atual: list, complemento: list) -> list:
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

    def _separar_faltas_situacao(self, texto: str) -> tuple[str, str]:
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
        celulas = [self._limpar_texto(c) for c in celulas if self._limpar_texto(c) != ""]

        if not celulas:
            return [""] * 8

        if len(celulas) == 1:
            tentativa = self._quebrar_linha_colapsada(celulas[0])
            if tentativa:
                return tentativa

        while len(celulas) < 8:
            celulas.append("")

        return celulas[:8]

    def _normalizar_continuacao(self, celulas: list) -> list:
        """Normaliza linha de continuação preservando posições de coluna do pdfplumber.

        Diferente de _normalizar_linha, NÃO remove células vazias do início —
        isso garante que dados no col 3 (Professor) permaneçam no col 3 após o merge,
        mesmo que os cols 0-2 estejam vazios.
        """
        resultado = [self._limpar_texto(c) for c in celulas]
        while len(resultado) < 8:
            resultado.append("")
        return resultado[:8]

    def _quebrar_linha_colapsada(self, linha_texto: str) -> list | None:
        """Extrai 8 campos de uma linha colapsada numa única string.

        Usa âncoras estruturais confiáveis (Ord, Semestre, CH, Média, Faltas)
        em vez de um único regex frágil. Suporta texto multi-linha causado por
        quebra de página — os \\n são achatados antes da análise.
        """
        flat = re.sub(r"\s+", " ", linha_texto).strip()

        m_inicio = re.match(r"^(\d+)\s+(20\d{2}\.\d)\s+", flat)
        if not m_inicio:
            return None

        ord_val = m_inicio.group(1)
        sem_val = m_inicio.group(2)
        resto   = flat[m_inicio.end():]

        # CH é o primeiro token "NúmeroH" (ex: 60h, 120h)
        m_ch = re.search(r"\b(\d+h)\b", resto, re.IGNORECASE)
        if not m_ch:
            return None

        disc_prof = resto[: m_ch.start()].strip()
        ch_val    = m_ch.group(1)
        apos_ch   = resto[m_ch.end() :].strip()

        # Após CH espera-se: Média Faltas Situação(+lixo de quebra)
        m_mf = re.match(r"([\d,]+)\s+(\d+)\s+(.*)", apos_ch, re.DOTALL)
        if not m_mf:
            return None

        media_val = m_mf.group(1)
        faltas_val = m_mf.group(2)
        situacao   = self._extrair_situacao_de_texto(m_mf.group(3))

        disc, prof = self._dividir_disc_prof(disc_prof)

        return [ord_val, sem_val, disc, prof, ch_val, media_val, faltas_val, situacao]

    def _extrair_situacao_de_texto(self, texto: str) -> str:
        """Extrai situação limpa de um texto que pode conter lixo de quebra de página."""
        t = re.sub(r"\s+", " ", texto).lower()
        # Testa do mais específico para o mais genérico
        if "reprovado" in t and ("falta" in t or "frequência" in t):
            return "Reprovado por Falta"
        if "reprovado" in t:
            return "Reprovado"
        if "aproveitada" in t:
            return "Disciplina Aproveitada"
        if "aprovado" in t:
            return "Aprovado"
        # Fallback: primeira palavra antes do lixo de quebra
        return texto.strip().split()[0] if texto.strip() else ""

    def _dividir_disc_prof(self, texto: str) -> tuple[str, str]:
        """Divide 'Disciplina Professor' usando prefixos de título acadêmico."""
        m = _TITULOS_PROFESSOR.search(texto)
        if m:
            return texto[: m.start()].strip(), texto[m.start() :].strip()
        return texto.strip(), ""
