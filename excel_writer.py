import re
import os
from dataclasses import dataclass, field

import openpyxl

from pdf_parser import PDFTranscriptParser

# Índices das colunas — espelham a ordem do cabeçalho escrito em gerar()
_COL_ALUNO      = 0
_COL_CPF        = 1
_COL_CURSO      = 2
_COL_MATRICULA  = 3
_COL_ORD        = 4
_COL_SEMESTRE   = 5
_COL_DISCIPLINA = 6
_COL_PROFESSOR  = 7
_COL_CH         = 8
_COL_MEDIA      = 9
_COL_FALTAS     = 10
_COL_SITUACAO   = 11


@dataclass
class FiltroSemestres:
    modo: str = "todos"           # "todos" | "primeiros_n" | "especificos"
    quantidade: int = 4
    especificos: list[str] = field(default_factory=list)


class ExcelReportGenerator:
    """Gera relatórios .xlsx a partir dos dados extraídos pelo PDFTranscriptParser."""

    def __init__(self):
        self._parser = PDFTranscriptParser()

    def gerar(
        self,
        dados_aluno: dict,
        linhas_brutas: list,
        pasta_destino: str,
        filtro: FiltroSemestres,
    ) -> str:
        """Salva o arquivo .xlsx e retorna o caminho completo."""
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Historico"
        ws.append([
            "Aluno", "CPF", "Curso", "Matrícula",
            "Ord", "Semestre", "Disciplina", "Professor",
            "CH", "Média", "Faltas", "Situação",
        ])

        for d in linhas_brutas:
            if len(d) >= 8 and d[7].strip():
                faltas, situacao = d[6], d[7]
            else:
                faltas, situacao = self._parser._separar_faltas_situacao(d[6])

            try:
                val_ord    = int(d[0]) if str(d[0]).isdigit() else d[0]
                media_str  = str(d[5]).replace(",", ".").strip()
                val_media  = float(media_str) if media_str and media_str != "None" else 0.0
                val_faltas = int(faltas) if str(faltas).isdigit() else 0
            except (ValueError, TypeError):
                val_ord, val_media, val_faltas = d[0], 0.0, 0

            ws.append([
                dados_aluno["nome"],
                dados_aluno["cpf"],
                dados_aluno["curso"],
                dados_aluno["mat"],
                val_ord,
                d[1],
                d[2],
                d[3],
                d[4],
                val_media,
                val_faltas,
                situacao,
            ])

        self._calcular_medias_semestrais(ws, dados_aluno["nome"], filtro)

        nome_saida = f"Relatorio_{dados_aluno['mat']}_{{{dados_aluno['nome']}}}.xlsx"
        caminho_saida = os.path.join(pasta_destino, nome_saida)
        try:
            wb.save(caminho_saida)
        except PermissionError as e:
            raise PermissionError(
                f"Sem permissão para salvar em '{pasta_destino}': {e}"
            ) from e
        return caminho_saida

    def _calcular_medias_semestrais(
        self, worksheet, nome_aluno: str, filtro: FiltroSemestres
    ) -> None:
        agrupamento: dict[str, list[float]] = {}

        for row in worksheet.iter_rows(min_row=2, values_only=True):
            if not row or row[0] is None or "APLICANTE" in str(row[0]):
                break

            sem_raw  = str(row[_COL_SEMESTRE] or "").strip()
            nota_raw = str(row[_COL_MEDIA] if row[_COL_MEDIA] is not None else "").strip().replace(",", ".")

            match_sem = re.search(r"(\d{4}[./]\d|\d+)", sem_raw)
            if not match_sem:
                continue

            sem_id = match_sem.group(0).replace("/", ".")
            try:
                nota = float(nota_raw)
                agrupamento.setdefault(sem_id, []).append(nota)
            except ValueError:
                continue

        periodos_ordenados = sorted(agrupamento.keys())

        if filtro.modo == "primeiros_n":
            periodos_selecionados = periodos_ordenados[: filtro.quantidade]
            label_total = f"MÉDIA GERAL ({len(periodos_selecionados)} SEMESTRES):"
        elif filtro.modo == "especificos":
            desejados = {s.strip().replace("/", ".") for s in filtro.especificos}
            periodos_selecionados = [p for p in periodos_ordenados if p in desejados]
            label_total = f"MÉDIA GERAL ({len(periodos_selecionados)} SEMESTRES SELECIONADOS):"
        else:
            periodos_selecionados = periodos_ordenados
            label_total = f"MÉDIA GERAL ({len(periodos_ordenados)} SEMESTRES):"

        worksheet.append([])
        worksheet.append([f"APLICANTE: {nome_aluno}"])
        worksheet.append(["Semestre", "Soma Notas", "Qtd. Disciplinas", "Média Período"])

        medias_finais = []
        for sem in periodos_selecionados:
            notas = agrupamento.get(sem, [])
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
                label_total,
                "",
                "",
                round(sum(medias_finais) / len(medias_finais), 2),
            ])
