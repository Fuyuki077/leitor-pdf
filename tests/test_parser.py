import sys
import os

sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))

from pdf_parser import PDFTranscriptParser

p = PDFTranscriptParser()


# ─── _limpar_texto ─────────────────────────────────────────────────────────────

def test_limpar_texto_none():
    assert p._limpar_texto(None) == ""

def test_limpar_texto_remove_caracteres_controle():
    assert p._limpar_texto("abc\x00def") == "abcdef"

def test_limpar_texto_strip():
    assert p._limpar_texto("  hello  ") == "hello"

def test_limpar_texto_normaliza_crlf():
    assert "\r" not in p._limpar_texto("linha\r\noutra")


# ─── _eh_cabecalho ─────────────────────────────────────────────────────────────

def test_eh_cabecalho_duas_palavras_suficientes():
    assert p._eh_cabecalho(["Ord", "Semestre", "Algo", "Outro"])

def test_eh_cabecalho_com_disciplina_e_professor():
    assert p._eh_cabecalho(["Disciplina", "Professor", "x", "y"])

def test_eh_cabecalho_insuficiente():
    assert not p._eh_cabecalho(["Disciplina", "Nota"])

def test_eh_cabecalho_lista_vazia():
    assert not p._eh_cabecalho([])


# ─── _inicio_de_registro ───────────────────────────────────────────────────────

def test_inicio_de_registro_valido():
    assert p._inicio_de_registro(["1 2022.1 Alguma Disciplina"])

def test_inicio_de_registro_multiplas_celulas():
    assert p._inicio_de_registro(["1", "2023.2", "Cálculo I"])

def test_inicio_de_registro_texto_nao_numerico():
    assert not p._inicio_de_registro(["Cálculo I Professor"])

def test_inicio_de_registro_lista_vazia():
    assert not p._inicio_de_registro([])

def test_inicio_de_registro_ano_invalido():
    # Ano antes de 2000 não deve ser reconhecido
    assert not p._inicio_de_registro(["1 1999.1 Disciplina"])


# ─── _separar_faltas_situacao ──────────────────────────────────────────────────

def test_separar_faltas_situacao_vazio():
    assert p._separar_faltas_situacao("") == ("", "")

def test_separar_faltas_situacao_aprovado():
    faltas, situacao = p._separar_faltas_situacao("5 Aprovado")
    assert faltas == "5"
    assert "Aprovado" in situacao

def test_separar_faltas_situacao_reprovado():
    faltas, situacao = p._separar_faltas_situacao("12Reprovado")
    assert faltas == "12"
    assert "Reprovado" in situacao

def test_separar_faltas_situacao_aproveitada():
    faltas, situacao = p._separar_faltas_situacao("Disciplina Aproveitada")
    assert faltas == ""
    assert situacao == "Disciplina Aproveitada"

def test_separar_faltas_situacao_so_numero():
    faltas, situacao = p._separar_faltas_situacao("3")
    assert faltas == "3"
    assert situacao == ""


# ─── _normalizar_linha ─────────────────────────────────────────────────────────

def test_normalizar_linha_preenche_ate_8():
    result = p._normalizar_linha(["a", "b", "c"])
    assert len(result) == 8

def test_normalizar_linha_trunca_acima_8():
    result = p._normalizar_linha(["x"] * 12)
    assert len(result) == 8

def test_normalizar_linha_vazia_retorna_8_vazios():
    result = p._normalizar_linha([])
    assert result == [""] * 8

def test_normalizar_linha_remove_celulas_vazias():
    result = p._normalizar_linha(["", "a", "", "b"])
    assert result[0] == "a"
    assert result[1] == "b"


# ─── _quebrar_linha_colapsada ──────────────────────────────────────────────────

def test_quebrar_linha_colapsada_formato_valido():
    linha = "1 2022.1 Cálculo I Prof Silva 60h 8,5 2 Aprovado"
    result = p._quebrar_linha_colapsada(linha)
    assert result is not None
    assert len(result) == 8
    assert result[0] == "1"
    assert result[1] == "2022.1"
    assert result[4] == "60h"
    assert result[5] == "8,5"
    assert result[6] == "2"
    assert result[7] == "Aprovado"

def test_quebrar_linha_colapsada_formato_invalido():
    result = p._quebrar_linha_colapsada("texto sem formato esperado")
    assert result is None

def test_quebrar_linha_colapsada_multiline_quebra_pagina():
    """Simula linha colapsada com \\n de quebra de página."""
    linha = (
        "8 2021.1 MTC(BAC) - Metodologia do Dr. John Eric Lemos 60h 6,00 36 Reprovado\n"
        "Trabalho Científico de Amorim por Falta\n"
        "Msc. Josuá Gomes\nda Silva"
    )
    result = p._quebrar_linha_colapsada(linha)
    assert result is not None
    assert len(result) == 8
    assert result[0] == "8"
    assert result[1] == "2021.1"
    assert result[4] == "60h"
    assert result[5] == "6,00"
    assert result[6] == "36"
    assert result[7] == "Reprovado por Falta"

def test_quebrar_linha_colapsada_situacao_reprovado_simples():
    linha = "3 2022.2 Física I Dr. Santos 60h 4,0 10 Reprovado"
    result = p._quebrar_linha_colapsada(linha)
    assert result is not None
    assert result[7] == "Reprovado"

def test_quebrar_linha_colapsada_sem_ch():
    result = p._quebrar_linha_colapsada("1 2022.1 Disciplina Professor nota resultado")
    assert result is None


# ─── _extrair_situacao_de_texto ────────────────────────────────────────────────

def test_extrair_situacao_aprovado():
    assert p._extrair_situacao_de_texto("Aprovado") == "Aprovado"

def test_extrair_situacao_reprovado():
    assert p._extrair_situacao_de_texto("Reprovado") == "Reprovado"

def test_extrair_situacao_reprovado_por_falta_com_lixo():
    assert p._extrair_situacao_de_texto("Reprovado Trabalho Científico por Falta lixo") == "Reprovado por Falta"

def test_extrair_situacao_aproveitada():
    assert p._extrair_situacao_de_texto("Disciplina Aproveitada") == "Disciplina Aproveitada"


# ─── _dividir_disc_prof ────────────────────────────────────────────────────────

def test_dividir_disc_prof_com_dr():
    disc, prof = p._dividir_disc_prof("Cálculo I Dr. Santos Silva")
    assert disc == "Cálculo I"
    assert prof == "Dr. Santos Silva"

def test_dividir_disc_prof_com_msc():
    disc, prof = p._dividir_disc_prof("Física Msc. João Pereira")
    assert disc == "Física"
    assert prof.startswith("Msc.")

def test_dividir_disc_prof_sem_titulo():
    disc, prof = p._dividir_disc_prof("Cálculo I Prof Silva")
    assert disc == "Cálculo I Prof Silva"
    assert prof == ""


# ─── _mesclar_continuacao ──────────────────────────────────────────────────────

def test_mesclar_preenche_campo_vazio():
    atual = ["1", "2022.1", "Cálculo", "", "", "", "", ""]
    comp  = ["",  "",        "",        "Prof Silva", "", "", "", ""]
    result = p._mesclar_continuacao(atual, comp)
    assert result[3] == "Prof Silva"

def test_mesclar_preserva_campo_existente():
    atual = ["1", "2022.1", "Cálculo", "", "", "", "", ""]
    comp  = ["",  "",        "Cálculo", "", "", "", "", ""]
    result = p._mesclar_continuacao(atual, comp)
    # Não duplica o conteúdo idêntico
    assert result[2] == "Cálculo"

def test_mesclar_concatena_conteudo_diferente():
    atual = ["1", "2022.1", "Parte 1", "", "", "", "", ""]
    comp  = ["",  "",        "Parte 2", "", "", "", "", ""]
    result = p._mesclar_continuacao(atual, comp)
    assert "Parte 1" in result[2]
    assert "Parte 2" in result[2]
