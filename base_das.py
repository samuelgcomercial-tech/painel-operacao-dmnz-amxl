# -*- coding: utf-8 -*-
"""
base_das.py (versao web)

Mesma logica do trecho "DAs DMNZ" do processa_csv_scc.py do robo desktop
(funcao revisa_cadastro_das) - so devolve os dados pra tela decidir o que
perguntar (via botoes/pills), em vez de usar input()/print() de terminal.
A base de DAs em si mora no GitHub (ve github_store.py) - esse arquivo so
cuida do "CSV em texto" pra dentro/fora, sem saber de onde ele veio.

Le o CSV bruto exportado do SCC (Pesquisar -> Export to CSV), acha quais
motoristas ("Last Scan By") rodaram hoje e ainda nao estao cadastrados
como DMNZ ou nao, e monta a base atualizada depois que o usuario escolhe.

Ainda NAO faz (fica pra depois, de proposito):
  - a excecao "esse DA nao rodou pela DMNZ hoje" (so por hoje / permanente)
  - a coluna "State Finalizador" (justificativa por TBR)
  - gerar o CSV final com as colunas adicionadas
"""

import csv
import io

# Nomes de sistema/robo que aparecem no "Last Scan By" no lugar de um DA de
# verdade (ex: quando a leitura foi feita por automacao, nao por uma
# pessoa) - mesma lista do robo desktop.
NOMES_SISTEMA = {"LASTMILEROUTEPLANNER"}


def normaliza_nome(nome):
    """Deixa o nome tolerante a espaco duplo, espaco nas pontas e diferenca
    de maiuscula/minuscula - os nomes que chegam do SCC vem bem
    inconsistentes nesse quesito. Mesma funcao do robo desktop."""
    return " ".join(nome.split()).upper()


def eh_da_de_verdade(last_scan_by):
    """Mesma regra do robo desktop: login de suporte (tem @) ou nome de
    sistema nao conta como DA de verdade."""
    if not last_scan_by:
        return False
    if "@" in last_scan_by:
        return False
    if normaliza_nome(last_scan_by) in NOMES_SISTEMA:
        return False
    return True


def _detecta_delimitador(texto_amostra):
    """Excel com configuracao regional BR costuma salvar CSV com ';' em vez
    de ',' - deixa a leitura tolerante aos dois. Mesma ideia do robo
    desktop, so que a partir de um texto ja em memoria em vez de reabrir
    o arquivo do disco."""
    return ";" if texto_amostra.count(";") > texto_amostra.count(",") else ","


def le_csv_scc(arquivo):
    """Le o CSV exportado do SCC a partir de um objeto de upload do
    Streamlit. Devolve lista de dicts (uma por linha), igual csv.DictReader
    do robo desktop."""
    texto = arquivo.getvalue().decode("utf-8-sig")
    delimitador = _detecta_delimitador(texto[:2000])
    reader = csv.DictReader(io.StringIO(texto), delimiter=delimitador)
    return list(reader)


def extrai_das_do_dia(linhas_csv):
    """Lista (sem repetir, em ordem alfabetica) os nomes de DA de verdade
    que aparecem no Last Scan By das linhas de hoje. Mesma funcao do robo
    desktop."""
    nomes = set()
    for l in linhas_csv:
        nome = (l.get("Last Scan By") or "").strip()
        if eh_da_de_verdade(nome):
            nomes.add(nome)
    return sorted(nomes)


def parseia_texto_base(texto):
    """Le a base de DAs a partir de um texto CSV já em memória (Nome do
    DA,DMNZ) - usado tanto pro que vem do GitHub quanto (se um dia
    precisar de novo) de um upload manual. Se o texto vier vazio/None
    (primeira vez, arquivo ainda não existe), devolve tudo vazio - não
    trava o app, só significa que todo mundo vai aparecer como "ainda
    não cadastrado".

    Devolve (base_dict, linhas_bruta):
      - base_dict: nome normalizado -> classificacao (ex.: "DMNZ")
      - linhas_bruta: lista de (nome_original, classificacao), preserva a
        grafia original (acentos/maiusculas) pra nao perder formatacao ao
        salvar de novo.
    """
    if not texto:
        return {}, []

    delimitador = _detecta_delimitador(texto[:2000])
    reader = csv.DictReader(io.StringIO(texto), delimiter=delimitador)
    fieldnames_norm = [(c or "").strip().lower() for c in (reader.fieldnames or [])]
    if "nome do da" not in fieldnames_norm or "dmnz" not in fieldnames_norm:
        raise ValueError(
            "A base de DAs no repositório não tem as colunas esperadas "
            f"'Nome do DA' e 'DMNZ' (encontrei: {reader.fieldnames})."
        )
    col_nome = reader.fieldnames[fieldnames_norm.index("nome do da")]
    col_dmnz = reader.fieldnames[fieldnames_norm.index("dmnz")]

    base_dict = {}
    linhas_bruta = []
    for row in reader:
        nome_bruto = row.get(col_nome)
        if not nome_bruto:
            continue
        # protecao: nome contaminado com virgula colada na mesma celula
        # (erro comum de digitacao no Excel) - usa só a parte antes dela.
        if "," in nome_bruto:
            nome_bruto = nome_bruto.split(",")[0]
        classificacao = (row.get(col_dmnz) or "").strip()
        base_dict[normaliza_nome(nome_bruto)] = classificacao
        linhas_bruta.append((nome_bruto.strip(), classificacao))

    return base_dict, linhas_bruta


def gera_csv_base_das(linhas_bruta):
    """Monta o texto do CSV da base de DAs (mesmo formato do robo desktop:
    'Nome do DA,DMNZ'), pronto pra oferecer como download."""
    buffer = io.StringIO()
    writer = csv.writer(buffer)
    writer.writerow(["Nome do DA", "DMNZ"])
    writer.writerows(linhas_bruta)
    return buffer.getvalue()
