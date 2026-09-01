# -*- coding: utf-8 -*-
"""
historico_tbr.py (versao web)

Mesma logica do historico do processa_csv_scc.py do robo desktop
(base_scc_analise.py: carrega_base/salva_base/decide_quem_precisa_reabrir/
atualiza_base) - guarda a classificacao (State Finalizador) de cada TBR
problematico (diferente de Delivered) ja analisado, pra so pedir revisao
de novo quando o TBR e novo OU mudou de state desde a ultima vez.

Isso evita perguntar de novo, todo santo dia, sobre TBR que continua
"preso" no mesmo state de ontem (ex: um MNR que so vai ser reentregue daqui
a 3 dias) - a imensa maioria dos TBRs problematicos de ontem continua
igual hoje.

A base em si mora no GitHub (ve github_store.py) - esse arquivo so cuida
do "CSV em texto" pra dentro/fora, igual base_das.py.
"""

import csv
import io

COLUNAS_HISTORICO = ["tbr", "classificacao", "state_scc", "primeira_vez_visto", "ultima_verificacao"]


def parseia_texto_historico(texto):
    """Le o historico a partir de um texto CSV ja em memoria (vindo do
    GitHub). Se vier vazio/None (primeira vez, arquivo ainda nao existe),
    devolve um dict vazio - nao trava o app, so significa que todo TBR
    nao-Delivered vai pedir revisao (comportamento esperado no primeiro
    dia).

    Devolve dict: tbr -> {tbr, classificacao, state_scc,
    primeira_vez_visto, ultima_verificacao}."""
    if not texto:
        return {}
    reader = csv.DictReader(io.StringIO(texto))
    return {row["tbr"]: row for row in reader if row.get("tbr")}


def gera_csv_historico(base):
    """Monta o texto do CSV do historico (mesmo formato do robo desktop),
    pronto pra salvar no GitHub."""
    buffer = io.StringIO()
    writer = csv.DictWriter(buffer, fieldnames=COLUNAS_HISTORICO)
    writer.writeheader()
    writer.writerows(base.values())
    return buffer.getvalue()


def decide_quem_precisa_reabrir(tbrs_hoje, base):
    """
    tbrs_hoje: lista de dicts {'tbr':..., 'state_scc':...} vindos do CSV de
    hoje (so os TBRs diferentes de Delivered).

    Devolve (precisam_reabrir, ja_conhecidos):
      - precisam_reabrir: TBRs novos ou que mudaram de state - precisam de
        revisao manual (digitar o State Finalizador de novo)
      - ja_conhecidos: TBRs que continuam com o mesmo state de antes -
        reaproveita a classificacao ja salva na base

    Mesma funcao do robo desktop (base_scc_analise.py), sem nenhuma
    mudanca de logica.
    """
    precisam_reabrir = []
    ja_conhecidos = []
    for r in tbrs_hoje:
        tbr = r["tbr"]
        state_hoje = r["state_scc"]
        anterior = base.get(tbr)
        if anterior is None:
            precisam_reabrir.append({**r, "motivo_reabertura": "novo"})
        elif anterior["state_scc"] != state_hoje:
            precisam_reabrir.append({
                **r, "motivo_reabertura": f"state mudou ({anterior['state_scc']} -> {state_hoje})"
            })
        else:
            ja_conhecidos.append(r)
    return precisam_reabrir, ja_conhecidos


def atualiza_base(base, tbr, classificacao, state_scc, data_hoje_str):
    """Atualiza (ou cria) a entrada de um TBR na base apos classifica-lo.
    Mesma funcao do robo desktop, so que 'data_hoje_str' ja vem formatada
    (dd/mm/aaaa) em vez de um objeto date - quem chama decide o formato."""
    existente = base.get(tbr)
    primeira_vez = existente["primeira_vez_visto"] if existente else data_hoje_str
    base[tbr] = {
        "tbr": tbr,
        "classificacao": classificacao,
        "state_scc": state_scc,
        "primeira_vez_visto": primeira_vez,
        "ultima_verificacao": data_hoje_str,
    }
