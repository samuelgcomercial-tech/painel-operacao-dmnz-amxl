# -*- coding: utf-8 -*-
"""
historico_tbr.py (versao web)

Mesma logica do historico do processa_csv_scc.py do robo desktop
(base_scc_analise.py: carrega_base/salva_base/decide_quem_precisa_reabrir/
atualiza_base) - guarda a classificacao (State Finalizador) de cada TBR
problematico (diferente de Delivered) ja analisado, pra so pedir revisao
de novo quando o TBR e novo, mudou de state desde a ultima vez, OU
voltou como "Received" com a classificacao salva de ontem sendo
"Insucesso" (ver eh_classificacao_de_insucesso e TEXTO_STATE_RECEIVED -
esse terceiro caso e ajuste do Samuel em 15/09/2026, ver
decide_quem_precisa_reabrir).

Isso evita perguntar de novo, todo santo dia, sobre TBR que continua
"preso" no mesmo state de ontem (ex: um MNR que so vai ser reentregue daqui
a 3 dias) - a imensa maioria dos TBRs problematicos de ontem continua
igual hoje. A excecao e "Received" depois de Insucesso: o SCC usa esse
MESMO texto generico pra "voltou pro hub" depois de QUALQUER tentativa
fracassada, seja de quem for, entao "state igual" nao quer dizer "nada
mudou de verdade" nesse caso - o pacote pode ter sido tentado por uma
empresa diferente da tentativa anterior (hoje foi a DMNZ, ontem foi a
MRIZ, por exemplo). Restrito a essa combinacao especifica (state
"Received" + classificacao "Insucesso") a pedido do Samuel - nem todo
state repetido e ambiguo assim ("cancelado" nunca volta pra rota, entao
fica de fora; "retorno"/"devolucao" tambem ficaram de fora depois de um
primeiro ajuste mais amplo, restrito de proposito so a "Insucesso").

A base em si mora no GitHub (ve github_store.py) - esse arquivo so cuida
do "CSV em texto" pra dentro/fora, igual base_das.py.
"""

import csv
import io

COLUNAS_HISTORICO = ["tbr", "classificacao", "state_scc", "primeira_vez_visto", "ultima_verificacao"]

# So "Insucesso" - ajuste do Samuel em 15/09/2026 (restringindo um ajuste
# anterior que tambem pegava "retorno"/"devolucao"): so insucesso tem a
# ambiguidade de "pode ter sido outra empresa dessa vez".
PALAVRAS_FORCAM_REABERTURA = ("insucesso",)

# So o state "Received" - "cancelado" NUNCA volta pra rota (uma vez
# cancelado, e definitivo - nao ha risco de outra empresa pegar ele
# depois, entao "state igual" continua sendo um sinal confiavel pra
# esses), entao a ambiguidade so existe pro "Received" (reentrada
# generica no hub depois de qualquer tentativa fracassada, seja qual for
# a empresa). Comparado sem acento/case - eh so um strip+lower, sem
# acento nenhum nesse texto especifico, entao nao precisa da funcao de
# normalizacao de calculo_fechamento.py (evita import circular).
TEXTO_STATE_RECEIVED = "received"


def eh_classificacao_de_insucesso(classificacao):
    """'Insucesso' de ONTEM nao garante nada sobre HOJE quando o State do
    SCC de hoje volta a ser 'Received' - um pacote pode falhar em dias
    diferentes com motoristas/empresas diferentes (hoje foi a DMNZ,
    ontem foi a MRIZ), e o SCC usa esse MESMO texto generico pra
    'voltou pro hub' depois de qualquer tentativa fracassada, seja de
    quem for - entao 'o state nao mudou' NAO significa 'nada mudou de
    verdade' nesse caso especifico (ao contrario do resto dos TBRs,
    onde e um sinal confiavel, "cancelado" incluido - ver
    TEXTO_STATE_RECEIVED acima). Comparacao por palavra (nao lista
    fechada de textos) porque a classificacao e digitada livre pelo
    usuario, com variacoes (ex: 'INSUCESSO - DMNZ', 'Insucesso -
    Cliente Ausente')."""
    txt = (classificacao or "").lower()
    return any(palavra in txt for palavra in PALAVRAS_FORCAM_REABERTURA)


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
      - precisam_reabrir: TBRs novos, que mudaram de state, OU que
        voltaram como "Received" com classificacao salva de ontem
        "Insucesso" (ver eh_classificacao_de_insucesso e
        TEXTO_STATE_RECEIVED - forcado a pedir revisao de novo mesmo
        com o state igual, ajuste do Samuel em 15/09/2026 - SO pra essa
        combinacao especifica, "cancelado" e "retorno"/"devolucao"
        ficam de fora de proposito, ve o comentario de
        PALAVRAS_FORCAM_REABERTURA) - precisam de revisao manual
        (digitar o State Finalizador de novo)
      - ja_conhecidos: TBRs que continuam com o mesmo state de antes E
        (nao estao no caso acima) - reaproveita a classificacao ja
        salva na base

    Baseada na funcao do robo desktop (base_scc_analise.py), com uma
    regra nova (o terceiro branch abaixo) - antes so olhava se o state
    mudou.
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
        elif (
            state_hoje.strip().lower() == TEXTO_STATE_RECEIVED
            and eh_classificacao_de_insucesso(anterior["classificacao"])
        ):
            precisam_reabrir.append({
                **r, "motivo_reabertura": (
                    f"classificado ontem como \"{anterior['classificacao']}\" e voltou "
                    "como \"Received\" - Insucesso pede revisão de novo mesmo com "
                    "o state igual (pode ter sido outra empresa/motorista dessa vez)"
                )
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
