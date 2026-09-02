# -*- coding: utf-8 -*-
"""
state_finalizador.py (versao web)

Mesma logica do trecho "State Finalizador" do processa_csv_scc.py do robo
desktop - preenche duas colunas por TBR:
  - "State Finalizador": Delivered vira "Entregue" automatico. In
    Transit/Failed com um DA de verdade vinculado viram "EM ROTA - DMNZ"
    ou "EM ROTA - PARCEIRO" automatico (cruzando com a base de DAs). O
    resto so pede revisao manual quando o TBR e novo ou mudou de state
    desde a ultima vez (reaproveita historico_tbr.py pro resto).
  - "DAs DMNZ": nome do motorista (Last Scan By) cruzado com a base de
    DAs, com fallback pro "motorista real da rota" quando o Last Scan By
    e login de suporte (tem @) - so pra entregas confirmadas (Delivered).

Ainda NAO faz (fica pra depois, de proposito, mesmo escopo ja combinado
pra base_das.py):
  - a excecao "esse DA nao rodou pela DMNZ hoje" (so por hoje/permanente)
  - a coluna "Reversa"/"Pacotes em Piso" (arquivo separado)
"""

from collections import Counter

from base_das import eh_da_de_verdade, normaliza_nome
from historico_tbr import decide_quem_precisa_reabrir


def calcula_motorista_real_da_rota(linhas_csv):
    """Pra cada Route Code, acha o DA de verdade que mais aparece nela -
    usado como 'motorista real da rota' quando o Last Scan By de uma
    entrega especifica veio como login de suporte (@) em vez do nome de
    um DA. Mesma funcao do robo desktop."""
    contagem_por_rota = {}
    for l in linhas_csv:
        rota = (l.get("Route Code") or "").strip()
        nome = (l.get("Last Scan By") or "").strip()
        if rota and eh_da_de_verdade(nome):
            contagem_por_rota.setdefault(rota, Counter())[nome] += 1
    return {
        rota: contador.most_common(1)[0][0]
        for rota, contador in contagem_por_rota.items()
    }


def classifica_dmnz_ou_parceiro(nome_da, rota, base_das, motorista_real_da_rota):
    """Devolve 'DMNZ' ou 'PARCEIRO' (nunca vazio) - usado pro
    preenchimento automatico de 'EM ROTA - ...'. Tenta achar o nome direto
    na base primeiro; se nao achar, cai pro motorista que mais aparece
    naquela mesma rota. Mesma funcao do robo desktop (sem a excecao do
    dia, que ainda nao existe na versao web)."""
    nome_norm = normaliza_nome(nome_da) if nome_da else ""
    if base_das.get(nome_norm) == "DMNZ":
        return "DMNZ"
    nome_real = motorista_real_da_rota.get(rota, "")
    nome_real_norm = normaliza_nome(nome_real) if nome_real else ""
    return "DMNZ" if base_das.get(nome_real_norm) == "DMNZ" else "PARCEIRO"


def eh_state_vazio(state_scc):
    """TBR que chegou do SCC sem nenhum State preenchido - acontece de vez
    em quando (falha de sincronizacao do proprio SCC). Nao tem State
    Finalizador nenhum pra deduzir dai, entao classifica automatico como
    'SEM STATE SCC' em vez de pedir revisao manual - nao ha nada pra
    revisar, o problema e o dado que nao veio."""
    return not (state_scc or "").strip()


def eh_state_auto_em_rota(state_scc, last_scan_by):
    """States que o robo classifica sozinho (sem perguntar) quando ja tem
    um DA de verdade vinculado: 'EM ROTA - DMNZ' ou 'EM ROTA - PARCEIRO'.
    Usa "contem" em vez de comparacao exata porque o state as vezes vem
    com variacao (ex: 'In Transit (DS -> Customer)', 'Delivery Failed').
    So classifica automatico se o Last Scan By for um DA de verdade (sem
    @) - com login de suporte ninguem confirmou quem esta com o pacote,
    entao cai pra revisao manual. Mesma regra do robo desktop."""
    txt = (state_scc or "").strip().lower()
    eh_in_transit_ou_failed = "in transit" in txt or "failed" in txt
    return eh_in_transit_ou_failed and eh_da_de_verdade(last_scan_by)


def processa(linhas_csv, base_tbr, base_das, data_hoje_str):
    """Centraliza o processamento do State Finalizador pra todas as linhas
    do CSV de hoje. Nao decide nada sozinho sobre o que precisa de
    revisao manual - so separa e devolve pra tela perguntar.

    Devolve um dict:
      - entregues: linhas com State=Delivered (State Finalizador = "Entregue" automatico)
      - auto_em_rota: lista de dicts {tbr, resposta, last_scan_by, route_code} - EM ROTA automatico
      - conhecidos: lista de dicts (do historico) reaproveitados sem mudanca
      - pendentes: lista de dicts {tbr, state_scc, motivo_reabertura, last_scan_by, route_code}
        agrupaveis por state_scc - precisam de revisao manual na tela
      - motorista_real_da_rota: dict rota -> nome (usado depois pra montar
        a coluna final "DAs DMNZ")
    """
    motorista_real_da_rota = calcula_motorista_real_da_rota(linhas_csv)

    entregues = [l for l in linhas_csv if (l.get("State") or "").strip().lower() == "delivered"]
    nao_entregues_raw = [l for l in linhas_csv if (l.get("State") or "").strip().lower() != "delivered"]
    nao_entregues = [
        {
            "tbr": l["Tracking ID"],
            "state_scc": l["State"],
            "last_scan_by": (l.get("Last Scan By") or "").strip(),
            "route_code": (l.get("Route Code") or "").strip(),
        }
        for l in nao_entregues_raw
    ]

    reabrir, conhecidos = decide_quem_precisa_reabrir(nao_entregues, base_tbr)

    auto_em_rota = []
    auto_sem_state = []
    pendentes = []
    for r in reabrir:
        if eh_state_vazio(r["state_scc"]):
            auto_sem_state.append({**r, "resposta": "SEM STATE SCC"})
        elif eh_state_auto_em_rota(r["state_scc"], r["last_scan_by"]):
            classificacao = classifica_dmnz_ou_parceiro(
                r["last_scan_by"], r["route_code"], base_das, motorista_real_da_rota
            )
            auto_em_rota.append({**r, "resposta": f"EM ROTA - {classificacao}"})
        else:
            pendentes.append(r)

    return {
        "entregues": entregues,
        "auto_em_rota": auto_em_rota,
        "auto_sem_state": auto_sem_state,
        "conhecidos": conhecidos,
        "pendentes": pendentes,
        "motorista_real_da_rota": motorista_real_da_rota,
    }


def _categoria_state(state_scc):
    """Agrupa o State bruto do SCC em 3 baldes pra conferencia por DA:
    'entregue' (State=Delivered), 'em_rota' (contem "in transit" ou
    "failed" - mesmo criterio ja usado em eh_state_auto_em_rota) e
    'a_analisar' (qualquer outro state - ex: Manifested, Arrived, Out
    for Delivery - ainda no inicio do fluxo, sem State Finalizador
    automatico possivel; precisa de revisao manual, igual o resto do
    State Finalizador)."""
    txt = (state_scc or "").strip().lower()
    if txt == "delivered":
        return "entregue"
    if "in transit" in txt or "failed" in txt:
        return "em_rota"
    return "a_analisar"


def resumo_das_por_rota(linhas_csv, base_das):
    """Monta um resumo Rota -> DA(s) que apareceram nela escaneando, com a
    classificacao DMNZ/PARCEIRO de cada um - serve pra conferencia visual
    manual (ex: perceber que um DA nosso apareceu numa rota que
    normalmente e do parceiro, ou o contrario - troca entre empresas,
    resgate ou divisao de rota). Ao contrario de
    calcula_motorista_real_da_rota, aqui aparece TODO DA de verdade que
    escaneou na rota, nao so o que mais aparece - senao um "resgate" ou
    "divisao" ficaria escondido atras do DA dominante.

    O total de pacotes de cada DA vem detalhado em 3 baldes (ve
    _categoria_state): entregues, em_rota, a_analisar - alem do total.

    Devolve uma lista de dicts {rota, da, entregues, em_rota, a_analisar,
    pacotes, classificacao, mista} - 'mista' e True em toda linha de uma
    rota que teve DA de DMNZ E de Parceiro ao mesmo tempo (o caso que
    mais vale a pena conferir)."""
    detalhe_por_rota = {}
    for l in linhas_csv:
        rota = (l.get("Route Code") or "").strip()
        nome = (l.get("Last Scan By") or "").strip()
        if not (rota and eh_da_de_verdade(nome)):
            continue
        categoria = _categoria_state(l.get("State"))
        contagem = detalhe_por_rota.setdefault(rota, {}).setdefault(
            nome, {"entregue": 0, "em_rota": 0, "a_analisar": 0}
        )
        contagem[categoria] += 1

    linhas = []
    for rota in sorted(detalhe_por_rota):
        # ordena pelo total de pacotes (maior primeiro) - mesmo efeito do
        # most_common() de antes, só que agora somando os 3 baldes.
        itens_nome = sorted(
            detalhe_por_rota[rota].items(),
            key=lambda kv: sum(kv[1].values()),
            reverse=True,
        )
        itens_rota = []
        for nome, contagem in itens_nome:
            nome_norm = normaliza_nome(nome)
            classificacao = "DMNZ" if base_das.get(nome_norm) == "DMNZ" else "PARCEIRO"
            itens_rota.append({
                "rota": rota,
                "da": nome,
                "entregues": contagem["entregue"],
                "em_rota": contagem["em_rota"],
                "a_analisar": contagem["a_analisar"],
                "pacotes": sum(contagem.values()),
                "classificacao": classificacao,
            })
        mista = len({item["classificacao"] for item in itens_rota}) > 1
        for item in itens_rota:
            item["mista"] = mista
        linhas.extend(itens_rota)
    return linhas


def lista_sem_da_de_verdade(linhas_csv, base_das, motorista_real_da_rota):
    """Pacotes cujo 'Last Scan By' NAO e um DA de verdade (login de
    suporte com @, ou nome de sistema tipo LastMileRoutePlanner) - o robo
    desktop mostra essa lista separada na prevoa final antes de salvar,
    porque e um caso especial: so cruza com o motorista real da rota (pra
    saber se conta como DMNZ) quando a entrega foi CONFIRMADA (Delivered)
    - pra outros states ninguem confirmou quem esta com o pacote de
    verdade. Mesma logica de monta_coluna_das_dmnz, so que devolvendo uma
    lista pronta pra mostrar em tabela em vez de um dict tbr->classificacao.

    Devolve lista de dicts {tbr, rota, state_scc, motorista_real_da_rota,
    das_dmnz} - 'das_dmnz' vem "(não é DMNZ)" quando vazio, só pra não
    aparecer uma célula em branco confundindo com "ainda não analisado"."""
    linhas = []
    for l in linhas_csv:
        nome_da = (l.get("Last Scan By") or "").strip()
        if eh_da_de_verdade(nome_da):
            continue
        rota = (l.get("Route Code") or "").strip()
        entregue = (l.get("State") or "").strip().lower() == "delivered"
        nome_real = motorista_real_da_rota.get(rota, "") if entregue else ""
        nome_real_norm = normaliza_nome(nome_real) if nome_real else ""
        classificacao = base_das.get(nome_real_norm, "") if entregue else ""
        linhas.append({
            "tbr": l.get("Tracking ID", ""),
            "rota": rota,
            "state_scc": l.get("State", ""),
            "motorista_real_da_rota": nome_real or "?",
            "das_dmnz": classificacao or "(não é DMNZ)",
        })
    return linhas


def monta_coluna_das_dmnz(linhas_csv, base_das, motorista_real_da_rota):
    """Monta a coluna final 'DAs DMNZ' pra cada linha do CSV (dict tbr ->
    classificacao). So o Last Scan By de uma entrega CONFIRMADA
    (Delivered) com login de suporte cai pro motorista real da rota -
    pra 'in transit'/outros states ninguem confirmou quem esta com o
    pacote, entao fica sem DA vinculado. Mesma regra do robo desktop."""
    resultado = {}
    for l in linhas_csv:
        tbr = l["Tracking ID"]
        nome_da = (l.get("Last Scan By") or "").strip()
        entregue = (l.get("State") or "").strip().lower() == "delivered"

        if entregue and not eh_da_de_verdade(nome_da):
            rota = (l.get("Route Code") or "").strip()
            nome_real = motorista_real_da_rota.get(rota, "")
            nome_da_norm = normaliza_nome(nome_real) if nome_real else ""
        elif not eh_da_de_verdade(nome_da):
            nome_da_norm = ""
        else:
            nome_da_norm = normaliza_nome(nome_da) if nome_da else ""

        resultado[tbr] = base_das.get(nome_da_norm, "")
    return resultado
