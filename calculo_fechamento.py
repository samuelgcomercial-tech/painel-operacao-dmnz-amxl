# -*- coding: utf-8 -*-
"""
calculo_fechamento.py (versao web)

Mesma logica de calculo do calcula_tudo/monta_painel3/eh_insucesso/
classifica_insucesso do robo desktop (fechamento_lrn9.py) - so a parte de
CALCULO (sem o desenho em PIL, que era so pro PNG fixo). O desenho vira
dashboard HTML depois, em cima do resultado daqui (Etapa 3 - Gerar
Fechamento).

Reaproveitado em DOIS momentos: na previa que aparece logo apos o State
Finalizador ser salvo, e depois de verdade na Etapa 3 - mesma conta, dois
lugares diferentes usando o mesmo resultado, pra nunca ficar dessincronizado.

Ainda NAO inclui (fica pra depois, de proposito):
  - Painel 6 (State Reversa) - arquivo separado, ainda nao entrou no fluxo web
  - avisos de limite visual do PNG antigo (so faziam sentido pro desenho
    fixo em pixel; no dashboard HTML o espaco se ajusta sozinho)
"""

import unicodedata
from collections import Counter, OrderedDict

from state_finalizador import monta_coluna_das_dmnz


def normaliza_texto_painel(txt):
    """Maiusculo e sem acento, só pra exibição/agrupamento nos painéis -
    pedido do Samuel pra tolerar o State Finalizador digitado com
    variação de maiuscula/minuscula/acento ('Insucesso - Dmnz' e
    'INSUCESSO - DMNZ' precisam contar como a MESMA coisa no painel).
    NÃO mexe no que fica salvo no histórico - lá continua a grafia
    original que o usuário digitou, isso é só uma cópia pra exibição."""
    txt = unicodedata.normalize("NFKD", txt or "")
    sem_acento = "".join(c for c in txt if not unicodedata.combining(c))
    return " ".join(sem_acento.upper().split())


def monta_dados_do_dia(linhas_csv, base_tbr, base_das, motorista_real_da_rota):
    """Monta a lista 'dados' que calcula_tudo espera, juntando os pedaços
    que já existem separados: State do SCC (direto do CSV), State
    Finalizador (Entregue automático pra quem foi Delivered, ou o que
    estiver salvo no histórico pra quem não foi - já reaproveitando
    normaliza_texto_painel) e a coluna DAs DMNZ (monta_coluna_das_dmnz, do
    state_finalizador.py - mesma fonte de verdade que já alimenta a tela,
    sem recalcular nada com regra diferente)."""
    dmnz_por_tbr = monta_coluna_das_dmnz(linhas_csv, base_das, motorista_real_da_rota)

    dados = []
    for l in linhas_csv:
        tbr = l["Tracking ID"]
        state_scc = (l.get("State") or "").strip()
        if state_scc.lower() == "delivered":
            finalizador = "Entregue"
        else:
            entrada = base_tbr.get(tbr)
            finalizador = normaliza_texto_painel(entrada["classificacao"]) if entrada else ""
        dados.append({
            "tbr": tbr,
            "state_scc": state_scc,
            "finalizador": finalizador,
            "dmnz": dmnz_por_tbr.get(tbr, "") == "DMNZ",
        })
    return dados


def monta_painel_state_scc(dados):
    """Painel 1 (contagem simples por State do SCC) e painel 3 (mesma
    contagem, detalhada por State Finalizador dentro de cada State do
    SCC). Mesma funcao do desktop (monta_painel3), sem mudanca de logica."""
    contagem_scc = Counter(r["state_scc"] for r in dados)
    nested = OrderedDict((cat, {}) for cat, _ in contagem_scc.most_common())
    for r in dados:
        cat, fin = r["state_scc"], r["finalizador"]
        nested[cat][fin] = nested[cat].get(fin, 0) + 1
    for cat in nested:
        nested[cat] = dict(sorted(nested[cat].items(), key=lambda kv: -kv[1]))
    return contagem_scc, nested


def eh_insucesso(finalizador):
    """Mesma regra do desktop: 'insucesso' no texto conta; 'em rota' junto
    NAO conta (segunda tentativa ainda no mesmo dia, nao e insucesso
    definitivo); 'cancelado' com '- dmnz'/'- mriz' tambem conta (cancelado
    ja em rota). 'PCT cancelado' sozinho (sem esses sufixos) nao conta -
    significa cancelado ANTES de ir pra rota."""
    txt = finalizador.lower()
    if "em rota" in txt:
        return False
    if "insucesso" in txt:
        return True
    if "cancelad" in txt and ("dmnz" in txt or "mriz" in txt):
        return True
    return False


def classifica_insucesso(finalizador):
    """'dmnz', 'parceiro' (qualquer nome especifico, ex: mriz/deluna) ou
    'sem_detalhe' (usuario digitou so 'Insucesso' sem dizer de quem).
    Mesma funcao do desktop."""
    txt = finalizador.strip().lower()
    if "dmnz" in txt:
        return "dmnz"
    if txt == "insucesso":
        return "sem_detalhe"
    return "parceiro"


PALAVRAS_AVARIA = ["avaria", "quebrado", "vazando", "rasgado"]


def frase_insucesso(dados, node):
    """Mesma função do desktop: detalha se o insucesso é referente a DMNZ,
    parceiro (qualquer nome específico) ou ficou sem detalhamento de a
    quem pertence."""
    insucessos = [r for r in dados if eh_insucesso(r["finalizador"])]
    if not insucessos:
        return f"Em {node} não houve registros de insucessos"

    tem_dmnz = any(classifica_insucesso(r["finalizador"]) == "dmnz" for r in insucessos)
    tem_parceiro = any(classifica_insucesso(r["finalizador"]) == "parceiro" for r in insucessos)
    tem_sem_detalhe = any(classifica_insucesso(r["finalizador"]) == "sem_detalhe" for r in insucessos)

    partes = []
    if tem_dmnz:
        partes.append("referentes a DMNZ")
    if tem_parceiro:
        partes.append("referentes a parceiro(s)")
    if tem_sem_detalhe:
        partes.append("sem detalhamento de a quem pertence")

    frase = f"Em {node} houve registro de insucesso(s), " + " e ".join(partes)
    if tem_parceiro and not tem_dmnz and not tem_sem_detalhe:
        frase += " (a DMNZ obteve 100% nas entregas)"
    return frase


def frase_mnr_outro_node(dados, n_outro_node):
    """Versão sem reincidência (isso ainda não existe na web - depende do
    histórico de dias anteriores, que é uma etapa futura). Mesma lógica do
    desktop no cenário simples (cenario='1')."""
    tem_mnr = any(r["finalizador"].upper().startswith("MNR") for r in dados)
    if not tem_mnr and n_outro_node == 0:
        return None
    if tem_mnr and n_outro_node > 0:
        base = "houve registros de MNRs e PCT de outro node"
    elif tem_mnr:
        base = "houve registro de MNRs"
    else:
        base = "houve registro de PCT de outro node"
    partes = [base]
    if n_outro_node > 0:
        partes.append("entre os pacotes de outro node, todos já sinalizados ao cliente e aguardando tratativas")
    return ", ".join(partes)


def monta_observacoes(dados, node):
    """Texto automático da caixa 'OBSERVAÇÕES' - mesma lógica do desktop
    (monta_observacoes), sem a parte de reincidência (cenário 2) e sem
    Reversa (painel 6 ainda não integrado na web)."""
    frases = [frase_insucesso(dados, node)]

    avariados = [r for r in dados if any(p in r["finalizador"].lower() for p in PALAVRAS_AVARIA)]
    if avariados:
        frases.append("houve registro de pct(s) avariado(s)")

    perdidos = [r for r in dados if "perdido" in r["finalizador"].lower()]
    if perdidos:
        frases.append("houve registro de pct(s) perdido(s)")

    n_outro_node = len([r for r in dados if "outro node" in r["finalizador"].lower()])
    frase_mnr = frase_mnr_outro_node(dados, n_outro_node)
    if frase_mnr:
        frases.append(frase_mnr)

    texto = ", ".join(frases) + "."
    return texto[0].upper() + texto[1:]


def calcula_tudo(dados, node="LRN9"):
    """Centraliza os calculos de todos os paineis (exceto o 6 - Reversa) a
    partir da lista de linhas ja finalizadas.

    dados: lista de dicts {tbr, state_scc, finalizador, dmnz(bool)} - mesmo
    formato do robo desktop (ler_dados), so que montado direto do CSV do
    SCC + historico de TBR + base de DAs, em vez de ler de um Excel
    intermediario.

    Devolve um dict com os numeros de cada painel, pronto tanto pra previa
    quanto pro dashboard final."""
    contagem_scc, nested = monta_painel_state_scc(dados)

    n_dmnz = sum(1 for r in dados if r["dmnz"])

    total_mnr = sum(1 for r in dados if r["finalizador"].upper().startswith("MNR"))

    outro_node = [r["tbr"] for r in dados if "outro node" in r["finalizador"].lower()]

    insucessos = [r for r in dados if eh_insucesso(r["finalizador"])]
    n_insucesso_dmnz = sum(1 for r in insucessos if classifica_insucesso(r["finalizador"]) == "dmnz")
    n_insucesso_parceiro = sum(1 for r in insucessos if classifica_insucesso(r["finalizador"]) == "parceiro")
    n_insucesso_sem_detalhe = sum(1 for r in insucessos if classifica_insucesso(r["finalizador"]) == "sem_detalhe")

    # auditoria: TBR nao-Delivered que chegou aqui sem State Finalizador
    # preenchido - o State Finalizador deveria ter classificado (automatico
    # ou manual) todo mundo antes de chegar nessa etapa; se sobrou alguem
    # em branco, alguma coisa passou batido.
    sem_finalizador = [
        r["tbr"] for r in dados
        if r["state_scc"].strip().lower() != "delivered" and not r["finalizador"].strip()
    ]

    return {
        "total_geral": len(dados),
        "contagem_scc": contagem_scc,
        "nested": nested,
        "n_dmnz": n_dmnz,
        "n_insucesso_dmnz": n_insucesso_dmnz,
        "n_insucesso_parceiro": n_insucesso_parceiro,
        "n_insucesso_sem_detalhe": n_insucesso_sem_detalhe,
        "n_insucesso_total": len(insucessos),
        "total_mnr": total_mnr,
        "outro_node": outro_node,
        "sem_finalizador": sem_finalizador,
        "observacoes": monta_observacoes(dados, node),
    }

