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

from state_finalizador import monta_coluna_das_dmnz, normaliza_state_scc
import parceiros as parceiros_mod


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
        # normaliza_state_scc troca State vazio por "SEM STATE SCC" -
        # pedido do Samuel em 14/09/2026, mesma funcao usada em
        # state_finalizador.processa() pra nunca sobrar rotulo de grupo
        # em branco (aqui, na categoria do painel 1/3 e no nome da
        # coluna "State" do Excel).
        state_scc = normaliza_state_scc(l.get("State"))
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


def classifica_insucesso(finalizador, parceiros):
    """Devolve o NOME de quem é o insucesso: o nome cadastrado como
    'própria empresa' (ex: 'DMNZ', ou o que o node cadastrar), o nome de
    um parceiro terceirizado cadastrado (ex: 'MRIZ'), ou 'sem_detalhe'
    quando o texto não cita nenhum nome cadastrado (usuário digitou só
    'Insucesso', ou citou algo que ainda não está na lista - possível
    erro de digitação ou parceiro novo ainda não cadastrado).

    'parceiros': lista de dicts {nome, propria_empresa} vinda de
    parceiros.py/parceiros.csv (por node). Antes disso a regra era fixa
    (só sabia dizer 'dmnz' vs 'parceiro' genérico, sem saber QUAL
    parceiro, e travada na palavra literal 'dmnz') - agora usa o cadastro
    de verdade, que o Samuel decidiu manter por node (cada node pode
    chamar a própria empresa de um jeito diferente, ex: 'Domina').

    Se 'parceiros' vier vazio (node ainda sem cadastro), cai pro
    comportamento antigo - só pra não quebrar quem ainda não cadastrou
    nada; assim que existir cadastro pra esse node, essa branch nunca
    mais roda."""
    if not parceiros:
        txt = finalizador.strip().lower()
        if "dmnz" in txt:
            return "DMNZ"
        if txt == "insucesso":
            return "sem_detalhe"
        return "parceiro"

    nome_propria = parceiros_mod.nome_propria_empresa(parceiros)
    parceiros_propria = [p for p in parceiros if p["propria_empresa"]]
    parceiros_terceiros = [p for p in parceiros if not p["propria_empresa"]]

    if parceiros_propria and parceiros_mod.identifica_parceiro_no_texto(finalizador, parceiros_propria):
        return nome_propria

    achado = parceiros_mod.identifica_parceiro_no_texto(finalizador, parceiros_terceiros)
    if achado:
        return achado

    return "sem_detalhe"


PALAVRAS_AVARIA = ["avaria", "quebrado", "vazando", "rasgado"]


def frase_insucesso(dados, node, parceiros):
    """Detalha o insucesso citando o nome de cada empresa envolvida de
    verdade (ex: 'referentes a DMNZ e referentes a MRIZ'), em vez do
    genérico 'parceiro(s)' de antes - agora que dá pra saber QUAL
    parceiro, graças ao cadastro."""
    insucessos = [r for r in dados if eh_insucesso(r["finalizador"])]
    if not insucessos:
        return f"Em {node} não houve registros de insucessos"

    classificados = [classifica_insucesso(r["finalizador"], parceiros) for r in insucessos]
    nomes_especificos = sorted({c for c in classificados if c != "sem_detalhe"})
    tem_sem_detalhe = "sem_detalhe" in classificados

    partes = [f"referentes a {nome}" for nome in nomes_especificos]
    if tem_sem_detalhe:
        partes.append("sem detalhamento de a quem pertence")

    frase = f"Em {node} houve registro de insucesso(s), " + " e ".join(partes)

    nome_propria = parceiros_mod.nome_propria_empresa(parceiros) or "DMNZ"
    so_terceiros = (
        nomes_especificos and nome_propria not in nomes_especificos and not tem_sem_detalhe
    )
    if so_terceiros:
        frase += f" (a {nome_propria} obteve 100% nas entregas)"
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


def monta_observacoes(dados, node, parceiros):
    """Texto automático da caixa 'OBSERVAÇÕES' - mesma lógica do desktop
    (monta_observacoes), sem a parte de reincidência (cenário 2) e sem
    Reversa (painel 6 ainda não integrado na web)."""
    frases = [frase_insucesso(dados, node, parceiros)]

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

    # Pct de NA cuja rota a Amazon já atribuiu (state_finalizador.py,
    # "PCT NA - <rota>") - só acompanhamento do dia (não é insucesso nem
    # confirma DMNZ/parceiro), mas lista os TBRs pra não ficar escondido
    # dia após dia (mesmo limite de 10 do aviso de auditoria, por
    # consistência). Só entra aqui quem já saiu de "pendente" (a versão
    # sem Route Code continua pedindo revisão manual - ver eh_na_com_rota).
    pct_na = [r["tbr"] for r in dados if r["finalizador"].upper().startswith("PCT NA")]
    if pct_na:
        frases.append(
            f"{len(pct_na)} pct(s) de NA em acompanhamento "
            f"({', '.join(pct_na[:10])}"
            + (", ..." if len(pct_na) > 10 else "") + ")"
        )

    texto = ", ".join(frases) + "."
    return texto[0].upper() + texto[1:]


def calcula_tudo(dados, node="LRN9", parceiros=None):
    """Centraliza os calculos de todos os paineis (exceto o 6 - Reversa) a
    partir da lista de linhas ja finalizadas.

    dados: lista de dicts {tbr, state_scc, finalizador, dmnz(bool)} - mesmo
    formato do robo desktop (ler_dados), so que montado direto do CSV do
    SCC + historico de TBR + base de DAs, em vez de ler de um Excel
    intermediario.

    parceiros: lista de dicts {nome, propria_empresa} do cadastro desse
    node (parceiros.py) - se vier vazio/None, a Análise Insucesso cai pro
    comportamento antigo (só DMNZ vs "parceiro" genérico, ver
    classifica_insucesso).

    Devolve um dict com os numeros de cada painel, pronto tanto pra previa
    quanto pro dashboard final."""
    parceiros = parceiros or []
    contagem_scc, nested = monta_painel_state_scc(dados)

    n_dmnz = sum(1 for r in dados if r["dmnz"])

    total_mnr = sum(1 for r in dados if r["finalizador"].upper().startswith("MNR"))

    outro_node = [r["tbr"] for r in dados if "outro node" in r["finalizador"].lower()]

    insucessos = [r for r in dados if eh_insucesso(r["finalizador"])]
    # Antes eram 3 contadores fixos (DMNZ / parceiro genérico / sem
    # detalhe) - agora é uma contagem por NOME (ex: {"DMNZ": 2, "MRIZ": 2}),
    # já que classifica_insucesso sabe dizer qual parceiro é, não só que
    # "é um parceiro".
    contagem_insucesso = Counter(
        classifica_insucesso(r["finalizador"], parceiros) for r in insucessos
    )

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
        "contagem_insucesso": contagem_insucesso,
        "n_insucesso_total": len(insucessos),
        "total_mnr": total_mnr,
        "outro_node": outro_node,
        "sem_finalizador": sem_finalizador,
        "observacoes": monta_observacoes(dados, node, parceiros),
    }
