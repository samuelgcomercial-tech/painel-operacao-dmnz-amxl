# -*- coding: utf-8 -*-
"""
consolida_recebimento.py (versao web)

Painel Recebimento - analisa o MESMO arquivo de rotas (planejado, o que
vem dos indianos, o mesmo que consolida_rotas.py usa) mas com outro
objetivo: em vez de gerar a lista pro SCC, calcula por rota a
QUANTIDADE DE PACOTES e a QUANTIDADE DE PARADAS.

Pacotes = 1 linha (1 Tracking ID) na aba da rota.
Parada = 1 local de entrega - varias entregas no MESMO endereco (ou no
mesmo predio/condominio, so mudando o apartamento) contam como UMA
parada so, nao uma por pacote.

De proposito SEM nenhuma logica de motorista/DA - o Samuel pediu
explicitamente pra ficar de fora, o motorista nao muda a conta de
pacotes/paradas.

--- Como decide "mesmo endereco" (ajuste do Samuel em 21/09/2026) ---
A ideia obvia seria usar Latitude/Longitude (o arquivo ja tem essas
duas colunas prontas). Testei essa ideia direto nos arquivos de rota
reais antes de escrever qualquer linha de codigo, e ela NAO se sustenta
sozinha:

  - Duas entregas do MESMO predio, aptos diferentes, ficam so uns
    3 a 12 metros uma da outra (confirmado com varios exemplos reais -
    ex: "Rua Nossa Senhora de Lourdes 282" com os apartamentos 2001,
    1502 e 402 do MESMO predio, a 4-7 metros de distancia entre si).
  - MAS duas casas DIFERENTES (numeros de rua diferentes, ex: "461" e
    "441" da mesma rua) tambem apareceram a so 2.6 metros uma da outra
    num caso real - vizinhos de porta, nao a mesma entrega.

Ou seja, nao existe uma distancia (metros) que separe com seguranca
"mesmo predio" de "vizinho de porta" - as duas faixas se sobrepoem nos
dados reais. Arredondar Latitude/Longitude pra um numero fixo de casas
decimais (uma alternativa mais simples que calcular distancia) tem o
mesmo problema, PIOR ainda: dois pontos a poucos metros um do outro
podem cair em "caixinhas" de arredondamento diferentes so por estarem
perto da borda (confirmado tambem com os mesmos exemplos reais acima).

Por isso a regra usada aqui e outra: o comeco do texto do endereco (rua
+ numero, tudo antes da primeira virgula - onde normalmente vem
complemento tipo "Apto 402", bairro, cidade), normalizado (sem acento,
sem maiusculo/minusculo, sem espaco duplicado). Testado nos mesmos
arquivos reais, esse metodo separou corretamente TODOS os casos de
"vizinho de porta com numero diferente" e uniu corretamente quase todos
os casos de "mesmo predio, apto diferente" - a unica excecao encontrada
foi um endereco onde o nome do condominio vem ANTES do numero da rua
num dos dois textos (ex: "Avenida X 1044, Ap 1" vs "Avenida X
Condominio Y 1044, Ap 2") - um caso raro de formatacao inconsistente do
proprio texto que veio dos indianos, nao um problema do metodo em si.

Resumindo: nenhum metodo automatico e 100% - mas o texto do endereco
errou bem menos que a distancia/coordenada nos dados reais que testei,
entao foi o escolhido. Se aparecerem muitos casos errados na pratica,
da pra revisar isso depois (o campo CEP/Postal tambem poderia ajudar
como reforco, mas nao foi necessario ate agora).
"""

import io
import re
import unicodedata

import openpyxl
from openpyxl.drawing.image import Image as ExcelImage
from openpyxl.styles import Alignment, Font, PatternFill
from PIL import Image as PILImage

from consolida_rotas import _carrega_planilha, detecta_data_do_arquivo, detecta_node_do_nome


def normaliza_endereco_base(endereco):
    """Pega so 'rua + numero' (tudo antes da primeira virgula) e
    normaliza (sem acento, minusculo, sem espaco duplicado) pra usar
    como chave de agrupamento. Enderecos identicos depois dessa
    normalizacao viram UMA parada so."""
    base = (endereco or "").split(",")[0]
    base = unicodedata.normalize("NFKD", base).encode("ascii", "ignore").decode("ascii")
    base = re.sub(r"\s+", " ", base).strip().lower()
    return base


def consolida_recebimento(arquivo, nome_arquivo):
    """arquivo: objeto vindo do st.file_uploader (mesmo formato do
    consolida_rotas.consolida). Devolve (por_rota, data_do_arquivo):

    por_rota: dict rota -> {
        "pacotes": int,
        "paradas": int,
        "detalhe_paradas": [ {"endereco_base": str, "qtd_pacotes": int,
                               "tbrs": [str, ...]}, ... ]  (so as com
                               mais de 1 pacote, pra revisao visual)
    }
    """
    abas, data_criacao_excel = _carrega_planilha(arquivo, nome_arquivo)

    abas_rota = [n for n in abas if n.lower().startswith("sequencedroute")]
    if not abas_rota:
        raise ValueError(
            "Nao encontrei nenhuma aba comecando com 'sequencedRoute' nesse arquivo. "
            "Confira se e o arquivo certo (o que vem dos indianos)."
        )

    por_rota = {}
    for nome_aba in abas_rota:
        linhas = abas[nome_aba]
        rota = nome_aba.split("_")[-1]  # "sequencedRoute_AX1" -> "AX1"

        paradas = {}  # endereco_base -> [tbr, ...], na ordem em que apareceram
        tbrs_da_rota = []  # todos os TBR da rota, na ordem do arquivo (pra cruzar com o CSV do SCC depois)
        itens = []  # [{"stop":, "tbr":, "endereco":}], na ordem do arquivo - pra consulta TBR+endereco na tela
        pacotes = 0
        for row in linhas[2:]:  # pula linha 1 (metadados) e 2 (cabecalho)
            stop = row[0] if len(row) > 0 else None
            tracking_id = row[1] if len(row) > 1 else None
            endereco = row[5] if len(row) > 5 else None
            if not tracking_id:
                continue
            tbr = str(tracking_id).strip()
            pacotes += 1
            tbrs_da_rota.append(tbr)
            itens.append({"stop": stop, "tbr": tbr, "endereco": endereco or ""})
            chave = normaliza_endereco_base(endereco)
            paradas.setdefault(chave, []).append(tbr)

        detalhe_paradas = [
            {"endereco_base": chave, "qtd_pacotes": len(tbrs), "tbrs": tbrs}
            for chave, tbrs in paradas.items()
            if len(tbrs) > 1
        ]
        detalhe_paradas.sort(key=lambda d: -d["qtd_pacotes"])

        por_rota[rota] = {
            "pacotes": pacotes,
            "paradas": len(paradas),
            "detalhe_paradas": detalhe_paradas,
            "tbrs": tbrs_da_rota,
            "itens": itens,
        }

    data_arquivo = detecta_data_do_arquivo(nome_arquivo, data_criacao_excel)
    node = detecta_node_do_nome(nome_arquivo)
    return por_rota, data_arquivo, node


# --- Checagem real (opcional) ---------------------------------------
#
# A contagem de pacotes/paradas acima e sobre o PLANO (arquivo de
# rotas) - so vira confiavel de verdade quando os pacotes ja passaram
# por todo o fluxo fisico do recebimento (confirmado pelo Samuel em
# 21/09/2026): Manifested -> Inducted -> Stowed -> (Em transito pra
# entrega). Nao da pra pular etapa (ex: nao tem como ficar "Stowed"
# sem antes ter sido "Inducted" - seria uma falha grave no processo).
#
# A fonte pra saber em que etapa cada TBR do plano esta e o MESMO CSV
# exportado do SCC que o resto do app ja usa (mesmas colunas -
# "Tracking ID" e "State" - mesma funcao base_das.le_csv_scc) - a
# unica diferenca e QUANDO ele e exportado: se for logo apos a
# inducao/o stow (em vez de no fim do dia), o "State" vem com esses
# states transitorios de recebimento em vez de Delivered/Received.
#
# Por causa da regra "nao pula etapa", nao precisa saber o texto exato
# de todo state possivel depois do Stow (In Transit, Delivered,
# Cancelado...) - qualquer state que NAO seja um dos "antes da inducao"
# abaixo so pode significar que o pacote ja passou do Stow (se nao
# tivesse passado, o SCC não teria como ter avançado ele pra la). Isso
# poupa ter que confirmar cada state do fim do fluxo com um arquivo de
# exemplo.
#
# BUG real encontrado testando com CSV de verdade do SCC em 23/09/2026
# (o exemplo sintetico anterior so tinha o texto puro "Manifested", que
# nunca apareceu no CSV real): a etapa "antes da inducao" tem MAIS
# variantes de texto do que só "Manifested" - viu-se no mesmo CSV
# "Manifested (FC -> DS)" (pacote ainda em transporte da FC pra
# estacao, nem chegou fisicamente ainda), "Arrived" (chegou no dock) e
# "Received" (registrado no sistema da estacao) - todos ainda ANTES da
# inducao de verdade. Com a comparacao exata antiga ("state ==
# 'manifested'"), esses 3 caiam por engano em 'stowed_ou_alem' -
# testado no CSV real de 524 pacotes: 14 pacotes (6 Manifested(FC->DS)
# + 5 Arrived + 3 Received) foram contados errado como "ja induzido ou
# além" quando na verdade nem chegaram na estacao ainda. Por isso a
# checagem de "antes da inducao" agora usa prefixo (pega qualquer
# variante que comece com "manifested") + essas duas palavras exatas,
# em vez de só uma comparacao exata com "manifested".
# --- MNR (ajuste do Samuel em 23/09/2026) ----------------------------
#
# Termo que ja existe no Fechamento (calculo_fechamento.py conta
# "finalizador" comecando com "MNR", digitado a mao pelo lider na
# revisao do State Finalizador) - explicado aqui pra usar a MESMA
# palavra no Recebimento, em vez de inventar outra ("nao encontrado")
# pra descrever a mesma coisa.
#
# O plano dos indianos (arquivo de rotas) puxa pacote pelo FATURAMENTO,
# nao pelo manifesto de transporte - ou seja, um TBR pode aparecer
# roteirizado no plano mesmo sem nunca ter sido de fato manifestado pro
# hub. Duas formas de isso virar MNR:
#
#   1. Roteirizado mas NUNCA aparece em nenhum CSV do SCC do
#      checkpoint (nem inducao, nem stow) - nunca chegou a ser
#      manifestado pro hub de verdade. E o "nao_encontrado" de
#      cruza_com_state_scc/soma_por_hub.
#   2. Aparece no CSV (foi manifestado) mas continua preso em
#      Manifested/Arrived/Received mesmo DEPOIS do checkpoint de stow
#      ja ter rodado - so significa "nao chegou a ser carregado" quando
#      o stow ja aconteceu (avaria, extravio, ou so nao deu tempo no
#      hub de origem); enquanto so o checkpoint de INDUCAO rodou, um
#      TBR em Manifested e normal (so ainda nao chegou a vez dele),
#      nao e sinal de MNR nenhum ainda.
#
# Os dois casos so viram rotulo/aviso na tela e no Excel - NENHUM dos
# dois trava a confirmacao do checkpoint (ver nota em cruza_com_state_scc
# e o botao de ressalva em app.py, que ja existe pra seguir em frente
# mesmo com essa divergencia).
ROTULO_MNR_NAO_MANIFESTADO = "Possível MNR — não veio no manifesto (roteirizado, mas nunca chegou a ser manifestado no hub)"
ROTULO_MNR_NAO_CARREGADO = "Possível MNR — manifestado, mas não chegou a ser carregado (avaria, extravio ou falta de tempo)"
ROTULO_AGUARDANDO_INDUCAO = "Ainda não induzido (aguardando — normal enquanto só o checkpoint de indução rodou)"

PREFIXOS_ANTES_DA_INDUCAO = ("manifested",)
ESTADOS_ANTES_DA_INDUCAO = ("arrived", "received")


def categoriza_state_recebimento(state_texto):
    """Devolve 'manifested' (ainda nao chegou/nao foi induzido - ver
    comentario acima pras variantes de texto que caem aqui), 'inducted',
    'stowed_ou_alem' (Stowed ou qualquer state posterior) ou
    'sem_state' (state vazio, mas TBR foi encontrado no CSV - raro,
    mesmo tratamento de state vazio que o resto do app da)."""
    state = (state_texto or "").strip().lower()
    if not state:
        return "sem_state"
    if state == "inducted":
        return "inducted"
    if state.startswith(PREFIXOS_ANTES_DA_INDUCAO) or state in ESTADOS_ANTES_DA_INDUCAO:
        return "manifested"
    return "stowed_ou_alem"


def cruza_com_state_scc(por_rota, linhas_csv, chave="situacao_real"):
    """linhas_csv: lista de dicts vinda de base_das.le_csv_scc, do MESMO
    CSV do SCC de sempre (so que exportado logo apos inducao/stow, com
    o State transitorio - ver comentario acima).

    'chave': nome da chave onde guardar o resultado dentro de cada rota
    - usado pra guardar a checagem da inducao ("situacao_inducao") e a
    do stow ("situacao_stow") separadas, ja que sao dois momentos
    diferentes do turno (ver PAINEL_RECEBIMENTO_WIZARD em app.py).

    Acrescenta em cada rota de por_rota a chave pedida:
        {
            "manifested": [tbr, ...],       # ainda nao foi induzido
            "inducted": [tbr, ...],         # induzido, falta armazenar (stow)
            "stowed_ou_alem": [tbr, ...],   # ja armazenado (ou mais adiante)
            "sem_state": [tbr, ...],        # achou o TBR no CSV mas sem State
            "nao_encontrado": [tbr, ...],   # TBR do plano que nem apareceu no CSV
            "confirmada": bool,             # ver nota abaixo
        }

    Nota sobre 'confirmada' (ajuste depois de testar com dados reais em
    21/09/2026, motivo corrigido em 23/09/2026 - ver comentario grande
    logo abaixo sobre MNR): SO fica False por causa de
    manifested/inducted/sem_state presos no meio do processo -
    "nao_encontrado" (planejado mas nunca apareceu em NENHUM CSV) NAO
    trava a confirmacao. Motivo real (explicado pelo Samuel em
    23/09/2026): o plano dos indianos puxa pacote pelo FATURAMENTO, nao
    pelo manifesto - um TBR pode estar roteirizado sem nunca ter sido
    de fato manifestado pro hub (isso e um MNR, mesmo termo ja usado no
    Fechamento - ver PENDENTE_ROTULO_MNR_NAO_MANIFESTADO abaixo). Isso
    e normal (nao e falha de processo do RECEBIMENTO - a causa e lá na
    ponta do faturamento/roteirizacao), entao nao devia impedir a rota
    de ser marcada como "concluida" pros que DE FATO chegaram. E
    exatamente por isso que "nao_encontrado" fica separado (pra dar pra
    ver quem sumiu, rotulado como possivel MNR no Excel), mas fora do
    calculo de 'confirmada'.

    Devolve o proprio por_rota (modificado in-place, alem de devolvido -
    fica explicito no retorno pra quem chama nao precisar adivinhar)."""
    estado_por_tbr = {}
    for linha in linhas_csv:
        tbr = (linha.get("Tracking ID") or "").strip()
        if tbr:
            estado_por_tbr[tbr] = linha.get("State")

    for info in por_rota.values():
        situacao = {
            "manifested": [], "inducted": [], "stowed_ou_alem": [],
            "sem_state": [], "nao_encontrado": [],
        }
        for tbr in info["tbrs"]:
            if tbr not in estado_por_tbr:
                situacao["nao_encontrado"].append(tbr)
            else:
                situacao[categoriza_state_recebimento(estado_por_tbr[tbr])].append(tbr)
        situacao["confirmada"] = not (
            situacao["manifested"] or situacao["inducted"] or situacao["sem_state"]
        )
        info[chave] = situacao

    return por_rota


def soma_por_hub(linhas_csv, hubs):
    """Conta o progresso de indução/stow POR HUB, usando a coluna
    'Source' que o próprio CSV do SCC já traz (REC9/FOR3 por TBR) - em
    vez de depender do arquivo de rotas, que NÃO distingue hub (ajuste
    do Samuel em 23/09/2026: cada hub manda um veículo separado, em
    horários bem diferentes - viu-se em dois Bill of Lading reais com
    mais de 4h de diferença entre a saída de um e a saída do outro).

    Isso existe porque a checagem por ROTA (cruza_com_state_scc) mistura
    os dois hubs - o que é correto pro total combinado, mas não ajuda a
    responder "o REC9 já terminou?" enquanto o FOR3 ainda nem chegou (e
    vice-versa). Como o progresso aqui vem DIRETO do CSV mais recente
    (não acumula entre uploads), quem chama deve guardar o resultado em
    session_state pra não perder o progresso de um hub quando o outro
    hub for atualizado depois (ver tela_recebimento em app.py).

    Devolve dict hub -> {"manifested": [tbr, ...] (ainda não induzido),
    "inducted_apenas": [tbr, ...] (induzido, falta armazenar),
    "inducido_ou_alem": int, "stowed_ou_alem": int, "sem_state": int} -
    TBRs cujo Source não bate com nenhum hub conhecido (ex.: NA,
    "CUSTOMER_ADDRESS") são ignorados, não contam em nenhum hub."""
    resultado = {
        hub: {"manifested": [], "inducted_apenas": [], "inducido_ou_alem": 0, "stowed_ou_alem": 0, "sem_state": 0}
        for hub in hubs
    }
    for linha in linhas_csv:
        hub = (linha.get("Source") or "").strip()
        if hub not in resultado:
            continue
        tbr = (linha.get("Tracking ID") or "").strip()
        categoria = categoriza_state_recebimento(linha.get("State"))
        if categoria == "manifested":
            resultado[hub]["manifested"].append(tbr)
        elif categoria == "sem_state":
            resultado[hub]["sem_state"] += 1
        elif categoria == "inducted":
            resultado[hub]["inducted_apenas"].append(tbr)
            resultado[hub]["inducido_ou_alem"] += 1
        elif categoria == "stowed_ou_alem":
            resultado[hub]["inducido_ou_alem"] += 1
            resultado[hub]["stowed_ou_alem"] += 1
    return resultado


# Limiar (%) acima do qual avisamos que o CSV subido num checkpoint
# pode ser, na verdade, um export de FIM DE TURNO (o mesmo tipo que
# alimenta o Fechamento) em vez de um checkpoint de recebimento de
# verdade - ver percentual_delivered() logo abaixo. Um valor baixo
# (poucos % de Delivered) pode acontecer de verdade se o checkpoint
# foi capturado um pouco atrasado; um valor alto e outra historia.
LIMITE_ALERTA_PCT_DELIVERED = 10


def percentual_delivered(por_rota, linhas_csv):
    """Calcula, entre os TBRs do PLANO (por_rota) que aparecem nesse CSV,
    qual percentual ja esta literalmente em 'Delivered'.

    Serve de sinal de alerta pro checkpoint de recebimento (pedido do
    Samuel em 22/09/2026, depois de testar com um CSV errado por
    engano): fisicamente NAO da pra ter pacote Delivered no ato do
    recebimento - a entrega so acontece horas depois, depois que a rota
    inteira (inducao -> stow -> saida) ja aconteceu. Entao se um
    percentual alto do CSV ja vem Delivered, o mais provavel e que esse
    arquivo seja um export de FIM DE TURNO (o mesmo tipo que ja
    alimenta o Fechamento), nao um checkpoint de verdade - so um aviso,
    nao trava nada (o usuario pode ter capturado o checkpoint atrasado
    de proposito, ou so estar testando).

    Devolve (percentual: float de 0 a 100, total_encontrados: int). Se
    nao achar nenhum TBR do plano nesse CSV, devolve (0.0, 0) - sem
    risco de divisao por zero."""
    tbrs_do_plano = {tbr for info in por_rota.values() for tbr in info["tbrs"]}
    estado_por_tbr = {}
    for linha in linhas_csv:
        tbr = (linha.get("Tracking ID") or "").strip()
        if tbr:
            estado_por_tbr[tbr] = linha.get("State")
    encontrados = [estado_por_tbr[tbr] for tbr in tbrs_do_plano if tbr in estado_por_tbr]
    if not encontrados:
        return 0.0, 0
    delivered = sum(1 for s in encontrados if (s or "").strip().lower() == "delivered")
    return (delivered / len(encontrados)) * 100, len(encontrados)


def soma_inducidos_ou_alem(por_rota, chave="situacao_inducao"):
    """Total (todas as rotas somadas) de TBR que ja passaram da etapa
    Manifested - ou seja, ja estao Inducted ou mais adiante (Stowed,
    etc). Usado pra comparar com o numero digitado a mao ("quantos
    pacotes chegaram no manifesto de transporte") no checkpoint de
    pos-inducao."""
    total = 0
    for info in por_rota.values():
        sit = info.get(chave)
        if sit:
            total += len(sit["inducted"]) + len(sit["stowed_ou_alem"])
    return total


def soma_stowed_ou_alem(por_rota, chave="situacao_stow"):
    """Total (todas as rotas somadas) de TBR que ja estao Stowed (ou
    mais adiante). Usado no checkpoint de pos-stow, comparado contra o
    total confirmado na inducao (ninguem deveria "sumir" entre um
    checkpoint e outro)."""
    total = 0
    for info in por_rota.values():
        sit = info.get(chave)
        if sit:
            total += len(sit["stowed_ou_alem"])
    return total


# --- Ressalva do líder (ajuste do Samuel em 21/09/2026) --------------
#
# Quando o Inducted vem MENOR que o Manifested (o numero digitado a
# mao, vindo do Bill of Lading/manifesto de transporte de verdade - o
# Samuel mandou foto de um), isso pode ser um erro de processo (faltou
# induzir alguem) OU pode ser que o fisico realmente veio a menos do
# que o proprio manifesto da Amazon diz (acontece: o caminhao chegou
# com menos pacote do que o documento afirma). Nesse segundo caso NAO
# FAZ SENTIDO travar o fluxo esperando um numero que nunca vai bater -
# o lider reconfere pacote a pacote e, se continuar diferente, registra
# uma RESSALVA (texto explicando + opcionalmente uma foto do proprio
# Manifested) pra seguir mesmo assim, em vez de ficar preso.
def monta_texto_ressalva_padrao(node, manifested, fisico):
    """Texto padrao da ressalva, ja preenchido com os numeros - o
    lider pode editar antes de confirmar (ver tela_recebimento em
    app.py)."""
    return (
        f"A quantidade de pacotes manifestada para o {node} veio a menos que o "
        "informado. Foi feita uma nova indução pacote a pacote para validar o físico "
        "x sistema e, ainda assim, o número se manteve. Segue a quantidade do "
        "Manifested e a que veio físico:\n\n"
        f"{manifested} Manifested\n"
        f"{fisico} físico"
    )


def _situacao_mais_recente(info):
    """Pega o checkpoint mais avancado que essa rota ja tem: stow (mais
    recente) > inducao > nenhum. Usado pra decidir o que mostrar no
    Excel sem o chamador precisar saber em que checkpoint o usuario
    parou."""
    return info.get("situacao_stow") or info.get("situacao_inducao")


def gera_workbook_recebimento(por_rota, node, data_arquivo, ressalvas=None, qtd_por_hub=None):
    """Monta o Excel final: aba 'Resumo' (pacotes/paradas por rota, com
    total no fim) + aba 'Detalhe paradas' (so as paradas com mais de 1
    pacote, pra dar pra conferir na mao se o agrupamento fez sentido -
    ver o comentario grande no topo do arquivo sobre a regra usada e a
    limitacao dela). Se alguma rota tiver checagem real (cruzou com o
    CSV do SCC na inducao e/ou no stow - ver cruza_com_state_scc), a
    aba Resumo ganha as colunas extras (do checkpoint mais avancado
    disponivel) e entra uma aba 'Pendentes' com o detalhe de quem
    ainda nao chegou em Stowed.

    ressalvas: lista opcional de dicts {"checkpoint": str, "texto": str,
    "foto_nome": str|None, "foto_bytes": bytes|None} - vira uma aba
    'Ressalvas' com o texto e a foto do Manifested embutida (quando
    tiver), pra registro/auditoria (ver monta_texto_ressalva_padrao).

    qtd_por_hub: dict opcional {hub: quantidade} (ex.: {"REC9": 314,
    "FOR3": 210}) - a quantidade fisica digitada por hub (ajuste do
    Samuel em 23/09/2026: o node recebe manifesto de hubs separados,
    cada um por veiculo proprio). Nao tem como quebrar isso por ROTA no
    Excel (o arquivo de rotas nao distingue hub por TBR), entao entra
    como uma linha de registro logo abaixo do titulo, so pra ficar
    documentado no arquivo final quanto veio de cada hub."""
    tem_situacao_real = any(_situacao_mais_recente(info) for info in por_rota.values())

    wb = openpyxl.Workbook()

    ws = wb.active
    ws.title = "Resumo"
    titulo = f"Painel Recebimento - {node or '?'} - {data_arquivo.strftime('%d/%m/%Y') if data_arquivo else '?'}"
    ws.append([titulo])
    ws["A1"].font = Font(bold=True, size=13)
    if qtd_por_hub:
        total_hubs = sum(qtd_por_hub.values())
        linha_hubs = " · ".join(f"{hub}: {qtd}" for hub, qtd in qtd_por_hub.items())
        ws.append([f"Recebido por hub — {linha_hubs} · Total: {total_hubs}"])
        ws[f"A{ws.max_row}"].font = Font(italic=True)
    ws.append([])
    cabecalho = ["Rota", "Pacotes", "Paradas"]
    if tem_situacao_real:
        cabecalho += ["Manifested", "Inducted", "Stowed ou além", "Não encontrado", "Confirmada?"]
    ws.append(cabecalho)
    for cel in ws[ws.max_row]:
        cel.font = Font(bold=True)
        cel.fill = PatternFill(fill_type="solid", fgColor="FFE8E8E8")

    total_pacotes = total_paradas = 0
    for rota in sorted(por_rota):
        info = por_rota[rota]
        linha = [rota, info["pacotes"], info["paradas"]]
        if tem_situacao_real:
            sit = _situacao_mais_recente(info)
            if sit:
                linha += [
                    len(sit["manifested"]), len(sit["inducted"]), len(sit["stowed_ou_alem"]),
                    len(sit["nao_encontrado"]) + len(sit["sem_state"]),
                    "Sim" if sit["confirmada"] else "Não",
                ]
            else:
                linha += ["-", "-", "-", "-", "-"]
        ws.append(linha)
        total_pacotes += info["pacotes"]
        total_paradas += info["paradas"]
    ws.append(["TOTAL", total_pacotes, total_paradas])
    for cel in ws[ws.max_row]:
        cel.font = Font(bold=True)

    ws.column_dimensions["A"].width = 14
    ws.column_dimensions["B"].width = 12
    ws.column_dimensions["C"].width = 12
    if tem_situacao_real:
        for letra in ["D", "E", "F", "G", "H"]:
            ws.column_dimensions[letra].width = 14

    if tem_situacao_real:
        ws_pend = wb.create_sheet("Pendentes")
        ws_pend.append(["Rota", "Situação", "TBR"])
        for cel in ws_pend[1]:
            cel.font = Font(bold=True)
            cel.fill = PatternFill(fill_type="solid", fgColor="FFE8E8E8")
        for rota in sorted(por_rota):
            info = por_rota[rota]
            sit = _situacao_mais_recente(info)
            if not sit:
                continue
            # Rótulo do "manifested" muda de acordo com o checkpoint mais
            # avançado que essa rota já tem (ver comentário sobre MNR
            # acima, perto de categoriza_state_recebimento): só DEPOIS do
            # stow já ter rodado é que "ainda em Manifested" vira sinal de
            # possível MNR de verdade - antes disso (só indução feita) é
            # normal, ainda não chegou a vez desse TBR.
            rotulo_manifested = (
                ROTULO_MNR_NAO_CARREGADO if "situacao_stow" in info else ROTULO_AGUARDANDO_INDUCAO
            )
            rotulos = {
                "manifested": rotulo_manifested,
                "inducted": "Induzido, falta armazenar (Inducted)",
                "sem_state": "Achado no CSV mas sem State",
                "nao_encontrado": ROTULO_MNR_NAO_MANIFESTADO,
            }
            for chave, rotulo in rotulos.items():
                for tbr in sit[chave]:
                    ws_pend.append([rota, rotulo, tbr])
        ws_pend.column_dimensions["A"].width = 10
        ws_pend.column_dimensions["B"].width = 55
        ws_pend.column_dimensions["C"].width = 18

    ws2 = wb.create_sheet("Detalhe paradas")
    ws2.append(["Rota", "Endereço (chave usada p/ agrupar)", "Qtd pacotes", "TBRs"])
    for cel in ws2[1]:
        cel.font = Font(bold=True)
        cel.fill = PatternFill(fill_type="solid", fgColor="FFE8E8E8")
    for rota in sorted(por_rota):
        for d in por_rota[rota]["detalhe_paradas"]:
            ws2.append([rota, d["endereco_base"], d["qtd_pacotes"], ", ".join(d["tbrs"])])
    ws2.column_dimensions["A"].width = 10
    ws2.column_dimensions["B"].width = 42
    ws2.column_dimensions["C"].width = 12
    ws2.column_dimensions["D"].width = 60

    if ressalvas:
        ws_res = wb.create_sheet("Ressalvas")
        linha_atual = 1
        for r in ressalvas:
            cel = ws_res.cell(row=linha_atual, column=1, value=f"Checkpoint: {r['checkpoint']}")
            cel.font = Font(bold=True)
            linha_atual += 1
            ws_res.cell(row=linha_atual, column=1, value=r["texto"])
            ws_res.cell(row=linha_atual, column=1).alignment = Alignment(wrap_text=True, vertical="top")
            ws_res.merge_cells(start_row=linha_atual, start_column=1, end_row=linha_atual, end_column=6)
            ws_res.row_dimensions[linha_atual].height = 90
            linha_atual += 2
            if r.get("foto_bytes"):
                try:
                    img_pil = PILImage.open(io.BytesIO(r["foto_bytes"]))
                    img_pil.thumbnail((500, 700))
                    img_buffer = io.BytesIO()
                    img_pil.convert("RGB").save(img_buffer, format="PNG")
                    img_buffer.seek(0)
                    img_excel = ExcelImage(img_buffer)
                    ws_res.add_image(img_excel, f"A{linha_atual}")
                    linha_atual += int(img_pil.height / 18) + 3
                except Exception:
                    ws_res.cell(row=linha_atual, column=1, value=f"(não consegui anexar a foto {r.get('foto_nome', '')})")
                    linha_atual += 2
            linha_atual += 2
        ws_res.column_dimensions["A"].width = 70

    buffer = io.BytesIO()
    wb.save(buffer)
    return buffer.getvalue()
