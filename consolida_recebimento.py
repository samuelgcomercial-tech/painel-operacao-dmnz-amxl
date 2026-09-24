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

import html
import io
import re
import unicodedata

import openpyxl
from openpyxl.drawing.image import Image as ExcelImage
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from PIL import Image as PILImage

from consolida_rotas import _carrega_planilha, detecta_data_do_arquivo, detecta_node_do_nome


def normaliza_endereco_base(endereco):
    """Pega so 'rua + numero' (tudo antes da primeira virgula) e
    normaliza (sem acento, minusculo, sem espaco duplicado) pra usar
    como chave de agrupamento. Enderecos identicos depois dessa
    normalizacao viram UMA parada so.

    Faz html.unescape() ANTES de normalizar (bug encontrado pelo Samuel
    em 24/09/2026: o arquivo de rotas as vezes traz o endereco com
    entidade HTML tipo "jer&ocirc;nimo c&acirc;mara" em vez de "jerônimo
    câmara" - sem o unescape, essas letras ficam intactas (ja sao ascii)
    e o endereco vira uma chave DIFERENTE de "avenida jeronimo camara",
    quebrando o agrupamento em 2 paradas quando era 1 so)."""
    base = html.unescape(endereco or "").split(",")[0]
    base = unicodedata.normalize("NFKD", base).encode("ascii", "ignore").decode("ascii")
    base = re.sub(r"\s+", " ", base).strip().lower()
    return base


def endereco_para_exibicao(endereco):
    """Versao do endereco pra MOSTRAR pro usuario (no Excel/tela) - so
    desfaz a entidade HTML e tira espaco/virgula sobrando nas pontas,
    mas MANTEM acento e maiusculo/minusculo (ao contrario da chave de
    agrupamento acima, que e só pra comparar, não pra ler)."""
    return html.unescape(endereco or "").strip().strip(",").strip()


def consolida_recebimento(arquivo, nome_arquivo):
    """arquivo: objeto vindo do st.file_uploader (mesmo formato do
    consolida_rotas.consolida). Devolve (por_rota, data_do_arquivo):

    por_rota: dict rota -> {
        "pacotes": int,
        "paradas": int,
        "detalhe_paradas": [ {"endereco_base": str, "endereco_exibicao": str,
                               "qtd_pacotes": int, "tbrs": [str, ...]}, ... ]
                               (so as paradas com mais de 1 pacote - usado
                               no expander "Conferir agrupamento" da tela,
                               que e so pra revisar se o agrupamento fez
                               sentido, nao precisa listar parada de 1 so),
        "todas_paradas": [ mesma coisa acima, mas TODAS as paradas,
                            inclusive as de 1 pacote so - usado na aba
                            "Detalhe paradas" do Excel exportado (ajuste
                            do Samuel em 24/09/2026: reparou que a aba
                            exportada nao trazia todos os TBR - o
                            expander da tela sempre foi só um resumo de
                            conferência, mas o Excel final precisa ter
                            TODOS, pra servir de registro completo) ]
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

        paradas = {}  # endereco_base -> {"endereco_exibicao": str, "tbrs": [tbr, ...]}, na ordem em que apareceram
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
            grupo = paradas.setdefault(
                chave, {"endereco_exibicao": endereco_para_exibicao(endereco), "tbrs": []}
            )
            grupo["tbrs"].append(tbr)

        todas_paradas = [
            {
                "endereco_base": chave,
                "endereco_exibicao": grupo["endereco_exibicao"],
                "qtd_pacotes": len(grupo["tbrs"]),
                "tbrs": grupo["tbrs"],
            }
            for chave, grupo in paradas.items()
        ]
        todas_paradas.sort(key=lambda d: -d["qtd_pacotes"])
        detalhe_paradas = [d for d in todas_paradas if d["qtd_pacotes"] > 1]

        por_rota[rota] = {
            "pacotes": pacotes,
            "paradas": len(paradas),
            "detalhe_paradas": detalhe_paradas,
            "todas_paradas": todas_paradas,
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
#      cruza_com_state_scc.
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


# Removida em 24/09/2026 (pedido do Samuel): existia aqui uma
# soma_por_hub(linhas_csv, hubs), que contava indução/stow separado por
# REC9/FOR3 usando a coluna "Source" do CSV do SCC. Motivo da remoção:
# "Source" não é um rótulo fixo de hub de origem - é o local do ÚLTIMO
# SCAN (um TBR já Delivered aparece com Source "CUSTOMER_ADDRESS", não
# o hub de onde saiu - visto de verdade num CSV real: 6 TBR assim + 1
# com "CGH7", nenhum de outro hub de verdade). Num arquivo com REC9 e
# FOR3 misturados não tem como saber a qual dos dois um TBR com Source
# "sujo" pertence - contar diferente por hub em cima dessa coluna não
# era fiel à realidade. A checagem por ROTA (cruza_com_state_scc, logo
# acima) já cobre o caso combinado e continua sendo a única forma de
# checagem real usada pelo app.


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


def monta_texto_copia_mnr(por_rota):
    """Monta o texto 'empilhado' (um TBR por linha, cada um com ':' no
    final) dos TBRs que ficaram como possivel MNR no checkpoint mais
    avancado disponivel de cada rota (STOW, se ja rodou - e so faz
    sentido de verdade DEPOIS do stow, ver comentario sobre MNR mais
    acima) - pensado pra:

      1. Mostrar num st.code() com icone de copiar (mesmo padrao ja
         usado na Etapa 1 do Fechamento), pro lider copiar a lista
         inteira de uma vez.
      2. Vir pre-preenchido na caixa de texto da ressalva final, pro
         lider so completar a analise depois de cada ':'.

    Ajuste do Samuel em 24/09/2026: substitui a aba 'Pendentes' que
    existia no Excel - em vez de so listar os TBR num Excel a parte
    (sem espaco pra explicar o motivo de cada um), agora o proprio
    lider detalha a analise de cada TBR na ressalva.

    So considera rotas que ja tem 'situacao_stow' (rota sem stow ainda
    nao entra aqui - nao faz sentido cobrar analise de MNR de quem
    ainda nem chegou nessa etapa). Devolve "" se nenhuma rota tiver
    stow ainda."""
    nao_carregado, nao_manifestado, sem_state = [], [], []
    for info in por_rota.values():
        sit = info.get("situacao_stow")
        if not sit:
            continue
        nao_carregado += sit["manifested"]
        nao_manifestado += sit["nao_encontrado"]
        sem_state += sit["sem_state"]

    blocos = []
    if nao_carregado:
        blocos.append(f"{ROTULO_MNR_NAO_CARREGADO}:\n" + "\n".join(f"{tbr}: " for tbr in nao_carregado))
    if nao_manifestado:
        blocos.append(f"{ROTULO_MNR_NAO_MANIFESTADO}:\n" + "\n".join(f"{tbr}: " for tbr in nao_manifestado))
    if sem_state:
        blocos.append("Achado no CSV mas sem State:\n" + "\n".join(f"{tbr}: " for tbr in sem_state))
    return "\n\n".join(blocos)


def _estima_altura_linha_ressalva(texto, largura_caracteres=95, altura_por_linha=15, minimo=90):
    """Estima a altura (em pontos) que a linha da aba 'Ressalvas' precisa
    pra mostrar o texto INTEIRO, sem cortar (bug real visto pelo Samuel
    em 24/09/2026: a altura ficava travada em 90 - dava pra ver só as
    primeiras linhas de uma ressalva mais longa, tipo quando junta a
    lista empilhada de vários TBR possível MNR, o resto ficava
    escondido).

    openpyxl NAO calcula altura de linha automaticamente pra texto com
    wrap - precisa estimar na mão quantas linhas o texto vai ocupar
    (contando tanto as quebras de linha explícitas quanto o texto
    "dando a volta" dentro da largura da célula) e converter em altura.

    largura_caracteres: quantos caracteres cabem numa linha antes do
    Excel quebrar sozinho - estimativa conservadora (arredonda PRA
    CIMA o número de linhas) pra largura das colunas A:F mescladas
    (A tem width=70 + 5 colunas no width padrão) - melhor estimar
    linha a mais (sobra espaço em branco) do que a menos (corta
    texto)."""
    linhas = 0
    for linha_texto in (texto or "").split("\n"):
        linhas += max(1, -(-len(linha_texto) // largura_caracteres))  # divisao arredondando pra cima
    return max(linhas * altura_por_linha + 20, minimo)


def gera_workbook_recebimento(por_rota, node, data_arquivo, ressalvas=None, qtd_chegou=None):
    """Monta o Excel final: aba 'Resumo' (pacotes/paradas/Em Stow por
    rota, com total no fim) + aba 'Detalhe paradas' (todos os TBR da
    rota, ver comentario grande no topo do arquivo sobre a regra usada
    pra agrupar parada e a limitacao dela) + aba 'Ressalvas', se tiver.

    Ajuste do Samuel em 24/09/2026 (reformulacao da aba Resumo e fim da
    aba Pendentes): a aba Resumo agora só traz a coluna 'Em Stow' (TBR
    ja Stowed ou alem, por rota) - é o que REALMENTE vai ser expedido
    pela manhã, entao é o unico numero que importa aqui; só aparece
    quando pelo menos uma rota já rodou o checkpoint de STOW (antes da
    indução sozinha, a coluna nem aparecia - nao faz sentido mostrar
    "Em Stow" de quem ainda nem foi induzido). A aba 'Pendentes' (que
    listava TBR x situação, sem espaço pra análise) foi REMOVIDA - o
    detalhe de cada possível MNR agora fica na aba 'Ressalvas', escrito
    pelo próprio líder (ver monta_texto_copia_mnr, chamada em
    app.py na tela, que monta a lista pra copiar/colar nessa ressalva).

    ressalvas: lista opcional de dicts {"checkpoint": str, "texto": str,
    "fotos": [(nome, bytes), ...]} - vira uma aba 'Ressalvas' com o
    texto e CADA foto/documento embutido (0, 1 ou varias - ajuste do
    Samuel em 24/09/2026: antes era só 1 foto, agora aceita várias -
    foto do Manifested/Bill of Lading, foto de avaria, etc., tudo no
    mesmo registro), pra registro/auditoria (ver
    monta_texto_ressalva_padrao e monta_texto_copia_mnr). Ajuste
    também em 24/09/2026: a ressalva virou um registro SÓ, feito no
    final (depois do checkpoint de stow), em vez de uma por checkpoint -
    ver tela_recebimento() em app.py.

    qtd_chegou: int opcional - a quantidade fisica digitada (contagem do
    turno inteiro). Ate 23/09/2026 isso era separado por hub (REC9/
    FOR3), mas o Samuel pediu pra tirar essa separacao em 24/09/2026: a
    coluna "Source" do CSV do SCC nao e um rotulo fixo de hub de origem
    (e o local do ULTIMO SCAN - um TBR ja Delivered aparece com Source
    "CUSTOMER_ADDRESS", nao o hub de onde saiu), entao nao da pra
    confiar nela pra separar automaticamente REC9 de FOR3 - virou um
    numero unico de novo, registrado como uma linha logo abaixo do
    titulo."""
    tem_stow_real = any("situacao_stow" in info for info in por_rota.values())

    wb = openpyxl.Workbook()

    ws = wb.active
    ws.title = "Resumo"
    titulo = f"Painel Recebimento - {node or '?'} - {data_arquivo.strftime('%d/%m/%Y') if data_arquivo else '?'}"
    ws.append([titulo])
    ws["A1"].font = Font(bold=True, size=13)
    if qtd_chegou:
        ws.append([f"Recebido: {qtd_chegou} pacotes"])
        ws[f"A{ws.max_row}"].font = Font(italic=True)
    ws.append([])
    cabecalho = ["Rota", "Pacotes", "Paradas"]
    if tem_stow_real:
        cabecalho.append("Em Stow (pronto p/ expedir)")
    ws.append(cabecalho)
    for cel in ws[ws.max_row]:
        cel.font = Font(bold=True)
        cel.fill = PatternFill(fill_type="solid", fgColor="FFE8E8E8")

    total_pacotes = total_paradas = total_em_stow = 0
    for rota in sorted(por_rota):
        info = por_rota[rota]
        linha = [rota, info["pacotes"], info["paradas"]]
        if tem_stow_real:
            sit = info.get("situacao_stow")
            if sit:
                em_stow = len(sit["stowed_ou_alem"])
                linha.append(em_stow)
                total_em_stow += em_stow
            else:
                linha.append("-")
        ws.append(linha)
        total_pacotes += info["pacotes"]
        total_paradas += info["paradas"]
    linha_total = ["TOTAL", total_pacotes, total_paradas]
    if tem_stow_real:
        linha_total.append(total_em_stow)
    ws.append(linha_total)
    for cel in ws[ws.max_row]:
        cel.font = Font(bold=True)

    ws.column_dimensions["A"].width = 14
    ws.column_dimensions["B"].width = 12
    ws.column_dimensions["C"].width = 12
    if tem_stow_real:
        ws.column_dimensions["D"].width = 22

    # Aba "Detalhe paradas" - reformulada a pedido do Samuel em
    # 24/09/2026: (1) antes só listava as paradas com mais de 1 pacote
    # (a aba usava "detalhe_paradas", que é só um resumo de conferência)
    # - agora usa "todas_paradas" e lista TODOS os TBR da rota, inclusive
    # parada de 1 pacote só; (2) cada rota vira uma "tabelinha" separada
    # (cabeçalho da rota + cabeçalho de coluna próprios), em vez de uma
    # tabela só com todas as rotas misturadas numa coluna "Rota"; (3) o
    # número da "Parada" fica numa coluna, repetido pra cada TBR daquela
    # parada, pra ficar fácil ver quais pacotes caem juntos sem precisar
    # ler o endereço inteiro de novo em cada linha.
    #
    # BUG real encontrado pelo Samuel em 24/09/2026 (rota AX8, "Rua das
    # Embarcações"): a coluna "Endereço" mostrava o mesmo texto (do
    # PRIMEIRO TBR do grupo) pra TODOS os TBR daquela parada - inclusive
    # quando o complemento de cada um era bem diferente (bloco/apto
    # diferente: "Bl 08 Apto 408", "AP 403 BLOCO 7", "Bl 29 ap 202" etc,
    # 5 blocos diferentes do mesmo condomínio agrupados como 1 parada só
    # porque a CHAVE de agrupamento - rua+número, antes da vírgula - é
    # igual pros 5). A chave de agrupamento continua a mesma (é a regra
    # documentada no topo do arquivo), mas a coluna "Endereço" agora
    # mostra o endereço de CADA TBR individualmente (com html.unescape,
    # mas sem tirar acento/complemento - ver endereco_para_exibicao),
    # em vez do endereço de um TBR só repetido pra todo mundo do grupo.
    ws2 = wb.create_sheet("Detalhe paradas")
    ws2.column_dimensions["A"].width = 10
    ws2.column_dimensions["B"].width = 55
    ws2.column_dimensions["C"].width = 16

    linha = 1
    for rota in sorted(por_rota):
        info = por_rota[rota]
        endereco_por_tbr = {
            item["tbr"]: endereco_para_exibicao(item["endereco"]) for item in info["itens"]
        }
        ws2.cell(row=linha, column=1, value=f"Rota {rota} — {info['pacotes']} pacotes / {info['paradas']} paradas")
        ws2.merge_cells(start_row=linha, start_column=1, end_row=linha, end_column=3)
        cel_titulo = ws2.cell(row=linha, column=1)
        cel_titulo.font = Font(bold=True, size=12)
        cel_titulo.fill = PatternFill(fill_type="solid", fgColor="FFD9D9D9")
        linha += 1

        ws2.cell(row=linha, column=1, value="Parada")
        ws2.cell(row=linha, column=2, value="Endereço")
        ws2.cell(row=linha, column=3, value="TBR")
        for col in range(1, 4):
            cel = ws2.cell(row=linha, column=col)
            cel.font = Font(bold=True)
            cel.fill = PatternFill(fill_type="solid", fgColor="FFE8E8E8")
        linha += 1

        for num_parada, d in enumerate(info["todas_paradas"], start=1):
            for tbr in d["tbrs"]:
                ws2.cell(row=linha, column=1, value=num_parada)
                ws2.cell(row=linha, column=2, value=endereco_por_tbr.get(tbr, d["endereco_exibicao"]))
                ws2.cell(row=linha, column=3, value=tbr)
                linha += 1

        linha += 1  # linha em branco separando a próxima rota

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
            ws_res.row_dimensions[linha_atual].height = _estima_altura_linha_ressalva(r["texto"])
            linha_atual += 2
            # Varias fotos/documentos por ressalva (ajuste do Samuel em
            # 24/09/2026 - antes era só 1, depois ficavam uma embaixo da
            # outra - agora ficam LADO A LADO, andando de coluna em vez
            # de linha, só descendo de linha depois de colocar TODAS as
            # fotos dessa ressalva). PDF não é imagem pro Pillow abrir
            # direto - cai no "except" e só anota o nome do arquivo,
            # mesmo fallback de antes.
            #
            # Começa na coluna B (não A) porque a coluna A é a larga
            # (width 70, lá embaixo) usada pro texto da ressalva - só a
            # partir da B que a largura de coluna é a padrão do Excel
            # (~64px), que é a base usada pra calcular quantas colunas
            # cada foto ocupa (PX_POR_COLUNA) e não sobrepor a próxima.
            PX_POR_COLUNA = 64
            coluna_atual = 2
            maior_altura_px = 0
            for nome_foto, dados_foto in (r.get("fotos") or []):
                try:
                    img_pil = PILImage.open(io.BytesIO(dados_foto))
                    img_pil.thumbnail((320, 480))
                    img_buffer = io.BytesIO()
                    img_pil.convert("RGB").save(img_buffer, format="PNG")
                    img_buffer.seek(0)
                    img_excel = ExcelImage(img_buffer)
                    ancora = f"{get_column_letter(coluna_atual)}{linha_atual}"
                    ws_res.add_image(img_excel, ancora)
                    coluna_atual += int(img_pil.width / PX_POR_COLUNA) + 1
                    maior_altura_px = max(maior_altura_px, img_pil.height)
                except Exception:
                    ws_res.cell(row=linha_atual, column=coluna_atual, value=f"(não consegui anexar {nome_foto})")
                    coluna_atual += 2
            linha_atual += (int(maior_altura_px / 18) + 3) if maior_altura_px else 2
            linha_atual += 2
        ws_res.column_dimensions["A"].width = 70

    buffer = io.BytesIO()
    wb.save(buffer)
    return buffer.getvalue()
