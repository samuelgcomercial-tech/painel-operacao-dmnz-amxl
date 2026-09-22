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
# Received, Cancelado...) - qualquer state que NAO seja "Manifested"
# nem "Inducted" so pode significar que o pacote ja passou do Stow (se
# nao tivesse passado, o SCC não teria como ter avançado ele pra la).
# Isso poupa ter que confirmar cada state do fim do fluxo com um
# arquivo de exemplo.
ESTADOS_ANTES_DO_STOW = ("manifested", "inducted")


def categoriza_state_recebimento(state_texto):
    """Devolve 'manifested', 'inducted', 'stowed_ou_alem' (Stowed ou
    qualquer state posterior - ver comentario acima) ou
    'sem_state' (state vazio, mas TBR foi encontrado no CSV - raro,
    mesmo tratamento de state vazio que o resto do app da)."""
    state = (state_texto or "").strip().lower()
    if not state:
        return "sem_state"
    if state in ESTADOS_ANTES_DO_STOW:
        return state
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
    21/09/2026): SO fica False por causa de manifested/inducted/sem_state
    presos no meio do processo - "nao_encontrado" (planejado mas nunca
    apareceu em NENHUM CSV) NAO trava a confirmacao. Motivo: nem todo
    pacote do plano chega fisicamente na estacao todo santo dia (atraso
    de caminhao, ficou em outra estacao etc) - isso e normal, nao e um
    problema de processo, entao nao devia impedir a rota de ser marcada
    como "concluida" pros que DE FATO chegaram. E exatamente por isso
    que "nao_encontrado" fica separado (pra dar pra ver quem sumiu),
    mas fora do calculo de 'confirmada'.

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


def gera_workbook_recebimento(por_rota, node, data_arquivo, ressalvas=None):
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
    tiver), pra registro/auditoria (ver monta_texto_ressalva_padrao)."""
    tem_situacao_real = any(_situacao_mais_recente(info) for info in por_rota.values())

    wb = openpyxl.Workbook()

    ws = wb.active
    ws.title = "Resumo"
    titulo = f"Painel Recebimento - {node or '?'} - {data_arquivo.strftime('%d/%m/%Y') if data_arquivo else '?'}"
    ws.append([titulo])
    ws["A1"].font = Font(bold=True, size=13)
    ws.append([])
    cabecalho = ["Rota", "Pacotes", "Paradas"]
    if tem_situacao_real:
        cabecalho += ["Manifested", "Inducted", "Stowed ou além", "Não encontrado", "Confirmada?"]
    ws.append(cabecalho)
    for cel in ws[3]:
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
        rotulos = {
            "manifested": "Ainda não induzido (Manifested)",
            "inducted": "Induzido, falta armazenar (Inducted)",
            "sem_state": "Achado no CSV mas sem State",
            "nao_encontrado": "Não encontrado no CSV do SCC",
        }
        for rota in sorted(por_rota):
            sit = _situacao_mais_recente(por_rota[rota])
            if not sit:
                continue
            for chave, rotulo in rotulos.items():
                for tbr in sit[chave]:
                    ws_pend.append([rota, rotulo, tbr])
        ws_pend.column_dimensions["A"].width = 10
        ws_pend.column_dimensions["B"].width = 34
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
