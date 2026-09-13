# -*- coding: utf-8 -*-
"""
exporta_fechamento.py (versao web)

Gera o Excel final do Fechamento (Etapa 3) pra download - pedido do
Samuel em 13/09/2026. Planejado em 3 abas, construidas uma de cada vez:

  1. "Analise do dia" (esta primeira leva): uma linha por TBR, com TODAS
     as colunas do CSV original do SCC + duas colunas calculadas no
     final (State DA's e State Finalizador). E a "base" pronta pra ele
     montar tabela dinamica de verdade em cima dela no Excel (criar a
     tabela dinamica em si via codigo exigiria mexer direto no XML
     interno do arquivo - fragil e dificil de manter; a base "achatada"
     e o que da trabalho de verdade pra montar a mao, e isso o robo ja
     automatiza).
  2. "Apresentacao" (proxima leva): versao formatada (sem linha de
     grade, cabecalho fixo) dos paineis que ja saem no dashboard HTML
     (calculo_fechamento.calcula_tudo), pro Samuel decidir).
  3. "Historico" (ultima leva): export do historico_tbr.py (base
     persistida no GitHub) pra auditoria - desde quando cada TBR
     problematico foi visto e quando foi verificado pela ultima vez.

Ainda NAO incluido (fica pra depois, junto com as abas 2 e 3 acima).
"""

import datetime as dt
import io

import openpyxl
from openpyxl.styles import Font, Alignment
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo

from state_finalizador import monta_coluna_das_dmnz

# Mesma ordem/nomes de coluna do CSV cru exportado pelo SCC - conferido
# com o arquivo real (SearchResults_4.csv) em 13/09/2026.
COLUNAS_CSV_ORIGINAL = [
    "Tracking ID", "Last Updated Time", "Source", "State", "Destination",
    "Reason", "Station", "Operation", "Route Code", "Last Scan By",
    "Address Type", "Ship Option", "Ship Method", "Delivery Destination Type",
]

COL_STATE_DAS = "State DA's"
COL_STATE_FINALIZADOR = "State Finalizador"

# Coluna de data/hora do CSV cru - pedido do Samuel em 13/09/2026: mostrar
# so dd/mm/aaaa, sem hh:mm:ss. O valor continua guardando a hora por
# baixo dos panos (so o FORMATO de exibicao muda) - assim, se um dia
# precisar ordenar/filtrar pela hora exata, o dado nao foi perdido, so
# nao aparece na tela por padrao.
COL_DATA_HORA = "Last Updated Time"
FORMATO_DATA_HORA_ORIGINAL = "%Y-%m-%d %H:%M:%S"


def _parseia_data_hora(texto):
    """Converte o texto do CSV ('2026-09-13 07:39:52') pra um datetime de
    verdade, pra o Excel tratar como data (e nao como texto solto). Se
    vier vazio ou num formato inesperado, devolve o texto original sem
    quebrar - melhor mostrar o dado cru do que sumir com a linha."""
    texto = (texto or "").strip()
    if not texto:
        return ""
    try:
        return dt.datetime.strptime(texto, FORMATO_DATA_HORA_ORIGINAL)
    except ValueError:
        return texto


def monta_linhas_analise_do_dia(linhas_csv, base_das, motorista_real_da_rota, base_tbr):
    """Uma linha por TBR: todas as colunas originais do CSV do SCC + State
    DA's (DMNZ ou vazio - mesma fonte que ja alimenta a tela,
    monta_coluna_das_dmnz) + State Finalizador (mesma classificacao que
    ja sai no dashboard: 'Entregue' automatico pra quem foi Delivered,
    ou o que estiver salvo no historico pra quem nao foi - mesma regra
    de calculo_fechamento.monta_dados_do_dia, so que aqui devolvendo
    linha completa em vez de so os 4 campos resumidos)."""
    dmnz_por_tbr = monta_coluna_das_dmnz(linhas_csv, base_das, motorista_real_da_rota)

    linhas = []
    for l in linhas_csv:
        tbr = l["Tracking ID"]
        state_scc = (l.get("State") or "").strip()
        if state_scc.lower() == "delivered":
            finalizador = "Entregue"
        else:
            entrada = base_tbr.get(tbr)
            finalizador = entrada["classificacao"] if entrada else ""

        linha = {col: l.get(col, "") for col in COLUNAS_CSV_ORIGINAL}
        linha[COL_DATA_HORA] = _parseia_data_hora(linha[COL_DATA_HORA])
        linha[COL_STATE_DAS] = dmnz_por_tbr.get(tbr, "")
        linha[COL_STATE_FINALIZADOR] = finalizador
        linhas.append(linha)
    return linhas


def monta_linhas_resumo(linhas, resultado_calculo, n_ja_conhecidos, nome_usuario, data_criacao_str):
    """Monta as frases do resumo no topo da aba, mesmo padrao do robo
    SSW 081 (celula mesclada, uma frase por linha). Pedido do Samuel em
    13/09/2026.

    Contas:
      - entregues: contagem_scc['Delivered'], ja calculado pelo
        calculo_fechamento.calcula_tudo.
      - em_rota: quem ficou com State Finalizador comecando com 'EM
        ROTA' (classificacao automatica de In Transit/Failed com DA de
        verdade vinculado - ver state_finalizador.py) - ainda nao
        confirmado como entregue nem como insucesso.
      - revalidados sem alteracao: n_ja_conhecidos, que PRECISA vir de
        fora (Etapa 2, historico_tbr.decide_quem_precisa_reabrir) - nao
        da pra saber só pelo resultado final quem reaproveitou a
        classificacao de ontem e quem foi digitado hoje, os dois
        terminam com o mesmo texto salvo.
      - validados por pessoa hoje: o resto de quem nao foi entregue
        nem ficou em EM ROTA automatico nem foi reaproveitado do
        historico - ou seja, quem passou por revisao manual de
        verdade nesta rodada.
      - insucesso sem parceiro: contagem_insucesso['sem_detalhe'], ja
        calculado pelo calculo_fechamento.calcula_tudo (usuario digitou
        so 'Insucesso', sem citar DMNZ/nome de parceiro cadastrado)."""
    total = len(linhas)
    entregues = resultado_calculo["contagem_scc"].get("Delivered", 0)
    em_rota = sum(
        1 for l in linhas if l.get(COL_STATE_FINALIZADOR, "").upper().startswith("EM ROTA")
    )
    nao_entregues = total - entregues
    validados_hoje = max(nao_entregues - em_rota - n_ja_conhecidos, 0)
    insucesso_sem_parceiro = resultado_calculo["contagem_insucesso"].get("sem_detalhe", 0)

    return [
        f"{entregues} pacotes entregues",
        f"{n_ja_conhecidos} pacotes revalidados com state sem alteração",
        f"{validados_hoje} pacotes validados por {nome_usuario} dia {data_criacao_str}",
        f"{em_rota} pacotes ainda em rota",
        f"{insucesso_sem_parceiro} pacotes de insucesso sem parceiro atrelado",
    ]


def escreve_resumo_topo_mesclado(ws, n_colunas, linhas_resumo):
    """Escreve o resumo numa unica celula mesclada cobrindo da coluna A
    ate a ultima coluna da tabela, uma linha de planilha por frase
    (quebra de linha dentro da propria celula). Mesmo padrao ja usado
    no robo SSW 081 (escreve_resumo_topo_mesclado de
    processa_sswweb.py). Devolve quantas linhas de planilha ocupou, pra
    quem chama saber em que linha o cabecalho da tabela deve comecar."""
    linhas_resumo = [l for l in linhas_resumo if l]
    if not linhas_resumo:
        return 0
    n_linhas = len(linhas_resumo)
    if n_linhas > 1:
        ws.merge_cells(start_row=1, start_column=1, end_row=n_linhas, end_column=max(n_colunas, 1))
    cell = ws.cell(row=1, column=1, value="\n".join(linhas_resumo))
    cell.font = Font(bold=True)
    cell.alignment = Alignment(wrap_text=True, vertical="top", horizontal="left")
    for i in range(1, n_linhas + 1):
        ws.row_dimensions[i].height = 15.5
    return n_linhas


def escreve_aba_analise_do_dia(wb, linhas, linhas_resumo=None):
    """Escreve a aba 'Analise do dia', formatada como Tabela de verdade
    do Excel (nao so texto em negrito): resumo mesclado no topo (se
    'linhas_resumo' vier preenchido), cabecalho da tabela com
    filtro/ordenacao automatico, linhas zebradas pelo estilo da tabela,
    congelado (linha do cabecalho fixa ao rolar), sem linha de grade,
    largura de coluna ajustada ao conteudo. Pedido do Samuel em
    13/09/2026. Isso tambem deixa a base pronta pra virar Tabela
    Dinamica de verdade no Excel (Inserir > Tabela Dinamica), ja que
    uma tabela nomeada e a fonte de dados ideal pra isso - cresce
    sozinha se amanha vier mais linha."""
    colunas = COLUNAS_CSV_ORIGINAL + [COL_STATE_DAS, COL_STATE_FINALIZADOR]
    ws = wb.create_sheet("Analise do dia")
    ws.sheet_view.showGridLines = False

    n_linhas_resumo = escreve_resumo_topo_mesclado(ws, len(colunas), linhas_resumo or [])
    linha_cabecalho = n_linhas_resumo + 2 if n_linhas_resumo else 1

    for col_idx, nome_col in enumerate(colunas, start=1):
        cel = ws.cell(row=linha_cabecalho, column=col_idx, value=nome_col)
        cel.font = Font(bold=True)
        cel.alignment = Alignment(horizontal="center")

    for offset, linha in enumerate(linhas):
        row_idx = linha_cabecalho + 1 + offset
        for col_idx, nome_col in enumerate(colunas, start=1):
            valor = linha.get(nome_col, "")
            cel = ws.cell(row=row_idx, column=col_idx, value=valor)
            if nome_col == COL_DATA_HORA and isinstance(valor, dt.datetime):
                cel.number_format = "DD/MM/YYYY"

    ws.freeze_panes = f"A{linha_cabecalho + 1}"

    for col_idx, nome_col in enumerate(colunas, start=1):
        if nome_col == COL_DATA_HORA:
            # largura fixa pro tamanho de "31/12/2026" exibido, nao do
            # texto original ("2026-09-13 07:39:52", mais longo) que nao
            # aparece mais na tela.
            largura = max(len(nome_col), 12)
        else:
            largura = max(
                [len(str(nome_col))] + [len(str(l.get(nome_col, ""))) for l in linhas]
            )
        ws.column_dimensions[get_column_letter(col_idx)].width = min(largura + 2, 40)

    # Tabela de verdade (ListObject) so faz sentido com pelo menos 1
    # linha de dado - com a base vazia, fica so o cabecalho formatado
    # (nao quebra, so nao vira Tabela).
    if linhas:
        ultima_coluna = get_column_letter(len(colunas))
        ultima_linha = linha_cabecalho + len(linhas)
        tabela = Table(
            displayName="Tab_AnaliseDoDia",
            ref=f"A{linha_cabecalho}:{ultima_coluna}{ultima_linha}",
        )
        tabela.tableStyleInfo = TableStyleInfo(
            name="TableStyleMedium9", showRowStripes=True,
            showFirstColumn=False, showLastColumn=False, showColumnStripes=False,
        )
        ws.add_table(tabela)

    return ws


def gera_workbook_fechamento(
    linhas_csv, base_das, motorista_real_da_rota, base_tbr,
    resultado_calculo, n_ja_conhecidos, nome_usuario, data_criacao_str,
):
    """Ponto de entrada. Por enquanto so a aba 'Analise do dia' - as abas
    'Apresentacao' e 'Historico' entram nas proximas levas, depois de
    validar essa primeira.

    resultado_calculo: dict devolvido por calculo_fechamento.calcula_tudo
    (usa contagem_scc e contagem_insucesso pro resumo).
    n_ja_conhecidos: quantidade de TBRs que reaproveitaram a
    classificacao de ontem sem revisao nova hoje (vem da Etapa 2,
    historico_tbr.decide_quem_precisa_reabrir - nao da pra calcular
    aqui dentro).
    nome_usuario: st.session_state.nome_usuario (quem esta logado no
    app nessa sessao).
    data_criacao_str: data de hoje ja formatada 'dd/mm/aaaa'.

    Devolve bytes prontos pro st.download_button do app.py."""
    wb = openpyxl.Workbook()
    wb.remove(wb.active)  # tira a aba padrao "Sheet" vazia

    linhas_base = monta_linhas_analise_do_dia(linhas_csv, base_das, motorista_real_da_rota, base_tbr)
    linhas_resumo = monta_linhas_resumo(
        linhas_base, resultado_calculo, n_ja_conhecidos, nome_usuario, data_criacao_str
    )
    escreve_aba_analise_do_dia(wb, linhas_base, linhas_resumo)

    buffer = io.BytesIO()
    wb.save(buffer)
    return buffer.getvalue()
