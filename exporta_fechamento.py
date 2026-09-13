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

import io

import openpyxl
from openpyxl.styles import Font, Alignment
from openpyxl.utils import get_column_letter

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
        linha[COL_STATE_DAS] = dmnz_por_tbr.get(tbr, "")
        linha[COL_STATE_FINALIZADOR] = finalizador
        linhas.append(linha)
    return linhas


def escreve_aba_analise_do_dia(wb, linhas):
    """Escreve a aba 'Analise do dia', formatada pra parecer uma base
    pronta pra tabela dinamica: cabecalho em negrito e centralizado,
    congelado (linha 1 fixa ao rolar), sem linha de grade, largura de
    coluna ajustada ao conteudo (nao a largura padrao do Excel)."""
    colunas = COLUNAS_CSV_ORIGINAL + [COL_STATE_DAS, COL_STATE_FINALIZADOR]
    ws = wb.create_sheet("Analise do dia")
    ws.sheet_view.showGridLines = False

    for col_idx, nome_col in enumerate(colunas, start=1):
        cel = ws.cell(row=1, column=col_idx, value=nome_col)
        cel.font = Font(bold=True)
        cel.alignment = Alignment(horizontal="center")

    for row_idx, linha in enumerate(linhas, start=2):
        for col_idx, nome_col in enumerate(colunas, start=1):
            ws.cell(row=row_idx, column=col_idx, value=linha.get(nome_col, ""))

    ws.freeze_panes = "A2"

    for col_idx, nome_col in enumerate(colunas, start=1):
        maior = max(
            [len(str(nome_col))] + [len(str(l.get(nome_col, ""))) for l in linhas]
        )
        ws.column_dimensions[get_column_letter(col_idx)].width = min(maior + 2, 40)

    return ws


def gera_workbook_fechamento(linhas_csv, base_das, motorista_real_da_rota, base_tbr):
    """Ponto de entrada. Por enquanto so a aba 'Analise do dia' - as abas
    'Apresentacao' e 'Historico' entram nas proximas levas, depois de
    validar essa primeira. Devolve bytes prontos pro
    st.download_button do app.py."""
    wb = openpyxl.Workbook()
    wb.remove(wb.active)  # tira a aba padrao "Sheet" vazia

    linhas_base = monta_linhas_analise_do_dia(linhas_csv, base_das, motorista_real_da_rota, base_tbr)
    escreve_aba_analise_do_dia(wb, linhas_base)

    buffer = io.BytesIO()
    wb.save(buffer)
    return buffer.getvalue()
