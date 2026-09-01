# -*- coding: utf-8 -*-
"""
consolida_rotas.py (versao web)

Mesma logica do consolida_rotas.py do robo desktop, so que le o arquivo
de rotas direto de um upload do Streamlit (em memoria) em vez de um
caminho no disco. Nao muda nenhuma regra de negocio.

Le o arquivo de rotas que vem dos indianos (varias abas
"sequencedRoute_AX1", "sequencedRoute_AX2" etc, uma pra cada rota) e
consolida todos os Tracking IDs numa lista unica - pronta pra colar na
caixa de pesquisa do SCC (Package Summary > Search).
"""

import re
import datetime as dt
import openpyxl

# nodes da Dominalog sempre comecam com "L" (LRN9, LPB9, LFO9, LRE9, LPA9,
# LSA8, LBH9...) - lista conhecida + padrao generico pra pegar nodes
# novos que ainda nao estao na lista
NODES_CONHECIDOS = {"LRN9", "LFO9", "LPB9", "LRE9", "LPA9", "LSA8", "LBH9", "SXL9"}
PADRAO_NODE = re.compile(r"L[A-Z]{2}\d")

# procura AAAAMMDD em qualquer parte do nome do arquivo (ex:
# "LRN9_CYCLE_1_20260812true.xlsx" -> 2026-08-12)
PADRAO_DATA_NOME = re.compile(r"(20\d{2})(\d{2})(\d{2})")


def detecta_node_do_nome(nome_arquivo):
    """Procura um codigo de node (comeca com L, ex LRN9, LPB9, LBH9...) em
    qualquer parte do nome do arquivo. Devolve o codigo encontrado ou None
    se nao achar nada parecido."""
    nome = nome_arquivo.upper()
    for node in NODES_CONHECIDOS:
        if node in nome:
            return node
    m = PADRAO_NODE.search(nome)
    return m.group(0) if m else None


def _carrega_planilha(arquivo, nome_arquivo):
    """Le o Excel a partir de um objeto de arquivo em memoria (o que o
    st.file_uploader devolve). Suporta .xlsx/.xlsm (openpyxl) e .xls
    antigo (xlrd) - mesmo comportamento do robo desktop, so trocando
    'caminho no disco' por 'bytes em memoria'.

    Devolve (abas, data_criacao) - data_criacao vem da propriedade
    interna do Excel (quando o arquivo foi salvo pela primeira vez),
    so existe pra .xlsx/.xlsm (o .xls antigo nao guarda isso de um
    jeito facil de ler, entao vem None)."""
    ext = nome_arquivo.lower().rsplit(".", 1)[-1]

    if ext == "xls":
        import xlrd
        livro = xlrd.open_workbook(file_contents=arquivo.read())
        abas = {
            aba.name: [aba.row_values(r) for r in range(aba.nrows)]
            for aba in livro.sheets()
        }
        return abas, None

    wb = openpyxl.load_workbook(arquivo, data_only=True)
    abas = {
        nome: [list(row) for row in wb[nome].iter_rows(values_only=True)]
        for nome in wb.sheetnames
    }
    criado = wb.properties.created
    return abas, (criado.date() if criado else None)


def detecta_data_do_arquivo(nome_arquivo, data_criacao_excel):
    """Data do arquivo, pra avisar se a pessoa subiu um arquivo de outro
    dia sem perceber. Prioridade: 1) AAAAMMDD no nome do arquivo (é o
    que os indianos usam, reflete o dia do ciclo de verdade) - 2) se não
    achar no nome, usa a data que o proprio Excel guarda de quando foi
    criado (so existe pra .xlsx/.xlsm)."""
    m = PADRAO_DATA_NOME.search(nome_arquivo)
    if m:
        ano, mes, dia = int(m.group(1)), int(m.group(2)), int(m.group(3))
        try:
            return dt.date(ano, mes, dia)
        except ValueError:
            pass  # numeros no nome nao formam uma data valida - ignora
    return data_criacao_excel


def consolida(arquivo, nome_arquivo):
    """arquivo: objeto vindo do st.file_uploader (ja aberto em memoria).
    nome_arquivo: arquivo.name, so pra saber a extensao e detectar o
    node. Devolve (tbrs_por_rota, data_do_arquivo)."""
    abas, data_criacao_excel = _carrega_planilha(arquivo, nome_arquivo)

    abas_rota = [n for n in abas if n.lower().startswith("sequencedroute")]
    if not abas_rota:
        raise ValueError(
            "Nao encontrei nenhuma aba comecando com 'sequencedRoute' nesse arquivo. "
            "Confira se e o arquivo certo (o que vem dos indianos)."
        )

    tbrs_por_rota = {}
    for nome_aba in abas_rota:
        linhas = abas[nome_aba]
        rota = nome_aba.split("_")[-1]  # "sequencedRoute_AX1" -> "AX1"
        tbrs = []
        for row in linhas[2:]:  # pula linha 1 (metadados) e 2 (cabecalho)
            tracking_id = row[1] if len(row) > 1 else None
            if tracking_id:
                tbrs.append(str(tracking_id).strip())
        tbrs_por_rota[rota] = tbrs

    data_arquivo = detecta_data_do_arquivo(nome_arquivo, data_criacao_excel)
    return tbrs_por_rota, data_arquivo


def le_tbrs_colados(texto):
    """Le TBRs colados direto numa caixa de texto (um por linha) -
    equivalente web de 'colar os TBRs de NA direto', sem precisar de
    arquivo. Ignora linhas em branco."""
    if not texto:
        return []
    return [linha.strip() for linha in texto.splitlines() if linha.strip()]
