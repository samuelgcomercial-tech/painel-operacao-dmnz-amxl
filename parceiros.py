# -*- coding: utf-8 -*-
"""
parceiros.py

Cadastro de quais empresas atuam em cada node: a propria empresa (o nome
que o node usar pra si mesmo - pode ser "DMNZ", "Dominalog", "Domina",
depende de quem estiver usando) e os parceiros terceirizados (ex: MRIZ).
Guardado em CSV por node (dados_nodes/<NODE>/parceiros.csv), no mesmo
padrao do base_das.py - editavel a qualquer momento, sem nada fixo no
codigo (cada node pode ter parceiros diferentes, e um node pode ganhar
parceiro novo no meio do caminho).

Ainda NAO integrado no fluxo principal (fica pra proximas etapas, de
proposito - uma mudanca de cada vez):
  - montar a lista suspensa do State Finalizador (evitar erro de
    digitacao, ideia original do Samuel)
  - decidir se um "Insucesso" e da propria empresa ou de parceiro sem
    depender de uma palavra fixa tipo "dmnz" no meio do texto - hoje
    calculo_fechamento.classifica_insucesso ainda usa "dmnz" fixo; isso
    muda numa proxima etapa, depois que essa base estiver rodando de
    verdade e o Samuel confirmar que o cadastro ta ok
  - o alerta de "insucesso sem parceiro especificado" + o link direto
    pra editar aquele TBR
"""

import csv
import io
import unicodedata


def normaliza(nome):
    """Maiusculo e sem acento, so pra comparacao - mesma ideia do
    normaliza_texto_painel do calculo_fechamento.py."""
    txt = unicodedata.normalize("NFKD", nome or "")
    sem_acento = "".join(c for c in txt if not unicodedata.combining(c))
    return " ".join(sem_acento.upper().split())


def parseia_texto_parceiros(texto):
    """Le o CSV 'Nome,Propria Empresa' (SIM/NAO) a partir de um texto ja
    em memoria (vindo do GitHub). Se vier vazio/None (node novo, arquivo
    ainda nao existe), devolve lista vazia - quem decide o que fazer
    (perguntar ao usuario) e a tela, nao esse modulo.

    Devolve lista de dicts {nome, propria_empresa (bool)}, na ordem em
    que apareceram no arquivo."""
    if not texto:
        return []

    delimitador = ";" if texto[:2000].count(";") > texto[:2000].count(",") else ","
    reader = csv.DictReader(io.StringIO(texto), delimiter=delimitador)
    fieldnames_norm = [(c or "").strip().lower() for c in (reader.fieldnames or [])]
    if "nome" not in fieldnames_norm or "propria empresa" not in fieldnames_norm:
        raise ValueError(
            "O arquivo de parceiros não tem as colunas esperadas 'Nome' e "
            f"'Propria Empresa' (encontrei: {reader.fieldnames})."
        )
    col_nome = reader.fieldnames[fieldnames_norm.index("nome")]
    col_propria = reader.fieldnames[fieldnames_norm.index("propria empresa")]

    parceiros = []
    for row in reader:
        nome = (row.get(col_nome) or "").strip()
        if not nome:
            continue
        propria = (row.get(col_propria) or "").strip().upper() == "SIM"
        parceiros.append({"nome": nome, "propria_empresa": propria})
    return parceiros


def gera_csv_parceiros(parceiros):
    """Monta o texto do CSV a partir da lista de dicts {nome,
    propria_empresa}, pronto pra salvar no GitHub."""
    buffer = io.StringIO()
    writer = csv.writer(buffer)
    writer.writerow(["Nome", "Propria Empresa"])
    for p in parceiros:
        writer.writerow([p["nome"], "SIM" if p["propria_empresa"] else "NAO"])
    return buffer.getvalue()


def nome_propria_empresa(parceiros):
    """Devolve o nome cadastrado como 'propria empresa' (ex: 'DMNZ'), ou
    None se ainda nao tiver nenhum marcado assim (node novo, ou usuario
    ainda nao cadastrou). So faz sentido ter UM marcado como propria - se
    por engano tiver mais de um, usa o primeiro e ignora o resto (nao
    trava o app por causa disso)."""
    for p in parceiros:
        if p["propria_empresa"]:
            return p["nome"]
    return None


def identifica_parceiro_no_texto(texto, parceiros):
    """Verifica se o texto digitado (ex: o que vem depois do '-' no State
    Finalizador) cita algum dos parceiros cadastrados - por CONTEM, nao
    igualdade exata (pra tolerar 'Insucesso - Mriz atrasou' e nao so
    'Insucesso - Mriz'). Devolve o nome cadastrado (grafia original) do
    primeiro que bater, ou None se nao encontrar nenhum - nesse caso e
    que entra o alerta de 'parceiro nao especificado' (ainda nao
    construido)."""
    txt_norm = normaliza(texto)
    for p in parceiros:
        if normaliza(p["nome"]) in txt_norm:
            return p["nome"]
    return None
