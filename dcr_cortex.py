# -*- coding: utf-8 -*-
"""
dcr_cortex.py: cálculo do DCR da DMNZ a partir do Route Planning e dos itinerários do Cortex.

Duas fontes, e cada uma responde uma pergunta:
  - Route Planning (um .xlsx com uma aba por rota): quantos pacotes foram roteirizados.
  - Itinerário do Cortex (um arquivo por rota): o que aconteceu com cada TBR (entregue, tentativa, em rota).

O Cortex que o Samuel acessa só mostra as rotas da DMNZ. Por isso o resultado é o DCR da DMNZ, não o da
estação inteira. Se ele souber o RTS total da estação (de todos os DSPs), dá pra informar e o robô calcula
também o DCR da estação e quanto do limite a DMNZ consumiu.

Fórmulas (as mesmas do relatório que o CTA usa):
  limite de RTS = piso(roteirizados x (1 - target))
  DCR           = 1 - RTS / roteirizados

O arquivo do Cortex mostra "Tentativa de entrega", não o RTS oficial. Uma tentativa pode virar entrega se o
motorista voltar, então o RTS daqui é uma estimativa do momento em que o arquivo foi baixado.
"""

import csv
import datetime
import io
import re
import unicodedata
from decimal import Decimal, ROUND_DOWN
from zoneinfo import ZoneInfo

import openpyxl
import pandas as pd

TARGET_PADRAO = Decimal("97.50")

# Motivos do Cortex que já foram conferidos contra o relatório do CTA, com o nome que ele usa.
# Motivo que não está aqui aparece com o texto original do Cortex (nada é inventado).
# Chave: texto do Cortex sem acento e em minúsculas.
MOTIVOS_CONHECIDOS = {
    "cliente nao estava disponivel": "Cliente ausente",
    "otp nao disponivel": "Cliente sem o código (OTP)",
    "nao foi possivel encontrar o endereco": "Endereço não encontrado",
}

# Colunas do itinerário. Chave: nome interno. Valor: cabeçalho do Cortex normalizado (ver _norm).
COLUNAS_ITINERARIO = {
    "rota": "codigo de rota",
    "da": "nome da da",
    "parada": "numero de parada",
    "etiqueta": "etiqueta de parada",
    "tbr": "tba",
    "status": "status da entrega",
    "motivo": "codigo do motivo",
    "hora": "hora da ultima acao",
}
COLUNAS_OBRIGATORIAS = ("rota", "tbr", "status")

BASE_COLETADOS = "coletados"
BASE_PLANO = "plano"

SITUACAO_ENTREGUE = "entregue"
SITUACAO_RTS = "rts"
SITUACAO_EM_ROTA = "em rota"
SITUACAO_OUTRA = "outra"


def agora_natal():
    """Hora atual de Natal. O servidor da nuvem roda em UTC, então não dá pra usar o relógio dele direto."""
    return datetime.datetime.now(ZoneInfo("America/Fortaleza"))


def _norm(texto):
    """Minúsculas, sem acento e com espaços normalizados. Serve pra comparar textos do Cortex."""
    t = unicodedata.normalize("NFKD", str(texto or ""))
    t = "".join(c for c in t if not unicodedata.combining(c))
    return re.sub(r"\s+", " ", t).strip().lower()


# ---- Leitura do itinerário do Cortex ----

def _decodifica_csv(conteudo):
    for codificacao in ("utf-8-sig", "cp1252"):
        try:
            return conteudo.decode(codificacao)
        except UnicodeDecodeError:
            continue
    return conteudo.decode("latin-1")


def _conserta_acentos(texto):
    """Desfaz o acento quebrado ('CÃ³digo' vira 'Código'), que aparece quando o CSV é salvo como xlsx."""
    for codificacao in ("cp1252", "latin-1"):
        try:
            return texto.encode(codificacao).decode("utf-8")
        except (UnicodeEncodeError, UnicodeDecodeError):
            continue
    return texto


def _linhas_do_xlsx(conteudo):
    """Lê a primeira aba. Se o arquivo veio com o CSV inteiro dentro da coluna A, separa as colunas."""
    wb = openpyxl.load_workbook(io.BytesIO(conteudo), read_only=True, data_only=True)
    linhas = [list(l) for l in wb.worksheets[0].iter_rows(values_only=True)]
    wb.close()
    if not linhas:
        return []
    celulas_primeira_linha = [c for c in linhas[0] if c is not None and str(c).strip()]
    if len(celulas_primeira_linha) == 1 and "," in str(celulas_primeira_linha[0]):
        textos = [_conserta_acentos(str(l[0])) for l in linhas if l and l[0] is not None]
        return list(csv.reader(io.StringIO("\n".join(textos))))
    return [["" if c is None else str(c) for c in l] for l in linhas]


def le_itinerario(nome_arquivo, conteudo):
    """Lê um itinerário do Cortex (.csv ou .xlsx) e devolve um DataFrame com colunas internas.

    Colunas: rota, da, parada, etiqueta, tbr, status, motivo, hora (as que existirem no arquivo).
    Levanta ValueError com mensagem clara se faltar coluna essencial.
    """
    if nome_arquivo.lower().endswith((".xlsx", ".xlsm")):
        linhas = _linhas_do_xlsx(conteudo)
    else:
        linhas = list(csv.reader(io.StringIO(_decodifica_csv(conteudo))))
    linhas = [l for l in linhas if any(str(c).strip() for c in l)]
    if len(linhas) < 2:
        raise ValueError(f"{nome_arquivo}: arquivo vazio ou sem linhas de dados.")

    cabecalho = [_norm(c) for c in linhas[0]]
    indice = {}
    for interno, nome_cortex in COLUNAS_ITINERARIO.items():
        if nome_cortex in cabecalho:
            indice[interno] = cabecalho.index(nome_cortex)
    faltando = [COLUNAS_ITINERARIO[c] for c in COLUNAS_OBRIGATORIAS if c not in indice]
    if faltando:
        raise ValueError(
            f"{nome_arquivo}: não parece um itinerário do Cortex (faltam as colunas: {', '.join(faltando)})."
        )

    registros = []
    for l in linhas[1:]:
        registro = {}
        for interno, i in indice.items():
            registro[interno] = str(l[i]).strip() if i < len(l) else ""
        registros.append(registro)
    df = pd.DataFrame(registros)
    for coluna in COLUNAS_ITINERARIO:
        if coluna not in df.columns:
            df[coluna] = ""
    df["hora_dt"] = pd.to_datetime(df["hora"], errors="coerce")
    return df[df["tbr"] != ""].reset_index(drop=True)


def classifica_entregas(df):
    """Fica só com a linha de entrega de cada TBR e classifica a situação.

    Cada TBR aparece duas vezes no Cortex: uma linha 'Pegar' (carregou no hub) e outra da entrega.
    O status final está na linha da entrega. Se a coluna de etiqueta não existir, a parada de número
    mais baixo é tratada como a do 'Pegar'.
    """
    if (df["etiqueta"] != "").any():
        entregas = df[df["etiqueta"].map(_norm) != "pegar"]
    else:
        parada = pd.to_numeric(df["parada"], errors="coerce")
        entregas = df[parada != parada.min()]
    entregas = entregas.drop_duplicates("tbr", keep="last").copy()

    status_n = entregas["status"].map(_norm)
    entregas["situacao"] = SITUACAO_OUTRA
    entregas.loc[status_n == "entregue", "situacao"] = SITUACAO_ENTREGUE
    entregas.loc[status_n == "tentativa de entrega", "situacao"] = SITUACAO_RTS
    entregas.loc[status_n.isin(["proximo", "coletado"]), "situacao"] = SITUACAO_EM_ROTA

    def nome_do_motivo(texto):
        if _norm(texto) in ("", "nenhum"):
            return ""
        return MOTIVOS_CONHECIDOS.get(_norm(texto), texto)

    entregas["motivo_exibido"] = entregas["motivo"].map(nome_do_motivo)
    entregas["otp"] = entregas["etiqueta"].map(_norm).str.contains("otp")
    return entregas


# ---- Leitura do Route Planning ----

def le_route_planning(conteudo):
    """Lê o xlsx do Route Planning e devolve {rota: [TBRs na ordem]}.

    Cada aba é uma rota. A primeira linha tem 'Route for AX1 ...'; depois vem o cabeçalho
    (Stop, Tracking ID, ...) e uma linha por parada.
    """
    wb = openpyxl.load_workbook(io.BytesIO(conteudo), read_only=True, data_only=True)
    plano = {}
    for ws in wb.worksheets:
        linhas = list(ws.iter_rows(values_only=True))
        rota = None
        if linhas and linhas[0] and linhas[0][0]:
            achou = re.search(r"Route for\s+(\S+)", str(linhas[0][0]))
            rota = achou.group(1) if achou else None
        coluna_tbr = None
        primeira_linha_dados = None
        for n, l in enumerate(linhas):
            textos = [_norm(c) for c in l]
            if "stop" in textos and "tracking id" in textos:
                coluna_tbr = textos.index("tracking id")
                primeira_linha_dados = n + 1
                break
        if coluna_tbr is None:
            continue
        if not rota:
            rota = re.sub(r"^sequencedRoute_", "", ws.title)
        tbrs = []
        for l in linhas[primeira_linha_dados:]:
            valor = l[coluna_tbr] if coluna_tbr < len(l) else None
            if valor and str(valor).strip() and str(valor).strip() not in tbrs:
                tbrs.append(str(valor).strip())
        plano[rota] = tbrs
    wb.close()
    if not plano:
        raise ValueError("Não parece um arquivo do Route Planning (não achei abas com 'Stop' e 'Tracking ID').")
    return plano


# ---- Cálculo ----

def _pct(valor):
    """Percentual com 2 casas, cortando (não arredondando). 98,837 vira 98,83, como no relatório do CTA."""
    return valor.quantize(Decimal("0.01"), rounding=ROUND_DOWN)


def calcula_limite(roteirizados, target):
    """Maior número de RTS que ainda mantém o DCR no target: piso(roteirizados x (1 - target))."""
    return int((Decimal(roteirizados) * (Decimal(100) - target) / Decimal(100)).to_integral_value(rounding=ROUND_DOWN))


def calcula_dcr(roteirizados, rts):
    if roteirizados <= 0:
        return None
    return _pct((Decimal(1) - Decimal(rts) / Decimal(roteirizados)) * 100)


def monta_resultado(plano, itinerarios, target=TARGET_PADRAO, rts_estacao=None, base=BASE_COLETADOS):
    """Junta plano e itinerários e calcula tudo.

    plano: {rota: [TBRs]} do Route Planning.
    itinerarios: lista de (nome_do_arquivo, DataFrame de le_itinerario).
    rts_estacao: RTS total da estação (todos os DSPs), opcional.
    base: o que entra como denominador do DCR da DMNZ. BASE_COLETADOS usa os TBRs que o motorista
    pegou no hub (o que de fato saiu pra rota). BASE_PLANO usa tudo que o Route Planning roteirizou.
    A diferença são pacotes planejados que não saíram (avaria, MNR e afins).

    Devolve um dicionário com as rotas lidas, os totais, os buckets de motivo e os avisos.
    """
    avisos = []
    por_rota_arquivo = {}  # rota -> (nome_arquivo, entregas)
    for nome, df in itinerarios:
        entregas_do_arquivo = classifica_entregas(df)
        for rota, grupo in entregas_do_arquivo.groupby("rota"):
            if rota in por_rota_arquivo:
                nome_antigo, anterior = por_rota_arquivo[rota]
                mais_recente_novo = grupo["hora_dt"].max()
                mais_recente_antigo = anterior["hora_dt"].max()
                if pd.notna(mais_recente_novo) and (pd.isna(mais_recente_antigo) or mais_recente_novo > mais_recente_antigo):
                    por_rota_arquivo[rota] = (nome, grupo)
                    usado, descartado = nome, nome_antigo
                else:
                    usado, descartado = nome_antigo, nome
                avisos.append(f"A rota {rota} veio em dois arquivos. Usei {usado} (dados mais recentes) e ignorei {descartado}.")
            else:
                por_rota_arquivo[rota] = (nome, grupo)

    rotas = []
    todos_rts = []
    for rota in sorted(por_rota_arquivo):
        nome, entregas = por_rota_arquivo[rota]
        planejados = plano.get(rota)
        if planejados is None:
            avisos.append(f"A rota {rota} tem itinerário mas não está no Route Planning. Usei os TBRs do itinerário como roteirizados.")
            planejados = list(entregas["tbr"])
        fora_do_itinerario = [t for t in planejados if t not in set(entregas["tbr"])]
        fora_do_plano = [t for t in entregas["tbr"] if t not in set(planejados)]
        if fora_do_plano:
            avisos.append(f"Rota {rota}: {len(fora_do_plano)} TBR(s) do itinerário não estão no Route Planning.")
        das = [d for d in entregas["da"] if d]
        rts = entregas[entregas["situacao"] == SITUACAO_RTS]
        otp = entregas[entregas["otp"]]
        rotas.append({
            "rota": rota,
            "da": das[0] if das else "",
            "arquivo": nome,
            "roteirizados": len(planejados),
            "no_itinerario": len(entregas),
            "entregues": int((entregas["situacao"] == SITUACAO_ENTREGUE).sum()),
            "rts": len(rts),
            "em_rota": int((entregas["situacao"] == SITUACAO_EM_ROTA).sum()),
            "outras": int((entregas["situacao"] == SITUACAO_OUTRA).sum()),
            "otp_total": len(otp),
            "otp_entregues": int((otp["situacao"] == SITUACAO_ENTREGUE).sum()),
            "fora_do_itinerario": fora_do_itinerario,
            "ultima_acao": entregas["hora_dt"].max(),
        })
        for _, linha in rts.iterrows():
            todos_rts.append({
                "rota": rota,
                "da": linha["da"],
                "tbr": linha["tbr"],
                "motivo": linha["motivo_exibido"] or "(sem motivo no arquivo)",
                "hora": linha["hora"],
            })
        outras = entregas[entregas["situacao"] == SITUACAO_OUTRA]
        for status in sorted(set(outras["status"])):
            avisos.append(f"Rota {rota}: {int((outras['status'] == status).sum())} TBR(s) com status '{status}', que o robô não classifica. Estão em 'Outras situações'.")

    roteirizados = sum(r["roteirizados"] for r in rotas)
    coletados = sum(r["no_itinerario"] for r in rotas)
    base_dcr = coletados if base == BASE_COLETADOS else roteirizados
    rts = len(todos_rts)
    limite = calcula_limite(base_dcr, target)
    ultimas = [r["ultima_acao"] for r in rotas if pd.notna(r["ultima_acao"])]

    buckets = {}
    for item in todos_rts:
        b = buckets.setdefault(item["motivo"], {"motivo": item["motivo"], "qtd": 0, "das": {}})
        b["qtd"] += 1
        b["das"][item["da"]] = b["das"].get(item["da"], 0) + 1
    lista_buckets = sorted(buckets.values(), key=lambda b: -b["qtd"])

    resultado = {
        "rotas": rotas,
        "rts_lista": todos_rts,
        "buckets": lista_buckets,
        "target": target,
        "roteirizados": roteirizados,
        "coletados": coletados,
        "base": base,
        "base_dcr": base_dcr,
        "rts": rts,
        "limite": limite,
        "folga": limite - rts,
        "excesso": max(rts - limite, 0),
        "dcr": calcula_dcr(base_dcr, rts),
        "dentro_do_target": rts <= limite,
        "entregues": sum(r["entregues"] for r in rotas),
        "em_rota": sum(r["em_rota"] for r in rotas),
        "outras": sum(r["outras"] for r in rotas),
        "dados_ate": max(ultimas) if ultimas else None,
        "avisos": avisos,
        "estacao": None,
    }

    if rts_estacao is not None:
        roteirizados_estacao = sum(len(t) for t in plano.values())
        limite_estacao = calcula_limite(roteirizados_estacao, target)
        resultado["estacao"] = {
            "roteirizados": roteirizados_estacao,
            "rotas": len(plano),
            "rts": int(rts_estacao),
            "limite": limite_estacao,
            "dcr": calcula_dcr(roteirizados_estacao, int(rts_estacao)),
            "dentro_do_target": int(rts_estacao) <= limite_estacao,
            "parte_da_dmnz": (rts / int(rts_estacao) * 100) if rts_estacao else None,
        }
        if int(rts_estacao) < rts:
            avisos.append("O RTS da estação informado é menor que o RTS da DMNZ calculado aqui. Confira o número.")
    return resultado


# ---- Texto pra copiar ----

def fmt_hora(momento):
    """HH:MM de um Timestamp, ou '-' se estiver vazio."""
    return momento.strftime("%H:%M") if pd.notna(momento) else "-"


def fmt_pct(valor):
    return f"{valor}%".replace(".", ",") if valor is not None else "-"


def monta_texto(resultado, data_texto):
    """Resumo em texto puro, pra colar no WhatsApp ou e-mail."""
    r = resultado
    ate = r["dados_ate"].strftime("%H:%M") if r["dados_ate"] is not None else "?"
    nomes_rotas = ", ".join(x["rota"] for x in r["rotas"])
    situacao = "DENTRO DO TARGET" if r["dentro_do_target"] else "FORA DO TARGET"
    linhas = [
        f"DCR DMNZ · LRN9 · {data_texto} · dados do Cortex até {ate}",
        f"Rotas: {nomes_rotas} · coletados {r['coletados']} · roteirizados {r['roteirizados']}",
        f"Base do DCR: {'coletados' if r['base'] == BASE_COLETADOS else 'roteirizados'} ({r['base_dcr']})",
        f"{situacao} (DMNZ): DCR {fmt_pct(r['dcr'])} · target {fmt_pct(r['target'])}",
        f"RTS estimado: {r['rts']} · limite {r['limite']} · "
        + (f"folga {r['folga']}" if r["dentro_do_target"] else f"excesso {r['excesso']}"),
        f"Entregues: {r['entregues']} · ainda em rota: {r['em_rota']}",
    ]
    if r["estacao"]:
        e = r["estacao"]
        linha_estacao = f"Estação: RTS {e['rts']} de limite {e['limite']} · DCR {fmt_pct(e['dcr'])}"
        if e["parte_da_dmnz"] is not None:
            linha_estacao += f" · DMNZ responde por {e['parte_da_dmnz']:.0f}% do RTS"
        linhas.append(linha_estacao)
    if r["buckets"]:
        linhas.append("")
        linhas.append("RTS por motivo:")
        for b in r["buckets"]:
            das = ", ".join(f"{d} ({q})" for d, q in b["das"].items())
            linhas.append(f"- {b['motivo']}: {b['qtd']} · {das}")
        linhas.append("")
        linhas.append("TBRs:")
        for item in r["rts_lista"]:
            linhas.append(f"- {item['rota']} · {item['da']} · {item['tbr']} · {item['motivo']}")
    return "\n".join(linhas)
