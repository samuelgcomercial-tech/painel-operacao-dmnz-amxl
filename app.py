# -*- coding: utf-8 -*-
"""
app.py - Fechamento LRN9 (versao web)

Primeiro pedaco de verdade do robo web, pra comecar a usar e ajustar na
pratica (em vez de so desenhar mockup). Por enquanto so tem:

  - Tela inicial (nome de quem esta usando + 4 botoes de acao, so
    "Iniciar Fechamento" funcionando ainda)
  - Etapa 1: upload do arquivo de rotas -> lista consolidada na tela
    (com botao de copiar) + campo de upload do CSV do SCC do lado,
    pronto pra etapa 2 (que ainda vamos construir)

O resto (processar o CSV do SCC, montar o dashboard final, salvar
historico no GitHub) entra depois, em cima disso - sem redesenhar do
zero de novo.
"""

import datetime as dt
import hashlib

import streamlit as st

from consolida_rotas import consolida, detecta_node_do_nome, le_tbrs_colados
from consolida_recebimento import (
    LIMITE_ALERTA_PCT_DELIVERED,
    ROTULO_MNR_NAO_CARREGADO,
    consolida_recebimento,
    cruza_com_state_scc,
    gera_workbook_recebimento,
    monta_texto_copia_mnr,
    monta_texto_ressalva_padrao,
    percentual_delivered,
    soma_inducidos_ou_alem,
    soma_stowed_ou_alem,
)
from base_das import (
    extrai_das_do_dia,
    gera_csv_base_das,
    le_csv_scc,
    normaliza_nome,
    parseia_texto_base,
)
from github_store import ConflitoDeSalvamento, le_arquivo, salva_arquivo, secrets_configurados
from historico_tbr import atualiza_base, gera_csv_historico, parseia_texto_historico
from state_finalizador import processa as processa_state_finalizador
from state_finalizador import (
    resumo_das_por_rota,
    lista_sem_da_de_verdade,
    opcoes_da_rota_do_dia,
    classifica_dmnz_ou_parceiro,
)
from calculo_fechamento import calcula_tudo, monta_dados_do_dia
from dashboard_fechamento import gera_html_fechamento
from exporta_fechamento import gera_workbook_fechamento
import parceiros as parceiros_mod

NODE_ATUAL = "LRN9"  # unico node desta primeira versao (decisao ja tomada)

# Onde a base de DAs e o histórico de TBRs moram no repositório do GitHub -
# mesmo padrão já decidido pros arquivos de memória entre dias (dados_nodes/<NODE>/).
CAMINHO_BASE_DAS = f"dados_nodes/{NODE_ATUAL}/base_das_dmnz.csv"
CAMINHO_HISTORICO_TBR = f"dados_nodes/{NODE_ATUAL}/historico_scc_analise.csv"
CAMINHO_PARCEIROS = f"dados_nodes/{NODE_ATUAL}/parceiros.csv"


def _monta_insumos_export(linhas_csv, base_tbr, base_das, resultado, data_hoje_str):
    """Empacota tudo que exporta_fechamento.gera_workbook_fechamento vai
    precisar na Etapa 3, pra guardar em st.session_state - sao variaveis
    locais da Etapa 2 (existem só durante essa execução do Streamlit),
    então precisam ser persistidas explicitamente pra sobreviver até a
    tela de Fechamento (outra tela, outro rerun). Pedido do Samuel em
    13/09/2026, junto com o Excel da aba 'Analise do dia'."""
    return {
        "linhas_csv": linhas_csv,
        "base_das": base_das,
        "motorista_real_da_rota": resultado["motorista_real_da_rota"],
        "base_tbr": base_tbr,
        "tbrs_ja_conhecidos": [c["tbr"] for c in resultado["conhecidos"]],
        "data_hoje_str": data_hoje_str,
    }


st.set_page_config(page_title="Painel Operação DMNZ - AMXL", page_icon="📦", layout="wide")

# Por padrão o Streamlit só mostra o ícone de copiar de um st.code()
# quando passa o mouse/toca dentro do quadro - no celular isso confunde
# (parece que não tem botão). Aqui deixa o ícone sempre visível, só nos
# quadros de código (não mexe em outros botões escondidos do app, tipo
# de tabela/gráfico, se um dia existirem).
st.markdown(
    """
    <style>
    [data-testid="stCode"] [data-testid="stBaseButton-elementToolbar"] {
        visibility: visible !important;
        opacity: 1 !important;
    }
    </style>
    """,
    unsafe_allow_html=True,
)

# ------------------------------------------------------------------
# ESTADO DA SESSAO
# ------------------------------------------------------------------
if "tela" not in st.session_state:
    st.session_state.tela = "home"
if "nome_usuario" not in st.session_state:
    st.session_state.nome_usuario = ""
if "tbrs_por_rota" not in st.session_state:
    st.session_state.tbrs_por_rota = None
if "data_arquivo_rotas" not in st.session_state:
    st.session_state.data_arquivo_rotas = None
if "node_detectado_rotas" not in st.session_state:
    st.session_state.node_detectado_rotas = None
if "tem_na" not in st.session_state:
    st.session_state.tem_na = False
if "dados_fechamento" not in st.session_state:
    st.session_state.dados_fechamento = None
if "export_fechamento_insumos" not in st.session_state:
    st.session_state.export_fechamento_insumos = None
if "recebimento_resultado" not in st.session_state:
    st.session_state.recebimento_resultado = None
if "recebimento_confirmado" not in st.session_state:
    st.session_state.recebimento_confirmado = False
if "recebimento_qtd_chegou" not in st.session_state:
    st.session_state.recebimento_qtd_chegou = 0
if "recebimento_rota_selecionada" not in st.session_state:
    st.session_state.recebimento_rota_selecionada = None
if "recebimento_ressalva" not in st.session_state:
    # Retirada a divisão por hub (pedido do Samuel em 24/09/2026): a
    # coluna "Source" do CSV do SCC não é um rótulo fixo de hub de
    # origem, é o local do ÚLTIMO SCAN - um pacote Delivered aparece
    # com Source "CUSTOMER_ADDRESS", não o hub de onde saiu (visto de
    # verdade num CSV real: 6 TBR com Source CUSTOMER_ADDRESS e 1 com
    # CGH7, nenhum desses era de outro hub de verdade). Separar
    # REC9/FOR3 automaticamente em cima dessa coluna não é confiável -
    # num arquivo com os dois hubs misturados não tem como saber a
    # qual dos dois um TBR com Source errado pertence. Voltou a ser um
    # número único combinado (sem escolha de hub, sem botão).
    #
    # "fotos" (lista, ajuste do Samuel em 24/09/2026 - antes era 1 foto
    # só): serve pra anexar mais de uma imagem/documento - foto do
    # Manifested (Bill of Lading), foto de pacote avariado, e afins -
    # tudo junto no mesmo registro de ressalva, sem limite de 1.
    #
    # Renomeada de "recebimento_ressalva_inducao" pra
    # "recebimento_ressalva" (ajuste do Samuel em 24/09/2026): não é
    # mais uma caixa exclusiva do Checkpoint 1 (indução) - virou uma
    # única ressalva "final", registrada depois do Checkpoint 2 (stow) -
    # ver tela_recebimento().
    st.session_state.recebimento_ressalva = {"texto": "", "confirmada": False, "fotos": []}
if "recebimento_tem_na" not in st.session_state:
    st.session_state.recebimento_tem_na = False
if "recebimento_estado_bruto_stow" not in st.session_state:
    # State BRUTO (texto exato do SCC) de cada TBR do último CSV de
    # stow subido - usado só pra mostrar no aviso de possível MNR (ver
    # tela_recebimento) qual state o SCC mostra pra cada TBR parado.
    st.session_state.recebimento_estado_bruto_stow = {}


def vai_para(tela):
    st.session_state.tela = tela


def _mostra_diferenca_checkpoint(rotulo_avancado, avancado, qtd_chegou):
    """Mostra o resultado de comparar 'quantos já avançaram nesse
    checkpoint' com 'quantos foram informados que chegaram' - usado nos
    dois checkpoints (indução/stow). Bug corrigido em 24/09/2026
    (apontado pelo Samuel com print real): quando avancado > qtd_chegou
    (ex: CSV mostra 307 induzido mas só foi digitado 298), a conta
    'faltam' dava NEGATIVO ("faltam -9"), sem sentido nenhum - esse caso
    normalmente não é pacote sumido, é o número digitado que ficou
    desatualizado (esquecido de um teste/turno anterior). Agora tem uma
    mensagem própria pra esse caso, avisando pra conferir o número
    digitado em vez de sugerir uma "falta" que não existe.

    Devolve a diferença (qtd_chegou - avancado) pra quem chamou decidir
    o que mostrar a mais (ex: a ressalva do líder só faz sentido quando
    realmente falta gente induzir/armazenar, ou seja diferenca > 0)."""
    diferenca = qtd_chegou - avancado
    if diferenca == 0:
        st.success(f"Bateu — {avancado} {rotulo_avancado} (ou além) = {qtd_chegou} que chegaram.")
    elif diferenca > 0:
        st.warning(
            f"Não bateu ainda — {avancado} {rotulo_avancado} (ou além) vs {qtd_chegou} "
            f"que chegaram (faltam {diferenca})."
        )
    else:
        st.warning(
            f"Tem MAIS TBR {rotulo_avancado} (ou além) do que a quantidade informada — "
            f"{avancado} vs {qtd_chegou} que chegaram ({-diferenca} a mais). Confere se o "
            'número em "Quantos pacotes chegaram" está certo (pode ter ficado de um '
            "teste ou turno anterior, por exemplo)."
        )
    return diferenca


# ------------------------------------------------------------------
# TELA INICIAL — escolha entre os dois painéis (Fechamento x
# Recebimento). Cada um é independente: Recebimento nem pede nome (não
# precisa, não salva nada com autoria), Fechamento continua exigindo
# como sempre. Ajuste do Samuel em 21/09/2026 — antes o Recebimento
# entrava como botão pequeno dentro do próprio painel de Fechamento;
# agora os dois ficam lado a lado, no mesmo topo, no mesmo app/deploy.
# ------------------------------------------------------------------
def tela_home():
    st.title("📦 Painel Operação DMNZ - AMXL")
    st.caption(f"{NODE_ATUAL} · Data: {dt.date.today().strftime('%d/%m/%Y')}")

    st.divider()

    col1, col2 = st.columns(2)
    with col1:
        st.markdown("#### 🚚 Painel Fechamento")
        st.caption("Rotas → lista pro SCC → CSV do SCC → dashboard de fechamento do dia.")
        if st.button("Abrir Painel Fechamento", use_container_width=True, type="primary"):
            vai_para("home_fechamento")
            st.rerun()
    with col2:
        st.markdown("#### 📥 Painel Recebimento")
        st.caption("Mesmo arquivo de rotas, mas pra ver pacotes e paradas por rota. Sem motorista.")
        if st.button("Abrir Painel Recebimento", use_container_width=True, type="primary"):
            vai_para("recebimento")
            st.rerun()


# ------------------------------------------------------------------
# PAINEL FECHAMENTO — home de verdade do fluxo de fechamento (nome de
# quem tá usando + os botões de ação). Era a tela_home() antiga.
# ------------------------------------------------------------------
def tela_home_fechamento():
    col_voltar, col_titulo = st.columns([1, 3], vertical_alignment="center")
    with col_voltar:
        if st.button("← Voltar"):
            vai_para("home")
            st.rerun()
    with col_titulo:
        st.markdown("**Painel Fechamento**")

    st.session_state.nome_usuario = st.text_input(
        "Seu nome",
        value=st.session_state.nome_usuario,
        placeholder="Ex: Samuel",
        help="Fica registrado nas edições (Editar TBR / Editar Reversa), pra saber quem mexeu em quê. Não precisa de senha.",
    )
    st.caption(f"Data: {dt.date.today().strftime('%d/%m/%Y')}")

    st.divider()

    col1, col2 = st.columns(2)
    with col1:
        if st.button(
            "▶️ Iniciar",
            use_container_width=True,
            type="primary",
            help="Fluxo normal, do zero: subir rotas → lista pro SCC → subir CSV → gerar fechamento.",
        ):
            if not st.session_state.nome_usuario.strip():
                st.warning("Digita seu nome antes de continuar.")
            else:
                vai_para("etapa1")
                st.rerun()
        st.button(
            "🔁 Reprocessar CSV SCC",
            use_container_width=True,
            disabled=True,
            help="Ainda não construído — em breve. Pra quando já rodou o Iniciar mas precisa "
            "subir um CSV corrigido (ou reaproveitar uma edição de TBR/Reversa) e gerar o "
            "fechamento de novo, sem refazer a consolidação de rotas do zero.",
        )
    with col2:
        st.button("✏️ Editar TBR", use_container_width=True, disabled=True,
                   help="Ainda não construído — em breve.")
        st.button("🔄 Editar Reversa", use_container_width=True, disabled=True,
                   help="Ainda não construído — em breve.")

    st.divider()
    if st.button("⚙️ Parceiros cadastrados", use_container_width=True):
        vai_para("parceiros")
        st.rerun()


# ------------------------------------------------------------------
# ETAPA 1 — rotas dos indianos -> lista pro SCC (+ upload do CSV do SCC
# ja na mesma tela, pronto pra quando a etapa 2 existir)
# ------------------------------------------------------------------
def tela_etapa1():
    col_voltar, col_titulo = st.columns([1, 3], vertical_alignment="center")
    with col_voltar:
        if st.button("← Voltar"):
            vai_para("home_fechamento")
            st.rerun()
    with col_titulo:
        st.markdown(
            f"**Etapa 1 — Consolidar rotas**  ·  {st.session_state.nome_usuario} · {NODE_ATUAL}"
        )

    # Depois que já leu um arquivo com sucesso, a caixa de upload não
    # tem mais utilidade na tela - some completamente (os dados já
    # foram extraídos e continuam guardados). Só um botão pequeno pra
    # trocar de arquivo, se precisar corrigir.
    ja_tem_dados = st.session_state.tbrs_por_rota is not None
    if ja_tem_dados:
        arquivo_rotas = None
        if st.button("🔁 Trocar arquivo de rotas"):
            st.session_state.tbrs_por_rota = None
            st.session_state.data_arquivo_rotas = None
            st.session_state.node_detectado_rotas = None
            st.rerun()
    else:
        arquivo_rotas = st.file_uploader(
            "Arquivo de rotas (dos indianos)", type=["xlsx", "xlsm", "xls"]
        )

    if arquivo_rotas is not None:
        try:
            tbrs_por_rota, data_arquivo = consolida(arquivo_rotas, arquivo_rotas.name)
            st.session_state.tbrs_por_rota = tbrs_por_rota
            st.session_state.data_arquivo_rotas = data_arquivo
            st.session_state.node_detectado_rotas = detecta_node_do_nome(arquivo_rotas.name)
        except ValueError as e:
            st.error(str(e))
            st.session_state.tbrs_por_rota = None
            st.session_state.data_arquivo_rotas = None
            st.session_state.node_detectado_rotas = None

    if st.session_state.tbrs_por_rota:
        tbrs_por_rota = st.session_state.tbrs_por_rota
        data_arquivo = st.session_state.data_arquivo_rotas
        node_detectado = st.session_state.node_detectado_rotas
        total_rotas = sum(len(v) for v in tbrs_por_rota.values())

        node_texto = node_detectado or "não identificado"
        data_texto = data_arquivo.strftime("%d/%m/%Y") if data_arquivo else "não identificada"

        # Se desmarcar de novo, não conta o que tinha digitado antes -
        # "não tem NA" precisa realmente zerar, mesmo que o texto ainda
        # esteja guardado por baixo dos panos. Calculado ANTES da grade
        # (a caixa "Lista de TBRs" já precisa desse total pronto).
        tbrs_na = le_tbrs_colados(st.session_state.get("texto_na", "")) if st.session_state.tem_na else []
        lista_final = [tbr for tbrs in tbrs_por_rota.values() for tbr in tbrs] + tbrs_na
        total_geral = len(lista_final)

        # Avisos de node/data errados ficam ANTES da grade de caixas (não
        # no meio dela) - senão abrem um vão entre as duas linhas e elas
        # deixam de se encostar.
        if node_detectado and node_detectado != NODE_ATUAL:
            st.warning(
                f"O nome do arquivo parece ser do node **{node_detectado}**, "
                f"mas essa versão do robô só trata **{NODE_ATUAL}**. Confira "
                "se é o arquivo certo antes de seguir."
            )
        if data_arquivo and data_arquivo != dt.date.today():
            st.warning(
                f"Esse arquivo é de **{data_texto}**, não é de hoje "
                f"({dt.date.today().strftime('%d/%m/%Y')}). Confira se não "
                "subiu o arquivo do dia errado antes de seguir."
            )

        # Grade 2x2: as 4 caixas com a MESMA altura e a MESMA proporção
        # (colunas 1:1 nas duas linhas, mesmo gap na horizontal e sem
        # texto solto entre as linhas) - título e legenda de cada caixa
        # ficam DENTRO dela (não soltos acima), pra elas se encostarem
        # de verdade, tipo um cruzamento. Só por CSS (min-height), nunca
        # pelo parâmetro height= do st.container - esse parâmetro
        # transforma a caixa numa área com scroll interno, e isso quebra
        # o toque no celular (o dedo sempre mexe um pouquinho entre
        # tocar e soltar, e o navegador interpreta esse movimento dentro
        # de uma área com scroll como "rolar a página" em vez de
        # "clicar"). Com min-height a caixa cresce à vontade se precisar
        # (nunca corta nada) e o toque funciona normal.
        ALTURA_CAIXAS = 260
        st.markdown(
            f"""
            <style>
            .st-key-caixa_arquivo, .st-key-caixa_na,
            .st-key-caixa_rotas, .st-key-caixa_lista {{
                min-height: {ALTURA_CAIXAS}px;
            }}
            </style>
            """,
            unsafe_allow_html=True,
        )

        col_info, col_na = st.columns([1, 1], gap="small")
        with col_info:
            with st.container(border=True, key="caixa_arquivo", vertical_alignment="center"):
                st.markdown("**Arquivo lido**")
                st.write(f"Node: **{node_texto}**  ·  Data: **{data_texto}**")
        with col_na:
            with st.container(border=True, key="caixa_na", vertical_alignment="center"):
                # Era um st.checkbox, mas ele simplesmente não respondia ao
                # toque no celular (testado várias vezes, sempre falhou -
                # é um problema conhecido desse componente em alguns
                # celulares/navegadores). Trocado por um botão normal, que
                # já provou funcionar em toda a tela (Voltar, Trocar
                # arquivo, Iniciar Fechamento) - só liga/desliga um
                # "interruptor" guardado no session_state.
                rotulo_botao = (
                    "✅ Tem TBRs de NA (toque para desmarcar)"
                    if st.session_state.tem_na
                    else "☐ Tem TBRs de NA para consulta no SCC? (opcional — toque para marcar)"
                )
                if st.button(rotulo_botao, key="botao_tem_na", use_container_width=True):
                    st.session_state.tem_na = not st.session_state.tem_na
                    st.rerun()
                if st.session_state.tem_na:
                    st.text_area(
                        "Cole os TBRs de NA aqui, um por linha",
                        key="texto_na",
                        height=80,
                        label_visibility="collapsed",
                    )

        # Segunda linha da grade, MESMA proporção (1:1) e MESMO gap da
        # primeira linha - assim as 4 caixas formam um quadriculado só,
        # não dois blocos separados.
        col_rotas, col_lista = st.columns([1, 1], gap="small")

        with col_rotas:
            # Essa caixa não tem nenhum botão/checkbox dentro (só texto
            # estático), então pode usar height= de verdade (com scroll
            # se tiver muitas rotas) sem o risco de travar toque - o
            # problema de scroll cancelando clique só existe quando tem
            # algo clicável dentro pra tocar. SEM vertical_alignment=
            # "center" aqui de propósito: com muitas rotas o conteúdo
            # passa da altura da caixa, e centralizar um conteúdo maior
            # que a caixa empurra o título ("Rotas x pacotes") pra fora
            # da área visível, escondendo-o. Alinhado no topo, o título
            # sempre aparece primeiro, e só o que sobra rola pra baixo.
            with st.container(border=True, key="caixa_rotas", height=ALTURA_CAIXAS):
                st.markdown("**Rotas x pacotes**")
                st.caption(f"Total: {len(tbrs_por_rota)} rota(s), {total_rotas} TBR(s)")
                # Conteúdo centralizado (vertical E horizontal) na caixa -
                # com poucas rotas (caso comum do LRN9) não fica "grudado"
                # no canto, sobrando vazio ao redor. O alinhamento horizontal
                # do st.container sozinho não centraliza o TEXTO dentro da
                # linha (só centralizaria o bloco, que já ocupa a largura
                # toda) - por isso o texto vai direto em HTML com
                # text-align:center.
                linhas_rotas = "<br>".join(
                    f"{rota}: {len(v)}" for rota, v in tbrs_por_rota.items()
                )
                st.markdown(
                    f"<div style='text-align:center'>{linhas_rotas}</div>",
                    unsafe_allow_html=True,
                )

        with col_lista:
            with st.container(border=True, key="caixa_lista"):
                legenda_na = f" ({total_rotas} + {len(tbrs_na)} de NA)" if tbrs_na else ""
                st.markdown(f"**Lista de TBRs — {total_geral} TBR(s){legenda_na}**")
                st.caption("Ícone de copiar no canto do quadro pega a lista inteira de uma vez.")
                # Altura do bloco de código um pouco menor que a caixa
                # (sobra espaço do título/legenda acima) - o resto da
                # lista continua acessível rolando dentro do quadro, ou
                # pelo ícone de copiar, que pega tudo de uma vez.
                st.code(
                    "\n".join(lista_final),
                    language=None,
                    height=ALTURA_CAIXAS - 130,
                )

        st.write("")
        st.markdown("**Etapa 2 — CSV do SCC**")
        st.caption("Cole a lista no SCC, exporte e suba o CSV aqui.")
        # SEM restrição de tipo (type=None) de propósito - com type=["csv"] o
        # seletor de arquivo do Android/Chrome às vezes deixa o próprio CSV
        # exportado do SCC APAGADO/bloqueado na lista, porque o navegador
        # salva o download com um "tipo" (MIME) genérico em vez de
        # reconhecer como CSV de verdade, e o seletor filtra por esse tipo
        # (não só pela extensão) - visto em vídeo de teste real no celular
        # (09/09/2026). Sem essa restrição o seletor mostra todos os
        # arquivos; quem valida se é um CSV de verdade é o le_csv_scc logo
        # abaixo (já mostra um st.error claro se não conseguir ler).
        arquivo_csv = st.file_uploader("CSV exportado do SCC", type=None, key="csv_scc")
        if arquivo_csv is not None and not arquivo_csv.name.lower().endswith(".csv"):
            st.warning(
                f"O arquivo **{arquivo_csv.name}** não parece ser um `.csv`. "
                "Confira se é o arquivo certo antes de continuar."
            )

        # Primeiro pedaço de verdade da Etapa 2: só a parte de manter a
        # base de DAs (quem é DMNZ) em dia. O resto (State Finalizador
        # por TBR, gerar o CSV final com as colunas adicionadas) ainda
        # não foi construído - de propósito, um pedaço de cada vez.
        #
        # A base mora no GitHub (não é upload manual) - o app lê sozinho
        # e, quando o usuário confirma DAs novos, salva direto lá. Se o
        # arquivo ainda não existir no repositório, a primeira gravação
        # já cria ele - não precisa criar nada na mão antes.
        if arquivo_csv is not None:
            if not secrets_configurados():
                st.warning(
                    "A base de DAs ainda não está configurada pra salvar no GitHub "
                    "(falta o secret `github` no Streamlit Cloud - token + repositório). "
                    "Até isso ser configurado, essa parte não funciona."
                )
            else:
                try:
                    linhas_csv = le_csv_scc(arquivo_csv)
                except Exception as e:
                    st.error(f"Não consegui ler esse CSV: {e}")
                    linhas_csv = None

                if linhas_csv is not None:
                    hash_csv = hashlib.md5(arquivo_csv.getvalue()).hexdigest()
                    das_dia = extrai_das_do_dia(linhas_csv)

                    if not das_dia:
                        st.info(
                            "Nenhum motorista de verdade apareceu nesse CSV (só leituras "
                            "de sistema/suporte, tipo LastMileRoutePlanner). Nada pra "
                            "classificar."
                        )
                    else:
                        try:
                            texto_base, sha_base = le_arquivo(CAMINHO_BASE_DAS)
                            base_dict, base_bruta = parseia_texto_base(texto_base)
                        except Exception as e:
                            st.error(f"Não consegui ler a base de DAs no GitHub: {e}")
                            base_dict, base_bruta, sha_base = {}, [], None

                        ja_cadastrados = {normaliza_nome(n) for n, _ in base_bruta}
                        desconhecidos = [
                            n for n in das_dia if normaliza_nome(n) not in ja_cadastrados
                        ]
                        qtd_dmnz = sum(
                            1 for n in das_dia if base_dict.get(normaliza_nome(n)) == "DMNZ"
                        )

                        st.markdown("**Base de DAs (quem é DMNZ)**")
                        st.caption(
                            f"Lida direto do repositório — {len(base_bruta)} DA(s) "
                            "cadastrado(s) no total."
                            if base_bruta
                            else "Base ainda vazia no repositório — é a primeira vez."
                        )
                        st.write(
                            f"**{len(das_dia)} motorista(s) diferentes rodaram hoje** "
                            f"— {len(das_dia) - len(desconhecidos)} já cadastrado(s) "
                            f"({qtd_dmnz} como DMNZ)."
                        )

                        if not desconhecidos:
                            st.success(
                                "Todos os motoristas de hoje já estão cadastrados na base."
                            )
                            das_revisados = True
                        else:
                            escolhidos = st.pills(
                                f"{len(desconhecidos)} motorista(s) ainda não "
                                "cadastrado(s) — quais são DMNZ?",
                                desconhecidos,
                                selection_mode="multi",
                                key="pills_desconhecidos",
                            )
                            st.caption(
                                f"{len(escolhidos)} selecionado(s). Quem não for tocado "
                                "fica de fora da base — nunca vai ser perguntado de novo "
                                "sobre esse CSV, mas se tocar sem querer dá pra desmarcar "
                                "de novo antes de continuar."
                            )

                            nova_base_bruta = None
                            if escolhidos:
                                nova_base_bruta = base_bruta + [
                                    (nome, "DMNZ") for nome in escolhidos
                                ]
                                with st.expander(
                                    f"Prévia da base atualizada "
                                    f"({len(nova_base_bruta)} DA(s) no total)"
                                ):
                                    st.dataframe(
                                        [
                                            {"Nome do DA": n, "DMNZ": c}
                                            for n, c in nova_base_bruta
                                        ],
                                        use_container_width=True,
                                        hide_index=True,
                                    )

                            rotulo_continuar = (
                                "💾 Salvar e continuar" if escolhidos
                                else "➡️ Continuar sem marcar ninguém como DMNZ"
                            )
                            if st.button(
                                rotulo_continuar,
                                key="continuar_base_das",
                                use_container_width=True,
                            ):
                                if nova_base_bruta is None:
                                    # ninguém marcado - não tem nada novo pra
                                    # gravar no GitHub, só avança pro próximo
                                    # passo (quem ficou de fora é tratado como
                                    # parceiro, igual no desktop)
                                    st.session_state["das_revisados_para"] = hash_csv
                                    st.rerun()
                                else:
                                    try:
                                        salva_arquivo(
                                            CAMINHO_BASE_DAS,
                                            gera_csv_base_das(nova_base_bruta),
                                            sha_base,
                                            mensagem=(
                                                f"Adiciona DA(s) DMNZ: "
                                                f"{', '.join(escolhidos)}"
                                            ),
                                        )
                                    except ConflitoDeSalvamento:
                                        st.error(
                                            "Alguém salvou a base ao mesmo tempo. Toca "
                                            "no botão de novo pra tentar com a versão "
                                            "mais recente."
                                        )
                                    except Exception as e:
                                        st.error(f"Não consegui salvar no GitHub: {e}")
                                    else:
                                        st.session_state["das_revisados_para"] = hash_csv
                                        st.success(
                                            f"Base atualizada! {len(escolhidos)} "
                                            "motorista(s) novo(s) salvo(s) como DMNZ."
                                        )
                                        st.rerun()

                            das_revisados = (
                                st.session_state.get("das_revisados_para") == hash_csv
                            )

                        # State Finalizador só entra depois que a revisão da
                        # base de DAs foi concluída pra esse CSV - mesma ordem
                        # do robô desktop, que termina toda a revisão de DAs
                        # antes de perguntar o State Finalizador de cada TBR.
                        # IMPORTANTE: não espera "desconhecidos" ficar vazio -
                        # DA de parceiro nunca entra na base (fica sempre como
                        # "desconhecido"), então isso travaria pra sempre.
                        if das_revisados:
                            try:
                                texto_historico, sha_historico = le_arquivo(CAMINHO_HISTORICO_TBR)
                                base_tbr = parseia_texto_historico(texto_historico)
                            except Exception as e:
                                st.error(f"Não consegui ler o histórico de TBRs no GitHub: {e}")
                                base_tbr, sha_historico = {}, None
                            else:
                                data_hoje_str = dt.date.today().strftime("%d/%m/%Y")
                                # Não bloqueia o fluxo se der erro/estiver
                                # vazio - cai pro comportamento antigo
                                # dentro de classifica_insucesso (ver
                                # calculo_fechamento.py). Cadastro ainda é
                                # opcional, não é pré-requisito pra fechar o dia.
                                try:
                                    texto_parceiros, _ = le_arquivo(CAMINHO_PARCEIROS)
                                    lista_parceiros = parceiros_mod.parseia_texto_parceiros(
                                        texto_parceiros
                                    )
                                except Exception:
                                    lista_parceiros = []
                                resultado = processa_state_finalizador(
                                    linhas_csv, base_tbr, base_dict, data_hoje_str,
                                    tbrs_marcados_na=tbrs_na,
                                )

                                # Conferência manual: DA de cada rota, pra
                                # pegar troca entre DMNZ x Parceiro (ex: um
                                # DA nosso resgatou uma rota do parceiro, ou
                                # o contrário) antes de seguir com o resto.
                                resumo_das = resumo_das_por_rota(linhas_csv, base_dict)
                                rotas_mistas = sorted({
                                    r["rota"] for r in resumo_das if r["mista"]
                                })
                                if rotas_mistas:
                                    st.warning(
                                        "⚠️ Rota(s) com DA de DMNZ e de Parceiro "
                                        f"misturados: {', '.join(rotas_mistas)}. "
                                        "Vale conferir se foi resgate ou divisão."
                                    )
                                with st.expander(
                                    f"🔍 Conferência: DA de cada rota (DMNZ x Parceiro) "
                                    f"— {len(resumo_das)} registro(s)"
                                ):
                                    st.dataframe(
                                        [
                                            {
                                                "Rota": r["rota"],
                                                "DA": r["da"],
                                                "Entregues": r["entregues"],
                                                "Em rota": r["em_rota"],
                                                "A analisar": r["a_analisar"],
                                                "Total": r["pacotes"],
                                                "Classificação": r["classificacao"],
                                            }
                                            for r in resumo_das
                                        ],
                                        use_container_width=True,
                                        hide_index=True,
                                    )

                                # Pacotes sem DA de verdade vinculado (login
                                # de suporte tipo e-mail, ou nome de sistema)
                                # - o robô desktop mostra isso separado na
                                # prévia final, igual aqui. Cruza com o
                                # motorista real da rota só quando a entrega
                                # foi confirmada (Delivered) - pros outros
                                # states ninguém confirmou quem está com o
                                # pacote de verdade.
                                sem_da = lista_sem_da_de_verdade(
                                    linhas_csv, base_dict,
                                    resultado["motorista_real_da_rota"],
                                )
                                if sem_da:
                                    with st.expander(
                                        f"📋 Pacotes sem DA de verdade vinculado "
                                        f"(login de suporte/sistema) "
                                        f"— {len(sem_da)} registro(s)"
                                    ):
                                        st.dataframe(
                                            [
                                                {
                                                    "TBR": r["tbr"],
                                                    "Rota": r["rota"],
                                                    "State": r["state_scc"],
                                                    "Motorista real da rota": (
                                                        r["motorista_real_da_rota"]
                                                    ),
                                                    "DAs DMNZ": r["das_dmnz"],
                                                }
                                                for r in sem_da
                                            ],
                                            use_container_width=True,
                                            hide_index=True,
                                        )

                                st.markdown("**State Finalizador (por TBR)**")
                                st.write(
                                    f"**{len(resultado['entregues'])} entregue(s)** "
                                    "(automático) · "
                                    f"**{len(resultado['auto_em_rota'])} em rota** "
                                    "(automático, DA de verdade vinculado) · "
                                    f"**{len(resultado['auto_retorno_insucesso'])} retorno(s)** "
                                    "(automático, era \"EM ROTA\" e voltou Received) · "
                                    f"**{len(resultado['auto_cancelado'])} cancelado(s)** "
                                    "(automático, antes da rota) · "
                                    f"**{len(resultado['auto_outro_node'])} outro node** "
                                    "(automático) · "
                                    f"**{len(resultado['auto_na'])} NA c/ rota** "
                                    "(automático, acompanhamento) · "
                                    f"**{len(resultado['precisa_validacao_na'])} NA** "
                                    "(validar motorista) · "
                                    f"**{len(resultado['conhecidos'])} reaproveitado(s)** "
                                    "do histórico · "
                                    f"**{len(resultado['pendentes'])} pendente(s)** "
                                    "de revisão."
                                )

                                # Reaproveitados do histórico ficam só no
                                # número do resumo acima (sem listar um a
                                # um) - ao contrário de "sem DA vinculado" ou
                                # dos pendentes, esses não pedem nenhuma ação
                                # do usuário (nada mudou neles), então uma
                                # caixinha a mais só ocuparia espaço na tela
                                # sem ajudar em nada.

                                # Alerta (só aviso, não muda pra onde o TBR vai) pros
                                # TBRs (pendentes OU já classificados automático como
                                # NA) cujo state mudou desde a última vez SEM ter
                                # vindo de um motorista de verdade - pode ser só o
                                # próprio sistema do SCC reprocessando sozinho (caso
                                # investigado pelo Samuel: TBR353806057, preso
                                # girando em Route Assignment sem confirmação de
                                # ninguém). Calculado ANTES do if/else porque pode
                                # acontecer mesmo sem nenhum pendente sobrando (só
                                # com NA automático).
                                suspeitos = [
                                    r for r in resultado["pendentes"] + resultado["auto_na"]
                                    if r.get("alerta_sistemico")
                                ]
                                if suspeitos:
                                    st.warning(
                                        f"⚠️ {len(suspeitos)} TBR(s) com mudança de "
                                        "state que parece ter vindo do próprio "
                                        "sistema do SCC (não de um motorista de "
                                        "verdade) — vale conferir com mais atenção "
                                        "antes de classificar:\n\n"
                                        + "\n".join(
                                            f"- **{r['tbr']}**: {r['motivo_reabertura']} "
                                            f"(leitura: {r['last_scan_by']})"
                                            for r in suspeitos
                                        )
                                    )

                                if not resultado["pendentes"] and not resultado["precisa_validacao_na"]:
                                    st.success(
                                        "Nenhum TBR pendente de revisão manual hoje."
                                    )
                                    automaticos_sem_pendentes = (
                                        resultado["auto_em_rota"]
                                        + resultado["auto_retorno_insucesso"]
                                        + resultado["auto_cancelado"]
                                        + resultado["auto_outro_node"]
                                        + resultado["auto_na"]
                                    )
                                    if automaticos_sem_pendentes:
                                        with st.expander(
                                            f"👁️ Prévia — {len(automaticos_sem_pendentes)} "
                                            "TBR(s) que serão salvos no histórico"
                                        ):
                                            st.dataframe(
                                                [
                                                    {
                                                        "TBR": r["tbr"],
                                                        "State Finalizador": r["resposta"],
                                                        "State SCC": r["state_scc"],
                                                    }
                                                    for r in automaticos_sem_pendentes
                                                ],
                                                use_container_width=True,
                                                hide_index=True,
                                            )
                                    if automaticos_sem_pendentes and st.button(
                                        "💾 Confirmar e salvar histórico",
                                        key="salvar_historico_sem_pendentes",
                                        use_container_width=True,
                                    ):
                                        for r in automaticos_sem_pendentes:
                                            atualiza_base(
                                                base_tbr, r["tbr"], r["resposta"],
                                                r["state_scc"], data_hoje_str,
                                            )
                                        try:
                                            salva_arquivo(
                                                CAMINHO_HISTORICO_TBR,
                                                gera_csv_historico(base_tbr),
                                                sha_historico,
                                                mensagem=(
                                                    f"Atualiza State Finalizador "
                                                    f"({len(automaticos_sem_pendentes)} "
                                                    "automático(s))"
                                                ),
                                            )
                                        except ConflitoDeSalvamento:
                                            st.error(
                                                "Alguém salvou o histórico ao mesmo "
                                                "tempo. Toca no botão de novo pra "
                                                "tentar com a versão mais recente."
                                            )
                                        except Exception as e:
                                            st.error(f"Não consegui salvar no GitHub: {e}")
                                        else:
                                            dados_dia = monta_dados_do_dia(
                                                linhas_csv, base_tbr, base_dict,
                                                resultado["motorista_real_da_rota"],
                                            )
                                            st.session_state.dados_fechamento = calcula_tudo(dados_dia, NODE_ATUAL, lista_parceiros)
                                            st.session_state.export_fechamento_insumos = _monta_insumos_export(
                                                linhas_csv, base_tbr, base_dict, resultado, data_hoje_str,
                                            )
                                            vai_para("previa_fechamento")
                                            st.rerun()
                                    elif st.button(
                                        "➡️ Ver prévia do fechamento",
                                        key="ver_previa_ja_salvo",
                                        use_container_width=True,
                                    ):
                                        # Já estava tudo salvo de uma visita anterior
                                        # (nada novo pra gravar) - só monta a prévia
                                        # de novo em cima do que já está no histórico.
                                        dados_dia = monta_dados_do_dia(
                                            linhas_csv, base_tbr, base_dict,
                                            resultado["motorista_real_da_rota"],
                                        )
                                        st.session_state.dados_fechamento = calcula_tudo(dados_dia, NODE_ATUAL, lista_parceiros)
                                        st.session_state.export_fechamento_insumos = _monta_insumos_export(
                                            linhas_csv, base_tbr, base_dict, resultado, data_hoje_str,
                                        )
                                        vai_para("previa_fechamento")
                                        st.rerun()
                                else:
                                    # Validação de NA (regra menos restritiva, pedida
                                    # pelo Samuel): TBR que o usuário colou na
                                    # caixinha "Tem TBRs de NA" da Etapa 1 e que
                                    # ainda voltou do SCC sem motorista confirmado -
                                    # em vez de tentar adivinhar pelo Route Code do
                                    # CSV (nem sempre vem preenchido), oferece uma
                                    # lista suspensa com quem já rodou hoje pro
                                    # próprio usuário escolher. Ao escolher, vira
                                    # "EM ROTA - DMNZ/PARCEIRO" igual um TBR normal
                                    # (mesma função classifica_dmnz_ou_parceiro).
                                    selecoes_na = {}
                                    if resultado["precisa_validacao_na"]:
                                        st.markdown(
                                            f"**Validação de NA — "
                                            f"{len(resultado['precisa_validacao_na'])} "
                                            "TBR(s)**"
                                        )
                                        st.caption(
                                            "Colados como NA na Etapa 1, ainda sem "
                                            "motorista confirmado no SCC. Se já souber "
                                            "quem pegou, escolhe na lista — quem ficar "
                                            "sem escolha continua pendente pra próxima "
                                            "vez."
                                        )
                                        opcoes_dropdown = opcoes_da_rota_do_dia(linhas_csv)
                                        rotulos_dropdown = ["— selecionar —"] + [
                                            f"{nome} — {rota}"
                                            for nome, rota in opcoes_dropdown
                                        ]
                                        for r in resultado["precisa_validacao_na"]:
                                            escolha = st.selectbox(
                                                r["tbr"],
                                                rotulos_dropdown,
                                                key=f"na_validacao_{r['tbr']}",
                                                help=f"State SCC: {r['state_scc']}",
                                            )
                                            indice = rotulos_dropdown.index(escolha)
                                            if indice > 0:
                                                selecoes_na[r["tbr"]] = opcoes_dropdown[
                                                    indice - 1
                                                ]

                                    validados_na = []
                                    for r in resultado["precisa_validacao_na"]:
                                        if r["tbr"] in selecoes_na:
                                            nome, rota = selecoes_na[r["tbr"]]
                                            classificacao = classifica_dmnz_ou_parceiro(
                                                nome, rota, base_dict,
                                                resultado["motorista_real_da_rota"],
                                            )
                                            validados_na.append({
                                                **r,
                                                "resposta": f"EM ROTA - {classificacao}",
                                            })
                                    faltam_na = (
                                        len(resultado["precisa_validacao_na"])
                                        - len(selecoes_na)
                                    )
                                    if faltam_na:
                                        st.caption(
                                            f"Faltam {faltam_na} TBR(s) de NA sem "
                                            "motorista escolhido — continuam pendentes "
                                            "pra próxima vez."
                                        )

                                    preenchidos = {}
                                    if resultado["pendentes"]:
                                        st.caption(
                                            "Copia a lista de cada grupo, pesquisa todos de "
                                            "uma vez no SCC, e vai digitando o State "
                                            "Finalizador de cada um. Agrupado por State do "
                                            "SCC, igual no desktop."
                                        )
                                        grupos = {}
                                        for r in resultado["pendentes"]:
                                            grupos.setdefault(r["state_scc"], []).append(r)

                                        # Grade de 3 colunas pros campos de texto - o rótulo
                                        # fica só o TBR (o motivo vira dica no ícone de
                                        # ajuda, "?"), e o texto digitado pode ficar
                                        # truncado dentro da caixa estreita sem problema
                                        # (o valor continua salvo inteiro por baixo, só a
                                        # exibição que corta, igual a lista de TBRs
                                        # consolidados que também não mostra tudo de uma
                                        # vez na caixa).
                                        COLUNAS_POR_LINHA = 3
                                        respostas = {}
                                        for state, itens in grupos.items():
                                            st.markdown(f"**{state}** ({len(itens)} TBR(s))")
                                            # Lista pra copiar e colar de uma vez no SCC -
                                            # separada por "," (testado no aparelho real:
                                            # com ";" o SCC dá erro na busca, só vírgula
                                            # funciona certo - correção de uma suposição
                                            # anterior que tinha ficado errada). Sem
                                            # height= fixo - como fica tudo numa linha só
                                            # (não quebra sozinha), uma altura fixa só
                                            # sobrava espaço vazio embaixo; sem ela a
                                            # caixa se ajusta à própria linha, rolando pro
                                            # lado se for comprida demais (o ícone de
                                            # copiar pega a lista inteira de qualquer
                                            # forma).
                                            st.code(
                                                ",".join(r["tbr"] for r in itens),
                                                language=None,
                                            )
                                            colunas = st.columns(COLUNAS_POR_LINHA)
                                            for i, r in enumerate(itens):
                                                with colunas[i % COLUNAS_POR_LINHA]:
                                                    respostas[r["tbr"]] = st.text_input(
                                                        r["tbr"],
                                                        key=f"finalizador_{r['tbr']}",
                                                        help=r["motivo_reabertura"],
                                                    )

                                        preenchidos = {
                                            tbr: v.strip()
                                            for tbr, v in respostas.items()
                                            if v.strip()
                                        }
                                        faltam = len(resultado["pendentes"]) - len(preenchidos)
                                        if faltam:
                                            st.caption(
                                                f"Faltam {faltam} TBR(s) sem State "
                                                "Finalizador preenchido — dá pra salvar só "
                                                "os que já foram preenchidos, os outros "
                                                "continuam pendentes pra próxima vez."
                                            )

                                    # Prévia mostra TUDO que vai ser processado e
                                    # salvo nessa confirmação - os automáticos (em
                                    # rota / sem state), os TBRs de NA validados
                                    # agora e o que acabou de ser digitado, com uma
                                    # coluna "Origem" pra diferenciar - não só o que
                                    # foi digitado agora, senão o usuário não teria
                                    # o vislumbre completo antes de confirmar.
                                    automaticos = (
                                        resultado["auto_em_rota"]
                                        + resultado["auto_retorno_insucesso"]
                                        + resultado["auto_cancelado"]
                                        + resultado["auto_outro_node"]
                                        + resultado["auto_na"]
                                    )
                                    linhas_previa = [
                                        {
                                            "TBR": r["tbr"],
                                            "State Finalizador": r["resposta"],
                                            "Origem": "Automático",
                                        }
                                        for r in automaticos
                                    ] + [
                                        {
                                            "TBR": r["tbr"],
                                            "State Finalizador": r["resposta"],
                                            "Origem": "Validado (NA)",
                                        }
                                        for r in validados_na
                                    ] + [
                                        {
                                            "TBR": tbr,
                                            "State Finalizador": v,
                                            "Origem": "Digitado agora",
                                        }
                                        for tbr, v in preenchidos.items()
                                    ]

                                    if linhas_previa:
                                        with st.expander(
                                            f"👁️ Prévia — {len(linhas_previa)} TBR(s) "
                                            "que serão salvos no histórico"
                                        ):
                                            st.dataframe(
                                                linhas_previa,
                                                use_container_width=True,
                                                hide_index=True,
                                            )

                                        if st.button(
                                            "💾 Confirmar e salvar histórico",
                                            key="salvar_historico",
                                            use_container_width=True,
                                        ):
                                            for r in automaticos + validados_na:
                                                atualiza_base(
                                                    base_tbr, r["tbr"], r["resposta"],
                                                    r["state_scc"], data_hoje_str,
                                                )
                                            for r in resultado["pendentes"]:
                                                if r["tbr"] in preenchidos:
                                                    atualiza_base(
                                                        base_tbr, r["tbr"],
                                                        preenchidos[r["tbr"]],
                                                        r["state_scc"], data_hoje_str,
                                                    )
                                            try:
                                                salva_arquivo(
                                                    CAMINHO_HISTORICO_TBR,
                                                    gera_csv_historico(base_tbr),
                                                    sha_historico,
                                                    mensagem=(
                                                        f"Atualiza State Finalizador de "
                                                        f"{len(linhas_previa)} TBR(s)"
                                                    ),
                                                )
                                            except ConflitoDeSalvamento:
                                                st.error(
                                                    "Alguém salvou o histórico ao mesmo "
                                                    "tempo. Toca no botão de novo pra "
                                                    "tentar com a versão mais recente."
                                                )
                                            except Exception as e:
                                                st.error(
                                                    f"Não consegui salvar no GitHub: {e}"
                                                )
                                            else:
                                                dados_dia = monta_dados_do_dia(
                                                    linhas_csv, base_tbr, base_dict,
                                                    resultado["motorista_real_da_rota"],
                                                )
                                                st.session_state.dados_fechamento = calcula_tudo(dados_dia, NODE_ATUAL, lista_parceiros)
                                                st.session_state.export_fechamento_insumos = _monta_insumos_export(
                                                    linhas_csv, base_tbr, base_dict, resultado, data_hoje_str,
                                                )
                                                vai_para("previa_fechamento")
                                                st.rerun()


# ------------------------------------------------------------------
# PRÉVIA GERAL DO FECHAMENTO — aparece depois que o State Finalizador
# já foi salvo no histórico. Mostra tudo que foi lido e processado (menos
# o painel de Reversa, que ainda não entrou no fluxo web) antes de liberar
# a Etapa 3 (Gerar Fechamento - o dashboard final, ainda não construído).
# ------------------------------------------------------------------
CSS_PAINEL_FECHAMENTO = """
<style>
.painel-lrn9 { border: 1px solid #e2e2e2; border-radius: 8px; overflow: hidden;
               margin-bottom: 14px; }
.painel-lrn9 .cabecalho { background: #E8590C; color: white; padding: 8px 14px;
               font-weight: 600; font-size: 0.95em; }
.painel-lrn9 .corpo { padding: 10px 14px 12px; background: white; color: #1a1a1a; }
.painel-lrn9 .linha { display:flex; justify-content:space-between; gap: 8px;
               padding: 3px 0; border-bottom: 1px dotted #ddd; font-size: 0.92em; }
.painel-lrn9 .linha.sub { padding-left: 14px; color:#555; }
.painel-lrn9 .linha.total { font-weight:700; border-top: 1px solid #999;
               border-bottom: none; margin-top:4px; padding-top:6px; }
.painel-lrn9 .linha.categoria { font-weight:600; border-bottom:none;
               padding-top: 8px; }
.painel-lrn9 .linha.categoria:first-child { padding-top: 0; }
</style>
"""


def _linha_painel(label, valor, sub=False, total=False):
    classe = "linha" + (" sub" if sub else "") + (" total" if total else "")
    return f'<div class="{classe}"><span>{label}</span><span>{valor}</span></div>'


def _linha_categoria(label):
    # Só o rótulo, sem número do lado - o número do template original só
    # aparece na linha do State Finalizador (dentro da categoria), não na
    # categoria em si. Repetir o total na categoria E na única linha de
    # baixo (ex: "Delivered 490" seguido de "Entregue 490") ficava
    # parecendo duplicado quando só tem um State Finalizador ali dentro.
    return f'<div class="linha categoria"><span>{label}</span></div>'


def _painel(titulo, linhas_html):
    return (
        f'<div class="painel-lrn9"><div class="cabecalho">{titulo}</div>'
        f'<div class="corpo">{linhas_html}</div></div>'
    )


def tela_previa_fechamento():
    col_voltar, col_titulo = st.columns([1, 3], vertical_alignment="center")
    with col_voltar:
        if st.button("← Voltar"):
            vai_para("home_fechamento")
            st.rerun()
    # Data do ARQUIVO de rotas (Etapa 1), não a data de hoje - o
    # fechamento pode ser rodado num dia diferente do dia do arquivo (ex:
    # rodou de madrugada, ou terminou de digitar o State Finalizador só
    # no dia seguinte). Cai pra hoje só se a data não tiver sido
    # detectada no arquivo.
    data_do_fechamento = st.session_state.data_arquivo_rotas or dt.date.today()
    with col_titulo:
        st.markdown(
            f"**Prévia do Fechamento**  ·  {st.session_state.nome_usuario} · "
            f"{NODE_ATUAL} · {data_do_fechamento.strftime('%d/%m/%Y')}"
        )
    st.caption(
        "State Finalizador já salvo no histórico. Confira os números antes de gerar "
        "o fechamento."
    )

    r = st.session_state.dados_fechamento
    if r is None:
        st.warning(
            "Não achei dados pra montar a prévia (a sessão pode ter expirado). "
            "Volta pro início e roda o State Finalizador de novo."
        )
        return

    st.markdown(CSS_PAINEL_FECHAMENTO, unsafe_allow_html=True)

    col1, col2, col3 = st.columns(3)

    with col1:
        linhas = "".join(
            _linha_painel(cat, qtd) for cat, qtd in r["contagem_scc"].most_common()
        )
        linhas += _linha_painel("Total Geral", r["total_geral"], total=True)
        st.markdown(_painel("1. STATE SCC", linhas), unsafe_allow_html=True)

        st.markdown(
            _painel("2. STATE ENTREGAS DMNZ", _linha_painel("DMNZ", r["n_dmnz"])),
            unsafe_allow_html=True,
        )

    with col2:
        blocos = []
        for cat, subitens in r["nested"].items():
            blocos.append(_linha_categoria(cat))
            for fin, qtd in subitens.items():
                blocos.append(_linha_painel(fin or "(sem State Finalizador)", qtd, sub=True))
        blocos.append(_linha_painel("Total Geral", r["total_geral"], total=True))
        st.markdown(
            _painel("3. STATE SCC & ANÁLISE DMNZ", "".join(blocos)),
            unsafe_allow_html=True,
        )

    with col3:
        # "sem_detalhe" (texto que nao citou nenhum parceiro cadastrado)
        # aparece por ultimo de proposito, mesmo que tenha mais TBRs que
        # os nomeados - eh o caso que precisa de atencao/correcao, faz
        # sentido ficar destacado no fim, nao competindo por ordem com os
        # nomes de verdade.
        itens_insucesso = sorted(
            r["contagem_insucesso"].items(),
            key=lambda kv: (kv[0] == "sem_detalhe", -kv[1]),
        )
        if itens_insucesso:
            linhas4 = "".join(
                _linha_painel(
                    "Insucesso (sem detalhe)" if nome == "sem_detalhe" else f"Insucesso {nome}",
                    qtd,
                )
                for nome, qtd in itens_insucesso
            )
        else:
            linhas4 = _linha_painel("Sem registros", "-")
        st.markdown(_painel("4. ANÁLISE INSUCESSO", linhas4), unsafe_allow_html=True)

        linhas5 = (
            _linha_painel("MNR", r["total_mnr"]) if r["total_mnr"]
            else _linha_painel("Sem registros", "-")
        )
        st.markdown(_painel("5. MNR's", linhas5), unsafe_allow_html=True)

        linhas_outro = (
            _linha_painel("Outro NODE", len(r["outro_node"])) if r["outro_node"]
            else _linha_painel("Sem registros", "-")
        )
        st.markdown(
            _painel("TBR's em outras estações", linhas_outro), unsafe_allow_html=True
        )

    if r["sem_finalizador"]:
        st.warning(
            f"⚠ {len(r['sem_finalizador'])} TBR(s) não-Delivered chegaram sem State "
            f"Finalizador preenchido: {', '.join(r['sem_finalizador'][:10])}"
            + (" ..." if len(r["sem_finalizador"]) > 10 else "")
        )

    st.info("Painel 6 (State Reversa) ainda não integrado — falta a base de Reversa entrar no fluxo web.")

    st.divider()
    if st.button(
        "📄 Gerar Fechamento",
        use_container_width=True,
        type="primary",
        help="Monta o dashboard final em cima desses mesmos números.",
    ):
        vai_para("fechamento_final")
        st.rerun()


# ------------------------------------------------------------------
# ETAPA 3 — GERAR FECHAMENTO (dashboard final)
# ------------------------------------------------------------------
def tela_fechamento_final():
    col_voltar, col_titulo = st.columns([1, 3], vertical_alignment="center")
    with col_voltar:
        if st.button("← Voltar à prévia"):
            vai_para("previa_fechamento")
            st.rerun()
    with col_titulo:
        st.markdown("**Etapa 3 — Fechamento**")

    r = st.session_state.dados_fechamento
    if r is None:
        st.warning("Não achei dados pra montar o fechamento. Volta e roda de novo.")
        return

    # Mesma correção da prévia: data do ARQUIVO de rotas, não a data de
    # hoje (o robô pode rodar num dia diferente do dia do arquivo).
    data_do_fechamento = st.session_state.data_arquivo_rotas or dt.date.today()
    html = gera_html_fechamento(
        r, NODE_ATUAL, st.session_state.nome_usuario, data_fechamento=data_do_fechamento
    )
    st.markdown(html, unsafe_allow_html=True)

    st.caption(
        "Visual ainda sem a marca d'água/logo do node e sem o painel de Reversa — "
        "faltam os arquivos de imagem e a integração da base de Reversa."
    )

    # Excel pra download - pedido do Samuel em 13/09/2026. Precisa dos
    # insumos guardados na Etapa 2 (ver _monta_insumos_export); se a
    # sessão foi reiniciada só com dados_fechamento sobrevivendo (não
    # deveria acontecer, os dois são salvos juntos, mas por garantia),
    # avisa em vez de quebrar a tela.
    insumos = st.session_state.export_fechamento_insumos
    st.divider()
    if insumos is None:
        st.info(
            "Excel não disponível pra essa sessão (dados de origem não encontrados) "
            "— volta e roda o State Finalizador de novo pra gerar o download."
        )
    else:
        xlsx_bytes = gera_workbook_fechamento(
            insumos["linhas_csv"], insumos["base_das"], insumos["motorista_real_da_rota"],
            insumos["base_tbr"], r, insumos["tbrs_ja_conhecidos"],
            st.session_state.nome_usuario, insumos["data_hoje_str"],
            NODE_ATUAL, data_do_fechamento,
        )
        # Nome pedido pelo Samuel em 13/09/2026: "LRN9 DD/MM" - troquei a
        # "/" por "-" porque barra não é permitida em nome de arquivo
        # (quebra no Windows e em outros sistemas).
        st.download_button(
            "⬇️ Baixar Excel (Análise do dia)",
            data=xlsx_bytes,
            file_name=f"{NODE_ATUAL} {data_do_fechamento.strftime('%d-%m')}.xlsx",
            use_container_width=True,
        )


# ------------------------------------------------------------------
# PARCEIROS DO NODE — cadastro de quem atua na base (a própria empresa +
# parceiros terceirizados). Ainda NÃO é usado em nenhum cálculo — é só a
# base de dados sendo montada agora; o dropdown do State Finalizador e o
# alerta de "insucesso sem parceiro especificado" entram numa próxima
# etapa, depois de confirmar que esse cadastro está funcionando.
# ------------------------------------------------------------------
def tela_parceiros():
    col_voltar, col_titulo = st.columns([1, 3], vertical_alignment="center")
    with col_voltar:
        if st.button("← Voltar"):
            vai_para("home_fechamento")
            st.rerun()
    with col_titulo:
        st.markdown(f"**Parceiros cadastrados — {NODE_ATUAL}**")

    st.caption(
        "Quem atua nessa base hoje: a própria empresa (o nome que vocês "
        "usam pra ela, ex: DMNZ) e os parceiros terceirizados (ex: MRIZ). "
        "Isso ainda não muda nada nos cálculos — é só o cadastro sendo "
        "criado; o resto (dropdown do State Finalizador, alerta de "
        "insucesso sem parceiro) vem depois."
    )

    if not secrets_configurados():
        st.warning(
            "Ainda não dá pra salvar isso — falta o secret `github` "
            "configurado no Streamlit Cloud."
        )
        return

    try:
        texto_parceiros, sha_parceiros = le_arquivo(CAMINHO_PARCEIROS)
        lista_parceiros = parceiros_mod.parseia_texto_parceiros(texto_parceiros)
    except Exception as e:
        st.error(f"Não consegui ler os parceiros no GitHub: {e}")
        return

    st.markdown("**Cadastrados atualmente**")
    if lista_parceiros:
        st.dataframe(
            [
                {
                    "Nome": p["nome"],
                    "Própria empresa?": "Sim" if p["propria_empresa"] else "Não",
                }
                for p in lista_parceiros
            ],
            use_container_width=True,
            hide_index=True,
        )
    else:
        st.info(
            f"Ainda não tem nenhum parceiro cadastrado pra {NODE_ATUAL}. "
            "Preenche abaixo pra começar — dá pra editar depois, isso "
            "nunca fica travado."
        )

    st.divider()
    st.markdown("**Adicionar parceiro**")
    nome_novo = st.text_input(
        "Nome (como deve aparecer no dropdown, ex: DMNZ, Dominalog, MRIZ...)",
        key="novo_parceiro_nome",
    )
    eh_propria = st.checkbox(
        "É a própria empresa (não é terceirizado)",
        key="novo_parceiro_propria",
    )

    if st.button("➕ Adicionar", key="add_parceiro", use_container_width=True):
        nome_norm_novo = parceiros_mod.normaliza(nome_novo)
        ja_marcado_propria = parceiros_mod.nome_propria_empresa(lista_parceiros)
        if not nome_novo.strip():
            st.warning("Digita um nome antes de adicionar.")
        elif any(nome_norm_novo == parceiros_mod.normaliza(p["nome"]) for p in lista_parceiros):
            st.warning(f"'{nome_novo}' já está cadastrado.")
        elif eh_propria and ja_marcado_propria:
            st.warning(
                f"Já tem **{ja_marcado_propria}** marcado como própria "
                "empresa — só pode ter um. Ainda não dá pra editar/remover "
                "marcação por aqui; me avisa se precisar trocar."
            )
        else:
            nova_lista = lista_parceiros + [
                {"nome": nome_novo.strip(), "propria_empresa": eh_propria}
            ]
            try:
                salva_arquivo(
                    CAMINHO_PARCEIROS,
                    parceiros_mod.gera_csv_parceiros(nova_lista),
                    sha_parceiros,
                    mensagem=f"Adiciona parceiro: {nome_novo.strip()}",
                )
            except ConflitoDeSalvamento:
                st.error("Alguém salvou ao mesmo tempo. Toca em Adicionar de novo.")
            except Exception as e:
                st.error(f"Não consegui salvar no GitHub: {e}")
            else:
                st.success(f"'{nome_novo.strip()}' adicionado!")
                st.rerun()


# ------------------------------------------------------------------
# PAINEL RECEBIMENTO - pacotes e paradas por rota, sem logica de
# motorista (ver consolida_recebimento.py pra regra de "mesma parada")
# ------------------------------------------------------------------
def _avisa_se_parece_csv_fechamento(por_rota, linhas_csv):
    """Mostra um st.warning se o CSV que acabou de subir num checkpoint
    de recebimento tiver um percentual alto de Delivered - sinal de que
    é, na verdade, um CSV de FIM DE TURNO (fechamento), não um
    checkpoint de verdade (ver percentual_delivered em
    consolida_recebimento.py). Só avisa, não trava nada."""
    pct_delivered, qtd_encontrada = percentual_delivered(por_rota, linhas_csv)
    if qtd_encontrada and pct_delivered >= LIMITE_ALERTA_PCT_DELIVERED:
        st.warning(
            f"⚠️ {pct_delivered:.0f}% dos pacotes do plano que aparecem nesse CSV já estão "
            "como **Delivered** — isso é sinal de que esse arquivo pode ser um CSV de "
            "**fechamento operacional (fim de turno)**, não um checkpoint de recebimento de "
            "verdade. No ato do recebimento ainda não dá pra ter pacote entregue (a entrega só "
            "acontece bem depois, já em rota). Confere se subiu o CSV certo antes de continuar."
        )
def tela_recebimento():
    col_voltar, col_titulo = st.columns([1, 3], vertical_alignment="center")
    with col_voltar:
        if st.button("← Voltar"):
            vai_para("home")
            st.rerun()
    with col_titulo:
        st.markdown("**Painel Recebimento**")

    st.caption(
        "Mesmo arquivo de rotas (planejado, dos indianos) — mas aqui a conta é "
        "pacotes e paradas por rota, sem entrar em motorista. Parada = mesmo "
        "endereço (rua + número); mais de uma entrega no mesmo prédio/condomínio, "
        "mudando só o apartamento, conta como 1 parada só."
    )

    ja_tem_dados = st.session_state.recebimento_resultado is not None
    if ja_tem_dados:
        arquivo_rotas = None
        if st.button("🔁 Trocar arquivo de rotas"):
            st.session_state.recebimento_resultado = None
            st.session_state.recebimento_confirmado = False
            st.session_state.recebimento_qtd_chegou = 0
            st.session_state.recebimento_rota_selecionada = None
            st.session_state.recebimento_ressalva = {"texto": "", "confirmada": False, "fotos": []}
            st.session_state.recebimento_tem_na = False
            st.session_state.recebimento_estado_bruto_stow = {}
            st.rerun()
    else:
        arquivo_rotas = st.file_uploader(
            "Arquivo de rotas (dos indianos)", type=["xlsx", "xlsm", "xls"], key="upload_recebimento"
        )

    if arquivo_rotas is not None:
        try:
            por_rota, data_arquivo, node = consolida_recebimento(arquivo_rotas, arquivo_rotas.name)
            st.session_state.recebimento_resultado = {
                "por_rota": por_rota,
                "data_arquivo": data_arquivo,
                "node": node,
            }
            st.session_state.recebimento_confirmado = False
        except ValueError as e:
            st.error(str(e))
            st.session_state.recebimento_resultado = None

    if not st.session_state.recebimento_resultado:
        return

    resultado = st.session_state.recebimento_resultado
    por_rota = resultado["por_rota"]
    data_arquivo = resultado["data_arquivo"]
    node = resultado["node"]

    data_texto = data_arquivo.strftime("%d/%m/%Y") if data_arquivo else "não identificada"
    node_texto = node or "não identificado"
    if node and node != NODE_ATUAL:
        st.warning(f"O arquivo parece ser do node **{node}**, não do {NODE_ATUAL}. Confere se é o arquivo certo.")

    # Removido de propósito (pedido do Samuel em 22/09/2026): tinha uma
    # tabela "Prévia" aqui (rota x pacotes x paradas) logo depois do
    # upload, mas isso já aparece nos cards do painel final por rota, lá
    # embaixo — era redundante e poluía a tela. Fica só o aviso de
    # node errado (não duplicado em lugar nenhum) e o expander de
    # conferência de agrupamento.
    st.divider()

    with st.expander("Conferir agrupamento (paradas com mais de 1 pacote)"):
        for rota in sorted(por_rota):
            detalhe = por_rota[rota]["detalhe_paradas"]
            if not detalhe:
                continue
            st.markdown(f"**{rota}**")
            st.dataframe(
                [{"Endereço": d["endereco_base"], "Qtd pacotes": d["qtd_pacotes"], "TBRs": ", ".join(d["tbrs"])}
                 for d in detalhe],
                use_container_width=True,
                hide_index=True,
            )

    # ------------------------------------------------------------
    # Lista unificada de TBRs (mesma ideia da Etapa 1 do Fechamento):
    # todas as rotas juntas num quadro só, com ícone de copiar, pra
    # colar direto na pesquisa do SCC - inclusive TBRs de NA, que não
    # vêm no arquivo de rotas (não foram atribuídos pela equipe da
    # noite, a Amazon manda a rota direto). Por enquanto essa lista é
    # só pra pesquisa/cópia - os TBRs de NA NÃO entram na conta dos
    # checkpoints abaixo (quantos chegaram x quantos induziram/stowed),
    # porque não têm rota nem "pacotes esperados" vindos do arquivo de
    # rotas pra comparar contra.
    # ------------------------------------------------------------
    st.divider()
    st.markdown("### Lista de TBRs (para pesquisar no SCC)")

    tbrs_todas_rotas = [tbr for info in por_rota.values() for tbr in info["tbrs"]]

    col_na, col_lista = st.columns([1, 1], gap="small")
    with col_na:
        with st.container(border=True):
            rotulo_botao_na = (
                "✅ Tem TBRs de NA (toque para desmarcar)"
                if st.session_state.recebimento_tem_na
                else "☐ Tem TBRs de NA para consulta no SCC? (opcional — toque para marcar)"
            )
            if st.button(rotulo_botao_na, key="botao_recebimento_tem_na", use_container_width=True):
                st.session_state.recebimento_tem_na = not st.session_state.recebimento_tem_na
                st.rerun()
            if st.session_state.recebimento_tem_na:
                st.text_area(
                    "Cole os TBRs de NA aqui, um por linha",
                    key="recebimento_texto_na",
                    height=80,
                    label_visibility="collapsed",
                )

    tbrs_na = (
        le_tbrs_colados(st.session_state.get("recebimento_texto_na", ""))
        if st.session_state.recebimento_tem_na
        else []
    )
    lista_final_tbrs = tbrs_todas_rotas + tbrs_na
    legenda_na = f" ({len(tbrs_todas_rotas)} + {len(tbrs_na)} de NA)" if tbrs_na else ""

    with col_lista:
        with st.container(border=True):
            st.markdown(f"**Lista de TBRs — {len(lista_final_tbrs)} TBR(s){legenda_na}**")
            st.caption("Ícone de copiar no canto do quadro pega a lista inteira de uma vez.")
            st.code("\n".join(lista_final_tbrs), language=None, height=150)

    # ------------------------------------------------------------
    # Checagem real, combinada (voltou a ser um total ÚNICO em
    # 24/09/2026, a pedido do Samuel - tinha uma versão anterior que
    # separava REC9/FOR3 automaticamente pela coluna "Source" do CSV do
    # SCC, mas o Samuel notou que "Source" não é um rótulo fixo de hub
    # de origem - é o local do ÚLTIMO SCAN, e vem "CUSTOMER_ADDRESS"
    # pra quem já foi entregue (visto de verdade: 6 TBR Delivered com
    # Source CUSTOMER_ADDRESS e 1 com CGH7 - nenhum de outro hub de
    # verdade, só o scan mais recente registrado em local diferente).
    # Num arquivo real com os dois hubs misturados não tem como saber a
    # qual dos dois um TBR com Source "sujo" pertence - separar por hub
    # em cima dessa coluna não é fiel à realidade. Sem escolha de hub,
    # sem botão - um número só, um checkpoint só.
    st.divider()
    st.markdown("### Checagem de recebimento (opcional)")
    st.caption(
        "A tabela acima é o PLANO — só vira confiável de verdade quando os pacotes que "
        "chegaram já passaram pelo fluxo físico (Manifested → Inducted → Stowed)."
    )

    st.session_state.recebimento_qtd_chegou = st.number_input(
        "Quantos pacotes chegaram? (contagem física, turno inteiro)",
        min_value=0,
        step=1,
        value=st.session_state.recebimento_qtd_chegou,
    )
    qtd_chegou = st.session_state.recebimento_qtd_chegou

    # Valores default (ajuste do Samuel em 24/09/2026: a ressalva saiu
    # daqui de dentro do Checkpoint 1 e virou uma única caixa "final",
    # depois do Checkpoint 2 - precisa desses dois sobreviverem pra lá
    # mesmo se o Checkpoint 1 ainda não tiver CSV subido).
    diferenca_inducao = None
    induzido = 0

    if qtd_chegou > 0:
        st.markdown("**Checkpoint 1 — pós-indução**")
        # SEM restrição de tipo (type=None) de propósito - mesmo motivo
        # da Etapa 1 do Fechamento (seletor do Android/Chrome às vezes
        # bloqueia o próprio CSV exportado do SCC por causa do "tipo"
        # que o navegador salvou no download) - e também aceita o CSV
        # salvo sem querer como .xlsx (visto de verdade em 23/09/2026 -
        # ver le_csv_scc em base_das.py, que detecta os dois formatos
        # pela assinatura do arquivo, não pela extensão).
        csv_inducao = st.file_uploader(
            "CSV (ou Excel) do SCC pós-indução (Tracking ID + State)", type=None, key="upload_csv_inducao"
        )
        if csv_inducao is not None:
            try:
                linhas_csv_inducao = le_csv_scc(csv_inducao)
            except Exception as e:
                st.error(f"Não consegui ler esse CSV: {e}")
            else:
                _avisa_se_parece_csv_fechamento(por_rota, linhas_csv_inducao)
                cruza_com_state_scc(por_rota, linhas_csv_inducao, chave="situacao_inducao")

        tem_situacao_inducao = any("situacao_inducao" in info for info in por_rota.values())
        if not tem_situacao_inducao:
            st.caption("Sobe o CSV acima pra ver o progresso de indução.")
        else:
            induzido = soma_inducidos_ou_alem(por_rota)
            manifested_pendentes = [
                tbr for info in por_rota.values()
                for tbr in info.get("situacao_inducao", {}).get("manifested", [])
            ]
            diferenca_inducao = _mostra_diferenca_checkpoint("induzido(s)", induzido, qtd_chegou)
            if diferenca_inducao != 0 and manifested_pendentes:
                with st.expander("Quem ainda está Manifested (falta induzir)"):
                    st.write(", ".join(manifested_pendentes))

        st.markdown("**Checkpoint 2 — pós-stow**")
        st.caption("Sempre disponível — não precisa esperar o Checkpoint 1 bater pra conferir o stow.")
        csv_stow = st.file_uploader(
            "CSV (ou Excel) do SCC pós-stow (Tracking ID + State)", type=None, key="upload_csv_stow"
        )
        if csv_stow is not None:
            try:
                linhas_csv_stow = le_csv_scc(csv_stow)
            except Exception as e:
                st.error(f"Não consegui ler esse CSV: {e}")
            else:
                _avisa_se_parece_csv_fechamento(por_rota, linhas_csv_stow)
                cruza_com_state_scc(por_rota, linhas_csv_stow, chave="situacao_stow")
                # Guarda o State BRUTO de cada TBR desse CSV (pedido do
                # Samuel em 24/09/2026) - cruza_com_state_scc só guarda
                # a CATEGORIA (manifested/inducted/...), não o texto
                # exato do SCC. Usado no aviso de possível MNR logo
                # abaixo, pra mostrar não só "esse TBR não avançou" mas
                # o state exato que o SCC mostra pra ele (ex: "Arrived"),
                # já que às vezes é útil ver se é o MESMO state de antes
                # (não mudou nada) ou um state diferente.
                st.session_state.recebimento_estado_bruto_stow = {
                    (l.get("Tracking ID") or "").strip(): (l.get("State") or "").strip()
                    for l in linhas_csv_stow
                }

        tem_situacao_stow = any("situacao_stow" in info for info in por_rota.values())
        if not tem_situacao_stow:
            st.caption("Sobe o CSV acima pra ver o progresso de stow.")
        else:
            stowed = soma_stowed_ou_alem(por_rota)
            inducted_pendentes = [
                tbr for info in por_rota.values()
                for tbr in info.get("situacao_stow", {}).get("inducted", [])
            ]
            manifested_pendentes_stow = [
                tbr for info in por_rota.values()
                for tbr in info.get("situacao_stow", {}).get("manifested", [])
            ]
            diferenca = _mostra_diferenca_checkpoint("armazenado(s)", stowed, qtd_chegou)
            if diferenca != 0 and inducted_pendentes:
                with st.expander("Quem ainda está Inducted (falta armazenar)"):
                    st.write(", ".join(inducted_pendentes))

            # Possível MNR (ajuste do Samuel em 23/09/2026): se um TBR
            # AINDA está em Manifested/Arrived/Received mesmo depois do
            # checkpoint de STOW já ter rodado, ele não chegou a ser
            # carregado de verdade (avaria, extravio ou falta de tempo
            # no hub de origem) - diferente de estar em Manifested só
            # no checkpoint de indução (isso aí é normal, só ainda não
            # chegou a vez). Só avisa, não trava - a ressalva final logo
            # abaixo já cobre seguir em frente mesmo com essa diferença.
            if manifested_pendentes_stow:
                st.warning(f"{len(manifested_pendentes_stow)} TBR(s) — {ROTULO_MNR_NAO_CARREGADO}")
                with st.expander("Possíveis MNR (TBRs)"):
                    estado_bruto_stow = st.session_state.get("recebimento_estado_bruto_stow") or {}
                    for tbr in manifested_pendentes_stow:
                        st.write(f"**{tbr}** — state atual no SCC: {estado_bruto_stow.get(tbr, '?')}")

        # Ressalva final (ajuste do Samuel em 24/09/2026): antes ficava
        # dentro do Checkpoint 1, sempre disponível - agora é uma caixa
        # SÓ, no final da checagem (depois do Checkpoint 2, "quando
        # entra tudo em stow"), e substitui a aba "Pendentes" do Excel:
        # em vez de só listar os TBR possível MNR num Excel à parte, o
        # próprio líder detalha a análise de cada um aqui - igual já
        # funciona na Etapa 1 do Fechamento (lista pra copiar com o
        # ícone do st.code) - e a lista já vem empilhada, um TBR por
        # linha, cada um com ":" no final, pronta pro líder completar.
        st.markdown("**Ressalva final**")
        texto_copia_mnr = monta_texto_copia_mnr(por_rota)
        if texto_copia_mnr:
            st.caption(
                "Possíveis MNR (o que não vai ser expedido pela manhã) — ícone de copiar "
                "no canto do quadro pega a lista inteira de uma vez:"
            )
            st.code(texto_copia_mnr, language=None)

        ressalva = st.session_state.recebimento_ressalva
        if ressalva["confirmada"]:
            qtd_fotos = len(ressalva["fotos"])
            st.info(
                "📋 Ressalva registrada"
                + (f" — {qtd_fotos} foto(s)/documento(s) anexado(s)." if qtd_fotos else ".")
            )
            with st.expander("Ver ressalva registrada"):
                st.text(ressalva["texto"])
                for nome, _ in ressalva["fotos"]:
                    st.caption(f"📎 {nome}")
            if st.button("✏️ Editar ressalva"):
                st.session_state.recebimento_ressalva["confirmada"] = False
                st.rerun()
        else:
            # Texto padrão junta duas coisas, se existirem: (1) o aviso
            # de "veio a menos que o Manifested" da indução (mesma regra
            # de antes: só pré-preenche se realmente faltou); (2) a
            # lista empilhada dos possíveis MNR do stow, pro líder
            # escrever a análise depois de cada ":".
            partes_texto_padrao = []
            if diferenca_inducao and diferenca_inducao > 0:
                partes_texto_padrao.append(monta_texto_ressalva_padrao(node_texto, qtd_chegou, induzido))
            if texto_copia_mnr:
                partes_texto_padrao.append(texto_copia_mnr)
            texto_padrao = "\n\n".join(partes_texto_padrao)

            with st.expander(
                "📋 Registrar ressalva / anexar fotos (opcional)", expanded=bool(texto_padrao)
            ):
                st.caption(
                    "Pra registrar a análise final do turno — número que não bateu com o "
                    "Manifested, detalhe de cada possível MNR, pacote avariado, ou qualquer "
                    "outra coisa que precise ficar documentada — com foto(s) se precisar. "
                    "Não trava nada, é só registro."
                )
                # Mesmo fix de key dinâmica de antes (bug real visto pelo
                # Samuel em 24/09/2026: key fixa faz o Streamlit ignorar
                # o value= nos reruns seguintes) - agora usando um hash
                # do texto padrão inteiro como key, já que ele pode mudar
                # por vários motivos diferentes (indução, stow, novo CSV)
                # e listar cada motivo à mão na key ia ficando frágil.
                texto_ressalva = st.text_area(
                    "Texto (observação, ressalva, o que for)",
                    value=ressalva["texto"] or texto_padrao,
                    height=220,
                    key=f"texto_area_ressalva_{hash(texto_padrao)}",
                )
                fotos_ressalva = st.file_uploader(
                    "Fotos/documentos (Manifested, avaria, etc.) — opcional, pode escolher várias",
                    type=["jpg", "jpeg", "png", "pdf"],
                    accept_multiple_files=True,
                    key="upload_fotos_ressalva",
                )
                if st.button("✅ Confirmar e seguir mesmo assim"):
                    st.session_state.recebimento_ressalva["texto"] = texto_ressalva
                    st.session_state.recebimento_ressalva["confirmada"] = True
                    st.session_state.recebimento_ressalva["fotos"] = [
                        (arq.name, arq.getvalue()) for arq in fotos_ressalva
                    ]
                    st.rerun()

    # ------------------------------------------------------------
    # Painel final por rota — cards clicáveis (pedido do Samuel em
    # 21/09/2026): clicar mostra os TBR + endereço daquela rota, pra
    # consultar direto da sala sem precisar abrir o arquivo de rotas
    # original. Os endereços já ficam guardados em por_rota[rota]["itens"]
    # desde a leitura do arquivo (ver consolida_recebimento.py).
    # ------------------------------------------------------------
    st.divider()
    st.markdown(f"### Painel por rota — {node_texto} · {data_texto}")
    tem_stow = any("situacao_stow" in info for info in por_rota.values())
    if tem_stow:
        rotas_confirmadas = sum(1 for info in por_rota.values() if info["situacao_stow"]["confirmada"])
        st.caption(
            f"{rotas_confirmadas} de {len(por_rota)} rota(s) confirmada(s) — 100% Stowed, "
            "considerando só quem chegou (quem nunca apareceu no CSV não trava a confirmação, "
            "só entra como possível MNR — detalhe na ressalva final, acima)."
        )

    rotas_ordenadas = sorted(por_rota)
    colunas = st.columns(3)
    for i, rota in enumerate(rotas_ordenadas):
        info = por_rota[rota]
        with colunas[i % 3]:
            with st.container(border=True):
                situacao_final = info.get("situacao_stow")
                if situacao_final:
                    st.markdown(f"**{'✅' if situacao_final['confirmada'] else '⏳'} {rota}**")
                else:
                    st.markdown(f"**{rota}**")
                st.markdown(f"{info['pacotes']} pacotes · {info['paradas']} paradas")
                if st.button("🔍 Ver TBRs / endereços", key=f"btn_ver_{rota}", use_container_width=True):
                    st.session_state.recebimento_rota_selecionada = (
                        None if st.session_state.recebimento_rota_selecionada == rota else rota
                    )
                    st.rerun()

    rota_sel = st.session_state.recebimento_rota_selecionada
    if rota_sel and rota_sel in por_rota:
        st.markdown(f"#### {rota_sel} — TBRs e endereços")
        st.dataframe(
            [
                {"Parada": item["stop"], "TBR": item["tbr"], "Endereço": item["endereco"]}
                for item in por_rota[rota_sel]["itens"]
            ],
            use_container_width=True,
            hide_index=True,
        )

    ressalvas = []
    ressalva = st.session_state.recebimento_ressalva
    if ressalva["confirmada"]:
        ressalvas.append({
            # Rotulada como "Pós-stow" (era "Pós-indução" antes do
            # ajuste do Samuel em 24/09/2026 que moveu a ressalva pro
            # final, depois do Checkpoint 2) - é o registro final do
            # turno, feito quando tudo já devia estar em Stow.
            "checkpoint": "Pós-stow",
            "texto": ressalva["texto"],
            "fotos": ressalva["fotos"],  # lista de (nome, bytes) - pode ter 0, 1 ou várias
        })

    st.divider()
    st.caption(
        "Confere a prévia acima antes de gerar o arquivo — o Excel final vai ter "
        "essa mesma tabela (aba Resumo, com a quantidade recebida registrada no topo) "
        "+ o detalhe do agrupamento (aba Detalhe paradas)"
        + (" + a ressalva final registrada (com foto(s), se tiver)." if ressalvas else ".")
    )
    if st.button("✅ Confirmar e gerar arquivo", type="primary"):
        st.session_state.recebimento_confirmado = True

    if st.session_state.recebimento_confirmado:
        excel_bytes = gera_workbook_recebimento(
            por_rota, node_texto, data_arquivo,
            ressalvas=ressalvas,
            qtd_chegou=st.session_state.recebimento_qtd_chegou,
        )
        nome_arquivo = f"painel_recebimento_{node_texto}_{data_arquivo.strftime('%Y%m%d') if data_arquivo else 'sem_data'}.xlsx"
        st.download_button(
            "⬇️ Baixar Excel",
            data=excel_bytes,
            file_name=nome_arquivo,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )


# ------------------------------------------------------------------
# LOGIN — pedido do Samuel em 22/09/2026: o link do app no Streamlit
# Cloud vai ficar aberto pra qualquer um (sem exigir conta
# Google/Streamlit, que é o que a opção "Only specific people" do
# próprio Streamlit Cloud exige), mas só quem souber usuário+senha
# consegue de fato ENTRAR e ver/mexer em alguma coisa - as credenciais
# NUNCA ficam no código (mesma regra do token do GitHub em
# github_store.py): moram nos secrets do Streamlit Cloud, seção
# [auth], configurada em "Manage app" -> "Settings" -> "Secrets":
#
#     [auth]
#     usuario = "escolha o que quiser aqui"
#     senha = "escolha uma senha forte aqui"
#
# Isso é so uma comparacao simples (nao tem "esqueci minha senha",
# nao tem varios usuarios, nao tem limite de tentativas) - de proposito,
# pra ficar simples de configurar e entender. Da pra evoluir depois se
# precisar de mais de uma pessoa com login proprio.
# ------------------------------------------------------------------
if "autenticado" not in st.session_state:
    st.session_state.autenticado = False

if not st.session_state.autenticado:
    st.title("📦 Painel Operação DMNZ - AMXL")
    st.caption("Acesso restrito — entra com usuário e senha.")
    usuario_digitado = st.text_input("Usuário")
    senha_digitada = st.text_input("Senha", type="password")
    if st.button("Entrar", type="primary"):
        try:
            usuario_certo = st.secrets["auth"]["usuario"]
            senha_certa = st.secrets["auth"]["senha"]
        except Exception:
            st.error(
                "Login ainda não configurado — falta o secret [auth] "
                "(usuario/senha) em Settings → Secrets no Streamlit Cloud."
            )
        else:
            if usuario_digitado == usuario_certo and senha_digitada == senha_certa:
                st.session_state.autenticado = True
                st.rerun()
            else:
                st.error("Usuário ou senha incorretos.")
    st.stop()


# ------------------------------------------------------------------
# ROTEADOR
# ------------------------------------------------------------------
if st.session_state.tela == "home":
    tela_home()
elif st.session_state.tela == "home_fechamento":
    tela_home_fechamento()
elif st.session_state.tela == "etapa1":
    tela_etapa1()
elif st.session_state.tela == "previa_fechamento":
    tela_previa_fechamento()
elif st.session_state.tela == "fechamento_final":
    tela_fechamento_final()
elif st.session_state.tela == "parceiros":
    tela_parceiros()
elif st.session_state.tela == "recebimento":
    tela_recebimento()
