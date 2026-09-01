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

import streamlit as st

from consolida_rotas import consolida, detecta_node_do_nome, le_tbrs_colados

NODE_ATUAL = "LRN9"  # unico node desta primeira versao (decisao ja tomada)

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


def vai_para(tela):
    st.session_state.tela = tela


# ------------------------------------------------------------------
# TELA INICIAL
# ------------------------------------------------------------------
def tela_home():
    st.title("📦 Painel Operação DMNZ - AMXL")
    st.caption("Fechamento LRN9")

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


# ------------------------------------------------------------------
# ETAPA 1 — rotas dos indianos -> lista pro SCC (+ upload do CSV do SCC
# ja na mesma tela, pronto pra quando a etapa 2 existir)
# ------------------------------------------------------------------
def tela_etapa1():
    col_voltar, col_titulo = st.columns([1, 3], vertical_alignment="center")
    with col_voltar:
        if st.button("← Voltar"):
            vai_para("home")
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

        # "Arquivo lido" e "Tem NA" lado a lado, dois retângulos na
        # mesma linha. A simetria de altura é feita só por CSS
        # (min-height), NÃO pelo parâmetro height= do st.container - esse
        # parâmetro transforma a caixa numa área com scroll interno, e
        # isso quebra o toque no celular (o dedo sempre mexe um
        # pouquinho entre tocar e soltar, e o navegador interpreta esse
        # movimento dentro de uma área com scroll como "rolar a página"
        # em vez de "clicar", cancelando o clique do checkbox). Com
        # min-height a caixa cresce à vontade se precisar (nunca corta
        # nada) e o toque funciona normal.
        ALTURA_CAIXAS = 170
        st.markdown(
            f"""
            <style>
            .st-key-caixa_arquivo, .st-key-caixa_na {{
                min-height: {ALTURA_CAIXAS}px;
            }}
            </style>
            """,
            unsafe_allow_html=True,
        )
        col_info, col_na = st.columns([1, 1], gap="medium")
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

        # Se desmarcar de novo, não conta o que tinha digitado antes -
        # "não tem NA" precisa realmente zerar, mesmo que o texto ainda
        # esteja guardado por baixo dos panos.
        tbrs_na = le_tbrs_colados(st.session_state.get("texto_na", "")) if st.session_state.tem_na else []
        lista_final = [tbr for tbrs in tbrs_por_rota.values() for tbr in tbrs] + tbrs_na
        total_geral = len(lista_final)

        # Avisos de node/data errados continuam de largura total, bem
        # visíveis - não fica escondido dentro de um retângulo pequeno.
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

        # Estilo paisagem: rotas x pacotes numa coluna, lista de TBRs na
        # coluna do lado - em vez de empilhado. Num monitor (o uso real
        # no trabalho) fica lado a lado de verdade; no celular estreito
        # o Streamlit empilha essas colunas sozinho, então nesse caso
        # continua parecido com antes - é o navegador se ajustando à
        # largura da tela, não um erro.
        col_rotas, col_lista = st.columns([1, 1.3], gap="medium")

        with col_rotas:
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
            with st.container(border=True, height=320, vertical_alignment="center"):
                st.markdown(
                    f"<div style='text-align:center'>{linhas_rotas}</div>",
                    unsafe_allow_html=True,
                )

        with col_lista:
            legenda_na = f" ({total_rotas} + {len(tbrs_na)} de NA)" if tbrs_na else ""
            st.markdown(f"**Lista de TBRs — {total_geral} TBR(s){legenda_na}**")
            st.caption("Ícone de copiar no canto do quadro pega a lista inteira de uma vez.")
            # Caixa mais alta - fica mais parecida em altura com o
            # bloco de upload + rotas ao lado, formando um bloco mais
            # "quadrado" no geral, em vez de uma tira baixa e larga.
            st.code("\n".join(lista_final), language=None, height=320)

        st.write("")
        col_esq, col_meio, col_dir = st.columns([1, 2, 1])
        with col_meio:
            st.markdown("**Etapa 2 — CSV do SCC**")
            st.caption("Cole a lista no SCC, exporte e suba o CSV aqui.")
            arquivo_csv = st.file_uploader("CSV exportado do SCC", type=["csv"], key="csv_scc")
        if arquivo_csv is not None:
            st.info(
                "Arquivo recebido — o processamento da Etapa 2 (classificação "
                "automática DMNZ/Parceiro, base de histórico) ainda vai ser "
                "construído. Por enquanto só confirma que o upload funciona."
            )


# ------------------------------------------------------------------
# ROTEADOR
# ------------------------------------------------------------------
if st.session_state.tela == "home":
    tela_home()
elif st.session_state.tela == "etapa1":
    tela_etapa1()
