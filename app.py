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

# ------------------------------------------------------------------
# ESTADO DA SESSAO
# ------------------------------------------------------------------
if "tela" not in st.session_state:
    st.session_state.tela = "home"
if "nome_usuario" not in st.session_state:
    st.session_state.nome_usuario = ""
if "tbrs_por_rota" not in st.session_state:
    st.session_state.tbrs_por_rota = None


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
        if st.button("▶️ Iniciar Fechamento", use_container_width=True, type="primary"):
            if not st.session_state.nome_usuario.strip():
                st.warning("Digita seu nome antes de continuar.")
            else:
                vai_para("etapa1")
                st.rerun()
        st.button("🔁 Reprocessar CSV", use_container_width=True, disabled=True,
                   help="Ainda não construído — em breve.")
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

    arquivo_rotas = st.file_uploader(
        "Arquivo de rotas (dos indianos)", type=["xlsx", "xlsm", "xls"]
    )

    if arquivo_rotas is not None:
        node_detectado = detecta_node_do_nome(arquivo_rotas.name)
        if node_detectado and node_detectado != NODE_ATUAL:
            st.warning(
                f"O nome do arquivo parece ser do node **{node_detectado}**, "
                f"mas essa versão do robô só trata **{NODE_ATUAL}**. Confira "
                "se é o arquivo certo antes de seguir."
            )

        try:
            tbrs_por_rota = consolida(arquivo_rotas, arquivo_rotas.name)
            st.session_state.tbrs_por_rota = tbrs_por_rota
        except ValueError as e:
            st.error(str(e))
            st.session_state.tbrs_por_rota = None

    if st.session_state.tbrs_por_rota:
        tbrs_por_rota = st.session_state.tbrs_por_rota
        total_rotas = sum(len(v) for v in tbrs_por_rota.values())

        tbrs_na = le_tbrs_colados(st.session_state.get("texto_na", ""))
        lista_final = [tbr for tbrs in tbrs_por_rota.values() for tbr in tbrs] + tbrs_na
        total_geral = len(lista_final)

        # Estilo paisagem: rotas x pacotes numa coluna, lista de TBRs na
        # coluna do lado - em vez de empilhado. Num monitor (o uso real
        # no trabalho) fica lado a lado de verdade; no celular estreito
        # o Streamlit empilha essas colunas sozinho, então nesse caso
        # continua parecido com antes - é o navegador se ajustando à
        # largura da tela, não um erro.
        col_rotas, col_lista = st.columns([1, 1.3], gap="medium")

        with col_rotas:
            st.markdown("**Rotas x pacotes**")
            for rota, v in tbrs_por_rota.items():
                st.write(f"{rota}: {len(v)}")
            st.caption(f"Total: {len(tbrs_por_rota)} rota(s), {total_rotas} TBR(s)")

            with st.expander("Tem NA para consulta no SCC? (opcional)"):
                st.text_area(
                    "Cole os TBRs de NA aqui, um por linha",
                    key="texto_na",
                    height=80,
                )

        with col_lista:
            legenda_na = f" ({total_rotas} + {len(tbrs_na)} de NA)" if tbrs_na else ""
            st.markdown(f"**Lista de TBRs — {total_geral} TBR(s){legenda_na}**")
            st.caption("Ícone de copiar no canto do quadro pega a lista inteira de uma vez.")
            st.code("\n".join(lista_final), language=None, height=260)

        st.markdown("**Etapa 2 — CSV do SCC:** cole a lista no SCC, exporte e suba o CSV aqui.")
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
