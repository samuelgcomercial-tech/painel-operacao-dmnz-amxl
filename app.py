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
from state_finalizador import resumo_das_por_rota, lista_sem_da_de_verdade
from calculo_fechamento import calcula_tudo, monta_dados_do_dia

NODE_ATUAL = "LRN9"  # unico node desta primeira versao (decisao ja tomada)

# Onde a base de DAs e o histórico de TBRs moram no repositório do GitHub -
# mesmo padrão já decidido pros arquivos de memória entre dias (dados_nodes/<NODE>/).
CAMINHO_BASE_DAS = f"dados_nodes/{NODE_ATUAL}/base_das_dmnz.csv"
CAMINHO_HISTORICO_TBR = f"dados_nodes/{NODE_ATUAL}/historico_scc_analise.csv"

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
        arquivo_csv = st.file_uploader("CSV exportado do SCC", type=["csv"], key="csv_scc")

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
                                resultado = processa_state_finalizador(
                                    linhas_csv, base_tbr, base_dict, data_hoje_str
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
                                    f"**{len(resultado['auto_sem_state'])} sem State** "
                                    "(automático) · "
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

                                if not resultado["pendentes"]:
                                    st.success(
                                        "Nenhum TBR pendente de revisão manual hoje."
                                    )
                                    automaticos_sem_pendentes = (
                                        resultado["auto_em_rota"]
                                        + resultado["auto_sem_state"]
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
                                            st.session_state.dados_fechamento = calcula_tudo(dados_dia)
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
                                        st.session_state.dados_fechamento = calcula_tudo(dados_dia)
                                        vai_para("previa_fechamento")
                                        st.rerun()
                                else:
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
                                    # rota / sem state) junto com o que acabou de
                                    # ser digitado, com uma coluna "Origem" pra
                                    # diferenciar - não só o que foi digitado
                                    # agora, senão o usuário não teria o vislumbre
                                    # completo antes de confirmar.
                                    automaticos = (
                                        resultado["auto_em_rota"]
                                        + resultado["auto_sem_state"]
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
                                            for r in automaticos:
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
                                                st.session_state.dados_fechamento = calcula_tudo(dados_dia)
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
</style>
"""


def _linha_painel(label, valor, sub=False, total=False):
    classe = "linha" + (" sub" if sub else "") + (" total" if total else "")
    return f'<div class="{classe}"><span>{label}</span><span>{valor}</span></div>'


def _painel(titulo, linhas_html):
    return (
        f'<div class="painel-lrn9"><div class="cabecalho">{titulo}</div>'
        f'<div class="corpo">{linhas_html}</div></div>'
    )


def tela_previa_fechamento():
    col_voltar, col_titulo = st.columns([1, 3], vertical_alignment="center")
    with col_voltar:
        if st.button("← Voltar"):
            vai_para("home")
            st.rerun()
    with col_titulo:
        st.markdown(
            f"**Prévia do Fechamento**  ·  {st.session_state.nome_usuario} · "
            f"{NODE_ATUAL} · {dt.date.today().strftime('%d/%m/%Y')}"
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
            blocos.append(_linha_painel(cat, sum(subitens.values())))
            for fin, qtd in subitens.items():
                blocos.append(_linha_painel(fin or "(sem State Finalizador)", qtd, sub=True))
        blocos.append(_linha_painel("Total Geral", r["total_geral"], total=True))
        st.markdown(
            _painel("3. STATE SCC & ANÁLISE DMNZ", "".join(blocos)),
            unsafe_allow_html=True,
        )

    with col3:
        linhas4 = ""
        if r["n_insucesso_dmnz"]:
            linhas4 += _linha_painel("Insucesso DMNZ", r["n_insucesso_dmnz"])
        if r["n_insucesso_parceiro"]:
            linhas4 += _linha_painel("Insucesso Parceiro", r["n_insucesso_parceiro"])
        if r["n_insucesso_sem_detalhe"]:
            linhas4 += _linha_painel("Insucesso (sem detalhe)", r["n_insucesso_sem_detalhe"])
        if not linhas4:
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
    st.button(
        "📄 Gerar Fechamento",
        use_container_width=True,
        type="primary",
        disabled=True,
        help="Ainda não construído — é a próxima etapa (Etapa 3), monta o dashboard final em cima desses mesmos números.",
    )


# ------------------------------------------------------------------
# ROTEADOR
# ------------------------------------------------------------------
if st.session_state.tela == "home":
    tela_home()
elif st.session_state.tela == "etapa1":
    tela_etapa1()
elif st.session_state.tela == "previa_fechamento":
    tela_previa_fechamento()

