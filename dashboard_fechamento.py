# -*- coding: utf-8 -*-
"""
dashboard_fechamento.py (versao web)

Monta o HTML do dashboard final do Fechamento (Etapa 3) a partir do
resultado de calculo_fechamento.calcula_tudo() - separado do cálculo em
si (calculo_fechamento.py) porque aqui é só "como desenhar na tela", sem
nenhuma regra de negócio nova.

Cores tiradas por amostragem de pixel do template PNG atual (Dominalog),
não são um valor "oficial" declarado - se tiver o hex certo da marca,
troca aqui.

Ainda NÃO tem (fica pra depois, de propósito):
  - Marca d'água/selo circular do node, e a foto de fundo (ponte/por do
    sol) - não tenho esses arquivos de imagem, só o print de referência.
  - Painel 6 (State Reversa) - arquivo de Reversa ainda não integrado.
  - Logo real da Dominalog/Amazon - uso só texto no lugar.
"""

import datetime as dt

COR_LARANJA = "#FD4701"
COR_NAVY = "#000D2D"

MESES_PT = {
    1: "janeiro", 2: "fevereiro", 3: "março", 4: "abril", 5: "maio",
    6: "junho", 7: "julho", 8: "agosto", 9: "setembro", 10: "outubro",
    11: "novembro", 12: "dezembro",
}


def _data_por_extenso(data):
    return f"{data.day} de {MESES_PT[data.month]}"


CSS_DASHBOARD = f"""
<style>
.fech-wrap {{ font-family: Arial, Helvetica, sans-serif; }}
.fech-header {{ background: {COR_NAVY}; color: white; padding: 16px 24px;
                border-radius: 8px 8px 0 0; display:flex;
                justify-content:space-between; align-items:center; }}
.fech-header .titulo {{ font-size: 1.3em; font-weight: 700; }}
.fech-header .titulo b {{ color: {COR_LARANJA}; }}
.fech-header .marca {{ font-size: 0.85em; opacity: 0.85; letter-spacing: 1px; }}
.fech-grid {{ display:grid; grid-template-columns: repeat(3, 1fr); gap: 14px;
              padding: 16px; background: #f2f2f2; border-radius: 0 0 8px 8px; }}
.fech-painel {{ border: 1px solid #e2e2e2; border-radius: 8px; overflow: hidden;
               margin-bottom: 14px; background: white; }}
.fech-painel .cabecalho {{ background: {COR_LARANJA}; color: white; padding: 8px 14px;
               font-weight: 600; font-size: 0.95em; }}
.fech-painel .corpo {{ padding: 10px 14px 12px; color: #1a1a1a; }}
.fech-painel .linha {{ display:flex; justify-content:space-between; gap: 8px;
               padding: 3px 0; border-bottom: 1px dotted #ddd; font-size: 0.92em; }}
.fech-painel .linha.sub {{ padding-left: 14px; color:#555; }}
.fech-painel .linha.total {{ font-weight:700; border-top: 1px solid #999;
               border-bottom: none; margin-top:4px; padding-top:6px; }}
.fech-painel .linha.categoria {{ font-weight:600; border-bottom:none;
               padding-top: 8px; }}
.fech-painel .linha.categoria:first-child {{ padding-top: 0; }}
.fech-obs {{ margin: 0 16px 16px; border-radius: 8px; overflow:hidden;
             border: 1px solid #e2e2e2; }}
.fech-obs .cabecalho {{ background: {COR_NAVY}; color: white; padding: 8px 14px;
               font-weight: 600; }}
.fech-obs .corpo {{ background: white; padding: 12px 14px; font-size: 0.95em;
               color: #1a1a1a; }}
.fech-aviso {{ margin: 0 16px 16px; font-size: 0.85em; color: #7a4b00;
               background: #fff3c4; border-radius: 6px; padding: 8px 12px; }}
</style>
"""


def _linha(label, valor, sub=False, total=False):
    classe = "linha" + (" sub" if sub else "") + (" total" if total else "")
    return f'<div class="{classe}"><span>{label}</span><span>{valor}</span></div>'


def _categoria(label):
    return f'<div class="linha categoria"><span>{label}</span></div>'


def _painel(titulo, linhas_html):
    return (
        f'<div class="fech-painel"><div class="cabecalho">{titulo}</div>'
        f'<div class="corpo">{linhas_html}</div></div>'
    )


def gera_html_fechamento(r, node, nome_usuario, data_fechamento=None):
    """r: resultado de calculo_fechamento.calcula_tudo(). Devolve um bloco
    HTML pronto pra embutir com st.markdown(..., unsafe_allow_html=True)."""
    data_fechamento = data_fechamento or dt.date.today()

    col1 = [
        _painel(
            "1. STATE SCC",
            "".join(_linha(cat, qtd) for cat, qtd in r["contagem_scc"].most_common())
            + _linha("Total Geral", r["total_geral"], total=True),
        ),
        _painel("2. STATE ENTREGAS DMNZ", _linha("DMNZ", r["n_dmnz"])),
    ]

    blocos = []
    for cat, subitens in r["nested"].items():
        blocos.append(_categoria(cat))
        for fin, qtd in subitens.items():
            blocos.append(_linha(fin or "(sem State Finalizador)", qtd, sub=True))
    blocos.append(_linha("Total Geral", r["total_geral"], total=True))
    col2 = [_painel("3. STATE SCC & ANÁLISE DMNZ", "".join(blocos))]

    # "sem_detalhe" (texto que não citou nenhum parceiro cadastrado) por
    # último de propósito - é o caso que precisa de atenção/correção,
    # não compete por ordem com os nomes de verdade (DMNZ, MRIZ, ...).
    itens_insucesso = sorted(
        r["contagem_insucesso"].items(),
        key=lambda kv: (kv[0] == "sem_detalhe", -kv[1]),
    )
    if itens_insucesso:
        linhas4 = "".join(
            _linha(
                "Insucesso (sem detalhe)" if nome == "sem_detalhe" else f"Insucesso {nome}",
                qtd,
            )
            for nome, qtd in itens_insucesso
        )
    else:
        linhas4 = _linha("Sem registros", "-")

    linhas5 = _linha("MNR", r["total_mnr"]) if r["total_mnr"] else _linha("Sem registros", "-")
    linhas_outro = (
        _linha("Outro NODE", len(r["outro_node"])) if r["outro_node"]
        else _linha("Sem registros", "-")
    )
    col3 = [
        _painel("4. ANÁLISE INSUCESSO", linhas4),
        _painel("5. MNR's", linhas5),
        _painel("TBR's em outras estações", linhas_outro),
    ]

    aviso_html = ""
    if r["sem_finalizador"]:
        aviso_html = (
            f'<div class="fech-aviso">⚠ AUDITORIA: {len(r["sem_finalizador"])} TBR(s) '
            f'não-Delivered sem State Finalizador preenchido: '
            f'{", ".join(r["sem_finalizador"][:10])}'
            + (" ..." if len(r["sem_finalizador"]) > 10 else "") + "</div>"
        )

    html = f"""
    {CSS_DASHBOARD}
    <div class="fech-wrap">
      <div class="fech-header">
        <div class="titulo">Fechamento <b>{node}</b> – {_data_por_extenso(data_fechamento)}</div>
        <div class="marca">DOMINALOG</div>
      </div>
      <div class="fech-grid">
        <div>{"".join(col1)}</div>
        <div>{"".join(col2)}</div>
        <div>{"".join(col3)}</div>
      </div>
      {aviso_html}
      <div class="fech-obs">
        <div class="cabecalho">OBSERVAÇÕES</div>
        <div class="corpo">{r["observacoes"]}</div>
      </div>
    </div>
    """

    # st.markdown() roda o texto por um parser de Markdown antes de exibir
    # - e Markdown tem duas regras que atrapalham HTML grande feito esse:
    # (1) linha começando com 4+ espaços vira BLOCO DE CÓDIGO (mostra a
    # tag como texto cru em vez de renderizar - foi exatamente o que
    # aconteceu no teste real, a tela mostrou as tags <div> na tela); (2)
    # linha em branco separa um bloco de HTML do próximo, podendo cortar
    # a estrutura no meio. Em vez de tentar acertar a identação/linhas em
    # branco na mão (frágil - qualquer ajuste futuro pode reintroduzir o
    # mesmo bug), a saída inteira vira UMA linha só, sem quebra de linha
    # nenhuma - HTML e CSS não ligam pra isso (só <pre>/<code> ligam, que
    # a gente não usa aqui), então elimina de vez o risco.
    return " ".join(linha.strip() for linha in html.splitlines() if linha.strip())

