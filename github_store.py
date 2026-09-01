# -*- coding: utf-8 -*-
"""
github_store.py

Leitura/escrita de arquivos direto no repositório do GitHub, via API REST
(Contents API) - usado pra persistir arquivos que precisam sobreviver de
um dia pro outro (base de DAs, e depois histórico/reincidência/reversa),
já que o Streamlit Cloud não guarda nada em disco de uma "acordada" do
app pra outra.

Precisa de um token de acesso pessoal do GitHub (fine-grained, só com
permissão de leitura/escrita de CONTEÚDO nesse repositório específico),
guardado como secret no Streamlit Cloud (nunca no código, nunca no
repositório) - configuração em "Manage app" -> "Settings" -> "Secrets":

    [github]
    token = "github_pat_xxx..."
    repo = "usuario/painel-operacao-dmnz-amxl"

Se o arquivo ainda não existir no repositório, é tratado como "vazio"
(NÃO é erro) - a primeira escrita já CRIA o arquivo automaticamente, sem
precisar criar nada manualmente antes.
"""

import base64

import requests
import streamlit as st

API_BASE = "https://api.github.com"


class ConflitoDeSalvamento(Exception):
    """Alguém (ou outra aba/sessão) salvou esse mesmo arquivo entre a
    hora que a gente leu e a hora que tentou salvar. Quem chamar precisa
    reler o arquivo mais recente e tentar de novo - nunca sobrescrever
    às cegas."""


def secrets_configurados():
    """Confere se os secrets do GitHub existem, sem estourar um erro feio
    de StreamlitSecretNotFoundError/KeyError direto pra tela - se ainda
    não existir NENHUM secret configurado no app (caso comum antes da
    primeira configuração), o próprio st.secrets já levanta erro só de
    tentar olhar pra dentro dele, então precisa do try/except mesmo pra
    esse caso mais básico. Quem chamar deve checar isso ANTES de usar
    le_arquivo/salva_arquivo, e mostrar uma instrução clara se vier False
    - em vez de deixar o app quebrar ou fingir que a base tá vazia quando
    na verdade é só falta de configuração."""
    try:
        return "github" in st.secrets and "token" in st.secrets["github"] and "repo" in st.secrets["github"]
    except Exception:
        return False


def _headers():
    token = st.secrets["github"]["token"]
    return {
        "Authorization": f"Bearer {token}",
        "Accept": "application/vnd.github+json",
        "X-GitHub-Api-Version": "2022-11-28",
    }


def _repo():
    return st.secrets["github"]["repo"]


def le_arquivo(caminho):
    """Lê um arquivo do repositório. Devolve (texto, sha).

    Se o arquivo ainda não existir (primeira vez), devolve (None, None) -
    não é erro."""
    url = f"{API_BASE}/repos/{_repo()}/contents/{caminho}"
    resp = requests.get(url, headers=_headers(), timeout=15)
    if resp.status_code == 404:
        return None, None
    resp.raise_for_status()
    dados = resp.json()
    conteudo = base64.b64decode(dados["content"]).decode("utf-8")
    return conteudo, dados["sha"]


def salva_arquivo(caminho, texto, sha_anterior, mensagem):
    """Cria ou atualiza um arquivo no repositório.

    Se sha_anterior for None, CRIA o arquivo (primeira vez). Se vier
    preenchido, ATUALIZA - só funciona se for o sha mais recente de
    verdade (o GitHub recusa - HTTP 409/422 - se alguém mexeu no arquivo
    depois que a gente leu, pra nunca sobrescrever uma mudança por
    cima da outra sem perceber). Nesse caso levanta ConflitoDeSalvamento,
    pra quem chamou reler o arquivo atual e tentar de novo.

    Devolve o sha novo do arquivo (pra usar numa próxima atualização)."""
    url = f"{API_BASE}/repos/{_repo()}/contents/{caminho}"
    payload = {
        "message": mensagem,
        "content": base64.b64encode(texto.encode("utf-8")).decode("ascii"),
    }
    if sha_anterior:
        payload["sha"] = sha_anterior

    resp = requests.put(url, headers=_headers(), json=payload, timeout=15)
    if resp.status_code in (409, 422):
        raise ConflitoDeSalvamento(
            "Esse arquivo foi salvo por outra sessão entre a leitura e agora."
        )
    resp.raise_for_status()
    return resp.json()["content"]["sha"]
