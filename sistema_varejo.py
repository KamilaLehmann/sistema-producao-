"""
Painel Executivo de Produção — Varejo
======================================
Requisitos: streamlit>=1.32, pandas, openpyxl, matplotlib

Tudo que já existia continua funcionando (leitura por cabeçalho, ausências,
movimentação, texto livre, horários salvos, imagem, Excel, e-mail em HTML,
adicionar manualmente). Novidades desta versão:

 17. NOVOS COLABORADORES SEM DIGITAR: nomes que aparecem na planilha e não
     estão na equipe viram uma lista na lateral ("🆕 Novos na planilha").
     Marque quem entra, escolha o cargo e clique em Cadastrar. Também há o
     botão Vincular (mesma pessoa com outro nome na planilha) e Ignorar
     (nomes que não são da equipe, para o aviso parar de aparecer).
 18. EQUIPE EDITÁVEL COMO PLANILHA (st.data_editor) no lugar do JSON: nome,
     cargo, nome na planilha, meta individual e horários padrão.
 19. PERSISTÊNCIA EM DISCO: equipe, pessoas adicionadas manualmente, metas e
     destinatários agora sobrevivem a recarregar a página/reiniciar o app.
 20. METAS DO SETOR editáveis na lateral (antes fixas no código).
 21. HISTÓRICO DIÁRIO (CSV) + gráficos de evolução na aba "Histórico".
 22. Bug corrigido: o campo de "Adicionar manualmente" agora usa st.form
     (clear_on_submit) e não gera mais StreamlitAPIException.
 23. SMTP: usuário/senha podem vir de st.secrets (SMTP_HOST, SMTP_PORT,
     SMTP_USUARIO, SMTP_SENHA). Destinatários ficam salvos.
 24. Tabela ordenada por cargo (Líder → Apoio → Operador) e depois por
     exemplares; cargos com selo colorido, avatar com iniciais e % de meta
     individual com cor de status.
 25. Novo layout: cabeçalho, 4 indicadores com barra de progresso e status,
     abas (Detalhamento / Imagem / Histórico / E-mail) e lateral organizada
     em seções recolhíveis.

Observação: em Streamlit Cloud o disco é temporário. Para persistência
definitiva em nuvem, guarde esses JSON/CSV em banco ou Google Sheets.
"""

import streamlit as st
import pandas as pd
import openpyxl
from datetime import datetime, time as dtime
from zoneinfo import ZoneInfo
import io
import os
import json
import html
import numbers
import textwrap
import unicodedata
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.image import MIMEImage
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches

# =============================================================================
# 0. CONSTANTES DE ARQUIVO
# =============================================================================
MOV_OVERRIDES_PATH = "movimentacoes_manuais.json"
HORARIOS_SALVOS_PATH = "horarios_salvos.json"
EQUIPE_CONFIG_PATH = "equipe_config.json"
PESSOAS_MANUAIS_PATH = "pessoas_manuais.json"
CONFIGURACOES_PATH = "configuracoes.json"
HISTORICO_PATH = "historico_diario.csv"


# =============================================================================
# 1. Configuração da página e estilo
# =============================================================================
st.set_page_config(page_title="Painel Executivo de Produção", page_icon="📊", layout="wide")
st.markdown("""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Sora:wght@500;600;700;800&family=Inter:wght@400;500;600;700&display=swap');

    html, body, [class*="css"] { font-family: 'Inter', sans-serif; }
    .stApp { background-color: #F4F6FA; }
    .block-container { padding-top: 1.4rem; padding-bottom: 2rem; max-width: 96%; }
    footer { visibility: hidden; }

    /* ---------- Lateral ---------- */
    section[data-testid="stSidebar"] {
        background-color: #FFFFFF;
        border-right: 1px solid #E5E7EB;
    }
    section[data-testid="stSidebar"] h2,
    section[data-testid="stSidebar"] h3 {
        font-family: 'Sora', sans-serif;
        color: #0F172A !important;
        letter-spacing: -0.2px;
    }
    section[data-testid="stSidebar"] h3 { font-size: 0.98rem !important; }
    section[data-testid="stSidebar"] details {
        border: 1px solid #E5E7EB !important;
        border-radius: 12px !important;
        background: #FBFCFE;
    }
    section[data-testid="stSidebar"] details summary p { font-weight: 600; color: #0F172A; }

    /* ---------- Cabeçalho ---------- */
    .hero {
        display: flex; justify-content: space-between; align-items: center; gap: 16px;
        background: #0B1B3A;
        background-image: linear-gradient(115deg, #0B1B3A 0%, #12306B 62%, #1D4ED8 130%);
        border-radius: 18px;
        padding: 22px 28px;
        margin-bottom: 22px;
        box-shadow: 0 10px 28px rgba(11, 27, 58, 0.22);
    }
    .hero-left { display: flex; align-items: center; gap: 16px; }
    .hero-logo { line-height: 0; }
    .hero-title {
        font-family: 'Sora', sans-serif; font-weight: 800; font-size: 1.55rem;
        letter-spacing: -0.5px; color: #FFFFFF; line-height: 1.15;
    }
    .hero-sub { color: #A9BBDD; font-size: 0.88rem; margin-top: 3px; }
    .hero-right { text-align: right; }
    .hero-date {
        font-family: 'Sora', sans-serif; font-weight: 700; font-size: 1.25rem; color: #FFFFFF;
    }

    /* ---------- Indicadores ---------- */
    .card-kpi {
        background: #FFFFFF;
        border: 1px solid #E5E7EB;
        border-top: 4px solid var(--accent-color, #2563EB);
        color: #0F172A;
        padding: 18px 20px 16px 20px;
        border-radius: 14px;
        box-shadow: 0 1px 2px rgba(15, 23, 42, 0.04), 0 6px 18px rgba(15, 23, 42, 0.04);
        margin-bottom: 14px;
        min-height: 168px;
    }
    .card-top { display: flex; justify-content: space-between; align-items: center; margin-bottom: 10px; }
    .card-icon {
        width: 34px; height: 34px; border-radius: 10px; display: flex; align-items: center;
        justify-content: center; font-size: 1.05rem;
        background: color-mix(in srgb, var(--accent-color, #2563EB) 12%, white);
    }
    .card-title { font-size: 0.86rem; font-weight: 600; color: #64748B; margin-bottom: 4px; }
    .card-value {
        font-family: 'Sora', sans-serif; font-size: 2.05rem; font-weight: 800;
        line-height: 1.05; color: #0F172A; margin-bottom: 6px; letter-spacing: -0.8px;
    }
    .card-sub { font-size: 0.82rem; font-weight: 500; color: #94A3B8; margin-bottom: 10px; }
    .bar { height: 7px; background: #E8ECF3; border-radius: 99px; overflow: hidden; }
    .bar-fill { height: 100%; border-radius: 99px; }
    .chip {
        font-size: 0.72rem; font-weight: 600; padding: 3px 10px; border-radius: 99px;
        white-space: nowrap;
    }
    .chip-ok   { background: #DCFCE7; color: #166534; }
    .chip-warn { background: #FEF3C7; color: #92400E; }
    .chip-bad  { background: #FEE2E2; color: #991B1B; }

    /* ---------- Abas ---------- */
    .stTabs [data-baseweb="tab-list"] { gap: 4px; border-bottom: 1px solid #E2E8F0; }
    .stTabs [data-baseweb="tab"] {
        height: 46px; padding: 0 18px; font-weight: 600; color: #64748B;
        border-radius: 10px 10px 0 0;
    }
    .stTabs [aria-selected="true"] { color: #0F172A; }
    .stTabs [data-baseweb="tab-highlight"] { background-color: #2563EB; height: 3px; }

    .secao-titulo {
        font-family: 'Sora', sans-serif; color: #0F172A; font-size: 1.05rem;
        font-weight: 700; margin: 6px 0 10px 0; letter-spacing: -0.2px;
    }

    hr { border-color: #E5E7EB !important; }

    /* ---------- Botões e campos ---------- */
    .stButton>button {
        background: #FFFFFF; border: 1px solid #CBD5E1; color: #0F172A;
        border-radius: 10px; font-weight: 600; transition: all 0.15s ease;
    }
    .stButton>button:hover { border-color: #2563EB; color: #2563EB; }
    div[data-testid="stDownloadButton"] button {
        background: #0F172A; border: 1px solid #0F172A; color: #FFFFFF; border-radius: 10px;
    }
    div[data-testid="stDownloadButton"] button:hover {
        background: #1E293B; border-color: #1E293B; color: #FFFFFF;
    }
    .stTextInput>div>div>input, .stDateInput input {
        background-color: #FFFFFF; color: #0F172A; border: 1px solid #E2E8F0; border-radius: 8px;
    }
    .stTextInput>div>div>input:focus { border-color: #2563EB; }
    div[data-testid="stDataFrame"] {
        border: 1px solid #E5E7EB; border-radius: 12px; overflow: hidden;
    }

    /* ---------- Tabela gerencial (HTML) ---------- */
    .tabela-wrapper {
        border: 1px solid #E2E8F0; border-radius: 16px; overflow: hidden;
        box-shadow: 0 1px 2px rgba(15, 23, 42, 0.04), 0 6px 18px rgba(15, 23, 42, 0.04);
        margin-bottom: 16px; background: #FFFFFF;
    }
    table.tabela-gerencial {
        width: 100%; border-collapse: collapse; font-size: 0.86rem; table-layout: fixed;
    }
    table.tabela-gerencial thead th {
        background: #0F172A; color: #FFFFFF; text-align: left; padding: 11px 14px;
        font-weight: 600; font-size: 0.8rem; letter-spacing: 0.1px;
    }
    table.tabela-gerencial tbody td {
        padding: 7px 14px; border-top: 1px solid #EEF2F6; color: #0F172A;
        vertical-align: middle; white-space: normal; overflow-wrap: break-word;
        line-height: 1.35; font-variant-numeric: tabular-nums;
    }
    table.tabela-gerencial td:nth-child(1),
    table.tabela-gerencial td:nth-child(2),
    table.tabela-gerencial th:nth-child(1),
    table.tabela-gerencial th:nth-child(2) { white-space: nowrap; }
    table.tabela-gerencial tbody tr:nth-child(even) { background: #FAFBFD; }
    table.tabela-gerencial tbody tr:hover { background: #EFF4FF; }
    .tabela-vazia {
        padding: 18px 16px; color: #64748B; font-size: 0.9rem;
        border: 1px solid #E2E8F0; border-radius: 16px; background: #FFFFFF;
    }
    .pill { font-size: 0.74rem; font-weight: 600; padding: 3px 10px; border-radius: 99px; }
    .pill-lider    { background: #DBEAFE; color: #1E40AF; }
    .pill-apoio    { background: #FEF3C7; color: #92400E; }
    .pill-operador { background: #E2E8F0; color: #334155; }
    .pill-outro    { background: #EDE9FE; color: #5B21B6; }
    .avatar {
        display: inline-flex; align-items: center; justify-content: center;
        width: 26px; height: 26px; border-radius: 50%; margin-right: 9px;
        background: #E0E7FF; color: #3730A3; font-size: 0.7rem; font-weight: 700;
        vertical-align: middle;
    }
    .nome-colab { font-weight: 600; vertical-align: middle; }
    .txt-ausente { color: #B91C1C; font-weight: 600; }

    /* ---------- Estado vazio ---------- */
    .vazio {
        background: #FFFFFF; border: 1px dashed #CBD5E1; border-radius: 18px;
        padding: 42px 30px; text-align: center; color: #475569;
    }
    .vazio h3 { font-family: 'Sora', sans-serif; color: #0F172A; margin: 0 0 8px 0; font-size: 1.2rem; }
    .vazio p { margin: 0; font-size: 0.95rem; }
    </style>
""", unsafe_allow_html=True)

LOGO_HERO_SVG = """<svg width="44" height="44" viewBox="0 0 34 34" xmlns="http://www.w3.org/2000/svg"><rect width="34" height="34" rx="9" fill="#FFFFFF"/><path d="M9 24V10h7.5c3.6 0 5.8 1.8 5.8 4.9 0 2.1-1.1 3.6-3 4.3l3.4 4.8h-3.4l-3-4.3H12V24H9zm3-6.7h4.2c1.8 0 2.8-.8 2.8-2.3s-1-2.3-2.8-2.3H12v4.6z" fill="#0B1B3A"/></svg>"""


# =============================================================================
# 2. Utilitários gerais
# =============================================================================
def normalizar(texto):
    """Remove acentos, espaços extras e padroniza para maiúsculas."""
    texto = str(texto).strip().upper()
    texto = unicodedata.normalize("NFKD", texto).encode("ASCII", "ignore").decode("ASCII")
    texto = " ".join(texto.split())
    return texto


def texto_seguro(valor):
    if valor is None:
        return ""
    try:
        if pd.isna(valor):
            return ""
    except (TypeError, ValueError):
        pass
    return str(valor)


def ler_json(caminho, padrao):
    if not os.path.exists(caminho):
        return padrao
    try:
        with open(caminho, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return padrao


def gravar_json(caminho, dados):
    try:
        with open(caminho, "w", encoding="utf-8") as f:
            json.dump(dados, f, ensure_ascii=False, indent=2)
        return True
    except Exception:
        return False


def segredo(chave, padrao=""):
    """Lê de st.secrets sem quebrar caso o arquivo secrets.toml não exista."""
    try:
        return st.secrets.get(chave, padrao)
    except Exception:
        return padrao


# =============================================================================
# 3. Equipe (padrão + persistência)
# =============================================================================
DEFAULT_EQUIPE = {
    "Líder": [
        {"nome": "Kamila Moraes", "alias_excel": "KAMILA"},
        {"nome": "Beatriz Alcantara", "alias_excel": "BEATRIZ"},
    ],
    "Apoio": [
        {"nome": "Alisson Lima"},
    ],
    "Operador(a)": [
        {"nome": "Rosana Delfino"},
        {"nome": "Ana Caroline", "alias_excel": "ANACAROLINE"},
        {"nome": "Karoline Gonçalves"},
        {"nome": "Gabriele"},
        {"nome": "Beatriz Mascarenhas"},
        {"nome": "Vinicius Silva", "alias_excel": "Vinicius Silva"},
        {"nome": "Paula Roberta", "alias_excel": "PAULA ROBERTA SANTOS DA SILVA"},
        {"nome": "Weliton"},
        {"nome": "Ellen Kelly"},
        {"nome": "Eloizem", "alias_excel": "Eloize Meire"},
    ],
}
CARGOS_BASE = ["Líder", "Apoio", "Operador(a)"]


def carregar_equipe_disco():
    dados = ler_json(EQUIPE_CONFIG_PATH, None)
    return dados if isinstance(dados, dict) and dados else None


def salvar_equipe_disco(config):
    return gravar_json(EQUIPE_CONFIG_PATH, config)


def salvar_pessoas_manuais():
    gravar_json(PESSOAS_MANUAIS_PATH, st.session_state["pessoas_manuais"])


def salvar_configuracoes():
    gravar_json(CONFIGURACOES_PATH, st.session_state["configuracoes"])


CONFIG_PADRAO = {"meta_exemplares": 55000, "meta_skus": 1200, "destinatarios": "", "ignorados": []}

if "equipe_config" not in st.session_state:
    st.session_state["equipe_config"] = carregar_equipe_disco() or json.loads(json.dumps(DEFAULT_EQUIPE))

if "pessoas_manuais" not in st.session_state:
    manuais_disco = ler_json(PESSOAS_MANUAIS_PATH, [])
    st.session_state["pessoas_manuais"] = manuais_disco if isinstance(manuais_disco, list) else []

if "configuracoes" not in st.session_state:
    cfg_inicial = dict(CONFIG_PADRAO)
    cfg_disco = ler_json(CONFIGURACOES_PATH, {})
    if isinstance(cfg_disco, dict):
        cfg_inicial.update(cfg_disco)
    st.session_state["configuracoes"] = cfg_inicial


def construir_estruturas_equipe(config):
    equipe = {}
    nomes_lista = []
    alias_excel = {}
    metas_individuais = {}
    defaults_mov = {}
    cargo_por_nome = {}

    for cargo, integrantes in config.items():
        nomes_cargo = []
        for pessoa in integrantes:
            nome = pessoa.get("nome", "").strip()
            if not nome:
                continue
            nomes_cargo.append(nome)
            nomes_lista.append(nome)
            cargo_por_nome[nome] = cargo
            if pessoa.get("alias_excel"):
                alias_excel[nome] = pessoa["alias_excel"]
            if pessoa.get("meta_exemplares"):
                try:
                    metas_individuais[nome] = int(pessoa["meta_exemplares"])
                except (TypeError, ValueError):
                    pass
            defaults_mov[nome] = {
                "saida": pessoa.get("saida_padrao", ""),
                "retorno": pessoa.get("retorno_padrao", ""),
                "local": pessoa.get("local_padrao", ""),
            }
        equipe[cargo] = nomes_cargo

    return equipe, nomes_lista, alias_excel, metas_individuais, defaults_mov, cargo_por_nome


def nome_excel(nome, alias_map):
    return normalizar(alias_map.get(nome, nome))


COLUNAS_EDITOR_EQUIPE = [
    "Cargo", "Nome", "Alias na planilha", "Meta individual",
    "Saída padrão", "Retorno padrão", "Local padrão",
]


def config_para_df(config):
    linhas = []
    for cargo, pessoas in config.items():
        for p in pessoas:
            linhas.append({
                "Cargo": cargo,
                "Nome": p.get("nome", ""),
                "Alias na planilha": p.get("alias_excel", ""),
                "Meta individual": p.get("meta_exemplares") or None,
                "Saída padrão": p.get("saida_padrao", ""),
                "Retorno padrão": p.get("retorno_padrao", ""),
                "Local padrão": p.get("local_padrao", ""),
            })
    df = pd.DataFrame(linhas, columns=COLUNAS_EDITOR_EQUIPE)
    df["Meta individual"] = pd.to_numeric(df["Meta individual"], errors="coerce")
    return df


def df_para_config(df):
    nova = {}
    for _, l in df.iterrows():
        nome = texto_seguro(l.get("Nome")).strip()
        if not nome:
            continue
        cargo = texto_seguro(l.get("Cargo")).strip() or "Operador(a)"
        pessoa = {"nome": nome}
        alias = texto_seguro(l.get("Alias na planilha")).strip()
        if alias:
            pessoa["alias_excel"] = alias
        meta = l.get("Meta individual")
        try:
            if meta is not None and not pd.isna(meta) and float(meta) > 0:
                pessoa["meta_exemplares"] = int(meta)
        except (TypeError, ValueError):
            pass
        for chave, coluna in (("saida_padrao", "Saída padrão"),
                              ("retorno_padrao", "Retorno padrão"),
                              ("local_padrao", "Local padrão")):
            valor = texto_seguro(l.get(coluna)).strip()
            if valor:
                pessoa[chave] = valor
        nova.setdefault(cargo, []).append(pessoa)
    return nova


# =============================================================================
# 4. Horários, overrides de texto e histórico
# =============================================================================
def parse_hora_str(valor_str):
    if not valor_str:
        return None
    try:
        h, m = valor_str.split(":")
        return dtime(int(h), int(m))
    except Exception:
        return None


def formatar_hora_editor(valor):
    if valor is None:
        return ""
    try:
        if pd.isna(valor):
            return ""
    except (TypeError, ValueError):
        pass
    if isinstance(valor, (dtime, datetime)):
        return valor.strftime("%Hh%M")
    if isinstance(valor, str):
        texto = valor.strip()
        if not texto:
            return ""
        for fmt in ("%H:%M:%S.%f", "%H:%M:%S", "%H:%M"):
            try:
                return datetime.strptime(texto, fmt).strftime("%Hh%M")
            except ValueError:
                continue
        try:
            return pd.to_datetime(texto).strftime("%Hh%M")
        except Exception:
            return texto
    if hasattr(valor, "strftime"):
        try:
            return valor.strftime("%Hh%M")
        except Exception:
            return ""
    return str(valor).strip()


def hora_para_iso(valor):
    if valor is None:
        return ""
    try:
        if pd.isna(valor):
            return ""
    except (TypeError, ValueError):
        pass
    if isinstance(valor, (dtime, datetime)):
        return valor.strftime("%H:%M")
    if isinstance(valor, str):
        texto = valor.strip()
        if not texto:
            return ""
        for fmt in ("%H:%M:%S.%f", "%H:%M:%S", "%H:%M"):
            try:
                return datetime.strptime(texto, fmt).strftime("%H:%M")
            except ValueError:
                continue
        try:
            return pd.to_datetime(texto).strftime("%H:%M")
        except Exception:
            return ""
    if hasattr(valor, "strftime"):
        try:
            return valor.strftime("%H:%M")
        except Exception:
            return ""
    return ""


def carregar_overrides_disco():
    if not os.path.exists(MOV_OVERRIDES_PATH):
        return {}
    try:
        with open(MOV_OVERRIDES_PATH, "r", encoding="utf-8") as f:
            bruto = json.load(f)
        return {tuple(chave.split("||", 1)): texto for chave, texto in bruto.items()}
    except Exception:
        return {}


def salvar_overrides_disco(overrides):
    try:
        bruto = {f"{data}||{nome}": texto for (data, nome), texto in overrides.items()}
        with open(MOV_OVERRIDES_PATH, "w", encoding="utf-8") as f:
            json.dump(bruto, f, ensure_ascii=False, indent=2)
    except Exception:
        pass


def carregar_horarios_disco():
    dados = ler_json(HORARIOS_SALVOS_PATH, {})
    return dados if isinstance(dados, dict) else {}


def salvar_horarios_disco(horarios):
    gravar_json(HORARIOS_SALVOS_PATH, horarios)


COLUNAS_HISTORICO = ["Data", "Exemplares", "SKUs", "Colaboradores"]


def carregar_historico():
    if not os.path.exists(HISTORICO_PATH):
        return pd.DataFrame(columns=COLUNAS_HISTORICO)
    try:
        df = pd.read_csv(HISTORICO_PATH)
        for c in COLUNAS_HISTORICO:
            if c not in df.columns:
                df[c] = 0
        df["Data"] = df["Data"].astype(str)
        return df[COLUNAS_HISTORICO]
    except Exception:
        return pd.DataFrame(columns=COLUNAS_HISTORICO)


def salvar_historico_dia(data_str, exemplares, skus, colaboradores):
    df = carregar_historico()
    df = df[df["Data"] != data_str]
    novo = pd.DataFrame([{
        "Data": data_str, "Exemplares": int(exemplares),
        "SKUs": int(skus), "Colaboradores": int(colaboradores),
    }])
    df = novo if df.empty else pd.concat([df, novo], ignore_index=True)
    df = df.sort_values("Data").reset_index(drop=True)
    try:
        df.to_csv(HISTORICO_PATH, index=False)
        return True
    except Exception:
        return False


if "mov_manual_overrides" not in st.session_state:
    st.session_state["mov_manual_overrides"] = carregar_overrides_disco()

if "horarios_salvos" not in st.session_state:
    st.session_state["horarios_salvos"] = carregar_horarios_disco()


# =============================================================================
# 5. Leitura da planilha — por CABEÇALHO de coluna, com cache
# =============================================================================
def localizar_colunas(sheet):
    mapeamento = {}
    for col in range(1, sheet.max_column + 1):
        valor = sheet.cell(row=1, column=col).value
        if valor is None:
            continue
        v = normalizar(valor)
        if v == "TOTAL" and "TOTAL" not in mapeamento:
            mapeamento["TOTAL"] = col
        if v == "USUARIO" and "USUARIO" not in mapeamento:
            mapeamento["USUARIO"] = col
    return mapeamento


@st.cache_data(show_spinner="Lendo planilha...")
def ler_planilha(bytes_arquivo):
    wb = openpyxl.load_workbook(io.BytesIO(bytes_arquivo), data_only=True)
    sheet = wb.active

    mapeamento = localizar_colunas(sheet)
    usando_fallback = ("TOTAL" not in mapeamento) or ("USUARIO" not in mapeamento)
    col_total = mapeamento.get("TOTAL", 9)
    col_usuario = mapeamento.get("USUARIO", 13)

    dados = []
    for row in range(2, sheet.max_row + 1):
        if sheet.row_dimensions[row].hidden:
            continue
        val_total = sheet.cell(row=row, column=col_total).value
        val_usuario = sheet.cell(row=row, column=col_usuario).value
        if val_total is not None and val_usuario is not None:
            dados.append({"TOTAL": val_total, "USUARIO": normalizar(val_usuario)})

    df = pd.DataFrame(dados, columns=["TOTAL", "USUARIO"])
    if not df.empty:
        df["TOTAL"] = pd.to_numeric(df["TOTAL"], errors="coerce").fillna(0)

    return df, usando_fallback


# =============================================================================
# 6. BARRA LATERAL
# =============================================================================
cfg = st.session_state["configuracoes"]

st.sidebar.header("🛠️ Controle Operacional")

# ---- Dados -----------------------------------------------------------------
uploaded_file = st.sidebar.file_uploader("Planilha de produção (.xlsx)", type=["xlsx"], key="uploaded_file")
data_produtividade = st.sidebar.date_input(
    "Data da produtividade:", datetime.now(ZoneInfo("America/Sao_Paulo")), key="data_produtividade"
)
data_formatada = data_produtividade.strftime("%d/%m")

# ---- Metas do setor --------------------------------------------------------
with st.sidebar.expander("🎯 Metas do setor"):
    nova_meta_ex = st.number_input(
        "Meta diária de exemplares:", min_value=1, value=int(cfg["meta_exemplares"]), step=500, key="meta_ex_input"
    )
    nova_meta_sku = st.number_input(
        "Meta diária de SKUs:", min_value=1, value=int(cfg["meta_skus"]), step=50, key="meta_sku_input"
    )
    if int(nova_meta_ex) != int(cfg["meta_exemplares"]) or int(nova_meta_sku) != int(cfg["meta_skus"]):
        cfg["meta_exemplares"] = int(nova_meta_ex)
        cfg["meta_skus"] = int(nova_meta_sku)
        salvar_configuracoes()

META_EXEMPLARES = int(cfg["meta_exemplares"])
META_SKUS = int(cfg["meta_skus"])

# ---- Equipe (editor em tabela) --------------------------------------------
with st.sidebar.expander("👥 Equipe (editar como planilha)"):
    st.caption(
        "Adicione linhas no fim da tabela para cadastrar alguém, edite células ou "
        "selecione uma linha e apague. 'Alias na planilha' é o nome exatamente como "
        "aparece na coluna USUARIO (só preencha se for diferente do nome). "
        "Clique em Aplicar para salvar."
    )
    cargos_disponiveis = list(dict.fromkeys(CARGOS_BASE + list(st.session_state["equipe_config"].keys())))
    editor_equipe = st.data_editor(
        config_para_df(st.session_state["equipe_config"]),
        num_rows="dynamic",
        hide_index=True,
        use_container_width=True,
        key="editor_equipe",
        column_config={
            "Cargo": st.column_config.SelectboxColumn("Cargo", options=cargos_disponiveis, required=True),
            "Nome": st.column_config.TextColumn("Nome", required=True),
            "Alias na planilha": st.column_config.TextColumn("Alias na planilha"),
            "Meta individual": st.column_config.NumberColumn("Meta individual", min_value=0, step=100),
            "Saída padrão": st.column_config.TextColumn("Saída padrão", help="Formato HH:MM"),
            "Retorno padrão": st.column_config.TextColumn("Retorno padrão", help="Formato HH:MM"),
            "Local padrão": st.column_config.TextColumn("Local padrão"),
        },
    )
    col_aplicar, col_restaurar = st.columns(2)
    with col_aplicar:
        if st.button("✅ Aplicar", use_container_width=True, key="btn_aplicar_equipe"):
            nova_config = df_para_config(editor_equipe)
            nomes_novos = [p["nome"] for pessoas in nova_config.values() for p in pessoas]
            duplicados = sorted({n for n in nomes_novos if nomes_novos.count(n) > 1})
            if not nova_config:
                st.error("A equipe não pode ficar vazia.")
            elif duplicados:
                st.error("Nomes repetidos: " + ", ".join(duplicados))
            else:
                st.session_state["equipe_config"] = nova_config
                salvar_equipe_disco(nova_config)
                st.session_state.pop("editor_equipe", None)
                st.rerun()
    with col_restaurar:
        if st.button("↩️ Padrão", use_container_width=True, key="btn_restaurar_equipe"):
            st.session_state["equipe_config"] = json.loads(json.dumps(DEFAULT_EQUIPE))
            salvar_equipe_disco(st.session_state["equipe_config"])
            st.session_state.pop("editor_equipe", None)
            st.rerun()

EQUIPE, NOMES_LISTA, ALIAS_EXCEL, METAS_INDIVIDUAIS, DEFAULTS_MOV, CARGO_POR_NOME = construir_estruturas_equipe(
    st.session_state["equipe_config"]
)
CARGOS_ORDEM = list(EQUIPE.keys())

# ---- Novos colaboradores detectados na planilha ---------------------------
if uploaded_file:
    df_previa, _ = ler_planilha(uploaded_file.getvalue())
    if not df_previa.empty:
        validos = {nome_excel(n, ALIAS_EXCEL) for n in NOMES_LISTA}
        ignorados = set(cfg.get("ignorados", []))
        skus_por_usuario = df_previa.groupby("USUARIO").size().to_dict()
        novos = sorted(set(skus_por_usuario) - validos - ignorados)

        if novos:
            with st.sidebar.expander(f"🆕 Novos na planilha ({len(novos)})", expanded=True):
                st.caption("Nomes que aparecem na planilha, mas não estão na equipe.")
                escolhidos = st.multiselect(
                    "Quem entra na equipe?",
                    novos,
                    format_func=lambda u: f"{u.title()} ({skus_por_usuario[u]} SKUs)",
                    key="novos_escolhidos",
                )
                cargo_novo = st.selectbox(
                    "Cargo:", list(dict.fromkeys(CARGOS_BASE + CARGOS_ORDEM)), key="cargo_novos"
                )
                col_cad, col_ign = st.columns(2)
                with col_cad:
                    if st.button("✅ Cadastrar", use_container_width=True, key="btn_cadastrar_novos"):
                        if not escolhidos:
                            st.warning("Marque ao menos um nome.")
                        else:
                            cfg_equipe = st.session_state["equipe_config"]
                            for u in escolhidos:
                                cfg_equipe.setdefault(cargo_novo, []).append(
                                    {"nome": u.title(), "alias_excel": u}
                                )
                            salvar_equipe_disco(cfg_equipe)
                            st.session_state.pop("novos_escolhidos", None)
                            st.session_state.pop("editor_equipe", None)
                            st.rerun()
                with col_ign:
                    if st.button("🙈 Ignorar", use_container_width=True, key="btn_ignorar_novos",
                                 help="Não são da equipe: o aviso deixa de aparecer para estes nomes."):
                        if not escolhidos:
                            st.warning("Marque ao menos um nome.")
                        else:
                            cfg["ignorados"] = sorted(set(cfg.get("ignorados", [])) | set(escolhidos))
                            salvar_configuracoes()
                            st.session_state.pop("novos_escolhidos", None)
                            st.rerun()

                st.markdown("---")
                st.caption("É a mesma pessoa com outro nome na planilha?")
                u_alias = st.selectbox("Nome na planilha:", ["—"] + novos, key="u_alias")
                p_alias = st.selectbox("Corresponde a:", ["—"] + NOMES_LISTA, key="p_alias")
                if st.button("🔗 Vincular", use_container_width=True, key="btn_vincular"):
                    if u_alias == "—" or p_alias == "—":
                        st.warning("Escolha o nome da planilha e a pessoa.")
                    else:
                        for pessoas in st.session_state["equipe_config"].values():
                            for p in pessoas:
                                if p["nome"] == p_alias:
                                    p["alias_excel"] = u_alias
                        salvar_equipe_disco(st.session_state["equipe_config"])
                        st.session_state.pop("editor_equipe", None)
                        st.rerun()

if cfg.get("ignorados"):
    with st.sidebar.expander(f"🙈 Nomes ignorados ({len(cfg['ignorados'])})"):
        st.caption(", ".join(n.title() for n in cfg["ignorados"]))
        if st.button("Voltar a avisar sobre todos", use_container_width=True, key="btn_limpar_ignorados"):
            cfg["ignorados"] = []
            salvar_configuracoes()
            st.rerun()

st.sidebar.markdown("<hr style='margin:14px 0px; border-color: #E5E7EB;'>", unsafe_allow_html=True)

# ---- Filtros ---------------------------------------------------------------
st.sidebar.markdown("### 👁️ Filtros gerenciais")
remover_do_setor = st.sidebar.multiselect("Ocultar do setor (tabela):", NOMES_LISTA, key="remover_do_setor")

# ---- Adicionar manualmente (corrigido: st.form + clear_on_submit) --------
with st.sidebar.expander("➕ Adicionar manualmente ao relatório"):
    st.caption(
        "Inclui uma pessoa na tabela do Detalhamento Gerencial mesmo que ela não "
        "esteja na equipe ou não tenha registro na planilha no dia."
    )
    with st.form("form_adicionar_manual", clear_on_submit=True):
        novo_nome_manual = st.text_input("Nome da pessoa:")
        novo_cargo_manual = st.selectbox("Cargo:", list(EQUIPE.keys()) + ["Outro"])
        enviado_manual = st.form_submit_button("➕ Adicionar à tabela", use_container_width=True)

    if enviado_manual:
        nome_limpo = novo_nome_manual.strip()
        if not nome_limpo:
            st.warning("Digite um nome antes de adicionar.")
        elif nome_limpo in NOMES_LISTA or any(p["nome"] == nome_limpo for p in st.session_state["pessoas_manuais"]):
            st.warning("Esse nome já está na equipe ou já foi adicionado.")
        else:
            st.session_state["pessoas_manuais"].append({"nome": nome_limpo, "cargo": novo_cargo_manual})
            salvar_pessoas_manuais()
            st.rerun()

    if st.session_state["pessoas_manuais"]:
        st.caption("Adicionados manualmente:")
        for i, pessoa in enumerate(st.session_state["pessoas_manuais"]):
            col_nome_add, col_remover_add = st.columns([3, 1])
            col_nome_add.markdown(f"👤 {pessoa['nome']} ({pessoa['cargo']})")
            if col_remover_add.button("🗑️", key=f"remover_manual_{i}", use_container_width=True):
                st.session_state["pessoas_manuais"].pop(i)
                salvar_pessoas_manuais()
                st.rerun()

st.sidebar.markdown("<hr style='margin:14px 0px; border-color: #E5E7EB;'>", unsafe_allow_html=True)

# ---- Ausências -------------------------------------------------------------
st.sidebar.markdown("### ❌ Ausências do dia")
faltas_selecionadas = st.sidebar.multiselect("Quem faltou hoje:", NOMES_LISTA, key="faltas_selecionadas")

st.sidebar.markdown("<hr style='margin:14px 0px; border-color: #E5E7EB;'>", unsafe_allow_html=True)

# ---- Movimentação ----------------------------------------------------------
st.sidebar.markdown("### ⏳ Movimentação de horários")
movimentados_selecionados = st.sidebar.multiselect(
    "🚚 Quem foi movimentado(a) hoje?",
    [n for n in NOMES_LISTA if n not in remover_do_setor and n not in faltas_selecionadas],
    key="movimentados_selecionados",
)

st.sidebar.markdown("<hr style='margin:14px 0px; border-color: #E5E7EB;'>", unsafe_allow_html=True)

MOTIVOS_FALTA_PADRAO = ["Falta administrativa", "Atestado médico", "Falta injustificada", "Folga compensatória", "Outro"]

dict_movimentacao = {}
dict_motivos_falta = {}

for cargo, integrantes in EQUIPE.items():
    integrantes_visiveis = [i for i in integrantes if i not in remover_do_setor]
    if integrantes_visiveis:
        st.sidebar.markdown(
            f"<h3 style='color:#1E3A8A; margin-top:10px; font-size:1.05rem;'>🔹 {html.escape(cargo)}</h3>",
            unsafe_allow_html=True,
        )

    for op in integrantes:
        if op in remover_do_setor:
            continue

        is_ausente = op in faltas_selecionadas
        is_movimentado = op in movimentados_selecionados

        if is_ausente:
            st.sidebar.markdown(f"❌ **{op} (ausente)**")
            motivo_escolhido = st.sidebar.selectbox(
                f"Motivo da falta de {op}:", MOTIVOS_FALTA_PADRAO, key=f"mot_falta_sel_{op}"
            )
            if motivo_escolhido == "Outro":
                motivo_escolhido = st.sidebar.text_input(
                    f"Descreva o motivo de {op}:", value="", key=f"mot_falta_txt_{op}"
                )
            dict_motivos_falta[op] = motivo_escolhido or "Falta administrativa"
            dict_movimentacao[op] = {"cargo": cargo, "movimentacoes": []}

        elif is_movimentado:
            st.sidebar.markdown(f"**👤 {op}**", unsafe_allow_html=True)

            salvo = st.session_state["horarios_salvos"].get(op)
            if salvo:
                linha_inicial = pd.DataFrame(
                    [
                        {
                            "Saída": parse_hora_str(linha.get("saida", "")),
                            "Retorno": parse_hora_str(linha.get("retorno", "")),
                            "Local": linha.get("local", ""),
                        }
                        for linha in salvo
                    ] or [{"Saída": None, "Retorno": None, "Local": ""}]
                )
            else:
                defaults = DEFAULTS_MOV.get(op, {})
                linha_inicial = pd.DataFrame(
                    [
                        {
                            "Saída": parse_hora_str(defaults.get("saida", "")),
                            "Retorno": parse_hora_str(defaults.get("retorno", "")),
                            "Local": defaults.get("local", ""),
                        }
                    ]
                )

            editado = st.sidebar.data_editor(
                linha_inicial,
                num_rows="dynamic",
                hide_index=True,
                use_container_width=True,
                key=f"mov_editor_{op}",
                column_config={
                    "Saída": st.column_config.TimeColumn("Saída", format="HH:mm", step=60),
                    "Retorno": st.column_config.TimeColumn("Retorno", format="HH:mm", step=60),
                    "Local": st.column_config.TextColumn("Local"),
                },
            )

            movimentacoes_op = []
            for _, linha in editado.iterrows():
                sai = linha.get("Saída")
                ret = linha.get("Retorno")
                loc = texto_seguro(linha.get("Local"))
                sai_txt = formatar_hora_editor(sai)
                ret_txt = formatar_hora_editor(ret)
                if sai_txt or ret_txt or loc.strip():
                    movimentacoes_op.append({"sai": sai_txt, "ret": ret_txt, "loc": loc})

            dict_movimentacao[op] = {"cargo": cargo, "movimentacoes": movimentacoes_op}

            col_salvar_hora, col_resetar_hora = st.sidebar.columns(2)
            with col_salvar_hora:
                if st.button("💾 Salvar", key=f"salvar_horario_{op}", use_container_width=True):
                    linhas_para_salvar = []
                    for _, linha_ed in editado.iterrows():
                        sai_iso = hora_para_iso(linha_ed.get("Saída"))
                        ret_iso = hora_para_iso(linha_ed.get("Retorno"))
                        loc_ed = texto_seguro(linha_ed.get("Local"))
                        if sai_iso or ret_iso or loc_ed.strip():
                            linhas_para_salvar.append({"saida": sai_iso, "retorno": ret_iso, "local": loc_ed})
                    st.session_state["horarios_salvos"][op] = linhas_para_salvar
                    salvar_horarios_disco(st.session_state["horarios_salvos"])
                    st.sidebar.success(f"Horário de {op} salvo.")
            with col_resetar_hora:
                if st.button("🔄 Resetar", key=f"resetar_horario_{op}", use_container_width=True):
                    st.session_state["horarios_salvos"].pop(op, None)
                    salvar_horarios_disco(st.session_state["horarios_salvos"])
                    st.session_state.pop(f"mov_editor_{op}", None)
                    st.rerun()

        else:
            st.sidebar.markdown(
                f"👤 {op} <span style='font-size:0.8rem; color:gray;'>(sem movimentação)</span>",
                unsafe_allow_html=True,
            )
            dict_movimentacao[op] = {"cargo": cargo, "movimentacoes": []}

        st.sidebar.markdown("<hr style='margin:6px 0px; border-color: #E5E7EB;'>", unsafe_allow_html=True)


# ---- E-mail (SMTP) ---------------------------------------------------------
with st.sidebar.expander("✉️ Configuração de e-mail (SMTP)"):
    st.caption(
        "Para mais segurança, defina SMTP_HOST, SMTP_PORT, SMTP_USUARIO e SMTP_SENHA "
        "em `.streamlit/secrets.toml`. Quando existirem, os campos correspondentes "
        "somem daqui."
    )
    smtp_host = st.text_input("Servidor SMTP:", value=str(segredo("SMTP_HOST", "smtp.gmail.com")), key="smtp_host")
    smtp_port = st.number_input("Porta:", value=int(segredo("SMTP_PORT", 587)), step=1, key="smtp_port")

    usuario_secret = segredo("SMTP_USUARIO", "")
    if usuario_secret:
        smtp_usuario = str(usuario_secret)
        st.caption(f"Remetente (via secrets): {smtp_usuario}")
    else:
        smtp_usuario = st.text_input("E-mail remetente:", key="smtp_usuario")

    senha_secret = segredo("SMTP_SENHA", "")
    if senha_secret:
        smtp_senha = str(senha_secret)
        st.caption("Senha carregada de st.secrets.")
    else:
        smtp_senha = st.text_input("Senha / senha de app:", type="password", key="smtp_senha")

    destinatarios_texto = st.text_input(
        "Destinatários (separados por vírgula):", value=cfg.get("destinatarios", ""), key="smtp_destinatarios"
    )
    if destinatarios_texto != cfg.get("destinatarios", ""):
        cfg["destinatarios"] = destinatarios_texto
        salvar_configuracoes()


def enviar_email_relatorio(assunto, corpo_texto, corpo_html, imagem_bytes, nome_imagem, cid_imagem):
    """Envia o relatório em HTML com a imagem do painel embutida no corpo
    (Content-ID), mais uma versão em texto simples como alternativa."""
    destinatarios = [d.strip() for d in destinatarios_texto.split(",") if d.strip()]
    if not (smtp_host and smtp_usuario and smtp_senha and destinatarios):
        st.error("Preencha servidor, remetente, senha e ao menos um destinatário na configuração de e-mail.")
        return
    try:
        msg = MIMEMultipart("related")
        msg["From"] = smtp_usuario
        msg["To"] = ", ".join(destinatarios)
        msg["Subject"] = assunto

        alternativo = MIMEMultipart("alternative")
        alternativo.attach(MIMEText(corpo_texto, "plain", "utf-8"))
        alternativo.attach(MIMEText(corpo_html, "html", "utf-8"))
        msg.attach(alternativo)

        img_part = MIMEImage(imagem_bytes, name=nome_imagem)
        img_part.add_header("Content-ID", f"<{cid_imagem}>")
        img_part.add_header("Content-Disposition", "inline", filename=nome_imagem)
        msg.attach(img_part)

        with smtplib.SMTP(smtp_host, int(smtp_port)) as servidor:
            servidor.starttls()
            servidor.login(smtp_usuario, smtp_senha)
            servidor.sendmail(smtp_usuario, destinatarios, msg.as_string())

        st.success(f"E-mail enviado com sucesso para: {', '.join(destinatarios)}")
    except Exception as e:
        st.error(f"Falha ao enviar e-mail: {e}")


# =============================================================================
# 7. Imagem do relatório, tabela HTML e Excel
# =============================================================================
def gerar_relatorio_imagem(total_exemplares, total_skus, pct_exemplares, pct_skus,
                            meta_exemplares, meta_skus, df_real, data_formatada=""):
    colunas_relatorio = ["Cargo", "Colaboradora", "Movimentação Operacional"]
    df_relatorio = df_real[colunas_relatorio].copy() if not df_real.empty else df_real

    LARGURA_QUEBRA = 78
    linhas_por_registro = []
    if not df_relatorio.empty:
        textos_quebrados = []
        for texto in df_relatorio["Movimentação Operacional"]:
            texto = "" if texto is None else str(texto)
            linhas_texto = textwrap.wrap(texto, width=LARGURA_QUEBRA) or [""]
            textos_quebrados.append("\n".join(linhas_texto))
            linhas_por_registro.append(len(linhas_texto))
        df_relatorio["Movimentação Operacional"] = textos_quebrados

    ALTURA_HEADER_IN = 0.62
    ESPACO_HEADER_CARDS_IN = 0.22
    ALTURA_CARDS_IN = 1.15
    ESPACO_CARDS_TABELA_IN = 0.22
    ALTURA_CABECALHO_TABELA_IN = 0.34
    ALTURA_LINHA_TABELA_IN = 0.24
    ESPACO_TABELA_RODAPE_IN = 0.16
    ALTURA_RODAPE_IN = 0.26
    MARGEM_INFERIOR_IN = 0.08

    total_linhas_texto = sum(max(n, 1) for n in linhas_por_registro) if linhas_por_registro else 1
    altura_tabela_in = ALTURA_CABECALHO_TABELA_IN + ALTURA_LINHA_TABELA_IN * total_linhas_texto

    altura_fig = (
        ALTURA_HEADER_IN + ESPACO_HEADER_CARDS_IN + ALTURA_CARDS_IN + ESPACO_CARDS_TABELA_IN
        + altura_tabela_in + ESPACO_TABELA_RODAPE_IN + ALTURA_RODAPE_IN + MARGEM_INFERIOR_IN
    )

    COR_NAVY = "#0F172A"
    COR_AZUL = "#2563EB"
    COR_TEAL = "#0D9488"
    COR_TEXTO = "#111827"
    COR_MUTED = "#64748B"

    fig = plt.figure(figsize=(11, altura_fig), dpi=200)
    fig.patch.set_facecolor("#F8FAFC")

    frac_header_altura = ALTURA_HEADER_IN / altura_fig
    ax_header = fig.add_axes([0, 1 - frac_header_altura, 1, frac_header_altura])
    ax_header.set_xlim(0, 1); ax_header.set_ylim(0, 1); ax_header.axis("off")
    ax_header.add_patch(mpatches.Rectangle((0, 0), 1, 1, facecolor="#FFFFFF", edgecolor="none"))
    ax_header.plot([0, 1], [0.02, 0.02], color="#E5E7EB", linewidth=1, transform=ax_header.transAxes)

    icon_x, icon_w = 0.028, 0.032
    ax_header.add_patch(mpatches.FancyBboxPatch(
        (icon_x, 0.28), icon_w, 0.44, boxstyle="round,pad=0,rounding_size=0.012",
        linewidth=0, facecolor=COR_NAVY, transform=ax_header.transAxes
    ))
    barra_larg = icon_w / 5.6
    for i, alt in enumerate([0.14, 0.22, 0.30]):
        ax_header.add_patch(mpatches.Rectangle(
            (icon_x + 0.006 + i * (barra_larg + 0.004), 0.36), barra_larg, alt,
            facecolor="#93C5FD", edgecolor="none", transform=ax_header.transAxes
        ))

    ax_header.text(icon_x + icon_w + 0.018, 0.66, "Painel Executivo de Produção",
                    fontsize=15.5, fontweight="bold", color=COR_NAVY, va="center")
    ax_header.text(icon_x + icon_w + 0.018, 0.30, "Varejo · acompanhamento diário de produtividade",
                    fontsize=8.5, color=COR_MUTED, va="center")

    horario_brasil = datetime.now(ZoneInfo("America/Sao_Paulo"))
    if data_formatada:
        ax_header.text(0.972, 0.66, f"Referente a {data_formatada}", fontsize=9.5,
                        fontweight="bold", color=COR_NAVY, va="center", ha="right")
        ax_header.text(0.972, 0.30, f"Gerado em {horario_brasil.strftime('%d/%m/%Y %H:%M')}",
                        fontsize=7.5, color=COR_MUTED, va="center", ha="right")

    y_cards_topo_in = altura_fig - ALTURA_HEADER_IN - ESPACO_HEADER_CARDS_IN
    y_cards_base_in = y_cards_topo_in - ALTURA_CARDS_IN
    frac_cards_base = y_cards_base_in / altura_fig
    frac_cards_altura = ALTURA_CARDS_IN / altura_fig

    def desenhar_card(x, largura, titulo, valor, sub, pct, cor_accent):
        ax = fig.add_axes([x, frac_cards_base, largura, frac_cards_altura])
        ax.set_xlim(0, 1); ax.set_ylim(0, 1); ax.axis("off")

        ax.add_patch(mpatches.FancyBboxPatch(
            (0.015, 0.03), 0.98, 0.92, boxstyle="round,pad=0,rounding_size=0.09",
            linewidth=0, facecolor="#E2E8F0", alpha=0.6, transform=ax.transAxes
        ))
        card = mpatches.FancyBboxPatch(
            (0.01, 0.06), 0.98, 0.92, boxstyle="round,pad=0,rounding_size=0.09",
            linewidth=1, edgecolor="#E5E7EB", facecolor="white", transform=ax.transAxes
        )
        ax.add_patch(card)
        barra = mpatches.FancyBboxPatch(
            (0.01, 0.06), 0.014, 0.92, boxstyle="round,pad=0,rounding_size=0.007",
            linewidth=0, facecolor=cor_accent, transform=ax.transAxes
        )
        ax.add_patch(barra)

        ax.text(0.09, 0.80, titulo, fontsize=8.5, fontweight="bold", color="#94A3B8", va="top")
        ax.text(0.09, 0.60, valor, fontsize=22, fontweight="bold", color=COR_TEXTO, va="top")
        ax.text(0.09, 0.35, sub, fontsize=7.8, color=COR_MUTED, va="top")

        largura_barra = 0.82
        ax.add_patch(mpatches.FancyBboxPatch(
            (0.09, 0.16), largura_barra, 0.055, boxstyle="round,pad=0,rounding_size=0.03",
            linewidth=0, facecolor="#E5E7EB", transform=ax.transAxes
        ))
        preenchido = max(min(pct, 1.0), 0.0) * largura_barra
        if preenchido > 0.02:
            ax.add_patch(mpatches.FancyBboxPatch(
                (0.09, 0.16), preenchido, 0.055, boxstyle="round,pad=0,rounding_size=0.03",
                linewidth=0, facecolor=cor_accent, transform=ax.transAxes
            ))

    desenhar_card(0.04, 0.44, "TOTAL DE EXEMPLARES", f"{total_exemplares:,} un",
                  f"Meta Diária: {meta_exemplares:,} un  ·  Atingido: {pct_exemplares:.1%}",
                  pct_exemplares, COR_AZUL)
    desenhar_card(0.52, 0.44, "TOTAL DE SKU", f"{total_skus:,}",
                  f"Meta Diária: {meta_skus:,}  ·  Atingido: {pct_skus:.1%}",
                  pct_skus, COR_TEAL)

    y_tabela_topo_in = y_cards_base_in - ESPACO_CARDS_TABELA_IN
    y_tabela_base_in = y_tabela_topo_in - altura_tabela_in
    frac_tabela_base = y_tabela_base_in / altura_fig
    frac_tabela_altura = altura_tabela_in / altura_fig

    ax_moldura = fig.add_axes([0.04, frac_tabela_base, 0.92, frac_tabela_altura])
    ax_moldura.set_xlim(0, 1); ax_moldura.set_ylim(0, 1); ax_moldura.axis("off")
    ax_moldura.add_patch(mpatches.FancyBboxPatch(
        (0, 0), 1, 1, boxstyle="round,pad=0,rounding_size=0.025",
        linewidth=1.1, edgecolor="#E2E8F0", facecolor="white", transform=ax_moldura.transAxes
    ))

    ax = fig.add_axes([0.04, frac_tabela_base, 0.92, frac_tabela_altura])
    ax.axis("off")

    if not df_relatorio.empty:
        tabela = ax.table(
            cellText=df_relatorio.values,
            colLabels=df_relatorio.columns,
            cellLoc="left",
            loc="upper left",
            colWidths=[0.14, 0.22, 0.64],
        )
        tabela.auto_set_font_size(False)
        tabela.set_fontsize(8.5)

        frac_por_linha_texto = ALTURA_LINHA_TABELA_IN / altura_tabela_in
        frac_cabecalho = ALTURA_CABECALHO_TABELA_IN / altura_tabela_in

        for (row, col), cell in tabela.get_celld().items():
            cell.set_edgecolor("#EEF2F6")
            cell.PAD = 0.025
            cell.get_text().set_verticalalignment("center")
            if row == 0:
                cell.set_facecolor(COR_NAVY)
                cell.set_text_props(color="white", fontweight="bold")
                cell.set_height(frac_cabecalho)
            else:
                cell.set_facecolor("#FFFFFF" if row % 2 == 0 else "#F8FAFC")
                cell.set_height(frac_por_linha_texto * max(linhas_por_registro[row - 1], 1))
    else:
        ax.text(0.02, 0.9, "Nenhum dado disponível.", fontsize=9, color=COR_MUTED)

    frac_rodape_altura = ALTURA_RODAPE_IN / altura_fig
    ax_rodape = fig.add_axes([0.04, 0, 0.92, frac_rodape_altura])
    ax_rodape.set_xlim(0, 1); ax_rodape.set_ylim(0, 1); ax_rodape.axis("off")
    ax_rodape.plot([0, 1], [0.92, 0.92], color="#E2E8F0", linewidth=1, transform=ax_rodape.transAxes)
    ax_rodape.text(0, 0.35, "Painel Executivo de Produção · Varejo", fontsize=7.5,
                    color="#94A3B8", va="center", ha="left")
    ax_rodape.text(1, 0.35, f"{total_skus} SKU · {total_exemplares:,} exemplares",
                    fontsize=7.5, color="#94A3B8", va="center", ha="right")

    buffer = io.BytesIO()
    fig.savefig(buffer, format="png", facecolor=fig.get_facecolor())
    plt.close(fig)
    buffer.seek(0)
    return buffer.getvalue()


def classe_cargo(cargo):
    n = normalizar(cargo)
    if n.startswith("LIDER"):
        return "lider"
    if n.startswith("APOIO"):
        return "apoio"
    if n.startswith("OPERADOR"):
        return "operador"
    return "outro"


def iniciais(nome):
    partes = [p for p in str(nome).split() if p]
    if not partes:
        return "?"
    return (partes[0][0] + (partes[1][0] if len(partes) > 1 else "")).upper()


def formatar_inteiro_br(valor):
    return f"{int(valor):,}".replace(",", ".")


def renderizar_tabela_html(df):
    if df.empty:
        return "<div class='tabela-vazia'>Nenhum dado disponível.</div>"

    larguras = {
        "Cargo": "13%",
        "Colaboradora": "19%",
        "Exemplares": "10%",
        "SKUs": "8%",
        "Meta Individual": "11%",
        "% Meta Individual": "11%",
    }

    colunas = list(df.columns)
    colgroup = "".join(
        f'<col style="width:{larguras[c]}">' if c in larguras else "<col>"
        for c in colunas
    )
    cabecalho = "".join(f"<th>{html.escape(str(c))}</th>" for c in colunas)

    linhas_html = []
    for _, linha in df.iterrows():
        celulas = []
        for c in colunas:
            valor = linha[c]
            texto_cru = texto_seguro(valor)

            if c == "Cargo":
                celula = f"<span class='pill pill-{classe_cargo(texto_cru)}'>{html.escape(texto_cru)}</span>"
            elif c == "Colaboradora":
                celula = (f"<span class='avatar'>{html.escape(iniciais(texto_cru))}</span>"
                          f"<span class='nome-colab'>{html.escape(texto_cru)}</span>")
            elif c == "% Meta Individual" and texto_cru.endswith("%"):
                try:
                    pct_num = int(texto_cru.rstrip("%"))
                except ValueError:
                    pct_num = None
                if pct_num is None:
                    celula = html.escape(texto_cru)
                else:
                    cls = "ok" if pct_num >= 100 else ("warn" if pct_num >= 70 else "bad")
                    celula = f"<span class='chip chip-{cls}'>{html.escape(texto_cru)}</span>"
            elif c == "Movimentação Operacional":
                if texto_cru.startswith("Ausente"):
                    celula = f"<span class='txt-ausente'>{html.escape(texto_cru)}</span>"
                else:
                    celula = html.escape(texto_cru)
            elif c in ("Exemplares", "SKUs", "Meta Individual") and isinstance(valor, numbers.Real) \
                    and not pd.isna(valor):
                celula = formatar_inteiro_br(valor)
            else:
                celula = html.escape(texto_cru)
            celulas.append(f"<td>{celula}</td>")
        linhas_html.append("<tr>" + "".join(celulas) + "</tr>")

    return (
        "<div class='tabela-wrapper'><table class='tabela-gerencial'>"
        f"<colgroup>{colgroup}</colgroup>"
        f"<thead><tr>{cabecalho}</tr></thead>"
        f"<tbody>{''.join(linhas_html)}</tbody>"
        "</table></div>"
    )


def gerar_excel_gerencial(df_real):
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        (df_real if not df_real.empty else pd.DataFrame(
            columns=["Cargo", "Colaboradora", "Exemplares", "SKUs", "Movimentação Operacional"]
        )).to_excel(writer, index=False, sheet_name="Produtividade")
    buffer.seek(0)
    return buffer.getvalue()


def status_meta(pct):
    if pct >= 1:
        return "Meta batida", "ok"
    if pct >= 0.7:
        return "No ritmo", "warn"
    return "Abaixo da meta", "bad"


def card_html(icone, titulo, valor, sub, pct=None, accent="#2563EB"):
    chip = ""
    barra = ""
    if pct is not None:
        texto_status, cls = status_meta(pct)
        chip = f"<span class='chip chip-{cls}'>{texto_status}</span>"
        largura = max(min(pct, 1.0), 0.0) * 100
        barra = (f"<div class='bar'><div class='bar-fill' "
                 f"style='width:{largura:.1f}%; background:{accent};'></div></div>")
    return (
        f"<div class='card-kpi' style='--accent-color:{accent};'>"
        f"<div class='card-top'><div class='card-icon'>{icone}</div>{chip}</div>"
        f"<div class='card-title'>{titulo}</div>"
        f"<div class='card-value'>{valor}</div>"
        f"<div class='card-sub'>{sub}</div>{barra}</div>"
    )


# =============================================================================
# 8. Cabeçalho
# =============================================================================
agora_br = datetime.now(ZoneInfo("America/Sao_Paulo"))
st.markdown(
    "<div class='hero'>"
    "<div class='hero-left'>"
    f"<div class='hero-logo'>{LOGO_HERO_SVG}</div>"
    "<div><div class='hero-title'>Painel Executivo de Produção</div>"
    "<div class='hero-sub'>Varejo — acompanhamento diário de produtividade</div></div>"
    "</div>"
    "<div class='hero-right'>"
    f"<div class='hero-date'>{data_produtividade.strftime('%d/%m/%Y')}</div>"
    f"<div class='hero-sub'>Atualizado às {agora_br.strftime('%H:%M')}</div>"
    "</div></div>",
    unsafe_allow_html=True,
)


# =============================================================================
# 9. LÓGICA PRINCIPAL
# =============================================================================
if uploaded_file:
    df_filtrado, usando_fallback = ler_planilha(uploaded_file.getvalue())

    if usando_fallback:
        st.warning(
            "⚠️ Não encontrei as colunas 'TOTAL' e 'USUARIO' pelo cabeçalho na primeira "
            "linha da planilha. Usando posição padrão (colunas I e M) por compatibilidade "
            "— verifique se o arquivo segue o modelo esperado."
        )

    if df_filtrado.empty:
        st.error(
            "❌ Nenhum dado válido foi encontrado na planilha enviada. Os totais abaixo "
            "estão zerados — **não envie este relatório para a diretoria sem antes "
            "verificar o arquivo**."
        )
        total_exemplares, total_skus = 0, 0
    else:
        total_exemplares = int(df_filtrado["TOTAL"].sum())
        total_skus = int(len(df_filtrado))

        nomes_excel_validos = {nome_excel(n, ALIAS_EXCEL) for n in NOMES_LISTA}
        nomes_manuais_norm = {normalizar(p["nome"]) for p in st.session_state["pessoas_manuais"]}
        nomes_nao_mapeados = sorted(
            set(df_filtrado["USUARIO"]) - nomes_excel_validos - nomes_manuais_norm - set(cfg.get("ignorados", []))
        )
        if nomes_nao_mapeados:
            st.warning(
                "⚠️ Encontrados na planilha, mas **não mapeados** para ninguém da equipe: "
                + ", ".join(nomes_nao_mapeados)
                + ". Esses registros entram no total geral, mas não aparecem na tabela individual. "
                "Use **🆕 Novos na planilha** na barra lateral para cadastrar, vincular ou ignorar."
            )

    pct_exemplares = (total_exemplares / META_EXEMPLARES) if META_EXEMPLARES else 0
    pct_skus = (total_skus / META_SKUS) if META_SKUS else 0

    data_str_atual = data_produtividade.strftime("%Y-%m-%d")

    # ------------------------------------------------------------------
    # Monta a tabela gerencial (mesma lógica de antes)
    # ------------------------------------------------------------------
    data_gerencial = []
    textos_automaticos_por_pessoa = {}
    for n in NOMES_LISTA:
        if n in remover_do_setor:
            continue

        is_ausente = n in faltas_selecionadas

        if is_ausente:
            if not df_filtrado.empty:
                df_func = df_filtrado[df_filtrado["USUARIO"] == nome_excel(n, ALIAS_EXCEL)]
                qtd_exemplares = int(df_func["TOTAL"].sum())
                qtd_skus = int(len(df_func))
            else:
                qtd_exemplares, qtd_skus = 0, 0
            motivo_individual = dict_motivos_falta.get(n, "Falta administrativa")
            justificativa_texto = f"Ausente. Motivo: {motivo_individual}."
            cargo_atual = CARGO_POR_NOME.get(n, "Operador(a)")
        else:
            mov = dict_movimentacao[n]
            cargo_atual = mov["cargo"]
            if not df_filtrado.empty:
                df_func = df_filtrado[df_filtrado["USUARIO"] == nome_excel(n, ALIAS_EXCEL)]
                qtd_exemplares = int(df_func["TOTAL"].sum())
                qtd_skus = int(len(df_func))
            else:
                qtd_exemplares, qtd_skus = 0, 0

            if qtd_skus == 0:
                continue

            historico_justificativas = []
            for m in mov["movimentacoes"]:
                sai, ret, loc = m["sai"].strip(), m["ret"].strip(), m["loc"].strip()
                if not (sai or ret or loc):
                    continue

                partes = []
                if loc:
                    partes.append(f"ao {loc}")
                if sai and ret:
                    partes.append(f"das {sai} às {ret}")
                elif sai:
                    partes.append(f"a partir das {sai}")
                elif ret:
                    partes.append(f"até às {ret}")

                prefixo = "Encaminhada" if not historico_justificativas else "encaminhada"
                complemento = " ".join(partes) if partes else "(sem horário/local informado)"
                historico_justificativas.append(f"{prefixo} {complemento}")

            if historico_justificativas:
                justificativa_texto = " ; ".join(historico_justificativas) + "."
            elif qtd_skus == 0:
                justificativa_texto = "Sem registros na planilha nesta data."
            else:
                justificativa_texto = "Atividade normal no setor."

        # A ausência tem prioridade sobre qualquer texto manual salvo antes.
        override_chave = (data_str_atual, n)
        if is_ausente:
            if override_chave in st.session_state["mov_manual_overrides"]:
                del st.session_state["mov_manual_overrides"][override_chave]
                salvar_overrides_disco(st.session_state["mov_manual_overrides"])
        else:
            textos_automaticos_por_pessoa[n] = justificativa_texto
            if override_chave in st.session_state["mov_manual_overrides"]:
                justificativa_texto = st.session_state["mov_manual_overrides"][override_chave]

        linha = {
            "Cargo": cargo_atual,
            "Colaboradora": n,
            "Exemplares": qtd_exemplares,
            "SKUs": qtd_skus,
            "Movimentação Operacional": justificativa_texto,
        }

        if METAS_INDIVIDUAIS:
            meta_pessoa = METAS_INDIVIDUAIS.get(n)
            if meta_pessoa:
                linha["Meta Individual"] = meta_pessoa
                linha["% Meta Individual"] = f"{(qtd_exemplares / meta_pessoa):.0%}"
            else:
                linha["Meta Individual"] = ""
                linha["% Meta Individual"] = ""

        data_gerencial.append(linha)

    # Pessoas adicionadas manualmente pela lateral
    for pessoa_manual in st.session_state["pessoas_manuais"]:
        nome_manual = pessoa_manual["nome"]
        cargo_manual = pessoa_manual["cargo"]
        if nome_manual in NOMES_LISTA:
            continue  # já foi cadastrada na equipe depois; evita linha duplicada

        if not df_filtrado.empty:
            df_func_manual = df_filtrado[df_filtrado["USUARIO"] == normalizar(nome_manual)]
            qtd_exemplares_manual = int(df_func_manual["TOTAL"].sum())
            qtd_skus_manual = int(len(df_func_manual))
        else:
            qtd_exemplares_manual, qtd_skus_manual = 0, 0

        override_chave_manual = (data_str_atual, nome_manual)
        if qtd_skus_manual > 0:
            texto_automatico_manual = "Atividade normal no setor."
        else:
            texto_automatico_manual = "Sem registros na planilha nesta data. (Adicionada manualmente)"
        textos_automaticos_por_pessoa[nome_manual] = texto_automatico_manual
        justificativa_manual = st.session_state["mov_manual_overrides"].get(
            override_chave_manual, texto_automatico_manual
        )

        linha_manual = {
            "Cargo": cargo_manual,
            "Colaboradora": nome_manual,
            "Exemplares": qtd_exemplares_manual,
            "SKUs": qtd_skus_manual,
            "Movimentação Operacional": justificativa_manual,
        }
        if METAS_INDIVIDUAIS:
            linha_manual["Meta Individual"] = ""
            linha_manual["% Meta Individual"] = ""

        data_gerencial.append(linha_manual)

    df_real = pd.DataFrame(data_gerencial)

    # Ordenação: cargo (ordem da equipe) e depois exemplares (maior primeiro)
    if not df_real.empty:
        ordem_cargo = {c: i for i, c in enumerate(CARGOS_ORDEM)}
        df_real["_ordem"] = df_real["Cargo"].map(ordem_cargo).fillna(len(ordem_cargo))
        df_real = (
            df_real.sort_values(["_ordem", "Exemplares"], ascending=[True, False], kind="stable")
            .drop(columns="_ordem")
            .reset_index(drop=True)
        )

    # ------------------------------------------------------------------
    # Indicadores
    # ------------------------------------------------------------------
    if not df_real.empty:
        ativos = int((df_real["SKUs"] > 0).sum())
        linha_destaque = df_real.loc[df_real["Exemplares"].idxmax()]
        destaque_txt = f"{linha_destaque['Colaboradora']} ({int(linha_destaque['Exemplares']):,} un)"
    else:
        ativos = 0
        destaque_txt = "—"
    media_por_pessoa = int(total_exemplares / ativos) if ativos else 0

    k1, k2, k3, k4 = st.columns(4)
    with k1:
        st.markdown(card_html(
            "📦", "Total de exemplares", f"{total_exemplares:,} un",
            f"Meta diária {META_EXEMPLARES:,} un — atingido {pct_exemplares:.1%}",
            pct_exemplares, "#2563EB"), unsafe_allow_html=True)
    with k2:
        st.markdown(card_html(
            "🏷️", "Total de SKU", f"{total_skus:,}",
            f"Meta diária {META_SKUS:,} — atingido {pct_skus:.1%}",
            pct_skus, "#0D9488"), unsafe_allow_html=True)
    with k3:
        st.markdown(card_html(
            "👥", "Colaboradores ativos", f"{ativos}",
            f"{len(NOMES_LISTA)} na equipe cadastrada", None, "#7C3AED"), unsafe_allow_html=True)
    with k4:
        st.markdown(card_html(
            "⚡", "Média por colaborador", f"{media_por_pessoa:,} un",
            f"Destaque: {html.escape(destaque_txt)}", None, "#EA580C"), unsafe_allow_html=True)

    # Imagem gerada uma vez e usada nas abas de imagem e e-mail
    imagem_relatorio = gerar_relatorio_imagem(
        total_exemplares, total_skus, pct_exemplares, pct_skus, META_EXEMPLARES, META_SKUS, df_real,
        data_formatada=data_formatada,
    )
    nome_arquivo_imagem = f"relatorio_producao_{data_produtividade.strftime('%Y-%m-%d')}.png"

    tab_det, tab_img, tab_hist, tab_mail = st.tabs(
        ["📋 Detalhamento", "🖼️ Relatório em imagem", "📈 Histórico", "✉️ E-mail"]
    )

    # ==================================================================
    # ABA 1 — Detalhamento gerencial
    # ==================================================================
    with tab_det:
        st.markdown("<div class='secao-titulo'>Detalhamento gerencial de produtividade</div>", unsafe_allow_html=True)

        col_toggle1, col_toggle2 = st.columns(2)
        with col_toggle1:
            mostrar_individual = st.checkbox(
                "👁️ Mostrar Exemplares/SKUs individuais", value=True, key="mostrar_individual"
            )
        with col_toggle2:
            modo_edicao = st.checkbox(
                "✏️ Editar tabela (como planilha)", value=False, key="modo_edicao_tabela"
            )

        colunas_ocultaveis = ["Exemplares", "SKUs"]
        if mostrar_individual:
            df_exibir = df_real.copy()
        else:
            df_exibir = df_real.drop(columns=[c for c in colunas_ocultaveis if c in df_real.columns])

        if modo_edicao and not df_exibir.empty:
            st.caption(
                "✍️ A coluna **Movimentação Operacional** é livre — escreva o que quiser. "
                "As demais colunas ficam bloqueadas para não conflitar com os dados da "
                "planilha. O texto tenta salvar sozinho ao sair do campo, mas para garantir "
                "que nada se perca clique em **💾 Salvar agora** antes de mexer nos horários "
                "na lateral ou atualizar a página."
            )

            colunas_bloqueadas = [c for c in df_exibir.columns if c != "Movimentação Operacional"]

            df_exibir_editado = st.data_editor(
                df_exibir,
                use_container_width=True,
                hide_index=True,
                num_rows="fixed",
                key="editor_tabela_gerencial",
                disabled=colunas_bloqueadas,
                column_config={
                    "Movimentação Operacional": st.column_config.TextColumn(
                        "Movimentação Operacional",
                        help="Escreva livremente o que quiser exibir para esta pessoa.",
                    ),
                },
            )

            for col in df_exibir_editado.columns:
                df_real[col] = df_exibir_editado[col].values

            def persistir_movimentacao_editada(df_editado):
                """Grava como override apenas o texto realmente editado à mão
                (diferente do automático). Texto igual ao automático remove o
                override antigo em vez de recriá-lo."""
                total_gravado = 0
                if "Movimentação Operacional" in df_editado.columns and "Colaboradora" in df_editado.columns:
                    for _, linha_editada in df_editado.iterrows():
                        nome_pessoa = linha_editada.get("Colaboradora")
                        texto_editado = texto_seguro(linha_editada.get("Movimentação Operacional"))
                        if not nome_pessoa:
                            continue
                        chave = (data_str_atual, nome_pessoa)
                        texto_automatico = textos_automaticos_por_pessoa.get(nome_pessoa)
                        if texto_automatico is not None and texto_editado == texto_automatico:
                            st.session_state["mov_manual_overrides"].pop(chave, None)
                        else:
                            st.session_state["mov_manual_overrides"][chave] = texto_editado
                            total_gravado += 1
                    salvar_overrides_disco(st.session_state["mov_manual_overrides"])
                return total_gravado

            persistir_movimentacao_editada(df_exibir_editado)

            col_salvar, col_restaurar_texto = st.columns([1, 2])
            with col_salvar:
                if st.button("💾 Salvar agora", use_container_width=True, type="primary"):
                    qtd = persistir_movimentacao_editada(df_exibir_editado)
                    st.success(f"✅ Salvo! ({qtd} linha(s) gravada(s) para {data_formatada})")
            with col_restaurar_texto:
                if st.button("🔄 Restaurar texto automático desta data", use_container_width=True):
                    chaves_para_remover = [
                        k for k in st.session_state["mov_manual_overrides"] if k[0] == data_str_atual
                    ]
                    for k in chaves_para_remover:
                        del st.session_state["mov_manual_overrides"][k]
                    salvar_overrides_disco(st.session_state["mov_manual_overrides"])
                    st.rerun()
        else:
            st.markdown(renderizar_tabela_html(df_exibir), unsafe_allow_html=True)

        if not df_real.empty:
            st.download_button(
                label="📥 Baixar tabela em Excel",
                data=gerar_excel_gerencial(df_real),
                file_name=f"produtividade_{data_produtividade.strftime('%Y-%m-%d')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )

    # ==================================================================
    # ABA 2 — Relatório em imagem
    # ==================================================================
    with tab_img:
        st.markdown("<div class='secao-titulo'>Relatório em imagem</div>", unsafe_allow_html=True)
        st.caption(
            "Clique com o botão direito na imagem e escolha **Copiar imagem** para colar "
            "direto no e-mail, ou baixe o arquivo abaixo."
        )
        st.image(imagem_relatorio, use_container_width=True)
        st.download_button(
            label="📥 Baixar relatório em imagem",
            data=imagem_relatorio,
            file_name=nome_arquivo_imagem,
            mime="image/png",
        )

    # ==================================================================
    # ABA 3 — Histórico
    # ==================================================================
    with tab_hist:
        st.markdown("<div class='secao-titulo'>Evolução diária</div>", unsafe_allow_html=True)
        col_h1, col_h2 = st.columns([1, 2])
        with col_h1:
            if st.button(f"💾 Salvar {data_formatada} no histórico", use_container_width=True, key="btn_salvar_hist"):
                if salvar_historico_dia(data_str_atual, total_exemplares, total_skus, ativos):
                    st.success(f"Dia {data_formatada} salvo no histórico.")
                else:
                    st.error("Não consegui gravar o arquivo de histórico.")
        with col_h2:
            st.caption("Salvar o mesmo dia de novo substitui o registro anterior daquela data.")

        df_hist = carregar_historico()
        if df_hist.empty:
            st.info("Ainda não há dias no histórico. Salve o primeiro com o botão acima.")
        else:
            df_hist["Data"] = pd.to_datetime(df_hist["Data"], errors="coerce")
            df_hist = df_hist.dropna(subset=["Data"]).sort_values("Data")

            m1, m2, m3 = st.columns(3)
            m1.metric("Dias registrados", len(df_hist))
            m2.metric("Média de exemplares", f"{int(df_hist['Exemplares'].mean()):,}")
            melhor = df_hist.loc[df_hist["Exemplares"].idxmax()]
            m3.metric("Melhor dia", melhor["Data"].strftime("%d/%m"), f"{int(melhor['Exemplares']):,} un")

            graf_ex = df_hist.set_index("Data")[["Exemplares"]].copy()
            graf_ex["Meta"] = META_EXEMPLARES
            st.markdown("**Exemplares por dia**")
            st.line_chart(graf_ex, color=["#2563EB", "#CBD5E1"])

            graf_sku = df_hist.set_index("Data")[["SKUs"]].copy()
            graf_sku["Meta"] = META_SKUS
            st.markdown("**SKUs por dia**")
            st.line_chart(graf_sku, color=["#0D9488", "#CBD5E1"])

            df_hist_exibir = df_hist.copy()
            df_hist_exibir["Data"] = df_hist_exibir["Data"].dt.strftime("%d/%m/%Y")
            st.dataframe(df_hist_exibir.iloc[::-1], use_container_width=True, hide_index=True)
            st.download_button(
                "📥 Baixar histórico (CSV)",
                data=df_hist_exibir.to_csv(index=False).encode("utf-8-sig"),
                file_name="historico_diario.csv",
                mime="text/csv",
            )

    # ==================================================================
    # ABA 4 — E-mail
    # ==================================================================
    with tab_mail:
        st.markdown("<div class='secao-titulo'>Texto do e-mail para a diretoria</div>", unsafe_allow_html=True)

        texto_final = (
            f"Boa tarde, Prezados.\n\nSegue abaixo o relatório de produção.\n"
            f"referente ao dia {data_formatada}.\n\nObservações do Dia:\n"
            f"(imagem do painel anexada/embutida neste e-mail)\n\n"
            f"--------------------------------\n"
            f"Resumo Varejo.\nSKU: {total_skus}\nExemplares: {total_exemplares:,}\n"
            f"--------------------------------\n\nAtenciosamente,"
        )

        CID_IMAGEM_RELATORIO = "relatorio_producao_imagem"
        corpo_html_email = f"""\
<html>
  <body style="font-family: Arial, Helvetica, sans-serif; font-size: 14px; color:#111827;">
    <p>Boa tarde, Prezados.</p>
    <p>Segue abaixo o relatório de produção.<br>
       referente ao dia {html.escape(data_formatada)}.</p>
    <p><strong>Observações do Dia:</strong></p>
    <p><img src="cid:{CID_IMAGEM_RELATORIO}" alt="Painel Executivo de Produção" style="max-width:700px; width:100%; border:1px solid #E5E7EB; border-radius:8px;"></p>
    <p>--------------------------------<br>
       Resumo Varejo.<br>
       SKU: {total_skus}<br>
       Exemplares: {total_exemplares:,}<br>
       --------------------------------</p>
    <p>Atenciosamente,</p>
  </body>
</html>
"""

        st.text_area("Selecione tudo abaixo e copie (Ctrl+A / Ctrl+C):", value=texto_final, height=220, key="texto_email")

        if st.button("📧 Enviar relatório por e-mail agora"):
            enviar_email_relatorio(
                assunto=f"Relatório de Produção - {data_formatada}",
                corpo_texto=texto_final,
                corpo_html=corpo_html_email,
                imagem_bytes=imagem_relatorio,
                nome_imagem=nome_arquivo_imagem,
                cid_imagem=CID_IMAGEM_RELATORIO,
            )

else:
    st.markdown(
        "<div class='vazio'>"
        "<h3>Envie a planilha para começar</h3>"
        "<p>Use o campo <b>Planilha de produção</b> na barra lateral. "
        "Os totais, a tabela e o relatório aparecem assim que o arquivo for lido.</p>"
        "</div>",
        unsafe_allow_html=True,
    )
