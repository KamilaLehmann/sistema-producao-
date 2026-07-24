"""
Painel Executivo de Produção — Varejo
======================================
Requisitos: streamlit>=1.32, pandas, openpyxl, matplotlib

Melhorias implementadas nesta versão em relação ao script original:
  (obs.: o login por senha foi removido a pedido — o painel fica sem
  proteção de acesso, igual ao comportamento original)
  2. Leitura da planilha por CABEÇALHO de coluna (TOTAL / USUARIO), com aviso
     caso precise cair para a posição fixa (colunas I/M) por compatibilidade.
  3. Remoção dos números de fallback fixos (50271 / 1104): agora, se a
     planilha não tiver dados válidos, o painel avisa claramente em vez de
     mostrar números "fantasmas".
  4. Alerta de nomes encontrados na planilha que não batem com ninguém da
     equipe cadastrada (evita gente "sumir" do relatório sem ninguém notar).
  5. Cache da leitura da planilha (st.cache_data) para não reprocessar tudo
     a cada clique no sidebar.
  6. Equipe configurável via JSON editável na própria interface (sem precisar
     mexer no código para adicionar/remover pessoas, alias de Excel ou metas
     individuais).
  7. Horários de movimentação via tabela editável (st.data_editor) com
     colunas de hora, no lugar de vários text_input soltos + botões de
     add/remover linha.
  8. Histórico diário salvo em CSV local + gráfico de evolução (tendência).
  9. Metas individuais opcionais por colaborador (além da meta do setor).
 10. Envio de e-mail direto por SMTP (opcional, configurável no sidebar).
 11. Exportação da tabela gerencial em Excel (.xlsx), além da imagem PNG.
 12. Tabela do detalhamento gerencial mais compacta (linhas menores, sem
     alterar o conteúdo do texto), com opção de editá-la como planilha e de
     ocultar/exibir Exemplares e SKUs individuais sob demanda.
"""

import streamlit as st
import pandas as pd
import openpyxl
from datetime import datetime, time as dtime
import io
import os
import json
import html
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
HIST_PATH = "historico_producao.csv"

# (constantes/função de senha removidas junto com o login — ver comentário
# na seção 2 abaixo caso queira reativar o acesso restrito no futuro)


# =============================================================================
# 1. Configuração e Estilização de Design Premium (HTML / CSS)
# =============================================================================
st.set_page_config(page_title="Dashboard Executivo Varejo", layout="wide")
st.markdown("""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Sora:wght@600;700;800&family=Inter:wght@400;500;600;700&display=swap');

    html, body, [class*="css"] { font-family: 'Inter', sans-serif; }

    .stApp { background-color: #F8FAFC; }

    .block-container {padding-top: 1.2rem; padding-bottom: 0rem; max-width: 96%;}

    section[data-testid="stSidebar"] {
        background-color: #FFFFFF;
        border-right: 1px solid #E5E7EB;
    }
    section[data-testid="stSidebar"] h3 {
        font-family: 'Sora', sans-serif;
        color: #0F172A !important;
        font-size: 0.95rem !important;
    }

    .card-kpi {
        background: #FFFFFF;
        border: 1px solid #E5E7EB;
        border-left: 4px solid var(--accent-color, #2563EB);
        color: #0F172A;
        padding: 20px 22px;
        border-radius: 10px;
        box-shadow: 0 1px 3px rgba(15, 23, 42, 0.06);
        text-align: left;
        margin-bottom: 14px;
        transition: box-shadow 0.2s ease;
    }
    .card-kpi:hover { box-shadow: 0 4px 14px rgba(15, 23, 42, 0.08); }
    .card-title {
        font-family: 'Sora', sans-serif;
        font-size: 0.8rem; font-weight: 700; opacity: 0.9; margin-bottom: 6px;
        letter-spacing: 1px; text-transform: uppercase; color: #64748B;
    }
    .card-value {
        font-family: 'Sora', sans-serif;
        font-size: 2.3rem; font-weight: 800; line-height: 1; margin-bottom: 8px;
        color: #0F172A;
    }
    .card-sub { font-size: 0.85rem; font-weight: 600; color: #94A3B8; }

    div.stProgress > div > div > div {
        background: #2563EB;
        height: 6px; border-radius: 4px;
    }
    div.stProgress > div > div { background: #E5E7EB; border-radius: 4px; }

    hr { border-color: #E5E7EB !important; }

    .stButton>button {
        background: #FFFFFF;
        border: 1px solid #CBD5E1;
        color: #0F172A;
        border-radius: 8px;
        font-family: 'Inter', sans-serif;
        font-weight: 600;
        transition: all 0.15s ease;
    }
    .stButton>button:hover {
        border-color: #2563EB;
        color: #2563EB;
    }
    div[data-testid="stDownloadButton"] button {
        background: #0F172A;
        border: 1px solid #0F172A;
        color: #FFFFFF;
    }
    div[data-testid="stDownloadButton"] button:hover {
        background: #1E293B;
        border-color: #1E293B;
        color: #FFFFFF;
    }

    .stTextInput>div>div>input, .stDateInput input {
        background-color: #FFFFFF;
        color: #0F172A;
        border: 1px solid #E2E8F0;
        border-radius: 6px;
    }
    .stTextInput>div>div>input:focus { border-color: #2563EB; }

    div[data-testid="stDataFrame"] {
        border: 1px solid #E5E7EB;
        border-radius: 10px;
        overflow: hidden;
    }

    /* Tabela gerencial em HTML (bordas arredondadas + quebra de texto) */
    .tabela-wrapper {
        border: 1px solid #E2E8F0;
        border-radius: 14px;
        overflow: hidden;
        box-shadow: 0 1px 4px rgba(15, 23, 42, 0.07);
        margin-bottom: 16px;
    }
    table.tabela-gerencial {
        width: 100%;
        border-collapse: collapse;
        font-family: 'Inter', sans-serif;
        font-size: 0.86rem;
        table-layout: fixed;
    }
    table.tabela-gerencial thead th {
        background: #0F172A;
        color: #FFFFFF;
        text-align: left;
        padding: 9px 14px;
        font-weight: 600;
        font-size: 0.78rem;
        text-transform: uppercase;
        letter-spacing: 0.3px;
    }
    table.tabela-gerencial tbody td {
        padding: 6px 14px;
        border-top: 1px solid #EEF2F6;
        color: #0F172A;
        vertical-align: middle;
        white-space: normal;      /* permite quebra de linha nas colunas de texto livre */
        overflow-wrap: break-word;
        line-height: 1.3;
        font-size: 0.86rem;
    }
    /* Cargo e Colaboradora são textos curtos/categóricos: nunca cortam no
       meio da palavra, ficam sempre em uma linha só. */
    table.tabela-gerencial td:nth-child(1),
    table.tabela-gerencial td:nth-child(2),
    table.tabela-gerencial th:nth-child(1),
    table.tabela-gerencial th:nth-child(2) {
        white-space: nowrap;
    }
    table.tabela-gerencial tbody tr:nth-child(even) { background: #FAFBFC; }
    table.tabela-gerencial tbody tr:hover { background: #F1F5F9; }
    .tabela-vazia {
        padding: 18px 16px;
        color: #64748B;
        font-size: 0.9rem;
        border: 1px solid #E2E8F0;
        border-radius: 14px;
        background: #FFFFFF;
    }
    </style>
""", unsafe_allow_html=True)

# Logo leve em SVG (substitui o PNG base64 do arquivo original — ver nota no
# topo do arquivo sobre como restaurar o logo original caso deseje).
LOGO_REALBRAS_SVG = """<svg width="34" height="34" viewBox="0 0 34 34" xmlns="http://www.w3.org/2000/svg">
  <rect width="34" height="34" rx="8" fill="#0F172A"/>
  <path d="M9 24V10h7.5c3.6 0 5.8 1.8 5.8 4.9 0 2.1-1.1 3.6-3 4.3l3.4 4.8h-3.4l-3-4.3H12V24H9zm3-6.7h4.2c1.8 0 2.8-.8 2.8-2.3s-1-2.3-2.8-2.3H12v4.6z" fill="#FFFFFF"/>
</svg>""".strip()


# =============================================================================
# 2. LOGIN — DESATIVADO
# =============================================================================
# O gate de senha foi removido a pedido. Se quiser reativar no futuro, basta
# restaurar a função `checar_senha()` (mantida como exemplo abaixo, comentada)
# e voltar a chamar `if not checar_senha(): st.stop()` antes do cabeçalho.
#
# def checar_senha():
#     if st.session_state.get("autenticado"):
#         return True
#
#     def senha_confirmada():
#         if st.session_state.get("senha_input") == obter_senha_configurada():
#             st.session_state["autenticado"] = True
#         else:
#             st.session_state["autenticado"] = False
#
#     st.markdown(f"""
#     <div style='display:flex; align-items:center; gap:10px; margin-bottom:18px;'>
#         {LOGO_REALBRAS_SVG}
#         <h2 style='font-family: "Sora", sans-serif; font-weight:800; color:#0F172A; margin:0;'>🔒 Acesso Restrito</h2>
#     </div>
#     """, unsafe_allow_html=True)
#     st.text_input("Senha de acesso:", type="password", on_change=senha_confirmada, key="senha_input")
#     if "autenticado" in st.session_state and not st.session_state["autenticado"]:
#         st.error("Senha incorreta. Tente novamente.")
#     st.caption("Dica: configure a senha em `.streamlit/secrets.toml` com a chave `SENHA_PAINEL`.")
#     return False

# =============================================================================
# 3. Cabeçalho
# =============================================================================
st.markdown(f"""
<div style='display:flex; align-items:center; gap:10px; margin-bottom:4px;'>
{LOGO_REALBRAS_SVG}
<h1 style='font-family: "Sora", sans-serif; font-weight:800; letter-spacing:-0.5px; color: #0F172A; margin:0;'>📊 Painel Executivo de Produção</h1>
</div>
<p style='text-align:left; font-family: "Inter", sans-serif; color:#64748B; font-size:0.95rem; margin-top:4px; margin-bottom:28px;'>Varejo · acompanhamento diário de produtividade</p>
""", unsafe_allow_html=True)


# =============================================================================
# 4. Normalização de texto (evita falha de correspondência de nomes)
# =============================================================================
def normalizar(texto):
    """Remove acentos, espaços extras e padroniza para maiúsculas, evitando
    falhas de correspondência entre o nome cadastrado no painel e o nome como
    aparece na planilha."""
    texto = str(texto).strip().upper()
    texto = unicodedata.normalize("NFKD", texto).encode("ASCII", "ignore").decode("ASCII")
    texto = " ".join(texto.split())
    return texto


# =============================================================================
# 5. EQUIPE CONFIGURÁVEL (editável pela interface, sem precisar mexer no código)
# =============================================================================
# Cada pessoa é um dicionário com:
#   nome            -> nome de exibição (obrigatório)
#   alias_excel      -> como o nome aparece na planilha, se for diferente (opcional)
#   meta_exemplares  -> meta individual diária de exemplares (opcional)
#   saida_padrao / retorno_padrao / local_padrao -> pré-preenche a tabela de
#                       movimentação quando a pessoa for marcada como movimentada (opcional)
DEFAULT_EQUIPE = {
    "Líder": [
        {"nome": "Kamila Moraes"},
        {"nome": "Beatriz Alcantara", "alias_excel": "BEATRIZ"},
    ],
    "Apoio": [
        {"nome": "Alisson Lima"},
    ],
    "Operador(a)": [
        {"nome": "Rosana Delfino", "saida_padrao": "06:15", "retorno_padrao": "07:30", "local_padrao": "Setor Loja"},
        {"nome": "Ana Caroline", "alias_excel": "ANACAROLINE", "saida_padrao": "06:15", "retorno_padrao": "10:30", "local_padrao": "Setor Loja"},
        {"nome": "Karoline Gonçalves", "saida_padrao": "06:15", "retorno_padrao": "10:30", "local_padrao": "Setor Loja"},
        {"nome": "Gabriele", "saida_padrao": "06:15", "retorno_padrao": "10:00", "local_padrao": "Setor Loja"},
        {"nome": "Beatriz Mascarenhas"},
        {"nome": "Graziela Pereira", "alias_excel": "GRAZIELA PEREIRA DO NASCIMENTO"},
        {"nome": "Paula Roberta", "alias_excel": "PAULA ROBERTA SANTOS DA SILVA"},
        {"nome": "Weliton"},
        {"nome": "Ellen Kelly"},
    ],
}

if "equipe_config" not in st.session_state:
    st.session_state["equipe_config"] = json.loads(json.dumps(DEFAULT_EQUIPE))  # cópia profunda


def construir_estruturas_equipe(config):
    """A partir do dicionário de configuração da equipe, monta as estruturas
    usadas pelo resto do app: lista de cargos->nomes, lista geral de nomes,
    mapa de alias do Excel, mapa de metas individuais e mapa de valores
    padrão de movimentação."""
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


def parse_hora_str(valor_str):
    """Converte 'HH:MM' em datetime.time; retorna None se vazio/ inválido."""
    if not valor_str:
        return None
    try:
        h, m = valor_str.split(":")
        return dtime(int(h), int(m))
    except Exception:
        return None


def formatar_hora_editor(valor):
    """Converte o valor de uma célula de hora vinda do st.data_editor para o
    texto 'HHhMM'. Não assume um único tipo de retorno (datetime.time,
    pandas.Timestamp ou string 'HH:MM'/'HH:MM:SS' já apareceram dependendo da
    versão do Streamlit/pandas) — se não conseguir reconhecer o formato,
    devolve vazio em vez de derrubar o app."""
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
            # Cobre variações não previstas acima (ex.: string ISO completa).
            return pd.to_datetime(texto).strftime("%Hh%M")
        except Exception:
            return texto  # formato não reconhecido: mantém como veio, em vez de sumir
    if hasattr(valor, "strftime"):
        try:
            return valor.strftime("%Hh%M")
        except Exception:
            return ""
    return str(valor).strip()


def texto_seguro(valor):
    """Converte o valor de uma célula de texto vinda do st.data_editor para
    string, tratando None/NaN como vazio — evita AttributeError quando a
    célula fica vazia e volta como float('nan') em vez de None ou ''."""
    if valor is None:
        return ""
    try:
        if pd.isna(valor):
            return ""
    except (TypeError, ValueError):
        pass
    return str(valor)


# =============================================================================
# 6. BARRA LATERAL — Upload, data, configuração de equipe e filtros
# =============================================================================
st.sidebar.header("🛠️ Controle Operacional")
uploaded_file = st.sidebar.file_uploader("Upload da Planilha Excel", type=["xlsx"], key="uploaded_file")

st.sidebar.markdown("<hr style='margin:14px 0px; border-color: #D1D5DB;'>", unsafe_allow_html=True)

data_produtividade = st.sidebar.date_input("Data da Produtividade:", datetime.now(), key="data_produtividade")
data_formatada = data_produtividade.strftime("%d/%m")

st.sidebar.markdown("<hr style='margin:14px 0px; border-color: #D1D5DB;'>", unsafe_allow_html=True)

# --- Configuração de equipe editável -----------------------------------------
with st.sidebar.expander("⚙️ Configurar Equipe (avançado)"):
    st.caption(
        "Edite o JSON abaixo para adicionar/remover pessoas, corrigir o nome "
        "como aparece na planilha (`alias_excel`) ou definir uma meta "
        "individual (`meta_exemplares`). Clique em Aplicar para salvar."
    )
    texto_config = st.text_area(
        "Configuração da equipe (JSON):",
        value=json.dumps(st.session_state["equipe_config"], ensure_ascii=False, indent=2),
        height=220,
        key="texto_config_equipe",
    )
    col_aplicar, col_restaurar = st.columns(2)
    with col_aplicar:
        if st.button("✅ Aplicar", use_container_width=True):
            try:
                nova_config = json.loads(texto_config)
                if not isinstance(nova_config, dict):
                    raise ValueError("O JSON precisa ser um objeto com cargos como chaves.")
                st.session_state["equipe_config"] = nova_config
                st.success("Configuração de equipe atualizada.")
                st.rerun()
            except Exception as e:
                st.error(f"JSON inválido: {e}")
    with col_restaurar:
        if st.button("↩️ Restaurar padrão", use_container_width=True):
            st.session_state["equipe_config"] = json.loads(json.dumps(DEFAULT_EQUIPE))
            st.session_state.pop("texto_config_equipe", None)  # força a caixa de texto a recarregar o padrão
            st.rerun()

EQUIPE, NOMES_LISTA, ALIAS_EXCEL, METAS_INDIVIDUAIS, DEFAULTS_MOV, CARGO_POR_NOME = construir_estruturas_equipe(
    st.session_state["equipe_config"]
)

st.sidebar.markdown("<hr style='margin:14px 0px; border-color: #D1D5DB;'>", unsafe_allow_html=True)

st.sidebar.markdown("### 👁️ Filtros Gerenciais")
remover_do_setor = st.sidebar.multiselect("Ocultar do Setor (Tabela):", NOMES_LISTA, key="remover_do_setor")

st.sidebar.markdown("<hr style='margin:14px 0px; border-color: #D1D5DB;'>", unsafe_allow_html=True)

st.sidebar.markdown("### ❌ Ausências do Dia")
faltas_selecionadas = st.sidebar.multiselect("Selecione quem faltou hoje:", NOMES_LISTA, key="faltas_selecionadas")

st.sidebar.markdown("<hr style='margin:14px 0px; border-color: #D1D5DB;'>", unsafe_allow_html=True)

st.sidebar.markdown("### ⏳ Movimentação de Horários")
movimentados_selecionados = st.sidebar.multiselect(
    "🚚 Quem foi movimentado(a) hoje?",
    [n for n in NOMES_LISTA if n not in remover_do_setor and n not in faltas_selecionadas],
    key="movimentados_selecionados",
)

st.sidebar.markdown("<hr style='margin:14px 0px; border-color: #D1D5DB;'>", unsafe_allow_html=True)

MOTIVOS_FALTA_PADRAO = ["Falta administrativa", "Atestado médico", "Falta injustificada", "Folga compensatória", "Outro"]

dict_movimentacao = {}
dict_motivos_falta = {}

for cargo, integrantes in EQUIPE.items():
    integrantes_visiveis = [i for i in integrantes if i not in remover_do_setor]
    if integrantes_visiveis:
        st.sidebar.markdown(
            f"<h3 style='color:#1E3A8A; margin-top:10px; font-size:1.1rem;'>🔹 {cargo.upper()}</h3>",
            unsafe_allow_html=True,
        )

    for op in integrantes:
        if op in remover_do_setor:
            continue

        is_ausente = op in faltas_selecionadas
        is_movimentado = op in movimentados_selecionados

        if is_ausente:
            st.sidebar.markdown(f"❌ **{op} (AUSENTE)**")
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

        else:
            st.sidebar.markdown(
                f"👤 {op} <span style='font-size:0.8rem; color:gray;'>(sem movimentação)</span>",
                unsafe_allow_html=True,
            )
            dict_movimentacao[op] = {"cargo": cargo, "movimentacoes": []}

        st.sidebar.markdown("<hr style='margin:6px 0px; border-color: #D1D5DB;'>", unsafe_allow_html=True)


# =============================================================================
# 7. Configuração de e-mail (SMTP) — opcional
# =============================================================================
with st.sidebar.expander("✉️ Configuração de E-mail (SMTP)"):
    st.caption(
        "Preencha para habilitar o envio direto do relatório por e-mail. "
        "Por segurança, prefira configurar a senha em `st.secrets` em vez de digitá-la aqui."
    )
    smtp_host = st.text_input("Servidor SMTP:", value="smtp.gmail.com", key="smtp_host")
    smtp_port = st.number_input("Porta:", value=587, step=1, key="smtp_port")
    smtp_usuario = st.text_input("E-mail remetente:", key="smtp_usuario")
    smtp_senha = st.text_input("Senha / senha de app:", type="password", key="smtp_senha")
    destinatarios_texto = st.text_input("Destinatários (separados por vírgula):", key="smtp_destinatarios")


def enviar_email_relatorio(assunto, corpo_texto, imagem_bytes, nome_imagem):
    destinatarios = [d.strip() for d in destinatarios_texto.split(",") if d.strip()]
    if not (smtp_host and smtp_usuario and smtp_senha and destinatarios):
        st.error("Preencha servidor, remetente, senha e ao menos um destinatário na configuração de e-mail.")
        return
    try:
        msg = MIMEMultipart()
        msg["From"] = smtp_usuario
        msg["To"] = ", ".join(destinatarios)
        msg["Subject"] = assunto
        msg.attach(MIMEText(corpo_texto, "plain"))

        img_part = MIMEImage(imagem_bytes, name=nome_imagem)
        msg.attach(img_part)

        with smtplib.SMTP(smtp_host, int(smtp_port)) as servidor:
            servidor.starttls()
            servidor.login(smtp_usuario, smtp_senha)
            servidor.sendmail(smtp_usuario, destinatarios, msg.as_string())

        st.success(f"E-mail enviado com sucesso para: {', '.join(destinatarios)}")
    except Exception as e:
        st.error(f"Falha ao enviar e-mail: {e}")


# =============================================================================
# 8. Leitura da planilha — por CABEÇALHO de coluna, com cache e avisos claros
# =============================================================================
def localizar_colunas(sheet):
    """Procura, na primeira linha da planilha, as colunas cujo cabeçalho é
    'TOTAL' e 'USUARIO' (sem diferenciar acento/maiúscula). Retorna um dict
    com os índices encontrados; chaves ausentes indicam que o cabeçalho não
    foi localizado."""
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
    """Lê a planilha e devolve (dataframe, usando_fallback_de_coluna).
    Considera apenas linhas visíveis (não ocultas), como no comportamento
    original."""
    wb = openpyxl.load_workbook(io.BytesIO(bytes_arquivo), data_only=True)
    sheet = wb.active

    mapeamento = localizar_colunas(sheet)
    usando_fallback = ("TOTAL" not in mapeamento) or ("USUARIO" not in mapeamento)
    col_total = mapeamento.get("TOTAL", 9)   # coluna I, por compatibilidade
    col_usuario = mapeamento.get("USUARIO", 13)  # coluna M, por compatibilidade

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
# 9. Geração de relatório em imagem, Excel e histórico
# =============================================================================
def gerar_relatorio_imagem(total_exemplares, total_skus, pct_exemplares, pct_skus,
                            meta_exemplares, meta_skus, df_real):
    """Gera uma imagem (PNG) com cards de KPI + tabela de detalhamento (sem os
    números individuais de Exemplares/SKUs, que a diretoria não precisa ver),
    pronta para copiar ou baixar."""
    colunas_relatorio = ["Cargo", "Colaboradora", "Movimentação Operacional"]
    df_relatorio = df_real[colunas_relatorio] if not df_real.empty else df_real

    n_linhas = max(len(df_relatorio), 1)
    altura_fig = 3.5 + 0.35 * n_linhas
    fig = plt.figure(figsize=(11, altura_fig), dpi=200)
    fig.patch.set_facecolor("#FAFAFA")

    fig.text(0.04, 0.975, "Painel Executivo de Produção", fontsize=18, fontweight="bold", color="#111827", va="top")

    def desenhar_card(x, largura, titulo, valor, sub, cor_accent):
        ax = fig.add_axes([x, 0.68, largura, 0.20])
        ax.set_xlim(0, 1); ax.set_ylim(0, 1); ax.axis("off")
        card = mpatches.FancyBboxPatch(
            (0.01, 0.04), 0.98, 0.92, boxstyle="round,pad=0,rounding_size=0.08",
            linewidth=1, edgecolor="#EAEAEA", facecolor="white", transform=ax.transAxes
        )
        ax.add_patch(card)
        barra = mpatches.FancyBboxPatch(
            (0.01, 0.04), 0.015, 0.92, boxstyle="round,pad=0,rounding_size=0.008",
            linewidth=0, facecolor=cor_accent, transform=ax.transAxes
        )
        ax.add_patch(barra)
        ax.text(0.09, 0.76, titulo, fontsize=9, fontweight="bold", color="#9CA3AF", va="top")
        ax.text(0.09, 0.53, valor, fontsize=23, fontweight="bold", color="#111827", va="top")
        ax.text(0.09, 0.20, sub, fontsize=8, color="#111827", va="top")

    desenhar_card(0.04, 0.44, "TOTAL DE EXEMPLARES", f"{total_exemplares:,} un",
                  f"Meta Diária: {meta_exemplares:,} un  ·  Atingido: {pct_exemplares:.1%}", "#2563EB")
    desenhar_card(0.52, 0.44, "TOTAL DE SKU", f"{total_skus:,}",
                  f"Meta Diária: {meta_skus:,}  ·  Atingido: {pct_skus:.1%}", "#0D9488")

    ax = fig.add_axes([0.04, 0.03, 0.92, 0.58])
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
        tabela.scale(1, 1.8)

        for (row, col), cell in tabela.get_celld().items():
            cell.set_edgecolor("#EFEFEF")
            if row == 0:
                cell.set_facecolor("#111827")
                cell.set_text_props(color="white", fontweight="bold")
            else:
                cell.set_facecolor("#FFFFFF" if row % 2 == 0 else "#FAFAFA")
    else:
        ax.text(0, 0.9, "Nenhum dado disponível.", fontsize=9, color="#64748B")

    buffer = io.BytesIO()
    fig.savefig(buffer, format="png", facecolor=fig.get_facecolor(), bbox_inches="tight")
    plt.close(fig)
    buffer.seek(0)
    return buffer.getvalue()


def renderizar_tabela_html(df):
    """Renderiza o DataFrame como uma tabela HTML própria (em vez do grid
    interativo do st.dataframe), permitindo bordas arredondadas e quebra de
    linha automática quando o texto de uma célula é longo (ex.: justificativas
    de movimentação), em vez de cortar o conteúdo. As colunas passadas em `df`
    definem o que aparece — para ocultar Exemplares/SKUs, basta não incluir
    essas colunas no DataFrame antes de chamar esta função."""
    if df.empty:
        return "<div class='tabela-vazia'>Nenhum dado disponível.</div>"

    larguras = {
        "Cargo": "13%",
        "Colaboradora": "17%",
        "Exemplares": "10%",
        "SKUs": "8%",
        "Meta Individual": "11%",
        "% Meta Individual": "11%",
        # "Movimentação Operacional" fica sem largura fixa -> ocupa o restante
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
            if isinstance(valor, (int,)) or (isinstance(valor, float) and float(valor).is_integer()):
                texto = f"{int(valor):,}".replace(",", ".") if c in ("Exemplares", "SKUs", "Meta Individual") else str(valor)
            else:
                texto = "" if valor is None else str(valor)
            celulas.append(f"<td>{html.escape(texto)}</td>")
        linhas_html.append(f"<tr>{''.join(celulas)}</tr>")

    return f"""
    <div class="tabela-wrapper">
      <table class="tabela-gerencial">
        <colgroup>{colgroup}</colgroup>
        <thead><tr>{cabecalho}</tr></thead>
        <tbody>{''.join(linhas_html)}</tbody>
      </table>
    </div>
    """


def gerar_excel_gerencial(df_real):
    """Gera um .xlsx com a tabela gerencial completa (incluindo Exemplares e
    SKUs individuais), para uso interno de análise — diferente da imagem/
    e-mail para diretoria, que oculta esses números."""
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        (df_real if not df_real.empty else pd.DataFrame(
            columns=["Cargo", "Colaboradora", "Exemplares", "SKUs", "Movimentação Operacional"]
        )).to_excel(writer, index=False, sheet_name="Produtividade")
    buffer.seek(0)
    return buffer.getvalue()


def salvar_historico(data_str, total_ex, total_sk, pct_ex, pct_sk):
    """Acrescenta (ou atualiza) o registro do dia no histórico local em CSV e
    devolve o histórico completo carregado. Observação: em ambientes com
    armazenamento efêmero (ex.: alguns serviços de deploy em nuvem), este
    arquivo pode não persistir entre reinicializações — para uso contínuo,
    considere apontar HIST_PATH para um disco persistente ou banco de dados."""
    novo = pd.DataFrame([{
        "data": data_str, "exemplares": total_ex, "skus": total_sk,
        "pct_exemplares": pct_ex, "pct_skus": pct_sk,
    }])
    if os.path.exists(HIST_PATH):
        try:
            hist = pd.read_csv(HIST_PATH)
            hist = hist[hist["data"] != data_str]
            hist = pd.concat([hist, novo], ignore_index=True)
        except Exception:
            hist = novo
    else:
        hist = novo
    hist = hist.sort_values("data")
    try:
        hist.to_csv(HIST_PATH, index=False)
    except Exception:
        pass  # ambiente sem permissão de escrita: segue sem persistir
    return hist


# =============================================================================
# 10. LÓGICA PRINCIPAL
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
        nomes_nao_mapeados = sorted(set(df_filtrado["USUARIO"]) - nomes_excel_validos)
        if nomes_nao_mapeados:
            st.warning(
                "⚠️ Encontrados na planilha, mas **não mapeados** para ninguém da equipe "
                "cadastrada: " + ", ".join(nomes_nao_mapeados) + ". Esses registros entram "
                "no TOTAL geral acima, mas não aparecem na tabela individual — confira se "
                "não é alguém novo que precisa ser cadastrado em '⚙️ Configurar Equipe'."
            )

    META_EXEMPLARES, META_SKUS = 55000, 1200
    pct_exemplares = (total_exemplares / META_EXEMPLARES) if META_EXEMPLARES else 0
    pct_skus = (total_skus / META_SKUS) if META_SKUS else 0

    c1, c2 = st.columns(2)
    with c1:
        st.markdown(
            f'<div class="card-kpi" style="--accent-color:#2563EB;">'
            f'<div class="card-title">📦 Total de Exemplares</div>'
            f'<div class="card-value">{total_exemplares:,} un</div>'
            f'<div class="card-sub">Meta Diária: {META_EXEMPLARES:,} un · Atingido: {pct_exemplares:.1%}</div></div>',
            unsafe_allow_html=True,
        )
        st.progress(min(pct_exemplares, 1.0))
    with c2:
        st.markdown(
            f'<div class="card-kpi" style="--accent-color:#0D9488;">'
            f'<div class="card-title">🏷️ Total de SKU</div>'
            f'<div class="card-value">{total_skus:,}</div>'
            f'<div class="card-sub">Meta Diária: {META_SKUS:,} · Atingido: {pct_skus:.1%}</div></div>',
            unsafe_allow_html=True,
        )
        st.progress(min(pct_skus, 1.0))

    st.markdown("<br>", unsafe_allow_html=True)

    # --- Tabela gerencial individual -----------------------------------------
    data_gerencial = []
    for n in NOMES_LISTA:
        if n in remover_do_setor:
            continue

        is_ausente = n in faltas_selecionadas

        if is_ausente:
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

            # Se o operador não aparece na planilha (nenhum registro encontrado) e
            # não está marcado como ausente, ele é ignorado e não entra na tabela.
            if qtd_skus == 0:
                continue

            historico_justificativas = []
            for m in mov["movimentacoes"]:
                sai, ret, loc = m["sai"].strip(), m["ret"].strip(), m["loc"].strip()
                if not (sai or ret or loc):
                    continue  # linha em branco na tabela de movimentação: ignora

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

            justificativa_texto = (
                " ; ".join(historico_justificativas) + "." if historico_justificativas else "Atividade normal no setor."
            )

        linha = {
            "Cargo": cargo_atual,
            "Colaboradora": n,
            "Exemplares": qtd_exemplares,
            "SKUs": qtd_skus,
            "Movimentação Operacional": justificativa_texto,
        }

        # Meta individual (opcional) — só aparece se configurada para ao menos alguém.
        if METAS_INDIVIDUAIS:
            meta_pessoa = METAS_INDIVIDUAIS.get(n)
            if meta_pessoa:
                linha["Meta Individual"] = meta_pessoa
                linha["% Meta Individual"] = f"{(qtd_exemplares / meta_pessoa):.0%}"
            else:
                linha["Meta Individual"] = ""
                linha["% Meta Individual"] = ""

        data_gerencial.append(linha)

    df_real = pd.DataFrame(data_gerencial)

    st.markdown(
        "<h3 style='font-family: \"Sora\", sans-serif; color: #0F172A; font-size: 1.05rem; "
        "font-weight: 700; margin-bottom:10px;'>📋 Detalhamento Gerencial de Produtividade</h3>",
        unsafe_allow_html=True,
    )

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

    if modo_edicao:
        df_exibir_editado = st.data_editor(
            df_exibir,
            use_container_width=True,
            hide_index=True,
            num_rows="fixed",
            key="editor_tabela_gerencial",
        )
        # Reincorpora as edições feitas na grade de volta ao df_real completo,
        # preservando as colunas que estiverem ocultas no momento (ex.: se
        # Exemplares/SKUs estiverem escondidos, seus valores originais não são
        # perdidos — só as colunas visíveis são atualizadas com o que foi editado).
        for col in df_exibir_editado.columns:
            df_real[col] = df_exibir_editado[col].values
    else:
        st.markdown(renderizar_tabela_html(df_exibir), unsafe_allow_html=True)

    col_exp1, col_exp2 = st.columns(2)
    with col_exp1:
        if not df_real.empty:
            st.download_button(
                label="📥 Baixar Tabela em Excel",
                data=gerar_excel_gerencial(df_real),
                file_name=f"produtividade_{data_produtividade.strftime('%Y-%m-%d')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )

    # --- Histórico / tendência ------------------------------------------------
    st.markdown("<br><hr>", unsafe_allow_html=True)
    st.markdown(
        "<h3 style='font-family: \"Sora\", sans-serif; color: #0F172A; font-size: 1.05rem; "
        "font-weight: 700;'>📈 Evolução da Produção</h3>",
        unsafe_allow_html=True,
    )
    if not df_filtrado.empty:
        hist = salvar_historico(
            data_produtividade.strftime("%Y-%m-%d"), total_exemplares, total_skus, pct_exemplares, pct_skus
        )
        hist_recente = hist.sort_values("data").tail(30).set_index("data")
        st.line_chart(hist_recente[["exemplares", "skus"]])
        st.caption("Últimos 30 registros salvos localmente em `historico_producao.csv`.")
    else:
        st.info("Sem dados válidos no dia — o histórico não foi atualizado.")

    # --- Relatório em imagem ---------------------------------------------------
    st.markdown("<h4 style='font-family: \"Sora\", sans-serif; color: #0F172A; font-size: 0.95rem; font-weight: 700; margin-top:18px;'>🖼️ Relatório em Imagem</h4>", unsafe_allow_html=True)
    st.caption("Clique com o botão direito na imagem e escolha **Copiar imagem** para colar direto no e-mail, ou baixe o arquivo abaixo.")
    imagem_relatorio = gerar_relatorio_imagem(
        total_exemplares, total_skus, pct_exemplares, pct_skus, META_EXEMPLARES, META_SKUS, df_real
    )
    st.image(imagem_relatorio, use_container_width=True)
    nome_arquivo_imagem = f"relatorio_producao_{data_produtividade.strftime('%Y-%m-%d')}.png"
    st.download_button(
        label="📥 Baixar Relatório em Imagem",
        data=imagem_relatorio,
        file_name=nome_arquivo_imagem,
        mime="image/png",
    )

    # --- Texto do e-mail ---------------------------------------------------------
    st.markdown("<br><hr>", unsafe_allow_html=True)
    st.markdown(
        "<h3 style='font-family: \"Sora\", sans-serif; color: #0F172A; font-size: 1.05rem; "
        "font-weight: 700;'>✉️ Texto do E-mail Pronto para a Diretoria</h3>",
        unsafe_allow_html=True,
    )

    texto_final = (
        f"Boa tarde, Prezados.\n\nSegue abaixo o relatório de produção.\n"
        f"referente ao dia {data_formatada}.\n\n--------------------------------\n"
        f"Resumo Varejo.\nSKU: {total_skus}\nExemplares: {total_exemplares:,}\n"
        f"--------------------------------\n\nAtenciosamente,"
    )

    st.text_area("Selecione tudo abaixo e copie (Ctrl+A / Ctrl+C):", value=texto_final, height=200, key="texto_email")

    if st.button("📧 Enviar relatório por e-mail agora"):
        enviar_email_relatorio(
            assunto=f"Relatório de Produção — {data_formatada}",
            corpo_texto=texto_final,
            imagem_bytes=imagem_relatorio,
            nome_imagem=nome_arquivo_imagem,
        )

else:
    st.info("👋 Painel atualizado com novas melhorias. Faça o upload da sua planilha Excel na barra lateral.")
