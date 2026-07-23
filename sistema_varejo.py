import streamlit as st
import pandas as pd
import openpyxl
from datetime import datetime
import io
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
# 1. Configuração e Estilização de Design Premium (HTML / CSS)
st.set_page_config(page_title="Dashboard Executivo Varejo", layout="wide")
st.markdown("""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Sora:wght@600;700;800&family=Inter:wght@400;500;600;700&display=swap');

    html, body, [class*="css"] { font-family: 'Inter', sans-serif; }

    .stApp { background-color: #F8FAFC; }

    .block-container {padding-top: 1.2rem; padding-bottom: 0rem; max-width: 96%;}

    /* Sidebar */
    section[data-testid="stSidebar"] {
        background-color: #FFFFFF;
        border-right: 1px solid #E5E7EB;
    }
    section[data-testid="stSidebar"] h3 {
        font-family: 'Sora', sans-serif;
        color: #0F172A !important;
        font-size: 0.95rem !important;
    }

    /* KPI Cards - minimalistas, com um detalhe de cor de assinatura */
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

    /* Botões */
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

    /* Campos de texto e data */
    .stTextInput>div>div>input, .stDateInput input {
        background-color: #FFFFFF;
        color: #0F172A;
        border: 1px solid #E2E8F0;
        border-radius: 6px;
    }
    .stTextInput>div>div>input:focus { border-color: #2563EB; }

    /* Tabela */
    div[data-testid="stDataFrame"] {
        border: 1px solid #E5E7EB;
        border-radius: 10px;
        overflow: hidden;
    }
    </style>
""", unsafe_allow_html=True)

st.markdown("""
<h1 style='text-align:left; font-family: "Sora", sans-serif; font-weight:800; letter-spacing:-0.5px;
color: #0F172A; margin-bottom: 4px;'>📊 Painel Executivo de Produção</h1>
<p style='text-align:left; font-family: "Inter", sans-serif; color:#64748B; font-size:0.95rem; margin-top:0; margin-bottom:28px;'>Varejo · acompanhamento diário de produtividade</p>
""", unsafe_allow_html=True)

import matplotlib.patches as mpatches

def gerar_relatorio_imagem(data_formatada, total_exemplares, total_skus, pct_exemplares, pct_skus, meta_exemplares, meta_skus, df_real):
    """Gera uma imagem (PNG) parecida com um print da tela: cards de KPI + tabela de
    detalhamento (sem os números individuais de Exemplares/SKUs, que a diretoria não
    precisa ver), pronta para copiar (clique direito > copiar imagem) ou baixar."""
    df_relatorio = df_real[["Cargo", "Colaboradora", "Movimentação Operacional"]] if not df_real.empty else df_real

    n_linhas = max(len(df_relatorio), 1)
    altura_fig = 3.5 + 0.35 * n_linhas
    fig = plt.figure(figsize=(11, altura_fig), dpi=200)
    fig.patch.set_facecolor("#FAFAFA")

    # Cabeçalho
    fig.text(0.04, 0.975, "Painel Executivo de Produção", fontsize=18, fontweight="bold", color="#111827", va="top")
    fig.text(0.04, 0.935, f"Varejo  ·  Referente ao dia {data_formatada}", fontsize=10, color="#9CA3AF", va="top")

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
        ax.text(0.09, 0.20, sub, fontsize=8, color="#B0B7C3", va="top")

    desenhar_card(0.04, 0.44, "TOTAL DE EXEMPLARES", f"{total_exemplares:,} un",
                  f"Meta Diária: {meta_exemplares:,} un  ·  Atingido: {pct_exemplares:.1%}", "#2563EB")
    desenhar_card(0.52, 0.44, "TOTAL DE SKU", f"{total_skus:,}",
                  f"Meta Diária: {meta_skus:,}  ·  Atingido: {pct_skus:.1%}", "#0D9488")

    # Tabela de detalhamento (sem colunas de produção individual)
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

def gerar_texto_detalhamento(df_real):
    """Monta o detalhamento por colaborador em texto simples, pronto para copiar e colar
    (por e-mail, WhatsApp, etc.), sem precisar tirar print da tela."""
    if df_real.empty:
        return "Nenhum dado disponível."
    linhas = []
    for _, row in df_real.iterrows():
        linhas.append(
            f"{row['Colaboradora']} ({row['Cargo']}): {row['Exemplares']} exemplares | {row['SKUs']} SKUs — {row['Movimentação Operacional']}"
        )
    return "\n".join(linhas)

# Organização da equipe por cargos oficiais
EQUIPE = {
    "Líder": ["Kamila Moraes", "Beatriz Alcantara"],
    "Apoio": ["Alisson Lima"],
    "Operador(a)": ["Rosana Delfino", "Ana Caroline", "Karoline Gonçalves", "Gabriele", "Beatriz Mascarenhas", "Graziela Pereira", "Paula Roberta", "Weliton", "Ellen Kelly"]
}
NOMES_LISTA = EQUIPE["Líder"] + EQUIPE["Apoio"] + EQUIPE["Operador(a)"]

# Mapeamento de nomes que aparecem de forma diferente (abreviada) na planilha Excel.
# A chave é o nome de exibição usado no painel; o valor é como o nome aparece na coluna USUARIO do Excel.
# Por padrão, o próprio nome (em maiúsculas) é usado para o cruzamento dos dados.
ALIAS_EXCEL = {
    "Beatriz Alcantara": "BEATRIZ",
    "Graziela Pereira": "GRAZIELA PEREIRA DO NASCIMENTO",
    "Paula Roberta": "PAULA ROBERTA SANTOS DA SILVA",
}

import unicodedata

def normalizar(texto):
    """Remove acentos, espaços extras e padroniza para maiúsculas, evitando falhas de
    correspondência entre o nome cadastrado no painel e o nome como aparece na planilha."""
    texto = str(texto).strip().upper()
    texto = unicodedata.normalize("NFKD", texto).encode("ASCII", "ignore").decode("ASCII")
    texto = " ".join(texto.split())
    return texto

def nome_excel(nome):
    return normalizar(ALIAS_EXCEL.get(nome, nome))

# 2. Barra Lateral de Controle Unificada
st.sidebar.header("🛠️ Controle Operacional")
uploaded_file = st.sidebar.file_uploader("Upload da Planilha Excel", type=["xlsx"])

# NOVO: Seleção da data da produtividade para atualizar o e-mail automaticamente
data_produtividade = st.sidebar.date_input("Data da Produtividade:", datetime.now())
data_formatada = data_produtividade.strftime("%d/%m")

# Filtros gerenciais limpos
st.sidebar.markdown("### 👁️ Filtros Gerenciais")
remover_do_setor = st.sidebar.multiselect("Ocultar do Setor (Tabela):", NOMES_LISTA)

st.sidebar.markdown("### ❌ Ausências do Dia")
faltas_selecionadas = st.sidebar.multiselect("Selecione quem faltou hoje:", NOMES_LISTA)

st.sidebar.markdown("### ⏳ Movimentação de Horários")
movimentados_selecionados = st.sidebar.multiselect(
    "🚚 Quem foi movimentado(a) hoje?",
    [n for n in NOMES_LISTA if n not in remover_do_setor and n not in faltas_selecionadas]
)

dict_movimentacao = {}
dict_motivos_falta = {}

for cargo, integrantes in EQUIPE.items():
    integrantes_visiveis = [i for i in integrantes if i not in remover_do_setor]
    if integrantes_visiveis:
        st.sidebar.markdown(f"<h3 style='color:#1E3A8A; margin-top:10px; font-size:1.1rem;'>🔹 {cargo.upper()}</h3>", unsafe_allow_html=True)
        
    for op in integrantes:
        if op in remover_do_setor:
            continue
            
        is_ausente = op in faltas_selecionadas
        is_movimentado = op in movimentados_selecionados
        
        if is_ausente:
            st.sidebar.markdown(f"❌ **{op} (AUSENTE)**")
            dict_motivos_falta[op] = st.sidebar.text_input(f"Motivo da falta de {op}:", value="Falta administrativa", key=f"mot_falta_{op}")
            dict_movimentacao[op] = {"cargo": cargo, "movimentacoes": []}
        elif is_movimentado:
            st.sidebar.markdown(f"**👤 {op}**", unsafe_allow_html=True)

            if f"num_mov_{op}" not in st.session_state:
                st.session_state[f"num_mov_{op}"] = 1  # começa com 1 saída, e cresce com o botão "+"

            movimentacoes_op = []
            for idx in range(1, st.session_state[f"num_mov_{op}"] + 1):
                st.sidebar.markdown(f"<span style='font-size:0.8rem; color:gray;'>Saída {idx}:</span>", unsafe_allow_html=True)
                c_sai, c_ret, c_loc = st.sidebar.columns(3)

                # Valores padrão (apenas na primeira saída, pra manter o comportamento de antes)
                init_sai, init_ret, init_loc = "", "", ""
                if idx == 1:
                    init_sai = "06h15" if op in ["Ana Caroline", "Rosana Delfino", "Karoline Gonçalves", "Gabriele"] else ""
                    init_ret = "10h00" if op == "Gabriele" else ("10h30" if op in ["Ana Caroline", "Karoline Gonçalves"] else ("07h30" if op == "Rosana Delfino" else ""))
                    init_loc = "Setor Loja" if op in ["Ana Caroline", "Rosana Delfino", "Karoline Gonçalves", "Gabriele"] else ""

                with c_sai: sai = st.text_input("Saída:", value=init_sai, key=f"sai_{op}_{idx}")
                with c_ret: ret = st.text_input("Retorno:", value=init_ret, key=f"ret_{op}_{idx}")
                with c_loc: loc = st.text_input("Local:", value=init_loc, key=f"loc_{op}_{idx}")

                movimentacoes_op.append({"sai": sai, "ret": ret, "loc": loc})

            c_add, c_rem = st.sidebar.columns(2)
            with c_add:
                if st.button("➕ Adicionar", key=f"add_mov_{op}", use_container_width=True):
                    st.session_state[f"num_mov_{op}"] += 1
                    st.rerun()
            with c_rem:
                if st.session_state[f"num_mov_{op}"] > 1:
                    if st.button("🗑️ Excluir", key=f"rem_mov_{op}", use_container_width=True):
                        idx_remover = st.session_state[f"num_mov_{op}"]
                        for campo in ["sai", "ret", "loc"]:
                            st.session_state.pop(f"{campo}_{op}_{idx_remover}", None)
                        st.session_state[f"num_mov_{op}"] -= 1
                        st.rerun()

            dict_movimentacao[op] = {"cargo": cargo, "movimentacoes": movimentacoes_op}
        else:
            st.sidebar.markdown(f"👤 {op} <span style='font-size:0.8rem; color:gray;'>(sem movimentação)</span>", unsafe_allow_html=True)
            dict_movimentacao[op] = {"cargo": cargo, "movimentacoes": []}
        st.sidebar.markdown("<hr style='margin:6px 0px; border-color: #D1D5DB;'>", unsafe_allow_html=True)

# 3. Lógica: Lendo Apenas Linhas Visíveis (Filtradas) do Excel
if uploaded_file:
    wb = openpyxl.load_workbook(uploaded_file, data_only=True)
    sheet = wb.active
    
    dados_visiveis = []
    for row in range(2, sheet.max_row + 1):
        if sheet.row_dimensions[row].hidden == False:
            val_i = sheet.cell(row=row, column=9).value   # Coluna I (TOTAL)
            val_m = sheet.cell(row=row, column=13).value  # Coluna M (USUARIO)
            
            if val_i is not None and val_m is not None:
                dados_visiveis.append({"TOTAL": val_i, "USUARIO": normalizar(val_m)})
                
    if dados_visiveis:
        df_filtrado = pd.DataFrame(dados_visiveis)
        df_filtrado["TOTAL"] = pd.to_numeric(df_filtrado["TOTAL"], errors='coerce').fillna(0)
        total_exemplares = int(df_filtrado["TOTAL"].sum())
        total_skus = int(len(df_filtrado))
    else:
        total_exemplares, total_skus = 50271, 1104
        df_filtrado = pd.DataFrame(columns=["TOTAL", "USUARIO"])

    META_EXEMPLARES, META_SKUS = 55000, 1200
    pct_exemplares = (total_exemplares / META_EXEMPLARES)
    pct_skus = (total_skus / META_SKUS)
    
    c1, c2 = st.columns(2)
    with c1:
        st.markdown(f'<div class="card-kpi" style="--accent-color:#2563EB;"><div class="card-title">📦 Total de Exemplares</div><div class="card-value">{total_exemplares:,} un</div><div class="card-sub">Meta Diária: {META_EXEMPLARES:,} un · Atingido: {pct_exemplares:.1%}</div></div>', unsafe_allow_html=True)
        st.progress(min(pct_exemplares, 1.0))
    with c2:
        st.markdown(f'<div class="card-kpi" style="--accent-color:#0D9488;"><div class="card-title">🏷️ Total de SKU</div><div class="card-value">{total_skus:,}</div><div class="card-sub">Meta Diária: {META_SKUS:,} · Atingido: {pct_skus:.1%}</div></div>', unsafe_allow_html=True)
        st.progress(min(pct_skus, 1.0))
        
    st.markdown("<br>", unsafe_allow_html=True)

    # Processamento dos indicadores individuais
    data_gerencial = []
    for n in NOMES_LISTA:
        if n in remover_do_setor:
            continue
            
        is_ausente = n in faltas_selecionadas
        
        if is_ausente:
            qtd_exemplares = 0
            qtd_skus = 0
            motivo_individual = dict_motivos_falta.get(n, "Falta administrativa")
            justificativa_texto = f"Ausente. Motivo: {motivo_individual}."
            cargo_atual = "Operador(a)"
            for c, ints in EQUIPE.items():
                if n in ints: cargo_atual = c
        else:
            mov = dict_movimentacao[n]
            cargo_atual = mov["cargo"]
            if not df_filtrado.empty:
                df_func = df_filtrado[df_filtrado["USUARIO"] == nome_excel(n)]
                qtd_exemplares = int(df_func["TOTAL"].sum())
                qtd_skus = int(len(df_func))
            else:
                qtd_exemplares, qtd_skus = 0, 0

            # Se o operador não aparece na planilha (nenhum registro encontrado) e não está
            # marcado como ausente, ele é ignorado e não entra na tabela gerencial.
            if qtd_skus == 0:
                continue
                
            historico_justificativas = []
            for idx, m in enumerate(mov["movimentacoes"]):
                if m["sai"].strip() != "" and m["sai"].strip().upper() != "N/A":
                    prefixo = "Encaminhada" if idx == 0 else "encaminhada"
                    historico_justificativas.append(f"{prefixo} ao {m['loc']} das {m['sai']} às {m['ret']}")
                
            justificativa_texto = " ; ".join(historico_justificativas) + "." if historico_justificativas else "Atividade normal no setor."
            
        data_gerencial.append({
            "Cargo": cargo_atual,
            "Colaboradora": n,
            "Exemplares": qtd_exemplares,
            "SKUs": qtd_skus,
            "Movimentação Operacional": justificativa_texto
        })
        
    df_real = pd.DataFrame(data_gerencial)
    st.markdown("<h3 style='font-family: \"Sora\", sans-serif; color: #0F172A; font-size: 1.05rem; font-weight: 700; margin-bottom:10px;'>📋 Detalhamento Gerencial de Produtividade</h3>", unsafe_allow_html=True)
    st.dataframe(df_real, use_container_width=True, hide_index=True)

    st.markdown("<h4 style='font-family: \"Sora\", sans-serif; color: #0F172A; font-size: 0.95rem; font-weight: 700; margin-top:18px;'>🖼️ Relatório em Imagem</h4>", unsafe_allow_html=True)
    st.caption("Clique com o botão direito na imagem e escolha **Copiar imagem** para colar direto no e-mail, ou baixe o arquivo abaixo.")
    imagem_relatorio = gerar_relatorio_imagem(
        data_formatada, total_exemplares, total_skus,
        pct_exemplares, pct_skus, META_EXEMPLARES, META_SKUS, df_real
    )
    st.image(imagem_relatorio, use_container_width=True)
    st.download_button(
        label="📥 Baixar Relatório em Imagem",
        data=imagem_relatorio,
        file_name=f"relatorio_producao_{data_produtividade.strftime('%Y-%m-%d')}.png",
        mime="image/png",
    )

    # 4. Caixa de Texto Gerada do E-mail Padronizado Conforme Solicitado
    st.markdown("<br><hr>", unsafe_allow_html=True)
    st.markdown("<h3 style='font-family: \"Sora\", sans-serif; color: #0F172A; font-size: 1.05rem; font-weight: 700;'>✉️ Texto do E-mail Pronto para a Diretoria</h3>", unsafe_allow_html=True)
    
    texto_final = f"Boa tarde, Prezados.\n\nSegue abaixo o relatório de produção.\nreferente ao dia {data_formatada}.\n\n--------------------------------\nResumo Varejo.\nSKU: {total_skus}\nExemplares: {total_exemplares:,}\n--------------------------------\n\nAtenciosamente,"
    
    st.text_area("Selecione tudo abaixo e copie (Ctrl+A / Ctrl+C):", value=texto_final, height=240)
else:
    st.info("👋 Padrão de e-mail atualizado. Faça o upload da sua planilha Excel na barra lateral.")
