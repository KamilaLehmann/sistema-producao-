import streamlit as st
import pandas as pd
import openpyxl
from datetime import datetime

# 1. Configuração e Estilização de Design Premium (HTML / CSS)
st.set_page_config(page_title="Dashboard Executivo Varejo", layout="wide")
st.markdown("""
    <style>
    .block-container {padding-top: 1rem; padding-bottom: 0rem; max-width: 96%;}
    .card-kpi {
        background: linear-gradient(135deg, #1E3A8A 0%, #3B82F6 100%);
        color: white;
        padding: 18px;
        border-radius: 12px;
        box-shadow: 0 4px 6px -1px rgba(0,0,0,0.1);
        text-align: center;
        margin-bottom: 12px;
    }
    .card-title { font-size: 1rem; font-weight: 700; opacity: 0.95; margin-bottom: 4px; letter-spacing: 0.5px; }
    .card-value { font-size: 2.2rem; font-weight: 800; line-height: 1; margin-bottom: 6px; color: #FFFFFF; }
    .card-sub { font-size: 0.85rem; opacity: 0.9; font-weight: 600; }
    div.stProgress > div > div > div { background-color: #10B981; height: 8px; }
    </style>
""", unsafe_allow_html=True)

st.markdown("<h1 style='text-align: center; color: #1E3A8A; font-weight: 800; margin-bottom: 25px;'>📊 Painel Executivo de Produção - Varejo</h1>", unsafe_allow_html=True)

# Lista Base Fixa da Equipe Atual
EQUIPE_BASE = {
    "Líder": ["Kamila Moraes"],
    "Apoio": ["Alisson Lima"],
    "Operadoras": ["Rosana Delfino", "Anacaroline", "Karoline Gonçalves", "Gabriele", "Beatriz Mascarenhas"]
}
NOMES_BASE = EQUIPE_BASE["Líder"] + EQUIPE_BASE["Apoio"] + EQUIPE_BASE["Operadoras"]

# 2. Barra Lateral de Controle
st.sidebar.header("🛠️ Controle Operacional")
uploaded_file = st.sidebar.file_uploader("Upload da Planilha Excel", type=["xlsx"])

# Inicializa a lista de nomes que vai rodar no dia
nomes_do_dia = NOMES_BASE.copy()
novos_nomes = []

# Identificação e captura automática de novas pessoas na planilha
if uploaded_file:
    try:
        wb_temp = openpyxl.load_workbook(uploaded_file, data_only=True)
        sheet_temp = wb_temp.active
        usuarios_identificados = set()
        
        for r in range(2, sheet_temp.max_row + 1):
            if sheet_temp.row_dimensions[r].hidden == False:
                u = sheet_temp.cell(row=r, column=13).value
                if u:
                    usuarios_identificados.add(str(u).strip())
                    
        for user in usuarios_identificados:
            if not any(user.upper() == n.upper() for n in NOMES_BASE):
                if user not in novos_nomes:
                    novos_nomes.append(user)
        nomes_do_dia += novos_nomes
    except:
        pass

# Data da Produtividade
data_produtividade = st.sidebar.date_input("Data da Produtividade:", datetime.now())
data_formatada = data_produtividade.strftime("%d/%m")

# Filtros gerenciais
st.sidebar.markdown("### 👁️ Filtros Gerenciais")
remover_do_setor = st.sidebar.multiselect("Ocultar do Setor (Tabela):", nomes_do_dia)
faltas_selecionadas = st.sidebar.multiselect("Selecione quem faltou hoje:", nomes_do_dia)

st.sidebar.markdown("### ⏳ Movimentação de Horários")
dict_movimentacao = {}
dict_motivos_falta = {}

equipe_lateral = EQUIPE_BASE.copy()
if novos_nomes:
    equipe_lateral["Novos Integrantes"] = novos_nomes

for cargo, integrantes in equipe_lateral.items():
    integrantes_visiveis = [i for i in integrantes if i not in remover_do_setor]
    if integrantes_visiveis:
        st.sidebar.markdown(f"<h3 style='color:#1E3A8A; margin-top:10px; font-size:1.1rem;'>🔹 {cargo.upper()}</h3>", unsafe_allow_html=True)
        
    for op in integrantes:
        if op in remover_do_setor:
            continue
            
        is_ausente = op in faltas_selecionadas
        
        if is_ausente:
            st.sidebar.markdown(f"❌ **{op} (AUSENTE)**")
            dict_motivos_falta[op] = st.sidebar.text_input(f"Motivo da falta de {op}:", value="Falta administrativa", key=f"mot_falta_{op}")
            dict_movimentacao[op] = {"sai1": "", "ret1": "", "loc1": "", "sai2": "", "ret2": "", "loc2": "", "cargo": cargo}
        else:
            st.sidebar.markdown(f"**👤 {op}**", unsafe_allow_html=True)
            st.sidebar.markdown("<span style='font-size:0.8rem; color:gray;'>Primeira Saída:</span>", unsafe_allow_html=True)
            c_sai1, c_ret1, c_loc1 = st.sidebar.columns(3)
            
            init_sai1 = "06h15" if op in ["Anacaroline", "Rosana Delfino", "Karoline Gonçalves", "Gabriele"] else ""
            init_ret1 = "10h00" if op == "Gabriele" else ("10h30" if op in ["Anacaroline", "Karoline Gonçalves"] else ("07h30" if op == "Rosana Delfino" else ""))
            init_loc1 = "Setor Loja" if op in ["Anacaroline", "Rosana Delfino", "Karoline Gonçalves", "Gabriele"] else ""
            
            with c_sai1: sai1 = st.text_input("Saída 1:", value=init_sai1, key=f"sai1_{op}")
            with c_ret1: ret1 = st.text_input("Retorno 1:", value=init_ret1, key=f"ret1_{op}")
            with c_loc1: loc1 = st.text_input("Local 1:", value=init_loc1, key=f"loc1_{op}")
            
            st.sidebar.markdown("<span style='font-size:0.8rem; color:gray;'>Segunda Saída (Se houver):</span>", unsafe_allow_html=True)
            c_sai2, c_ret2, c_loc2 = st.sidebar.columns(3)
            with c_sai2: sai2 = st.text_input("Saída 2:", value="", key=f"sai2_{op}")
            with c_ret2: ret2 = st.text_input("Retorno 2:", value="", key=f"ret2_{op}")
            with c_loc2: loc2 = st.text_input("Local 2:", value="", key=f"loc2_{op}")
                
            dict_movimentacao[op] = {
                "sai1": sai1, "ret1": ret1, "loc1": loc1,
                "sai2": sai2, "ret2": ret2, "loc2": loc2,
                "cargo": cargo
            }
        st.sidebar.markdown("<hr style='margin:6px 0px; border-color: #D1D5DB;'>", unsafe_allow_html=True)

# 3. Lógica principal
if uploaded_file:
    wb = openpyxl.load_workbook(uploaded_file, data_only=True)
    sheet = wb.active
    
    dados_visiveis = []
    for row in range(2, sheet.max_row + 1):
        if sheet.row_dimensions[row].hidden == False:
            val_i = sheet.cell(row=row, column=9).value   # Coluna I (TOTAL)
            val_m = sheet.cell(row=row, column=13).value  # Coluna M (USUARIO)
            
            if val_i is not None and val_m is not None:
                dados_visiveis.append({"TOTAL": val_i, "USUARIO": str(val_m).strip()})
                
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
    
    # CORREÇÃO: Títulos limpos sem menção a colunas
    c1, c2 = st.columns(2)
    with c1:
        st.markdown(f'<div class="card-kpi"><div class="card-title">📦 TOTAL DE EXEMPLARES</div><div class="card-value">{total_exemplares:,} un</div><div class="card-sub">Meta Diária: {META_EXEMPLARES:,} un | Atingido: {pct_exemplares:.1%}</div></div>', unsafe_allow_html=True)
        st.progress(min(pct_exemplares, 1.0))
    with c2:
        st.markdown(f'<div class="card-kpi" style="background: linear-gradient(135deg, #0F766E 0%, #14B8A6 100%);"><div class="card-title">🏷️ SKUs FILTRADOS</div><div class="card-value">{total_skus:,}</div><div class="card-sub">Meta Diária: {META_SKUS:,} | Atingido: {pct_skus:.1%}</div></div>', unsafe_allow_html=True)
        st.progress(min(pct_skus, 1.0))
        
    st.markdown("<br>", unsafe_allow_html=True)

    # Processamento individual
    data_gerencial = []
    for n in nomes_do_dia:
        if n in remover_do_setor:
            continue
            
        is_ausente = n in faltas_selecionadas
        
        if is_ausente:
            qtd_exemplares = 0
            qtd_skus = 0
            motivo_individual = dict_motivos_falta.get(n, "Falta administrativa")
            justificativa_texto = f"Ausente. Motivo: {motivo_individual}."
            cargo_atual = "Operadoras"
            for c, ints in EQUIPE_BASE["Líder"]: 
                if n in ints: cargo_atual = "Líder"
            for c, ints in EQUIPE_BASE["Apoio"]: 
                if n in ints: cargo_atual = "Apoio"
            if n in novos_nomes: cargo_atual = "Novos Integrantes"
        else:
            mov = dict_movimentacao.get(n, {"sai1": "", "ret1": "", "loc1": "", "sai2": "", "ret2": "", "loc2": "", "cargo": "Operadoras"})
            cargo_atual = mov.get("cargo", "Novos Integrantes") if n not in novos_nomes else "Novos Integrantes"
            
            if not df_filtrado.empty:
                df_func = df_filtrado[df_filtrado["USUARIO"].str.upper() == n.upper()]
                qtd_exemplares = int(df_func["TOTAL"].sum())
                qtd_skus = int(len(df_func))
            else:
                qtd_exemplares, qtd_skus = 0, 0
                
            historico_justificativas = []
            if mov["sai1"].strip() != "" and mov["sai1"].strip().upper() != "N/A":
                historico_justificativas.append(f"Encaminhada ao {mov['loc1']} das {mov['sai1']} às {mov['ret1']}")
            if mov["sai2"].strip() != "" and mov["sai2"].strip().upper() != "N/A":
                historico_justificativas.append(f"encaminhada ao {mov['loc2']} das {mov['sai2']} às {mov['ret2']}")
                
            justificativa_texto = " ; ".join(historico_justificativas) + "." if historico_justificativas else "Atividade normal no setor."
            
        data_gerencial.append({
            "Cargo": cargo_atual,
            "Colaboradora": n,
            "Exemplares": qtd_exemplares,
            "SKUs": qtd_skus,
            "Movimentação Operacional": justificativa_texto
        })
        
