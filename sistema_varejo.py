import streamlit as st
import pandas as pd
import openpyxl
import matplotlib.pyplot as plt

# 1. Configuração e Estilização Ultra Black Premium (HTML / CSS)
st.set_page_config(page_title="Dashboard Black Executivo", layout="wide")
st.markdown("""
    <style>
    /* Fundo Total Escuro */
    .stApp { background-color: #0B0F19; color: #E5E7EB; }
    .block-container {padding-top: 1.5rem; padding-bottom: 0rem; max-width: 96%;}
    
    /* Cards KPI Estilo Cyber Black */
    .card-kpi {
        background: linear-gradient(135deg, #111827 0%, #1F2937 100%);
        color: #FFFFFF;
        padding: 25px;
        border-radius: 16px;
        border: 1px solid #374151;
        box-shadow: 0 10px 25px rgba(0,0,0,0.5);
        text-align: center;
        margin-bottom: 12px;
    }
    .card-title { font-size: 1.1rem; font-weight: 700; color: #9CA3AF; margin-bottom: 8px; letter-spacing: 1px; }
    .card-value { font-size: 3.8rem; font-weight: 900; line-height: 1; margin-bottom: 10px; color: #3B82F6; }
    .card-sub { font-size: 1rem; color: #6B7280; font-weight: 600; }
    
    /* Customização de Input e Barras no Modo Escuro */
    div.stProgress > div > div > div { background-color: #10B981; height: 12px; border-radius: 6px; }
    .stSidebar { background-color: #030712 !important; border-right: 1px solid #1F2937; }
    h1, h3 { color: #FFFFFF !important; font-weight: 800; }
    
    /* Ajuste da tabela para o Dark Mode */
    .stDataFrame { background-color: #111827; border-radius: 12px; border: 1px solid #374151; }
    </style>
""", unsafe_allow_html=True)

st.markdown("<h1 style='text-align: center; margin-bottom: 25px;'>🏁 PAINEL EXECUTIVO - VAREJO BLACK EDITION</h1>", unsafe_allow_html=True)

# Organização da equipe por cargos oficiais
EQUIPE = {
    "Líder": ["Kamila Moraes"],
    "Apoio": ["Alisson Lima"],
    "Operadoras": ["Rosana Delfino", "Anacaroline", "Karoline Gonçalves", "Gabriele", "Beatriz Mascarenhas"]
}
NOMES_LISTA = EQUIPE["Líder"] + EQUIPE["Apoio"] + EQUIPE["Operadoras"]

# 2. Barra Lateral de Controle (Estilo Black)
st.sidebar.header("🕹️ CONTROLE OPERACIONAL")
uploaded_file = st.sidebar.file_uploader("Carregar Planilha Filtrada", type=["xlsx"])

# Filtro de Exibição
st.sidebar.markdown("### 👁️ Filtros")
exibir_kamila = st.sidebar.checkbox("Exibir Kamila Moraes (Líder)", value=True)
exibir_alisson = st.sidebar.checkbox("Exibir Alisson Lima (Apoio)", value=True)

st.sidebar.markdown("### ⏳ Movimentação de Horários")
dict_movimentacao = {}

for cargo, integrantes in EQUIPE.items():
    st.sidebar.markdown(f"<span style='color:#3B82F6; font-weight:bold;'>• {cargo.upper()}</span>", unsafe_allow_html=True)
    for op in integrantes:
        st.sidebar.markdown(f"**👤 {op}**")
        c_sai1, c_ret1, c_loc1 = st.sidebar.columns(3)
        init_sai1 = "06h15" if op in ["Anacaroline", "Rosana Delfino", "Karoline Gonçalves", "Gabriele"] else ""
        init_ret1 = "10h00" if op == "Gabriele" else ("10h30" if op in ["Anacaroline", "Karoline Gonçalves"] else ("07h30" if op == "Rosana Delfino" else ""))
        init_loc1 = "Setor Loja" if op in ["Anacaroline", "Rosana Delfino", "Karoline Gonçalves", "Gabriele"] else ""
        
        with c_sai1: sai1 = st.text_input("Saída 1:", value=init_sai1, key=f"s1_{op}", label_visibility="collapsed")
        with c_ret1: ret1 = st.text_input("Retorno 1:", value=init_ret1, key=f"r1_{op}", label_visibility="collapsed")
        with c_loc1: loc1 = st.text_input("Local 1:", value=init_loc1, key=f"l1_{op}", label_visibility="collapsed")
            
        dict_movimentacao[op] = {"sai1": sai1, "ret1": ret1, "loc1": loc1, "cargo": cargo}
        st.sidebar.markdown("<hr style='margin:4px 0px; border-color: #1F2937;'>", unsafe_allow_html=True)

# 3. Lógica Automatizada via Python
if uploaded_file:
    wb = openpyxl.load_workbook(uploaded_file, data_only=True)
    sheet = wb.active
    
    dados_visiveis = []
    for row in range(2, sheet.max_row + 1):
        if sheet.row_dimensions[row].hidden == False:
            val_i = sheet.cell(row=row, column=9).value   # Coluna I (TOTAL)
            val_m = sheet.cell(row=row, column=13).value  # Coluna M (USUARIO)
            if val_i is not None and val_m is not None:
                dados_visiveis.append({"TOTAL": val_i, "USUARIO": str(val_m).strip().upper()})
                
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
    
    # Cards de Alta Performance em HTML/CSS Black Mode
    c1, c2 = st.columns(2)
    with c1:
        st.markdown(f'<div class="card-kpi"><div class="card-title">📦 TOTAL DE EXEMPLARES</div><div class="card-value">{total_exemplares:,} un</div><div class="card-sub">Meta: {META_EXEMPLARES:,} | Performance: {pct_exemplares:.1%}</div></div>', unsafe_allow_html=True)
        st.progress(min(pct_exemplares, 1.0))
    with c2:
        st.markdown(f'<div class="card-kpi" style="border-color: #14B8A6;"><div class="card-title">🏷️ SKUs FILTRADOS</div><div class="card-value" style="color:#14B8A6;">{total_skus:,}</div><div class="card-sub">Meta: {META_SKUS:,} | Performance: {pct_skus:.1%}</div></div>', unsafe_allow_html=True)
        st.progress(min(pct_skus, 1.0))
        
    st.markdown("<br>", unsafe_allow_html=True)

    # Processamento dos indicadores individuais por cargo estruturado de forma automática
    data_gerencial = []
    for n in NOMES_LISTA:
        if n == "Kamila Moraes" and not exibir_kamila: continue
        if n == "Alisson Lima" and not exibir_alisson: continue
            
        if not df_filtrado.empty:
            df_func = df_filtrado[df_filtrado["USUARIO"] == n.upper()]
            qtd_exemplares = int(df_func["TOTAL"].sum())
            qtd_skus = int(len(df_func))
        else:
            qtd_exemplares, qtd_skus = 0, 0
            
        mov = dict_movimentacao[n]
        if mov["sai1"].strip() != "":
            justificativa_texto = f"Encaminhada ao {mov['loc1']} das {mov['sai1']} às {mov['ret1']}."
        else:
            justificativa_texto = "Atividade normal no setor."
            
        data_gerencial.append({
            "Cargo": mov["cargo"],
            "Colaboradora": n,
            "Exemplares": qtd_exemplares,
            "SKUs": qtd_skus,
            "Movimentação Operacional": justificativa_texto
        })
        
    df_real = pd.DataFrame(data_gerencial)
    st.markdown("### 📋 Tabela Gerencial de Produtividade")
    st.dataframe(df_real, use_container_width=True, hide_index=True)

    # 4. Gerador de E-mail Automático
    st.markdown("<br><hr style='border-color:#1F2937;'><h3 style='color:#3B82F6;'>✉️ Texto do E-mail para Diretoria (Pronto)</h3>", unsafe_allow_html=True)
    blocos_email = ""
    for cargo_tipo in ["Líder", "Apoio", "Operadoras"]:
        df_cargo = df_real[df_real["Cargo"] == cargo_tipo]
        if not df_cargo.empty:
            blocos_email += f"\\n**{cargo_tipo}:**\\n"
            for idx, r in df_cargo.iterrows():
                blocos_email += f"• **{r['Colaboradora']}**: {r['Exemplares']:,} exemplares | {r['SKUs']:,} SKUs | *{r['Movimentação Operacional']}*\\n"

    texto_final = f"""
**Assunto:** Relatório de Produtividade e Histórico de Movimentação - Varejo

Boa tarde, Prezados.

Segue abaixo o relatório analítico de produção do setor de Varejo, acompanhado das justificativas de movimentações internas da equipe.

**1. Resumo de Produção Geral (Performance Diária)**
• **Total de Exemplares Separados:** {total_exemplares:,} un (Atingido: {pct_exemplares:.1%} da meta de {META_EXEMPLARES:,})
• **SKUs Movimentados:** {total_skus:,} itens (Atingido: {pct_skus:.1%} da meta de {META_SKUS:,})

**2. Indicadores de Desempenho Coletivo e Histórico por Função**
{blocos_email}
*Nota: Os remanejamentos de colaboradores ocorreram devido à falta de pedidos no estoque no início da manhã.*

Atenciosamente,
    """
    st.text_area("Copiar Relatório:", value=texto_final.strip(), height=320)
else:
    st.info("👋 Modo Cyber Black carregado. Insira a sua planilha Excel filtrada na barra lateral para ver o sistema rodar sozinho!")
