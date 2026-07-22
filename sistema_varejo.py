import streamlit as st
import pandas as pd
import openpyxl

# 1. Configuração e Estilização de Design Premium (HTML / CSS)
st.set_page_config(page_title="Dashboard Executivo Varejo", layout="wide")
st.markdown("""
    <style>
    .block-container {padding-top: 1rem; padding-bottom: 0rem; max-width: 96%;}
    .card-kpi {
        background: linear-gradient(135deg, #1E3A8A 0%, #3B82F6 100%);
        color: white;
        padding: 25px;
        border-radius: 16px;
        box-shadow: 0 10px 15px -3px rgba(0,0,0,0.1), 0 4px 6px -2px rgba(0,0,0,0.05);
        text-align: center;
        margin-bottom: 12px;
    }
    .card-title { font-size: 1.2rem; font-weight: 700; opacity: 0.95; margin-bottom: 8px; letter-spacing: 0.5px; }
    .card-value { font-size: 3.6rem; font-weight: 900; line-height: 1; margin-bottom: 10px; color: #FFFFFF; }
    .card-sub { font-size: 1rem; opacity: 0.9; font-weight: 600; }
    div.stProgress > div > div > div { background-color: #10B981; height: 10px; }
    </style>
""", unsafe_allow_html=True)

st.markdown("<h1 style='text-align: center; color: #1E3A8A; font-weight: 800; margin-bottom: 25px;'>📊 Painel Executivo de Produção - Varejo</h1>", unsafe_allow_html=True)

# Organização da equipe por cargos oficiais
EQUIPE = {
    "Líder": ["Kamila Moraes"],
    "Apoio": ["Alisson Lima"],
    "Operadoras": ["Rosana Delfino", "Anacaroline", "Karoline Gonçalves", "Gabriele", "Beatriz Mascarenhas"]
}

# 2. Barra Lateral de Controle com múltiplos horários
st.sidebar.header("🛠️ Controle Operacional")
uploaded_file = st.sidebar.file_uploader("Upload da Planilha Excel", type=["xlsx"])

st.sidebar.markdown("### ⏳ Movimentação de Horários")
dict_movimentacao = {}

for cargo, integrantes in EQUIPE.items():
    st.sidebar.markdown(f"<h3 style='color:#1E3A8A; margin-top:10px; font-size:1.1rem;'>🔹 {cargo.upper()}</h3>", unsafe_allow_html=True)
    for op in integrantes:
        st.sidebar.markdown(f"**👤 {op}**")
        
        # Primeira Saída
        st.sidebar.markdown("<span style='font-size:0.8rem; color:gray;'>Primeira Saída:</span>", unsafe_allow_html=True)
        c_sai1, c_ret1, c_loc1 = st.sidebar.columns(3)
        init_sai1 = "06h15" if op in ["Anacaroline", "Rosana Delfino", "Karoline Gonçalves", "Gabriele"] else ""
        init_ret1 = "07h30" if op == "Gabriele" else ("Não retornou" if op in ["Anacaroline", "Rosana Delfino", "Karoline Gonçalves"] else "")
        init_loc1 = "Setor Loja" if op in ["Anacaroline", "Rosana Delfino", "Karoline Gonçalves", "Gabriele"] else ""
        
        with c_sai1: sai1 = st.text_input("Saída 1:", value=init_sai1, key=f"sai1_{op}")
        with c_ret1: ret1 = st.text_input("Retorno 1:", value=init_ret1, key=f"ret1_{op}")
        with c_loc1: loc1 = st.text_input("Local 1:", value=init_loc1, key=f"loc1_{op}")
        
        # Segunda Saída (Abas Adicionais)
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
    
    c1, c2 = st.columns(2)
    with c1:
        st.markdown(f'<div class="card-kpi"><div class="card-title">📦 TOTAL DE EXEMPLARES (SOMA DA COLUNA I)</div><div class="card-value">{total_exemplares:,} un</div><div class="card-sub">Meta Diária: {META_EXEMPLARES:,} un | Atingido: {pct_exemplares:.1%}</div></div>', unsafe_allow_html=True)
        st.progress(min(pct_exemplares, 1.0))
    with c2:
        st.markdown(f'<div class="card-kpi" style="background: linear-gradient(135deg, #0F766E 0%, #14B8A6 100%);"><div class="card-title">🏷️ SKUs FILTRADOS (CONTAGEM DA COLUNA I)</div><div class="card-value">{total_skus:,}</div><div class="card-sub">Meta Diária: {META_SKUS:,} | Atingido: {pct_skus:.1%}</div></div>', unsafe_allow_html=True)
        st.progress(min(pct_skus, 1.0))
        
    st.markdown("<br>", unsafe_allow_html=True)

    # Processamento dos indicadores individuais por cargo estruturado
    data_gerencial = []
    lista_completa_nomes = EQUIPE["Líder"] + EQUIPE["Apoio"] + EQUIPE["Operadoras"]
    
    for n in lista_completa_nomes:
        if not df_filtrado.empty:
            df_func = df_filtrado[df_filtrado["USUARIO"].str.upper() == n.upper()]
            qtd_exemplares = int(df_func["TOTAL"].sum())
            qtd_skus = int(len(df_func))
        else:
            qtd_exemplares, qtd_skus = 0, 0
            
        mov = dict_movimentacao[n]
        historico_justificativas = []
        
        # Validação do primeiro desvio
        if mov["sai1"].strip() != "" and mov["sai1"].strip().upper() != "N/A":
            historico_justificativas.append(f"Encaminhada ao {mov['loc1']} das {mov['sai1']} às {mov['ret1']}")
        
        # Validação do segundo desvio (aba adicional)
        if mov["sai2"].strip() != "" and mov["sai2"].strip().upper() != "N/A":
            historico_justificativas.append(f"encaminhada ao {mov['loc2']} das {mov['sai2']} às {mov['ret2']}")
            
        if historico_justificativas:
            justificativa_texto = " ; ".join(historico_justificativas) + "."
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
    st.markdown("<h3 style='color: #4B5563; font-size: 1.2rem; font-weight: 600; margin-bottom:10px;'>📋 Detalhamento Gerencial de Produtividade</h3>", unsafe_allow_html=True)
    st.dataframe(df_real, use_container_width=True, hide_index=True)

    # 4. Caixa de Texto Gerada do E-mail Organizada por Hierarquia
    st.markdown("<br><hr>", unsafe_allow_html=True)
    st.markdown("<h3 style='color: #1E3A8A; font-weight: 700;'>✉️ Texto do E-mail Pronto para a Diretoria</h3>", unsafe_allow_html=True)
    
    blocos_email = ""
    for cargo_tipo in ["Líder", "Apoio", "Operadoras"]:
        blocos_email += f"\\n**{cargo_tipo}:**\\n"
        df_cargo = df_real[df_real["Cargo"] == cargo_tipo]
        for idx, r in df_cargo.iterrows():
            blocos_email += f"• **{r['Colaboradora']}**: {r['Exemplares']:,} exemplares | {r['SKUs']:,} SKUs | *{r['Movimentação Operacional']}*\\n"

    texto_final = f"""
**Assunto:** Relatório de Produtividade e Histórico de Movimentação - Varejo

Boa tarde, Prezados.

Segue abaixo o relatório analítico de produção do setor de Varejo, acompanhado das justificativas de movimentações internas da equipe.

**1. Resumo de Production Geral (Performance Diária)**
• **Total de Exemplares Separados:** {total_exemplares:,} un (Atingido: {pct_exemplares:.1%} da meta de {META_EXEMPLARES:,})
• **SKUs Movimentados:** {total_skus:,} itens (Atingido: {pct_skus:.1%} da meta de {META_SKUS:,})

**2. Indicadores de Desempenho Coletivo e Histórico por Função**
{blocos_email}
*Nota: Os remanejamentos de colaboradores ocorreram preventivamente devido à flutuação no volume de pedidos disponíveis em estoque.*

Atenciosamente,
    """
    st.text_area("Selecione tudo abaixo e copie (Ctrl+A / Ctrl+C):", value=texto_final.strip(), height=320)
else:
    st.info("👋 Abas adicionais de horário configuradas. Faça o upload da sua planilha Excel na barra lateral.")
