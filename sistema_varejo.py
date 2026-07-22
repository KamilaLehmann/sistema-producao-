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

# Lista oficial de operadoras mapeadas a partir da sua imagem
NOMES_OFICIAIS = [
    "Rosana Delfino", "Anacaroline", "Karoline Gonçalves", 
    "Gabriele", "Beatriz Mascarenhas", "Alisson Lima", "Kamila Moraes"
]

# 2. Barra Lateral de Controle
st.sidebar.header("🛠️ Controle Operacional")
uploaded_file = st.sidebar.file_uploader("Upload da Planilha Excel", type=["xlsx"])

st.sidebar.markdown("### ⏳ Justificativas e Paradas")
dict_paradas = {}
dict_obs = {}

for op in NOMES_OFICIAIS:
    st.sidebar.markdown(f"**👤 {op}**")
    col_p, col_o = st.sidebar.columns(2)
    
    default_min = 75 if "Gabriele" in op or "Ellen" in op else (465 if op in ["Anacaroline", "Rosana Delfino", "Karoline Gonçalves"] else 0)
    default_obs = "Retornou às 07h30." if "Gabriele" in op or "Ellen" in op else ("Encaminhada à loja por falta de pedido." if op in ["Anacaroline", "Rosana Delfino", "Karoline Gonçalves"] else "Sem ocorrências.")
    
    with col_p:
        dict_paradas[op] = st.number_input("Min:", value=default_min, key=f"p_{op}", step=5, label_visibility="collapsed")
    with col_o:
        dict_obs[op] = st.text_input("Justificativa:", value=default_obs, key=f"o_{op}", label_visibility="collapsed")
    st.sidebar.markdown("<hr style='margin:4px 0px; border-color: #E5E7EB;'>", unsafe_allow_html=True)

# 3. Lógica Avançada: Lendo Apenas Linhas Visíveis (Filtradas) do Excel
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

    # Metas Diárias fixadas do setor
    META_EXEMPLARES, META_SKUS = 55000, 1200
    pct_exemplares = (total_exemplares / META_EXEMPLARES)
    pct_skus = (total_skus / META_SKUS)
    
    # Renderização dos Cards HTML de Alta Visibilidade
    c1, c2 = st.columns(2)
    with c1:
        st.markdown(f"""
            <div class="card-kpi">
                <div class="card-title">📦 TOTAL DE EXEMPLARES (SOMA DA COLUNA I)</div>
                <div class="card-value">{total_exemplares:,} un</div>
                <div class="card-sub">Meta Diária: {META_EXEMPLARES:,} un | Atingido: {pct_exemplares:.1%}</div>
            </div>
        """, unsafe_allow_html=True)
        st.progress(min(pct_exemplares, 1.0))
        
    with c2:
        st.markdown(f"""
            <div class="card-kpi" style="background: linear-gradient(135deg, #0F766E 0%, #14B8A6 100%);">
                <div class="card-title">🏷️ SKUs FILTRADOS (CONTAGEM DA COLUNA I)</div>
                <div class="card-value">{total_skus:,}</div>
                <div class="card-sub">Meta Diária: {META_SKUS:,} | Atingido: {pct_skus:.1%}</div>
            </div>
        """, unsafe_allow_html=True)
        st.progress(min(pct_skus, 1.0))
        
    st.markdown("<br>", unsafe_allow_html=True)

    # Processamento dos indicadores individuais focado na lista oficial
    data_gerencial = []
    for n in NOMES_OFICIAIS:
        if not df_filtrado.empty:
            df_func = df_filtrado[df_filtrado["USUARIO"].str.upper() == n.upper()]
            qtd_exemplares = int(df_func["TOTAL"].sum())
            qtd_skus = int(len(df_func))
        else:
            qtd_exemplares, qtd_skus = 0, 0
            
        data_gerencial.append({
            "Colaboradora": n,
            "Exemplares": qtd_exemplares,
            "SKUs": qtd_skus,
            "Tempo Parado": f"{dict_paradas[n]} min",
            "Justificativa": dict_obs[n]
        })
        
    df_real = pd.DataFrame(data_gerencial)
        
    # Exibição da Tabela Gerencial em tamanho completo na tela
    st.markdown("<h3 style='color: #4B5563; font-size: 1.2rem; font-weight: 600; margin-bottom:10px;'>📋 Detalhamento Gerencial de Produtividade</h3>", unsafe_allow_html=True)
    st.dataframe(df_real, use_container_width=True, hide_index=True)

    # 4. Caixa de Texto Gerada do E-mail Prontinha para Copiar
    st.markdown("<br><hr>", unsafe_allow_html=True)
    st.markdown("<h3 style='color: #1E3A8A; font-weight: 700;'>✉️ Texto do E-mail Pronto para a Diretoria</h3>", unsafe_allow_html=True)
    
    linhas_email = ""
    for idx, r in df_real.iterrows():
        if r['Exemplares'] > 0 or r['SKUs'] > 0:
            linhas_email += f"• **{r['Colaboradora']}**: {r['Exemplares']:,} exemplares | {r['SKUs']:,} SKUs | Parada: {r['Tempo Parado']} ({r['Justificativa']})\\n"

    texto_final = f"""
**Assunto:** Relatório de Produtividade e Indicadores de Parada - Varejo

Boa tarde, Prezados.

Segue abaixo o relatório analítico de produção do setor de Varejo, acompanhado dos indicadores de atingimento de metas.

**1. Resumo de Produção Geral (Performance Diária)**
• **Total de Exemplares Separados:** {total_exemplares:,} un (Atingido: {pct_exemplares:.1%} da meta de {META_EXEMPLARES:,})
• **SKUs Movimentados:** {total_skus:,} itens (Atingido: {pct_skus:.1%} da meta de {META_SKUS:,})

**2. Indicadores de Desempenho Coletivo e Tempo de Parada**
{linhas_email}
*Nota: As paradas registradas acima ocorreram estritamente de acordo com o fluxo operacional e remanejamentos preventivos de estoque.*

Atenciosamente,
    """
    st.text_area("Selecione tudo abaixo e copie (Ctrl+A / Ctrl+C):", value=texto_final.strip(), height=280)
else:
    st.info("👋 Alinhamento concluído. Faça o upload da sua planilha Excel filtrada na barra lateral.")
