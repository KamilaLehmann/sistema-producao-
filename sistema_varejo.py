import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt

# Configuração visual da página (Estilo Diretor)
st.set_page_config(page_title="Dashboard Varejo", layout="wide")

# Estilização CSS personalizada para um design mais limpo
st.markdown("""
    <style>
    .block-container {padding-top: 1rem; padding-bottom: 1rem;}
    </style>
""", unsafe_allow_html=True)

st.title("📊 Painel Executivo de Produção")

# Sidebar compacta
st.sidebar.header("🛠️ Controle")
arquivo = st.sidebar.file_uploader("Upload Excel", type=["xlsx", "xls"])

if arquivo:
    df_excel = pd.read_excel(arquivo)
    total_itens = len(df_excel)
    total_skus = df_excel.iloc[:, 1].nunique() if len(df_excel.columns) > 1 else 0
    
    # Metas e Cálculo de Porcentagem
    META_SKU = 1200
    META_EXEMPLARES = 55000
    pct_sku = min(float(total_skus / META_SKU), 1.0)
    pct_exemplares = min(float(total_itens / META_EXEMPLARES), 1.0)
    
    # 1. KPIs com Progresso
    col1, col2 = st.columns(2)
    with col1:
        st.metric(label=f"Exemplares ({pct_exemplares*100:.1f}%)", value=f"{total_itens:,}", delta=f"Meta: {META_EXEMPLARES:,}")
        st.progress(pct_exemplares)
        
    with col2:
        st.metric(label=f"SKUs ({pct_sku*100:.1f}%)", value=f"{total_skus:,}", delta=f"Meta: {META_SKU:,}")
        st.progress(pct_sku)
        
    st.divider()
    # 2. Resumo de Produção (Simplificado)
    st.write("### Resumo Operacional")
    st.dataframe(df_excel.head(10), use_container_width=True) # Exemplo resumido
        
else:
    st.info("👋 Aguardando upload do arquivo Excel.")
