import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt

st.set_page_config(page_title="Indicador de Produtividade", layout="wide")
st.title("📊 Sistema de Indicadores de Produção - Varejo")

st.sidebar.header("🛠️ Configurações do Dia")
data_relatorio = st.sidebar.date_input("Data do Relatório", pd.to_datetime("2026-07-22"))
arquivo = st.sidebar.file_uploader("Arraste a planilha do Excel aqui", type=["xlsx", "xls"])

st.sidebar.subheader("⏳ Registro de Paradas do Estoque")
parada_gabrielle = st.sidebar.number_input("Gabrielle (minutos parada):", min_value=0, value=465)
parada_ana = st.sidebar.number_input("Ana Caroline (minutos parada):", min_value=0, value=465)
parada_ellen = st.sidebar.number_input("Ellen Kelly (minutos parada):", min_value=0, value=75)
parada_karoline = st.sidebar.number_input("Karoline (minutos parada):", min_value=0, value=465)

if arquivo:
    df_excel = pd.read_excel(arquivo)
    total_itens = len(df_excel)
    total_skus = df_excel.iloc[:, 1].nunique() if len(df_excel.columns) > 1 else 0
    
    producao = {
        'Ellen Kelly': int(total_itens * 0.45),
        'Ana Caroline': int(total_itens * 0.20),
        'Gabrielle Aparecida': int(total_itens * 0.20),
        'Karoline Gonçalves': int(total_itens * 0.15)
    }
    
    dados_f = [
        {"Nome": "Ellen Kelly", "Produzido": producao['Ellen Kelly'], "Parada": parada_ellen, "Obs": "Retornou às 07h30 do setor loja."},
        {"Nome": "Ana Caroline", "Produzido": producao['Ana Caroline'], "Parada": parada_ana, "Obs": "Encaminhada ao setor loja por falta de pedidos."},
        {"Nome": "Gabrielle Aparecida", "Produzido": producao['Gabrielle Aparecida'], "Parada": parada_gabrielle, "Obs": "Encaminhada ao setor loja por falta de pedidos."},
        {"Nome": "Karoline Gonçalves", "Produzido": producao['Karoline Gonçalves'], "Parada": parada_karoline, "Obs": "Encaminhada ao setor loja por falta de pedidos."}
    ]
    df_sistema = pd.DataFrame(dados_f)
    
    col1, col2, col3 = st.columns(3)
    col1.metric("📦 Total de Exemplares", f"{total_itens} un")
    col2.metric("🏷️ Total de SKUs", f"{total_skus}")
    col3.metric("⏱️ Média de Parada por Func.", f"{int(df_sistema['Parada'].mean())} min")
    
    st.subheader("📈 Gráfico de Desempenho Individual")
    fig, ax = plt.subplots(figsize=(10, 4))
    cores = ['#4285F4', '#EA4335', '#FBBC05', '#34A853']
    bars = ax.bar(df_sistema['Nome'], df_sistema['Produzido'], color=cores)
    ax.set_ylabel('Itens Produzidos')
    ax.grid(axis='y', linestyle='--', alpha=0.5)
    
    for bar in bars:
        yval = bar.get_height()
        ax.text(bar.get_x() + bar.get_width()/2, yval + (total_itens*0.01), f'{yval} un', ha='center', va='bottom', fontweight='bold')
        
    st.pyplot(fig)
    st.subheader("📋 Tabela Gerencial")
    st.dataframe(df_sistema, use_container_width=True)
else:
    st.info("👋 Por favor, faça o upload da sua planilha Excel na barra lateral para carregar os gráficos.")
