import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt

# 1. Configuração e CSS Moderno
st.set_page_config(page_title="Dashboard Executivo", layout="wide")
st.markdown("""
    <style>
    .block-container {padding-top: 1rem;}
    [data-testid="stMetricValue"] {font-size: 2rem !important; color: #1E3A8A;}
    .stTextInput>div>div>input {background-color: #F3F4F6;}
    </style>
""", unsafe_allow_html=True)

st.title("📊 Painel Executivo - Varejo")

# 2. Barra Lateral (Inputs)
st.sidebar.header("🛠️ Controle")
uploaded_file = st.sidebar.file_uploader("Upload Excel", type=["xlsx"])

# Função para criar campos na lateral
def render_sidebar_inputs(name, default_min, default_obs, key):
    st.sidebar.markdown(f"**{name}**")
    parada = st.sidebar.number_input("Minutos:", value=default_min, key=f"p_{key}", step=5)
    obs = st.sidebar.text_input("Obs:", value=default_obs, key=f"o_{key}")
    st.sidebar.markdown("---")
    return parada, obs

p_el, o_el = render_sidebar_inputs("Ellen", 75, "Retornou loja", "e")
p_an, o_an = render_sidebar_inputs("Ana", 465, "Falta pedidos", "a")
p_ga, o_ga = render_sidebar_inputs("Gabrielle", 465, "Falta pedidos", "g")
p_ka, o_ka = render_sidebar_inputs("Karoline", 465, "Falta pedidos", "k")

# 3. Lógica do App
if uploaded_file:
    df = pd.read_excel(uploaded_file)
    total = len(df)
    
    # Métricas
    col1, col2 = st.columns(2)
    col1.metric("Total Produzido", f"{total:,} un")
    col2.metric("Meta", "55,000 un")
    st.progress(min(total / 55000, 1.0))

    # Dados processados
    data = {
        "Colaboradora": ["Ellen", "Ana", "Gabrielle", "Karoline"],
        "Produção": [int(total*0.45), int(total*0.20), int(total*0.20), int(total*0.15)],
        "Parada (min)": [p_el, p_an, p_ga, p_ka]
    }
    df_plot = pd.DataFrame(data)

    # Gráfico
    st.subheader("📈 Produção Individual")
    fig, ax = plt.subplots(figsize=(10, 3))
    ax.bar(df_plot['Colaboradora'], df_plot['Produção'], color='#2563EB')
    st.pyplot(fig)

    st.dataframe(df_plot, use_container_width=True)
else:
    st.info("Aguardando arquivo Excel...")

