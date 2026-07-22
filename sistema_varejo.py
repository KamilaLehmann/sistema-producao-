import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt

# 1. Configuração de Layout e Estilo CSS Compacto
st.set_page_config(page_title="Dashboard Varejo", layout="wide")
st.markdown("""
    <style>
    .block-container {padding-top: 1rem; padding-bottom: 0rem; max-width: 95%;}
    [data-testid="stMetricValue"] {font-size: 1.8rem !important; font-weight: bold; color: #1E3A8A;}
    [data-testid="stMetricLabel"] {font-size: 1rem !important; color: #4B5563;}
    div.stProgress > div > div > div {background-image: linear-gradient(to right, #3B82F6, #10B981);}
    </style>
""", unsafe_allow_html=True)

st.title("📊 Painel Executivo de Produção - Varejo")

# 2. Barra Lateral de Controle
st.sidebar.header("🛠️ Controle Operacional")
uploaded_file = st.sidebar.file_uploader("Upload da Planilha Excel", type=["xlsx", "xls"])

def render_sidebar_inputs(name, default_min, default_obs, key):
    st.sidebar.markdown(f"**{name}**")
    col_p, col_o = st.sidebar.columns([1, 2])
    with col_p:
        parada = st.number_input("Min:", value=default_min, key=f"p_{key}", step=5, label_visibility="collapsed")
    with col_o:
        obs = st.text_input("Justificativa:", value=default_obs, key=f"o_{key}", label_visibility="collapsed")
    st.sidebar.markdown("<hr style='margin:4px 0px'>", unsafe_allow_html=True)
    return parada, obs

p_el, o_el = render_sidebar_inputs("Ellen Kelly", 75, "Retornou às 07h30 do setor loja.", "e")
p_an, o_an = render_sidebar_inputs("Ana Caroline", 465, "Encaminhada à loja por falta de pedido.", "a")
p_ga, o_ga = render_sidebar_inputs("Gabrielle Aparecida", 465, "Encaminhada à loja por falta de pedido.", "g")
p_ka, o_ka = render_sidebar_inputs("Karoline Gonçalves", 465, "Encaminhada à loja por falta de pedido.", "k")

# 3. Processamento e Exibição
if uploaded_file:
    df = pd.read_excel(uploaded_file)
    
    # Cálculo SKUs (Edições) e Exemplares (Total Livros)
    total_exemplares = len(df)
    total_skus = df.iloc[:, 1].nunique() if len(df.columns) > 1 else df.iloc[:, 0].nunique()
    
    META_EXEMPLARES, META_SKUS = 55000, 1200
    
    # Métricas
    c1, c2 = st.columns(2)
    with c1:
        st.metric(label=f"📦 Exemplares ({total_exemplares:,} un)", value=f"{total_exemplares/META_EXEMPLARES:.1%}", delta=f"Meta: {META_EXEMPLARES:,}")
        st.progress(min(total_exemplares / META_EXEMPLARES, 1.0))
    with c2:
        st.metric(label=f"🏷️ SKUs ({total_skus:,} itens)", value=f"{total_skus/META_SKUS:.1%}", delta=f"Meta: {META_SKUS:,}")
        st.progress(min(total_skus / META_SKUS, 1.0))
        
    # Tabela com distribuição estimada
    data_meninas = {
        "Colaboradora": ["Ellen Kelly", "Ana Caroline", "Gabrielle Aparecida", "Karoline Gonçalves"],
        "Exemplares": [int(total_exemplares*0.45), int(total_exemplares*0.20), int(total_exemplares*0.20), int(total_exemplares*0.15)],
        "SKUs": [int(total_skus*0.45), int(total_skus*0.20), int(total_skus*0.20), int(total_skus*0.15)],
        "Tempo Parado": [f"{p_el} min", f"{p_an} min", f"{p_ga} min", f"{p_ka} min"]
    }
    df_res = pd.DataFrame(data_meninas)
    
    # Layout Lado a Lado: Gráfico e Tabela
    col_graf, col_tab = st.columns([4, 6])
    with col_graf:
        st.subheader("📈 Produção (Exemplares)")
        fig, ax = plt.subplots(figsize=(5, 2.5)) # Gráfico menor
        ax.bar(df_res['Colaboradora'], df_res['Exemplares'], color='#2563EB')
        ax.tick_params(axis='both', labelsize=8)
        plt.xticks(rotation=15, ha='right')
        st.pyplot(fig)
    with col_tab:
        st.subheader("📋 Detalhamento")
        st.dataframe(df_res, use_container_width=True, hide_index=True)
else:
    st.info("Aguardando arquivo Excel...")
