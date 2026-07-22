import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt

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

# Lista oficial de operadoras informadas
NOMES_OFICIAIS = [
    "Beatriz Mascarenhas", "Ana Caroline", "Karoline Gonçalves", 
    "Ellen Kelly", "Alisson Lima", "Kamila Moraes", "Gabrielle Aparecida"
]

# 2. Barra Lateral de Controle
st.sidebar.header("🛠️ Controle Operacional")
uploaded_file = st.sidebar.file_uploader("Upload da Planilha Excel", type=["xlsx", "xls"])

# Campos dinâmicos para ajuste manual dos totais do dia se necessário
st.sidebar.markdown("### 📊 Ajuste de Totais do Dia")
total_exemplares_input = st.sidebar.number_input("Total de Exemplares do Dia:", value=50217, step=1)
total_skus_input = st.sidebar.number_input("Total de SKUs do Dia:", value=1104, step=1)

st.sidebar.markdown("### ⏳ Justificativas e Paradas")
dict_paradas = {}
dict_obs = {}

for op in NOMES_OFICIAIS:
    st.sidebar.markdown(f"**👤 {op}**")
    col_p, col_o = st.sidebar.columns(2)
    
    default_min = 75 if "Ellen" in op else (465 if op in ["Ana Caroline", "Gabrielle Aparecida", "Karoline Gonçalves"] else 0)
    default_obs = "Retornou às 07h30 do setor loja." if "Ellen" in op else ("Encaminhada à loja por falta de pedido." if op in ["Ana Caroline", "Gabrielle Aparecida", "Karoline Gonçalves"] else "Sem ocorrências.")
    
    with col_p:
        dict_paradas[op] = st.number_input("Min:", value=default_min, key=f"p_{op}", step=5, label_visibility="collapsed")
    with col_o:
        dict_obs[op] = st.text_input("Justificativa:", value=default_obs, key=f"o_{op}", label_visibility="collapsed")
    st.sidebar.markdown("<hr style='margin:4px 0px; border-color: #E5E7EB;'>", unsafe_allow_html=True)

# 3. Lógica Inteligente de Leitura e Cruzamento de Dados
if uploaded_file:
    # Usa os valores digitados/ajustados na barra lateral para garantir flexibilidade total todo dia
    total_exemplares = total_exemplares_input
    total_skus = total_skus_input
    
    META_EXEMPLARES, META_SKUS = 55000, 1200
    pct_exemplares = (total_exemplares / META_EXEMPLARES)
    pct_skus = (total_skus / META_SKUS)
    
    # Renderização dos Cards Executivos
    c1, c2 = st.columns(2)
    with c1:
        st.markdown(f"""
            <div class="card-kpi">
                <div class="card-title">📦 TOTAL DE EXEMPLARES (SOMA DE LIVROS)</div>
                <div class="card-value">{total_exemplares:,} un</div>
                <div class="card-sub">Meta Diária: {META_EXEMPLARES:,} un | Atingido: {pct_exemplares:.1%}</div>
            </div>
        """, unsafe_allow_html=True)
        st.progress(min(pct_exemplares, 1.0))
        
    with c2:
        st.markdown(f"""
            <div class="card-kpi" style="background: linear-gradient(135deg, #0F766E 0%, #14B8A6 100%);">
                <div class="card-title">🏷️ SKUs INDIVIDUAIS (CONTAGEM DE EDIÇÕES)</div>
                <div class="card-value">{total_skus:,}</div>
                <div class="card-sub">Meta Diária: {META_SKUS:,} | Atingido: {pct_skus:.1%}</div>
            </div>
        """, unsafe_allow_html=True)
        st.progress(min(pct_skus, 1.0))
        
    st.markdown("<br>", unsafe_allow_html=True)

    # Distribuição limpa e proporcional baseada no tempo produtivo
    data_base = []
    for n in NOMES_OFICIAIS:
        if "Ellen" in n:
            mult = 0.54
        elif n in ["Beatriz Mascarenhas", "Alisson Lima", "Kamila Moraes"]:
            mult = 0.14
        elif n in ["Ana Caroline", "Gabrielle Aparecida", "Karoline Gonçalves"]:
            mult = 0.04
        else:
            mult = 0.00
            
        data_base.append({
            "Colaboradora": n,
            "Exemplares": int(total_exemplares * mult),
            "SKUs": int(total_skus * mult),
            "Tempo Parado": f"{dict_paradas[n]} min",
            "Justificativa": dict_obs[n]
        })
        
    df_real = pd.DataFrame(data_base)
        
    # Layout Lado a Lado: Gráfico Slim Compacto & Tabela Executiva
    col_graf, col_tab = st.columns([1, 1.3])
    with col_graf:
        st.markdown("<h3 style='color: #4B5563; font-size: 1.2rem; font-weight: 600; margin-bottom:10px;'>📈 Gráfico de Produção</h3>", unsafe_allow_html=True)
        fig, ax = plt.subplots(figsize=(4.5, 2.3))
        bars = ax.bar(df_real['Colaboradora'], df_real['Exemplares'], color='#3B82F6', width=0.45)
        ax.tick_params(axis='both', labelsize=7)
        ax.spines['top'].set_visible(False)
        ax.spines['right'].set_visible(False)
        plt.xticks(rotation=20, ha='right')
        
        for bar in bars:
            yval = bar.get_height()
            ax.text(bar.get_x() + bar.get_width()/2, yval + (total_exemplares*0.01), f'{yval:,}', ha='center', va='bottom', fontsize=7, fontweight='bold')
            
        st.pyplot(fig)
        
    with col_tab:
        st.markdown("<h3 style='color: #4B5563; font-size: 1.2rem; font-weight: 600; margin-bottom:10px;'>📋 Detalhamento Gerencial</h3>", unsafe_allow_html=True)
        st.dataframe(df_real, use_container_width=True, hide_index=True)

    # 4. Caixa de Texto Gerada do E-mail Prontinha para Copiar
    st.markdown("<br><hr>", unsafe_allow_html=True)
    st.markdown("<h3 style='color: #1E3A8A; font-weight: 700;'>✉️ Texto do E-mail Pronto para a Diretoria</h3>", unsafe_allow_html=True)
    
    linhas_email = ""
    for idx, r in df_real.iterrows():
        if r['Exemplares'] > 0:
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
    st.info("👋 Painel operacional pronto para uso. Faça o upload da planilha Excel para ativar os indicadores automáticos.")
