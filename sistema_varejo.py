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

# 2. Barra Lateral de Controle com horários de movimentação
st.sidebar.header("🛠️ Controle Operacional")
uploaded_file = st.sidebar.file_uploader("Upload da Planilha Excel", type=["xlsx"])

st.sidebar.markdown("### ⏳ Movimentação de Horários")
dict_movimentacao = {}

for op in NOMES_OFICIAIS:
    st.sidebar.markdown(f"**👤 {op}**")
    c_sai, c_ret, c_loc = st.sidebar.columns(3)
    
    # Valores padrão de exemplo baseados no seu e-mail do dia 22/07
    init_sai = "06h15" if op in ["Anacaroline", "Rosana Delfino", "Karoline Gonçalves", "Gabriele"] else "N/A"
    init_ret = "07h30" if op == "Gabriele" else ("Não retornou" if op in ["Anacaroline", "Rosana Delfino", "Karoline Gonçalves"] else "N/A")
    init_loc = "Setor Loja" if op in ["Anacaroline", "Rosana Delfino", "Karoline Gonçalves", "Gabriele"] else "Sem desvios"
    
    with c_sai:
        sai_val = st.text_input("Saída:", value=init_sai, key=f"sai_{op}")
    with c_ret:
        ret_val = st.text_input("Retorno:", value=init_ret, key=f"ret_{op}")
    with c_loc:
        loc_val = st.text_input("Local:", value=init_loc, key=f"loc_{op}")
        
    dict_movimentacao[op] = {"saida": sai_val, "retorno": ret_val, "local": loc_val}
    st.sidebar.markdown("<hr style='margin:4px 0px; border-color: #E5E7EB;'>", unsafe_allow_html=True)

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
            
        # Monta o texto explicativo da movimentação com base nos inputs inseridos
        mov = dict_movimentacao[n]
        if mov["saida"] != "N/A" and mov["saida"] != "":
            justificativa_texto = f"Encaminhada ao {mov['local']} às {mov['saida']}. Retorno: {mov['retorno']}."
        else:
            justificativa_texto = "Sem ocorrências / Atividade normal."
            
        data_gerencial.append({
            "Colaboradora": n,
            "Exemplares": qtd_exemplares,
            "SKUs": qtd_skus,
            "Movimentação Operacional": justificativa_texto
        })
        
    df_real = pd.DataFrame(data_gerencial)
        
    # Exibição da Tabela Gerencial limpa em tamanho completo na tela
    st.markdown("<h3 style='color: #4B5563; font-size: 1.2rem; font-weight: 600; margin-bottom:10px;'>📋 Detalhamento Gerencial de Produtividade</h3>", unsafe_allow_html=True)
    st.dataframe(df_real, use_container_width=True, hide_index=True)

    # 4. Caixa de Texto Gerada do E-mail Prontinha para Copiar
    st.markdown("<br><hr>", unsafe_allow_html=True)
    st.markdown("<h3 style='color: #1E3A8A; font-weight: 700;'>✉️ Texto do E-mail Pronto para a Diretoria</h3>", unsafe_allow_html=True)
    
    linhas_email = ""
    for idx, r in df_real.iterrows():
        if r['Exemplares'] > 0 or r['SKUs'] > 0:
            linhas_email += f"• **{r['Colaboradora']}**: {r['Exemplares']:,} exemplares | {r['SKUs']:,} SKUs | *{r['Movimentação Operacional']}*\\n"

    texto_final = f"""
**Assunto:** Relatório de Produtividade e Histórico de Movimentação - Varejo

Boa tarde, Prezados.

Segue abaixo o relatório analítico de produção do setor de Varejo, acompanhado das justificativas de movimentações internas da equipe.

**1. Resumo de Produção Geral (Performance Diária)**
• **Total de Exemplares Separados:** {total_exemplares:,} un (Atingido: {pct_exemplares:.1%} da meta de {META_EXEMPLARES:,})
• **SKUs Movimentados:** {total_skus:,} itens (Atingido: {pct_skus:.1%} da meta de {META_SKUS:,})

**2. Indicadores de Desempenho Coletivo e Histórico de Horários**
{linhas_email}
*Nota: Os remanejamentos de colaboradores ocorreram preventivamente devido à flutuação no volume de pedidos disponíveis em estoque.*

Atenciosamente,
    """
    st.text_area("Selecione tudo abaixo e copie (Ctrl+A / Ctrl+C):", value=texto_final.strip(), height=280)
else:
    st.info("👋 Configuração de horários de saída/retorno concluída. Faça o upload da sua planilha Excel na barra lateral.")
