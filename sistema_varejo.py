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

# Organização da equipe por cargos oficiais
EQUIPE = {
    "Líder": ["Kamila Moraes", "Beatriz Alcantara"],
    "Apoio": ["Alisson Lima"],
    "Operador(a)": ["Rosana Delfino", "Anacaroline", "Karoline Gonçalves", "Gabriele", "Beatriz Mascarenhas", "Graziela Pereira", "Paula Roberta", "Weliton"]
}
NOMES_LISTA = EQUIPE["Líder"] + EQUIPE["Apoio"] + EQUIPE["Operador(a)"]

# Mapeamento de nomes que aparecem de forma diferente (abreviada) na planilha Excel.
# A chave é o nome de exibição usado no painel; o valor é como o nome aparece na coluna USUARIO do Excel.
# Por padrão, o próprio nome (em maiúsculas) é usado para o cruzamento dos dados.
ALIAS_EXCEL = {
    "Beatriz Alcantara": "BEATRIZ",
    "Graziela Pereira": "GRAZIELA PEREIRA DO NASCIMENTO",
    "Paula Roberta": "PAULA ROBERTA SANTOS DA SILVA",
}

import unicodedata

def normalizar(texto):
    """Remove acentos, espaços extras e padroniza para maiúsculas, evitando falhas de
    correspondência entre o nome cadastrado no painel e o nome como aparece na planilha."""
    texto = str(texto).strip().upper()
    texto = unicodedata.normalize("NFKD", texto).encode("ASCII", "ignore").decode("ASCII")
    texto = " ".join(texto.split())
    return texto

def nome_excel(nome):
    return normalizar(ALIAS_EXCEL.get(nome, nome))

# 2. Barra Lateral de Controle Unificada
st.sidebar.header("🛠️ Controle Operacional")
uploaded_file = st.sidebar.file_uploader("Upload da Planilha Excel", type=["xlsx"])

# NOVO: Seleção da data da produtividade para atualizar o e-mail automaticamente
data_produtividade = st.sidebar.date_input("Data da Produtividade:", datetime.now())
data_formatada = data_produtividade.strftime("%d/%m")

# Filtros gerenciais limpos
st.sidebar.markdown("### 👁️ Filtros Gerenciais")
remover_do_setor = st.sidebar.multiselect("Ocultar do Setor (Tabela):", NOMES_LISTA)

st.sidebar.markdown("### ❌ Ausências do Dia")
faltas_selecionadas = st.sidebar.multiselect("Selecione quem faltou hoje:", NOMES_LISTA)

st.sidebar.markdown("### ⏳ Movimentação de Horários")
movimentados_selecionados = st.sidebar.multiselect(
    "🚚 Quem foi movimentado(a) hoje?",
    [n for n in NOMES_LISTA if n not in remover_do_setor and n not in faltas_selecionadas]
)

dict_movimentacao = {}
dict_motivos_falta = {}

for cargo, integrantes in EQUIPE.items():
    integrantes_visiveis = [i for i in integrantes if i not in remover_do_setor]
    if integrantes_visiveis:
        st.sidebar.markdown(f"<h3 style='color:#1E3A8A; margin-top:10px; font-size:1.1rem;'>🔹 {cargo.upper()}</h3>", unsafe_allow_html=True)
        
    for op in integrantes:
        if op in remover_do_setor:
            continue
            
        is_ausente = op in faltas_selecionadas
        is_movimentado = op in movimentados_selecionados
        
        if is_ausente:
            st.sidebar.markdown(f"❌ **{op} (AUSENTE)**")
            dict_motivos_falta[op] = st.sidebar.text_input(f"Motivo da falta de {op}:", value="Falta administrativa", key=f"mot_falta_{op}")
            dict_movimentacao[op] = {"cargo": cargo, "movimentacoes": []}
        elif is_movimentado:
            st.sidebar.markdown(f"**👤 {op}**", unsafe_allow_html=True)

            if f"num_mov_{op}" not in st.session_state:
                st.session_state[f"num_mov_{op}"] = 1  # começa com 1 saída, e cresce com o botão "+"

            movimentacoes_op = []
            for idx in range(1, st.session_state[f"num_mov_{op}"] + 1):
                st.sidebar.markdown(f"<span style='font-size:0.8rem; color:gray;'>Saída {idx}:</span>", unsafe_allow_html=True)
                c_sai, c_ret, c_loc = st.sidebar.columns(3)

                # Valores padrão (apenas na primeira saída, pra manter o comportamento de antes)
                init_sai, init_ret, init_loc = "", "", ""
                if idx == 1:
                    init_sai = "06h15" if op in ["Anacaroline", "Rosana Delfino", "Karoline Gonçalves", "Gabriele"] else ""
                    init_ret = "10h00" if op == "Gabriele" else ("10h30" if op in ["Anacaroline", "Karoline Gonçalves"] else ("07h30" if op == "Rosana Delfino" else ""))
                    init_loc = "Setor Loja" if op in ["Anacaroline", "Rosana Delfino", "Karoline Gonçalves", "Gabriele"] else ""

                with c_sai: sai = st.text_input("Saída:", value=init_sai, key=f"sai_{op}_{idx}")
                with c_ret: ret = st.text_input("Retorno:", value=init_ret, key=f"ret_{op}_{idx}")
                with c_loc: loc = st.text_input("Local:", value=init_loc, key=f"loc_{op}_{idx}")

                movimentacoes_op.append({"sai": sai, "ret": ret, "loc": loc})

            if st.sidebar.button(f"➕ Adicionar saída para {op}", key=f"add_mov_{op}"):
                st.session_state[f"num_mov_{op}"] += 1
                st.rerun()

            dict_movimentacao[op] = {"cargo": cargo, "movimentacoes": movimentacoes_op}
        else:
            st.sidebar.markdown(f"👤 {op} <span style='font-size:0.8rem; color:gray;'>(sem movimentação)</span>", unsafe_allow_html=True)
            dict_movimentacao[op] = {"cargo": cargo, "movimentacoes": []}
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
                dados_visiveis.append({"TOTAL": val_i, "USUARIO": normalizar(val_m)})
                
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
        st.markdown(f'<div class="card-kpi"><div class="card-title">📦 TOTAL DE EXEMPLARES</div><div class="card-value">{total_exemplares:,} un</div><div class="card-sub">Meta Diária: {META_EXEMPLARES:,} un | Atingido: {pct_exemplares:.1%}</div></div>', unsafe_allow_html=True)
        st.progress(min(pct_exemplares, 1.0))
    with c2:
        st.markdown(f'<div class="card-kpi" style="background: linear-gradient(135deg, #0F766E 0%, #14B8A6 100%);"><div class="card-title">🏷️ TOTAL DE SKU</div><div class="card-value">{total_skus:,}</div><div class="card-sub">Meta Diária: {META_SKUS:,} | Atingido: {pct_skus:.1%}</div></div>', unsafe_allow_html=True)
        st.progress(min(pct_skus, 1.0))
        
    st.markdown("<br>", unsafe_allow_html=True)

    # Processamento dos indicadores individuais
    data_gerencial = []
    for n in NOMES_LISTA:
        if n in remover_do_setor:
            continue
            
        is_ausente = n in faltas_selecionadas
        
        if is_ausente:
            qtd_exemplares = 0
            qtd_skus = 0
            motivo_individual = dict_motivos_falta.get(n, "Falta administrativa")
            justificativa_texto = f"Ausente. Motivo: {motivo_individual}."
            cargo_atual = "Operador(a)"
            for c, ints in EQUIPE.items():
                if n in ints: cargo_atual = c
        else:
            mov = dict_movimentacao[n]
            cargo_atual = mov["cargo"]
            if not df_filtrado.empty:
                df_func = df_filtrado[df_filtrado["USUARIO"] == nome_excel(n)]
                qtd_exemplares = int(df_func["TOTAL"].sum())
                qtd_skus = int(len(df_func))
            else:
                qtd_exemplares, qtd_skus = 0, 0

            # Se o operador não aparece na planilha (nenhum registro encontrado) e não está
            # marcado como ausente, ele é ignorado e não entra na tabela gerencial.
            if qtd_skus == 0:
                continue
                
            historico_justificativas = []
            for idx, m in enumerate(mov["movimentacoes"]):
                if m["sai"].strip() != "" and m["sai"].strip().upper() != "N/A":
                    prefixo = "Encaminhada" if idx == 0 else "encaminhada"
                    historico_justificativas.append(f"{prefixo} ao {m['loc']} das {m['sai']} às {m['ret']}")
                
            justificativa_texto = " ; ".join(historico_justificativas) + "." if historico_justificativas else "Atividade normal no setor."
            
        data_gerencial.append({
            "Cargo": cargo_atual,
            "Colaboradora": n,
            "Exemplares": qtd_exemplares,
            "SKUs": qtd_skus,
            "Movimentação Operacional": justificativa_texto
        })
        
    df_real = pd.DataFrame(data_gerencial)
    st.markdown("<h3 style='color: #4B5563; font-size: 1.2rem; font-weight: 600; margin-bottom:10px;'>📋 Detalhamento Gerencial de Produtividade</h3>", unsafe_allow_html=True)
    st.dataframe(df_real, use_container_width=True, hide_index=True)

    # 4. Caixa de Texto Gerada do E-mail Padronizado Conforme Solicitado
    st.markdown("<br><hr>", unsafe_allow_html=True)
    st.markdown("<h3 style='color: #1E3A8A; font-weight: 700;'>✉️ Texto do E-mail Pronto para a Diretoria</h3>", unsafe_allow_html=True)
    
    texto_final = f"Boa tarde, Prezados.\n\nSegue abaixo o relatório de produção.\nreferente ao dia {data_formatada}.\n\n--------------------------------\nResumo Varejo.\nSKU: {total_skus}\nExemplares: {total_exemplares:,}\n--------------------------------\n\nAtenciosamente,"
    
    st.text_area("Selecione tudo abaixo e copie (Ctrl+A / Ctrl+C):", value=texto_final, height=240)
else:
    st.info("👋 Padrão de e-mail atualizado. Faça o upload da sua planilha Excel na barra lateral.")
