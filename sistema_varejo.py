"""
Painel Executivo de Produção — Varejo
======================================
Requisitos: streamlit>=1.50, pandas, openpyxl, matplotlib

Tudo que já existia continua funcionando (leitura por cabeçalho, ausências,
movimentação, texto livre, horários salvos, imagem, Excel, e-mail em HTML,
adicionar manualmente). Novidades desta versão:

 17. NOVOS COLABORADORES SEM DIGITAR: nomes que aparecem na planilha e não
     estão na equipe viram uma lista na lateral ("🆕 Novos na planilha").
     Marque quem entra, escolha o cargo e clique em Cadastrar. Também há o
     botão Vincular (mesma pessoa com outro nome na planilha) e Ignorar
     (nomes que não são da equipe, para o aviso parar de aparecer).
 18. EQUIPE EDITÁVEL COMO PLANILHA (st.data_editor) no lugar do JSON: nome,
     cargo, nome na planilha, meta individual e horários padrão.
 19. PERSISTÊNCIA EM DISCO: equipe, pessoas adicionadas manualmente, metas e
     destinatários agora sobrevivem a recarregar a página/reiniciar o app.
 20. METAS DO SETOR editáveis na lateral (antes fixas no código).
 21. HISTÓRICO DIÁRIO (CSV) + gráficos de evolução na aba "Histórico".
 22. Bug corrigido: o campo de "Adicionar manualmente" agora usa st.form
     (clear_on_submit) e não gera mais StreamlitAPIException.
 23. SMTP: usuário/senha podem vir de st.secrets (SMTP_HOST, SMTP_PORT,
     SMTP_USUARIO, SMTP_SENHA). Destinatários ficam salvos.
 24. Tabela ordenada por cargo (Líder → Apoio → Operador) e depois por
     exemplares; cargos com selo colorido, avatar com iniciais e % de meta
     individual com cor de status.
 25. Novo layout: cabeçalho, indicadores com barra de progresso,
     abas (Detalhamento / Imagem / Histórico / E-mail) e lateral organizada
     em seções recolhíveis.
 26. Indicadores simplificados: apenas Total de exemplares e Total de SKU,
     sem selos de status e sem os cards de colaboradores/média.
 27. Novas colaboradoras: Ingrid Pereira e Kerolay Vitoria, cargo Operador(a).
     Quem já tem equipe salva em disco recebe as duas automaticamente uma
     única vez.
 28. Compatibilidade com Streamlit novo: use_container_width foi trocado
     por width="stretch".
 29. CORREÇÃO DE ALIAS: Ingrid e Kerolay aparecem na planilha com o nome
     completo (INGRID PEREIRA / KEROLAY VITORIA), então não precisam de alias.
     Os logins errados (ingridp / kerolayv) são removidos automaticamente,
     uma única vez, de equipes já salvas em disco.

 30. CADASTRO AUTOMÁTICO: quem aparece na planilha e não está na equipe é
     cadastrado sozinho (cargo padrão configurável), usando o nome da planilha
     como alias, então o vínculo nunca falha. Pode ser desligado na lateral.

Observação: em Streamlit Cloud o disco é temporário. Para persistência
definitiva em nuvem, guarde esses JSON/CSV em banco ou Google Sheets.
"""

import streamlit as st
import pandas as pd
import openpyxl
from datetime import datetime, time as dtime
from zoneinfo import ZoneInfo
import io
import os
import json
import html
import numbers
import textwrap
import unicodedata
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.image import MIMEImage
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches

# =============================================================================
# 0. CONSTANTES DE ARQUIVO
# =============================================================================
MOV_OVERRIDES_PATH = "movimentacoes_manuais.json"
HORARIOS_SALVOS_PATH = "horarios_salvos.json"
EQUIPE_CONFIG_PATH = "equipe_config.json"
PESSOAS_MANUAIS_PATH = "pessoas_manuais.json"
CONFIGURACOES_PATH = "configuracoes.json"
HISTORICO_PATH = "historico_diario.csv"


# =============================================================================
# 1. Configuração da página e estilo
# =============================================================================
st.set_page_config(page_title="Painel Executivo de Produção", page_icon="📊", layout="wide")
st.markdown("""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Sora:wght@500;600;700;800&family=Inter:wght@400;500;600;700&display=swap');

    html, body, [class*="css"] { font-family: 'Inter', sans-serif; }
    .stApp { background-color: #F4F6FA; }
    .block-container { padding-top: 1.4rem; padding-bottom: 2rem; max-width: 96%; }
    footer { visibility: hidden; }

    /* ---------- Lateral ---------- */
