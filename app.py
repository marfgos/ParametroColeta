import streamlit as st
import pandas as pd
import os

st.set_page_config(page_title="Parâmetros de Coleta", layout="wide")

st.title("📦 Configuração de Parâmetros de Coleta")

# =========================
# 📁 ARQUIVOS INTERNOS
# =========================

# Arquivos internos
arquivo_base_padrao = "parametros_contrato.xlsx"
arquivo_parametros_usuario = "parametros_usuario.xlsx"

# =========================
# 🧾 CARREGAR BASE PADRÃO
# =========================
try:
    df_padrao = pd.read_excel(arquivo_base_padrao)
except Exception as e:
    st.error(f"Erro ao carregar a base padrão: {e}")
    df_padrao = pd.DataFrame()

# =========================
# 📁 CARREGAR BASE DO USUÁRIO
# =========================
if os.path.exists(arquivo_parametros_usuario):
    df_usuario = pd.read_excel(arquivo_parametros_usuario)
else:
    df_usuario = pd.DataFrame(columns=df_padrao.columns)

# =========================
# ➕ FORMULÁRIO PARA NOVA REGRA
# =========================
st.markdown("## ➕ Adicionar Nova Regra de Redirecionamento")

with st.form("form_nova_regra"):
    col1, col2 = st.columns(2)
    with col1:
        substituta = st.text_input("Substituta")
        inicial = st.text_input("Inicial")
        recebe = st.text_input("Recebe")
        uf = st.text_input("UF")
    with col2:
        grupo = st.text_input("Grupo Econômico")
        modalidade = st.selectbox("Modalidade", ["", "FCA", "EXW"])
        tipo = st.selectbox("Tipo de Carga", ["", "Fracionado", "Lotação"])

    submitted = st.form_submit_button("✅ Adicionar Regra")

if submitted:
    nova_regra = {
        'Substituta': substituta,
        'Inicial': inicial,
        'Recebe': recebe,
        'UF': uf,
        'Grupo Economico': grupo,
        'Modalidade': modalidade,
        'Tipo de carga': tipo
    }

    df_usuario = pd.concat([df_usuario, pd.DataFrame([nova_regra])], ignore_index=True)
    df_usuario.to_excel(arquivo_parametros_usuario, index=False)
    st.success("✅ Regra adicionada com sucesso! Ela será usada nas próximas execuções.")

# =========================
# 🔗 BASE FINAL UNIFICADA
# =========================
df_completo = pd.concat([df_padrao, df_usuario], ignore_index=True)

st.markdown("### 📄 Base Final de Regras (Internas + Usuário)")
st.dataframe(df_completo, use_container_width=True)
