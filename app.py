import streamlit as st
import pandas as pd
import os
import logging
from tqdm import tqdm
from io import BytesIO

# === CONFIGURAÇÕES INICIAIS ===
st.set_page_config(page_title="Roteirização com Substituição", layout="wide")
st.title("📦 Roteirização com Regras de Substituição")

# === CONFIGURAÇÃO DE LOG ===
log_file = 'log_filiais_proximas.log'
logging.basicConfig(filename=log_file, level=logging.INFO, format='%(asctime)s - %(message)s')

# === CAMINHO DO ARQUIVO DE PARÂMETROS DENTRO DO REPOSITÓRIO ===
CAMINHO_PARAMETROS = "parametros_contrato.xlsx"

# === CARREGAMENTO DAS BASES FIXAS INTERNAS ===
try:
    df_dist = pd.read_excel("municipios_distanciasreais.xlsx")
    df_filiais = pd.read_excel("filiais_geocodificadas.xlsx")
    st.success("📍 Bases internas de distâncias e filiais carregadas com sucesso.")
except Exception as e:
    st.error(f"❌ Erro ao carregar arquivos internos: {e}")
    st.stop()

# === TÍTULO DO BLOCO DE PARÂMETROS ===
st.header("📥 Parâmetros Contratuais")

# === CARREGAMENTO DO ARQUIVO DE PARÂMETROS DO GIT ===
colunas_base = ['Substituta', 'Inicial', 'Recebe', 'UF', 'Grupo Economico', 'Modalidade', 'Tipo de carga']

if os.path.exists(CAMINHO_PARAMETROS):
    try:
        df_grupos = pd.read_excel(CAMINHO_PARAMETROS)
        st.success("📄 Parâmetros contratuais carregados do repositório.")
    except Exception as e:
        st.error(f"Erro ao carregar o arquivo de parâmetros: {e}")
        df_grupos = pd.DataFrame(columns=colunas_base)
else:
    st.warning("Arquivo de parâmetros não encontrado. Criando base vazia.")
    df_grupos = pd.DataFrame(columns=colunas_base)

# === EDIÇÃO INTERATIVA DOS PARÂMETROS ===
df_grupos_editado = st.data_editor(
    df_grupos,
    num_rows="dynamic",
    use_container_width=True,
    key="regras_editadas"
)

# === BOTÃO PARA SALVAR OS PARÂMETROS ===
if st.button("📤 Enviar alterações"):
    try:
        df_grupos_editado.to_excel(CAMINHO_PARAMETROS, index=False)
        st.success("💾 Alterações salvas com sucesso no repositório!")
    except Exception as e:
        st.error(f"Erro ao salvar alterações: {e}")

st.divider()

# === BOTÃO PARA PROCESSAR ===
if st.button("🚀 Rodar Roteirização"):
    if df_grupos_editado.empty:
        st.error("Por favor, preencha os parâmetros contratuais antes de processar.")
    else:
        with st.spinner("Processando..."):
            df_grupos = df_grupos_editado.copy()

            modalidades = [
                ("FCA", "Fracionado", "FCA/Fracionado"),
                ("FCA", "Lotação", "FCA/Lotação"),
                ("EXW", "Fracionado", "EXW/Fracionado"),
                ("EXW", "Lotação", "EXW/Lotação")
            ]

            def buscar_regras_substituicao(df_regras, uf, modalidade, tipo_carga):
                regras = df_regras[(df_regras['UF'] == uf) & (df_regras['Recebe'] == 'S')]
                regras = regras[
                    (regras['Modalidade'].isna() | (regras['Modalidade'] == modalidade)) &
                    (regras['Tipo de carga'].isna() | (regras['Tipo de carga'] == tipo_carga))
                ]
                return regras

            municipios = df_dist['MunicipioOrigem'].unique()
            resultados = []

            for municipio in tqdm(municipios, desc="Processando municípios"):
                uf_municipio = municipio.split('-')[-1].strip()

                for incoterm, tipo_carga, coluna_param in modalidades:
                    try:
                        filial_encontrada = False

                        regras_subs = buscar_regras_substituicao(df_grupos, uf_municipio, incoterm, tipo_carga)
                        if not regras_subs.empty:
                            regra = regras_subs.iloc[0]
                            try:
                                cod_filial_subs = df_filiais[df_filiais['Filial'] == regra['Substituta']]['Codigo'].iloc[0]
                            except IndexError:
                                logging.warning(f"Código não encontrado para filial substituta {regra['Substituta']}")
                                cod_filial_subs = '0000'

                            descricao_regra = (
                                f"Regra de Substituição: {regra['Substituta']} recebe coletas de {regra['Grupo Economico']} "
                                f"({regra['Modalidade'] if pd.notna(regra['Modalidade']) else 'Todas'}, "
                                f"{regra['Tipo de carga'] if pd.notna(regra['Tipo de carga']) else 'Todos'})"
                            )
                            if pd.notna(regra.get('Inicial')) and str(regra['Inicial']).strip():
                                descricao_regra += f" ao invés de {regra['Inicial']}"

                            resultados.append({
                                'Origem': municipio,
                                'Incoterm': incoterm,
                                'Tipo_Carga': tipo_carga,
                                'Filial': regra['Substituta'],
                                'Codigo_Filial': f"{int(cod_filial_subs):04}",
                                'KM_ID': None,
                                'Condicao_Atribuicao': descricao_regra,
                                'GRUPO ECONOMICO': f"{int(regra['Grupo Economico']):04}" if pd.notna(regra['Grupo Economico']) else None
                            })
                            filial_encontrada = True

                        if not filial_encontrada:
                            filiais_ativas = df_filiais[df_filiais[coluna_param] == "S"]
                            if not filiais_ativas.empty:
                                dist_filiais = df_dist[
                                    (df_dist['MunicipioOrigem'] == municipio) &
                                    (df_dist['Filial'].isin(filiais_ativas['Filial']))
                                ]
                                dist_filiais_validas = dist_filiais[dist_filiais['KM_ID'].notna()]
                                if not dist_filiais_validas.empty:
                                    mais_proxima = dist_filiais_validas.loc[dist_filiais_validas['KM_ID'].idxmin()]
                                    cod_filial = df_filiais[df_filiais['Filial'] == mais_proxima['Filial']]['Codigo'].values[0]
                                    resultados.append({
                                        'Origem': municipio,
                                        'Incoterm': incoterm,
                                        'Tipo_Carga': tipo_carga,
                                        'Filial': mais_proxima['Filial'],
                                        'Codigo_Filial': f"{int(cod_filial):04}",
                                        'KM_ID': mais_proxima['KM_ID'],
                                        'Condicao_Atribuicao': "Filial compatível com a modalidade",
                                        'GRUPO ECONOMICO': None
                                    })
                                    filial_encontrada = True

                        if not filial_encontrada:
                            filiais_uf = df_filiais[df_filiais['UF'] == uf_municipio]
                            if len(filiais_uf) == 1:
                                filial_unica = filiais_uf.iloc[0]
                                dist_filial = df_dist[
                                    (df_dist['MunicipioOrigem'] == municipio) &
                                    (df_dist['Filial'] == filial_unica['Filial'])
                                ]
                                if not dist_filial.empty and pd.notna(dist_filial['KM_ID'].iloc[0]):
                                    mais_proxima = dist_filial.iloc[0]
                                    cod_filial = filial_unica['Codigo']
                                    resultados.append({
                                        'Origem': municipio,
                                        'Incoterm': incoterm,
                                        'Tipo_Carga': tipo_carga,
                                        'Filial': filial_unica['Filial'],
                                        'Codigo_Filial': f"{int(cod_filial):04}",
                                        'KM_ID': mais_proxima['KM_ID'],
                                        'Condicao_Atribuicao': "Filial única no estado",
                                        'GRUPO ECONOMICO': None
                                    })
                                    filial_encontrada = True

                        if not filial_encontrada:
                            dist_filiais = df_dist[df_dist['MunicipioOrigem'] == municipio]
                            dist_filiais_validas = dist_filiais[dist_filiais['KM_ID'].notna()]
                            if not dist_filiais_validas.empty:
                                mais_proxima = dist_filiais_validas.loc[dist_filiais_validas['KM_ID'].idxmin()]
                                cod_filial = df_filiais[df_filiais['Filial'] == mais_proxima['Filial']]['Codigo'].values[0]
                                resultados.append({
                                    'Origem': municipio,
                                    'Incoterm': incoterm,
                                    'Tipo_Carga': tipo_carga,
                                    'Filial': mais_proxima['Filial'],
                                    'Codigo_Filial': f"{int(cod_filial):04}",
                                    'KM_ID': mais_proxima['KM_ID'],
                                    'Condicao_Atribuicao': "Filial mais próxima (sem restrição)",
                                    'GRUPO ECONOMICO': None
                                })
                                filial_encontrada = True

                        if not filial_encontrada:
                            resultados.append({
                                'Origem': municipio,
                                'Incoterm': incoterm,
                                'Tipo_Carga': tipo_carga,
                                'Filial': None,
                                'Codigo_Filial': None,
                                'KM_ID': None,
                                'Condicao_Atribuicao': "Sem filial disponível",
                                'GRUPO ECONOMICO': None
                            })

                    except Exception as e:
                        logging.error(f"Erro processando {municipio} - {incoterm} - {tipo_carga}: {str(e)}")
                        resultados.append({
                            'Origem': municipio,
                            'Incoterm': incoterm,
                            'Tipo_Carga': tipo_carga,
                            'Filial': None,
                            'Codigo_Filial': None,
                            'KM_ID': None,
                            'Condicao_Atribuicao': "Erro de processamento",
                            'GRUPO ECONOMICO': None
                        })

            df_resultado = pd.DataFrame(resultados)
            buffer = BytesIO()
            df_resultado.to_excel(buffer, index=False)
            buffer.seek(0)

            st.success("✅ Processamento concluído!")
            st.dataframe(df_resultado)

            st.download_button(
                label="📥 Baixar Resultado em Excel",
                data=buffer,
                file_name="resultado_final.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

            st.info(f"📄 Log salvo em: {os.path.abspath(log_file)}")
