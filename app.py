import streamlit as st
import pandas as pd
import os
import logging
from tqdm import tqdm
from io import BytesIO

st.set_page_config(page_title="Roteirização com Substituição", layout="wide")

# === CONFIGURAÇÕES ===
st.title("📦 Roteirização com Regras de Substituição")
log_file = 'log_filiais_proximas.log'
logging.basicConfig(filename=log_file, level=logging.INFO, format='%(asctime)s - %(message)s')

# === UPLOAD DE ARQUIVOS ===
st.header("📥 Upload de Arquivos")

dist_path = st.file_uploader("1. Distâncias reais (municipios_distanciasreais.xlsx)", type=["xlsx"])
filial_path = st.file_uploader("2. Filiais geocodificadas (filiais_geocodificadas.xlsx)", type=["xlsx"])
grupo_economico_file = st.file_uploader("3. Regras de Substituição (parametros_contrato.xlsx)", type=["xlsx"], help="Ou edite a base manualmente abaixo.")

# === MODELO DE BASE DE SUBSTITUIÇÃO ===
colunas_base = ['Substituta', 'Inicial', 'Recebe', 'UF', 'Grupo Economico', 'Modalidade', 'Tipo de carga']

if grupo_economico_file:
    df_grupos = pd.read_excel(grupo_economico_file)
    st.success("📄 Regras de substituição carregadas com sucesso!")
else:
    st.info("✏️ Edite a base de substituição abaixo.")
    df_grupos = pd.DataFrame(columns=colunas_base)

df_grupos_editado = st.data_editor(df_grupos, num_rows="dynamic", use_container_width=True, key="regras_editadas")

st.divider()

# === BOTÃO PARA PROCESSAR ===
if st.button("🚀 Rodar Roteirização"):
    if not dist_path or not filial_path:
        st.error("Por favor, envie os arquivos de distâncias e filiais.")
    else:
        with st.spinner("Processando..."):
            # Carregar dados
            df_dist = pd.read_excel(dist_path)
            df_filiais = pd.read_excel(filial_path)
            df_grupos = df_grupos_editado.copy()

            modalidades = [
                ("FCA", "Fracionado", "FCA/Fracionado"),
                ("FCA", "Lotação", "FCA/Lotação"),
                ("EXW", "Fracionado", "EXW/Fracionado"),
                ("EXW", "Lotação", "EXW/Lotação")
            ]

            def buscar_regras_substituicao(df_regras, uf, modalidade, tipo_carga):
                regras = df_regras[
                    (df_regras['UF'] == uf) & (df_regras['Recebe'] == 'S')
                ]
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

                        # 1. Compatível com a modalidade
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
                                filial_encontrada = True
                                condicao = "Filial compatível com a modalidade"

                        # 2. Única filial no estado
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
                                    filial_encontrada = True
                                    condicao = "Filial única no estado"

                        # 3. Filial mais próxima (sem restrição)
                        if not filial_encontrada:
                            dist_filiais = df_dist[df_dist['MunicipioOrigem'] == municipio]
                            dist_filiais_validas = dist_filiais[dist_filiais['KM_ID'].notna()]
                            if not dist_filiais_validas.empty:
                                mais_proxima = dist_filiais_validas.loc[dist_filiais_validas['KM_ID'].idxmin()]
                                cod_filial = df_filiais[df_filiais['Filial'] == mais_proxima['Filial']]['Codigo'].values[0]
                                filial_encontrada = True
                                condicao = "Filial mais próxima (sem restrição)"

                        if filial_encontrada:
                            resultados.append({
                                'Origem': municipio,
                                'Incoterm': incoterm,
                                'Tipo_Carga': tipo_carga,
                                'Filial': mais_proxima['Filial'],
                                'Codigo_Filial': f"{int(cod_filial):04}",
                                'KM_ID': mais_proxima['KM_ID'],
                                'Condicao_Atribuicao': condicao,
                                'GRUPO ECONOMICO': None
                            })

                            # Regras de Substituição
                            regras_subs = buscar_regras_substituicao(df_grupos, uf_municipio, incoterm, tipo_carga)
                            for _, regra in regras_subs.iterrows():
                                try:
                                    cod_filial_subs = df_filiais[df_filiais['Filial'] == regra['Substituta']]['Codigo'].iloc[0]
                                except:
                                    logging.warning(f"Código não encontrado para filial {regra['Substituta']}")
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
                        else:
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

            # Resultado final
            df_resultado = pd.DataFrame(resultados)

            # Download
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
