import streamlit as st
import pandas as pd
import os
import logging
from io import BytesIO, StringIO
import zipfile # Ainda necessário para o zip se você decidir voltar
import datetime

# === CONFIGURAÇÕES INICIAIS DO STREAMLIT ===
st.set_page_config(page_title="Roteirização com Substituição", layout="wide")

# === CONFIGURAÇÃO DE LOG ===
log_stream = StringIO()
for handler in logging.root.handlers[:]:
    logging.root.removeHandler(handler)
logging.basicConfig(stream=log_stream, level=logging.INFO, format='%(asctime)s - %(message)s')
logger = logging.getLogger()

st.title("📦 Roteirização com Regras de Substituição")

# Colunas esperadas para a base de parâmetros
colunas_base_parametros = {
    'Substituta': str,
    'Inicial': str,
    'Recebe': str,
    'UF': str,
    'Grupo Economico': str,
    'Modalidade': str,
    'Tipo de carga': str,
    'Data': 'datetime64[ns]'
}

# === CARREGAMENTO DAS BASES FIXAS INTERNAS (Distâncias e Filiais) ===
try:
    df_dist = pd.read_excel("municipios_distanciasreais.xlsx")
    df_filiais = pd.read_excel("filiais_geocodificadas.xlsx")
    st.success("📍 Bases internas de distâncias e filiais carregadas com sucesso.")
    logger.info("Bases internas de distâncias e filiais carregadas com sucesso.")
except Exception as e:
    st.error(f"❌ Erro ao carregar arquivos internos (distâncias/filiais): {e}")
    logger.error(f"Erro ao carregar arquivos internos (distâncias/filiais): {e}")
    st.stop()

# --- Lista de Filiais para Dropdowns (se necessário, mantém a lógica) ---
if 'Filial' in df_filiais.columns:
    lista_filiais = [''] + sorted(df_filiais['Filial'].astype(str).unique().tolist())
else:
    lista_filiais = ['']
    st.warning("Coluna 'Filial' não encontrada em 'filiais_geocodificadas.xlsx'. Dropdowns de Filial podem não funcionar.")

# === CARREGAMENTO E EXIBIÇÃO DA BASE DE PARÂMETROS PADRÃO VIA UPLOAD ===
st.header("📄 Carregar Parâmetros Contratuais")

# Widget de upload de arquivo
uploaded_file = st.file_uploader(
    "Faça o upload do arquivo 'parametros_contrato.xlsx'",
    type=["xlsx"],
    help="O arquivo Excel deve conter as regras de substituição."
)

# Inicializa df_padrao_parametros como um DataFrame vazio com as colunas corretas
df_padrao_parametros = pd.DataFrame(columns=list(colunas_base_parametros.keys()))
for col, dtype in colunas_base_parametros.items():
    df_padrao_parametros[col] = pd.Series(dtype=dtype)
    if dtype == 'datetime64[ns]':
        df_padrao_parametros[col] = pd.NaT
    else:
        df_padrao_parametros[col].fillna('', inplace=True)

if uploaded_file is not None:
    try:
        # Ler o arquivo Excel carregado
        df_padrao_parametros = pd.read_excel(uploaded_file)
        
        # Processamento das colunas (mesma lógica de antes para garantir tipos e preencher vazios)
        for col, dtype in colunas_base_parametros.items():
            if col in df_padrao_parametros.columns:
                if dtype == 'datetime64[ns]':
                    df_padrao_parametros[col] = pd.to_datetime(df_padrao_parametros[col], errors='coerce')
                    if df_padrao_parametros[col].isnull().any():
                        logger.warning(f"Coluna '{col}' do arquivo carregado contém valores que não puderam ser convertidos para data (NaN/NaT).")
                else:
                    df_padrao_parametros[col] = df_padrao_parametros[col].astype(str).fillna('')
            else:
                logger.warning(f"Coluna '{col}' esperada mas não encontrada no arquivo carregado. Adicionando-a com valores vazios.")
                df_padrao_parametros[col] = pd.Series(dtype=dtype, index=df_padrao_parametros.index)
                if dtype == 'datetime64[ns]':
                    df_padrao_parametros[col] = pd.NaT
                else:
                    df_padrao_parametros[col].fillna('', inplace=True)

        if df_padrao_parametros.empty:
            st.warning("O arquivo carregado está vazio ou não contém dados válidos após o processamento.")
            logger.warning("DataFrame de parâmetros padrão vazio ou inválido após carregamento e processamento do arquivo.")
            # Reseta df_padrao_parametros para vazio com colunas corretas
            df_padrao_parametros = pd.DataFrame(columns=list(colunas_base_parametros.keys()))
            for col, dtype in colunas_base_parametros.items():
                df_padrao_parametros[col] = pd.Series(dtype=dtype)
                if dtype == 'datetime64[ns]':
                    df_padrao_parametros[col] = pd.NaT
                else:
                    df_padrao_parametros[col].fillna('', inplace=True)
        else:
            st.success(f"Parâmetros contratuais carregados com sucesso! {df_padrao_parametros.shape[0]} regras.")
            logger.info(f"Parâmetros contratuais carregados com sucesso! {df_padrao_parametros.shape[0]} regras.")

    except Exception as e:
        st.error(f"Erro ao ler o arquivo Excel carregado: {e}")
        logger.error(f"Erro ao ler o arquivo Excel carregado: {e}")
        df_padrao_parametros = pd.DataFrame(columns=list(colunas_base_parametros.keys()))
        for col, dtype in colunas_base_parametros.items():
            df_padrao_parametros[col] = pd.Series(dtype=dtype)
            if dtype == 'datetime64[ns]':
                df_padrao_parametros[col] = pd.NaT
            else:
                df_padrao_parametros[col].fillna('', inplace=True)
else:
    st.info("Aguardando o upload do arquivo de parâmetros contratuais.")

st.dataframe(df_padrao_parametros, use_container_width=True, height=200)
st.divider()

# === BOTÃO PARA PROCESSAR ===
if st.button("🚀 Rodar Roteirização"):
    log_stream.seek(0)
    log_stream.truncate(0)

    # Verifica se o df_padrao_parametros foi carregado
    if df_padrao_parametros.empty:
        st.warning("Por favor, faça o upload do arquivo de parâmetros contratuais antes de rodar a roteirização.")
        logger.warning("Tentativa de rodar roteirização sem arquivo de parâmetros contratuais carregado.")
        st.stop()
    
    df_grupos_final = df_padrao_parametros.copy()
    logger.info("Base de parâmetros padrão utilizada para processamento (sem regras de usuário).")

    # Validação e limpeza das regras (mesma lógica anterior)
    df_grupos_final_validado = df_grupos_final.copy()
    for col, dtype in colunas_base_parametros.items():
        if col in df_grupos_final_validado.columns:
            if dtype == 'datetime64[ns]':
                df_grupos_final_validado[col] = pd.to_datetime(df_grupos_final_validado[col], errors='coerce')
            else:
                df_grupos_final_validado[col] = df_grupos_final_validado[col].astype(str).replace(r'^\s*$', '', regex=True)
        else:
            df_grupos_final_validado[col] = pd.Series(dtype=dtype, index=df_grupos_final_validado.index)
            if dtype == 'datetime64[ns]':
                df_grupos_final_validado[col] = pd.NaT
            else:
                df_grupos_final_validado[col].fillna('', inplace=True)

    for col in ['Substituta', 'Recebe', 'UF']:
        df_grupos_final_validado[col] = df_grupos_final_validado[col].replace('', pd.NA)

    df_grupos_final_validado = df_grupos_final_validado.dropna(subset=['Substituta', 'Recebe', 'UF'])

    for col, dtype in colunas_base_parametros.items():
        if col in df_grupos_final_validado.columns and dtype == str:
            df_grupos_final_validado[col].fillna('', inplace=True)

    if df_grupos_final_validado.empty:
        st.error("A base de parâmetros carregada não contém regras válidas após a validação. Certifique-se de que 'Substituta', 'Recebe' e 'UF' não estejam vazios para suas regras.")
        logger.error("Base de parâmetros final sem regras válidas para processamento após validação.")
        st.stop()

    with st.spinner("Processando..."):
        progress_text = "Processando roteirização... Aguarde."
        my_bar = st.progress(0, text=progress_text)

        modalidades = [
            ("FCA", "Fracionado", "FCA/Fracionado"),
            ("FCA", "Lotação", "FCA/Lotação"),
            ("EXW", "Fracionado", "EXW/Fracionado"),
            ("EXW", "Lotação", "EXW/Lotação")
        ]
        
        def buscar_regras_substituicao_multiplas(df_regras_concat, uf, modalidade, tipo_carga, grupo_economico_origem=None):
            regras_aplicaveis = df_regras_concat[
                (df_regras_concat['UF'] == uf) &
                (df_regras_concat['Recebe'] == 'S')
            ].copy()

            regras_aplicaveis = regras_aplicaveis[
                (regras_aplicaveis['Modalidade'].isin(['', modalidade])) &
                (regras_aplicaveis['Tipo de carga'].isin(['', tipo_carga]))
            ]
            
            if grupo_economico_origem and str(grupo_economico_origem).strip() != '':
                regras_aplicaveis = regras_aplicaveis[
                    (regras_aplicaveis['Grupo Economico'] == '') |
                    (regras_aplicaveis['Grupo Economico'] == grupo_economico_origem)
                ]
            else:
                regras_aplicaveis = regras_aplicaveis[regras_aplicaveis['Grupo Economico'] == '']

            return regras_aplicaveis

        municipios = df_dist['MunicipioOrigem'].unique()
        total_municipios = len(municipios)
        resultados = []

        for i, municipio in enumerate(municipios):
            uf_municipio = municipio.split('-')[-1].strip()
            grupo_economico_municipio = None # Mantido None para replicar o comportamento do codigo 2

            for incoterm, tipo_carga, coluna_param in modalidades:
                try:
                    filial_encontrada_padrao = False
                    mais_proxima = None
                    cod_filial_padrao = None
                    condicao_padrao = None

                    # 1. Filial compatível com a modalidade (mais próxima)
                    filiais_ativas = df_filiais[df_filiais[coluna_param] == "S"]
                    if not filiais_ativas.empty:
                        dist_filiais = df_dist[
                            (df_dist['MunicipioOrigem'] == municipio) &
                            (df_dist['Filial'].isin(filiais_ativas['Filial']))
                        ]
                        dist_filiais_validas = dist_filiais[dist_filiais['KM_ID'].notna()]
                        
                        if not dist_filiais_validas.empty:
                            mais_proxima = dist_filiais_validas.loc[dist_filiais_validas['KM_ID'].idxmin()]
                            cod_filial_padrao = df_filiais[df_filiais['Filial'] == mais_proxima['Filial']]['Codigo'].values[0]
                            filial_encontrada_padrao = True
                            condicao_padrao = "Filial compatível com a modalidade"
                            logger.info(f"Filial compatível com modalidade encontrada para {municipio} ({incoterm}/{tipo_carga}): {mais_proxima['Filial']} (KM_ID: {mais_proxima['KM_ID']}).")

                    # 2. Única filial no estado (se ainda não encontrou)
                    if not filial_encontrada_padrao:
                        filiais_uf = df_filiais[df_filiais['UF'] == uf_municipio]
                        if len(filiais_uf) == 1:
                            filial_unica = filiais_uf.iloc[0]
                            dist_filial = df_dist[
                                (df_dist['MunicipioOrigem'] == municipio) &
                                (df_dist['Filial'] == filial_unica['Filial'])
                            ]
                            if not dist_filial.empty and pd.notna(dist_filial['KM_ID'].iloc[0]):
                                mais_proxima = dist_filial.iloc[0]
                                cod_filial_padrao = filial_unica['Codigo']
                                filial_encontrada_padrao = True
                                condicao_padrao = "Filial única no estado"
                                logger.info(f"Filial única no estado encontrada para {municipio} ({incoterm}/{tipo_carga}): {filial_unica['Filial']} (KM_ID: {mais_proxima['KM_ID']}).")

                    # 3. Filial mais próxima independente de tudo (se ainda não encontrou)
                    if not filial_encontrada_padrao:
                        dist_filiais = df_dist[df_dist['MunicipioOrigem'] == municipio]
                        dist_filiais_validas = dist_filiais[dist_filiais['KM_ID'].notna()]
                        if not dist_filiais_validas.empty:
                            mais_proxima = dist_filiais_validas.loc[dist_filiais_validas['KM_ID'].idxmin()]
                            cod_filial_padrao = df_filiais[df_filiais['Filial'] == mais_proxima['Filial']]['Codigo'].values[0]
                            filial_encontrada_padrao = True
                            condicao_padrao = "Filial mais próxima (sem restrição)"
                            logger.info(f"Filial mais próxima (sem restrição) encontrada para {municipio} ({incoterm}/{tipo_carga}): {mais_proxima['Filial']} (KM_ID: {mais_proxima['KM_ID']}).")

                    # --- Adiciona o resultado padrão (se encontrado) ---
                    if filial_encontrada_padrao:
                        resultados.append({
                            'Origem': municipio,
                            'Incoterm': incoterm,
                            'Tipo_Carga': tipo_carga,
                            'Filial': mais_proxima['Filial'],
                            'Codigo_Filial': f"{int(cod_filial_padrao):04}",
                            'KM_ID': mais_proxima['KM_ID'],
                            'Condicao_Atribuicao': condicao_padrao,
                            'GRUPO ECONOMICO': None
                        })
                    else:
                        resultados.append({
                            'Origem': municipio,
                            'Incoterm': incoterm,
                            'Tipo_Carga': tipo_carga,
                            'Filial': None,
                            'Codigo_Filial': None,
                            'KM_ID': None,
                            'Condicao_Atribuicao': "Sem filial disponível (Padrão)",
                            'GRUPO ECONOMICO': None
                        })
                        logger.warning(f"Nenhuma filial padrão encontrada para {municipio} ({incoterm}/{tipo_carga}).")

                    # --- Agora, aplica as regras de substituição como resultados ADICIONAIS ---
                    regras_subs = buscar_regras_substituicao_multiplas(
                        df_grupos_final_validado, uf_municipio, incoterm, tipo_carga, grupo_economico_municipio
                    )

                    if not regras_subs.empty:
                        for _, regra in regras_subs.iterrows():
                            try:
                                cod_filial_subs = df_filiais[df_filiais['Filial'].astype(str) == str(regra['Substituta'])]['Codigo'].iloc[0]
                                logger.info(f"Regra de substituição aplicável encontrada para {municipio} ({incoterm}/{tipo_carga}): Filial {regra['Substituta']} (Código: {int(cod_filial_subs):04}).")
                            except IndexError:
                                logger.warning(f"Código não encontrado para filial substituta {regra['Substituta']} para {municipio} ({incoterm}/{tipo_carga}). Usando '0000'.")
                                cod_filial_subs = '0000'

                            grupo_economico_str = str(regra['Grupo Economico']) if pd.notna(regra['Grupo Economico']) and str(regra['Grupo Economico']).strip() != '' else 'qualquer grupo'
                            modalidade_str = regra['Modalidade'] if pd.notna(regra['Modalidade']) and str(regra['Modalidade']).strip() != '' else 'Todas as modalidades'
                            tipo_carga_str = regra['Tipo de carga'] if pd.notna(regra['Tipo de carga']) and str(regra['Tipo de carga']).strip() != '' else 'Todos os tipos de carga'

                            descricao_regra = (
                                f"Regra de Substituição: {regra['Substituta']} recebe coletas de {grupo_economico_str} "
                                f"({modalidade_str}, {tipo_carga_str})"
                            )
                            if pd.notna(regra.get('Inicial')) and str(regra['Inicial']).strip() != '':
                                descricao_regra += f" ao invés de {regra['Inicial']}"
                            
                            if pd.notna(regra.get('Data')):
                                descricao_regra += f" (Regra de {regra['Data'].strftime('%d/%m/%Y')})"
                            
                            resultados.append({
                                'Origem': municipio,
                                'Incoterm': incoterm,
                                'Tipo_Carga': tipo_carga,
                                'Filial': regra['Substituta'],
                                'Codigo_Filial': f"{int(cod_filial_subs):04}",
                                'KM_ID': None,
                                'Condicao_Atribuicao': descricao_regra,
                                'GRUPO ECONOMICO': f"{int(float(grupo_economico_str)):04}" if grupo_economico_str != 'qualquer grupo' and str(grupo_economico_str).replace('.', '', 1).isdigit() else None
                            })
                    else:
                        logger.info(f"Nenhuma regra de substituição aplicável para {municipio} ({incoterm}/{tipo_carga}).")

                except Exception as e:
                    logger.error(f"Erro processando {municipio} - {incoterm} - {tipo_carga}: {str(e)}")
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
                    
                percent_complete = min(100, int((i + 1) / total_municipios * 100))
                my_bar.progress(percent_complete, text=f"{progress_text} {percent_complete}% Concluído.")
                
        my_bar.progress(100, text=f"{progress_text} 100% Concluído.")

        df_resultado = pd.DataFrame(resultados)
        
        st.success("✅ Processamento concluído!")
        st.dataframe(df_resultado, use_container_width=True)

        # === NOVO: Botoões de Download Separados ===
        
        # Preparar o Excel para download
        excel_buffer = BytesIO()
        df_resultado.to_excel(excel_buffer, index=False)
        excel_buffer.seek(0) # Retorna o ponteiro para o início do buffer

        # Preparar o Log para download
        log_content = log_stream.getvalue()
        log_buffer = BytesIO(log_content.encode('utf-8')) # Codificar para bytes

        col_excel, col_log = st.columns(2) # Cria duas colunas para os botões

        with col_excel:
            st.download_button(
                label="📥 Baixar Resultados (Excel)",
                data=excel_buffer,
                file_name="resultado_roteirizacao.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key="download_excel" # Chave única para o botão
            )

        with col_log:
            st.download_button(
                label="📄 Baixar Log",
                data=log_buffer,
                file_name="log_roteirizacao.log",
                mime="text/plain",
                key="download_log" # Chave única para o botão
            )
        # === FIM DOS NOVOS BOTÕES ===

        st.info("Log de processamento:")
        st.text_area("Visualizar Log", log_content, height=200)
        logger.info("Aplicação concluída.")
