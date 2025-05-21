import streamlit as st
import pandas as pd
import os
import logging
from io import BytesIO, StringIO
import zipfile
import datetime # Importa datetime para pd.Timestamp.now()

# === CONFIGURAÇÕES INICIAIS DO STREAMLIT (DEVE SER A PRIMEIRA COISA APÓS OS IMPORTS) ===
st.set_page_config(page_title="Roteirização com Substituição", layout="wide")

# === CONFIGURAÇÃO DE LOG ===
log_stream = StringIO()
for handler in logging.root.handlers[:]:
    logging.root.removeHandler(handler)
logging.basicConfig(stream=log_stream, level=logging.INFO, format='%(asctime)s - %(message)s')
logger = logging.getLogger()

st.title("📦 Roteirização com Regras de Substituição")

# === CAMINHOS DOS ARQUIVOS DE PARÂMETROS ===
CAMINHO_PARAMETROS_PADRAO = "parametros_contrato.xlsx"
# CAMINHO_PARAMETETROS_USUARIO foi removido

# Colunas esperadas para as bases de parâmetros, incluindo a nova coluna 'Data'
colunas_base_parametros = {
    'Substituta': str,
    'Inicial': str,
    'Recebe': str,
    'UF': str,
    'Grupo Economico': str,
    'Modalidade': str,
    'Tipo de carga': str,
    'Data': 'datetime64[ns]' # Adicionada a coluna de data aqui
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

# --- ESTA É A PARTE QUE PEGA AS FILIAIS DA BASE "Filiais_Geocodificadas" ---
if 'Filial' in df_filiais.columns:
    # Garante que a lista contenha strings e ordena
    lista_filiais = [''] + sorted(df_filiais['Filial'].astype(str).unique().tolist())
else:
    lista_filiais = ['']
    st.warning("Coluna 'Filial' não encontrada em 'filiais_geocodificadas.xlsx'. Dropdowns de Filial podem não funcionar.")
# --- FIM DA PARTE ---


# === CARREGAMENTO E EXIBIÇÃO DA BASE DE PARÂMETROS PADRÃO ===
st.header("📄 Parâmetros Contratuais Padrão")
try:
    # Tenta ler o arquivo Excel completamente primeiro
    df_padrao_parametros = pd.read_excel(CAMINHO_PARAMETROS_PADRAO)
    
    # Adiciona log para verificar o DataFrame após a leitura inicial
    logger.info(f"DataFrame df_padrao_parametros lido inicialmente. Formato: {df_padrao_parametros.shape}, Colunas: {df_padrao_parametros.columns.tolist()}")
    
    # Itera sobre as colunas esperadas para garantir os tipos e preencher vazios
    for col, dtype in colunas_base_parametros.items():
        if col in df_padrao_parametros.columns:
            if dtype == 'datetime64[ns]':
                # Converte para datetime, forçando erros para NaT
                df_padrao_parametros[col] = pd.to_datetime(df_padrao_parametros[col], errors='coerce')
                # Log para verificar o resultado da conversão de data
                if df_padrao_parametros[col].isnull().any():
                    logger.warning(f"Coluna '{col}' em parametros_contrato.xlsx contém valores que não puderam ser convertidos para data (NaN/NaT).")
            else:
                # Converte para string e preenche quaisquer NaN/None com string vazia
                df_padrao_parametros[col] = df_padrao_parametros[col].astype(str).fillna('')
        else:
            # Se a coluna esperada não existir no DataFrame lido, adicione-a
            logger.warning(f"Coluna '{col}' esperada mas não encontrada em '{CAMINHO_PARAMETROS_PADRAO}'. Adicionando-a com valores vazios.")
            df_padrao_parametros[col] = pd.Series(dtype=dtype, index=df_padrao_parametros.index)
            if dtype == 'datetime64[ns]':
                df_padrao_parametros[col] = pd.NaT # Not a Time para datas vazias
            else:
                df_padrao_parametros[col].fillna('', inplace=True)

    # Verifica se o DataFrame padrão está vazio após o processamento
    if df_padrao_parametros.empty:
        st.warning(f"O arquivo '{CAMINHO_PARAMETROS_PADRAO}' foi carregado, mas está vazio ou não contém dados válidos após o processamento.")
        logger.warning(f"DataFrame de parâmetros padrão vazio ou inválido após carregamento e processamento.")
        # Reinicia o DataFrame para um estado vazio com colunas corretas
        df_padrao_parametros = pd.DataFrame(columns=list(colunas_base_parametros.keys()))
        for col, dtype in colunas_base_parametros.items():
            df_padrao_parametros[col] = pd.Series(dtype=dtype)
            if dtype == 'datetime64[ns]':
                df_padrao_parametros[col] = pd.NaT
            else:
                df_padrao_parametros[col].fillna('', inplace=True)
    else:
        st.info(f"Parâmetros contratuais padrão carregados de '{CAMINHO_PARAMETROS_PADRAO}'. {df_padrao_parametros.shape[0]} regras carregadas.")
        logger.info(f"Parâmetros contratuais padrão carregados de '{CAMINHO_PARAMETROS_PADRAO}'. {df_padrao_parametros.shape[0]} regras carregadas.")

except FileNotFoundError:
    st.warning(f"Arquivo '{CAMINHO_PARAMETROS_PADRAO}' não encontrado. Criando DataFrame padrão vazio.")
    logger.warning(f"Arquivo '{CAMINHO_PARAMETROS_PADRAO}' não encontrado. Criando DataFrame padrão vazio.")
    df_padrao_parametros = pd.DataFrame(columns=list(colunas_base_parametros.keys()))
    for col, dtype in colunas_base_parametros.items():
        df_padrao_parametros[col] = pd.Series(dtype=dtype)
        if dtype == 'datetime64[ns]':
            df_padrao_parametros[col] = pd.NaT
        else:
            df_padrao_parametros[col].fillna('', inplace=True)
except Exception as e:
    st.error(f"Erro ao carregar o arquivo de parâmetros padrão '{CAMINHO_PARAMETROS_PADRAO}': {e}")
    logger.error(f"Erro ao carregar o arquivo de parâmetros padrão '{CAMINHO_PARAMETROS_PADRAO}': {e}")
    df_padrao_parametros = pd.DataFrame(columns=list(colunas_base_parametros.keys()))
    for col, dtype in colunas_base_parametros.items():
        df_padrao_parametros[col] = pd.Series(dtype=dtype)
        if dtype == 'datetime64[ns]':
            df_padrao_parametros[col] = pd.NaT
        else:
            df_padrao_parametros[col].fillna('', inplace=True)

st.dataframe(df_padrao_parametros, use_container_width=True, height=200)
st.divider()

# === BOTÃO PARA PROCESSAR (AGORA COM AS DUAS BASES CONCATENADAS) ===
if st.button("🚀 Rodar Roteirização"):
    log_stream.seek(0)
    log_stream.truncate(0)

    # A base de parâmetros final para o processamento será APENAS df_padrao_parametros
    df_grupos_final = df_padrao_parametros.copy()
    logger.info("Base de parâmetros padrão utilizada para processamento (sem regras de usuário).")

    # Remova linhas que não tenham os campos obrigatórios para uma regra
    df_grupos_final_validado = df_grupos_final.copy()
    for col, dtype in colunas_base_parametros.items():
        if col in df_grupos_final_validado.columns:
            if dtype == 'datetime64[ns]':
                df_grupos_final_validado[col] = pd.to_datetime(df_grupos_final_validado[col], errors='coerce')
            else:
                # Converte para string e depois trata strings vazias para NA para dropna
                df_grupos_final_validado[col] = df_grupos_final_validado[col].astype(str).replace(r'^\s*$', '', regex=True)
        else:
            df_grupos_final_validado[col] = pd.Series(dtype=dtype, index=df_grupos_final_validado.index)
            if dtype == 'datetime64[ns]':
                df_grupos_final_validado[col] = pd.NaT
            else:
                df_grupos_final_validado[col].fillna('', inplace=True)

    # Trata as strings vazias para NA para facilitar o dropna subset
    for col in ['Substituta', 'Recebe', 'UF']:
        df_grupos_final_validado[col] = df_grupos_final_validado[col].replace('', pd.NA)

    df_grupos_final_validado = df_grupos_final_validado.dropna(subset=['Substituta', 'Recebe', 'UF'])

    # Volta as colunas de string de NA para vazio, mantendo NaT para datas
    for col, dtype in colunas_base_parametros.items():
        if col in df_grupos_final_validado.columns and dtype == str:
            df_grupos_final_validado[col].fillna('', inplace=True)


    if df_grupos_final_validado.empty:
        st.error("A base de parâmetros final não contém regras válidas. Certifique-se de que 'Substituta', 'Recebe' e 'UF' não estejam vazios.")
        logger.error("Base de parâmetros final sem regras válidas para processamento.")
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
        
        # --- FUNÇÃO MODIFICADA PARA BUSCAR REGRAS DE SUBSTITUIÇÃO (RETORNA TODAS AS APLICÁVEIS) ---
        # ATENÇÃO: A lógica da função foi ajustada para ser mais similar ao `codigo 2`
        # para a verificação de valores nulos/vazios em Modalidade e Tipo de Carga.
        # No `codigo 2` eles verificavam `isna()`, aqui estamos verificando `''` (string vazia)
        # que é o que Streamlit tende a produzir para campos não preenchidos de selectbox.
        def buscar_regras_substituicao_multiplas(df_regras_concat, uf, modalidade, tipo_carga, grupo_economico_origem=None):
            regras_aplicaveis = df_regras_concat[
                (df_regras_concat['UF'] == uf) &
                (df_regras_concat['Recebe'] == 'S')
            ].copy()

            # Filtro para Modalidade e Tipo de Carga (vazio '' significa "qualquer")
            # Este é o ajuste chave para replicar o comportamento de isna() do codigo 2
            regras_aplicaveis = regras_aplicaveis[
                (regras_aplicaveis['Modalidade'].isin(['', modalidade])) &
                (regras_aplicaveis['Tipo de carga'].isin(['', tipo_carga]))
            ]
            
            # Filtro para Grupo Economico (vazio '' significa "qualquer")
            # Se grupo_economico_origem não for vazio, busca por regras com esse grupo OU regras sem grupo especificado.
            # Se grupo_economico_origem for vazio, busca APENAS regras sem grupo especificado.
            # No contexto do `codigo 2`, o Grupo Econômico era sempre `None` na busca,
            # então este filtro resultaria em `regras_aplicaveis['Grupo Economico'] == ''`.
            # Para replicar o `codigo 2` exatamente, certifique-se de que `grupo_economico_origem` seja sempre `None`.
            if grupo_economico_origem and str(grupo_economico_origem).strip() != '':
                regras_aplicaveis = regras_aplicaveis[
                    (regras_aplicaveis['Grupo Economico'] == '') |
                    (regras_aplicaveis['Grupo Economico'] == grupo_economico_origem)
                ]
            else:
                regras_aplicaveis = regras_aplicaveis[regras_aplicaveis['Grupo Economico'] == '']

            return regras_aplicaveis
        # --- FIM DA FUNÇÃO MODIFICADA ---

        municipios = df_dist['MunicipioOrigem'].unique()
        total_municipios = len(municipios)
        resultados = []

        for i, municipio in enumerate(municipios):
            uf_municipio = municipio.split('-')[-1].strip()
            
            # Para replicar o `codigo 2` exatamente, o Grupo Econômico do município é None
            # porque não há uma fonte de dados para ele no `codigo 2` no loop de processamento.
            grupo_economico_municipio = None 

            for incoterm, tipo_carga, coluna_param in modalidades:
                try:
                    filial_encontrada_padrao = False
                    mais_proxima = None # Variável para armazenar a filial padrão encontrada
                    cod_filial_padrao = None
                    condicao_padrao = None

                    # --- Lógica de Atribuição Padrão (Prioridades do segundo script) ---

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
                            'GRUPO ECONOMICO': None # Grupo Econômico é da regra, não da atribuição padrão
                        })
                    else:
                        # Se nenhuma filial padrão for encontrada, adiciona uma linha indicando isso
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
                    # Usando apenas df_padrao_parametros, replicando o comportamento do codigo 2 com df_grupos
                    regras_subs = buscar_regras_substituicao_multiplas(
                        df_grupos_final_validado, uf_municipio, incoterm, tipo_carga, grupo_economico_municipio
                    )

                    if not regras_subs.empty:
                        for _, regra in regras_subs.iterrows():
                            try:
                                # Ajuste para garantir que 'Substituta' seja string antes da comparação
                                cod_filial_subs = df_filiais[df_filiais['Filial'].astype(str) == str(regra['Substituta'])]['Codigo'].iloc[0]
                                logger.info(f"Regra de substituição aplicável encontrada para {municipio} ({incoterm}/{tipo_carga}): Filial {regra['Substituta']} (Código: {int(cod_filial_subs):04}).")
                            except IndexError:
                                logger.warning(f"Código não encontrado para filial substituta {regra['Substituta']} para {municipio} ({incoterm}/{tipo_carga}). Usando '0000'.")
                                cod_filial_subs = '0000'

                            # Ajuste para formatação da descrição ser mais fiel ao codigo 2
                            # Nota: O codigo 2 usa `pd.notna()` e o `codigo 1` usa `str().strip() != ''`
                            # para campos de texto. Mantenho a do `codigo 1` por ser mais robusta com Streamlit.
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
                                'KM_ID': None, # Regra de substituição não tem KM_ID diretamente
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
        
        zip_buffer = BytesIO()
        with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zf:
            excel_buffer = BytesIO()
            df_resultado.to_excel(excel_buffer, index=False)
            excel_buffer.seek(0)
            zf.writestr("resultado_roteirizacao.xlsx", excel_buffer.getvalue())
            logger.info("Resultado da roteirização adicionado ao ZIP.")

            log_content = log_stream.getvalue()
            zf.writestr("log_roteirizacao.log", log_content)
            logger.info("Log da roteirização adicionado ao ZIP.")
            
        zip_buffer.seek(0)

        st.success("✅ Processamento concluído!")
        st.dataframe(df_resultado)

        st.download_button(
            label="📥 Baixar Resultado (Excel + Log .zip)",
            data=zip_buffer,
            file_name="roteirizacao_completa.zip",
            mime="application/zip"
        )

        st.info("Log de processamento:")
        st.text_area("Visualizar Log", log_content, height=200)
        logger.info("Aplicação concluída.")
