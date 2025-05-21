import streamlit as st
import pandas as pd
import os
import logging
from io import BytesIO, StringIO
import zipfile

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
CAMINHO_PARAMETETROS_USUARIO = "parametros_usuario.xlsx"

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

# === CARREGAMENTO E EDIÇÃO DA BASE DE PARÂMETROS DO USUÁRIO ===
st.header("✏️ Parâmetros Contratuais do Usuário (Editável)")

df_grupos_usuario = pd.DataFrame(columns=list(colunas_base_parametros.keys()))
for col, dtype in colunas_base_parametros.items():
    df_grupos_usuario[col] = pd.Series(dtype=dtype)
    if dtype == 'datetime64[ns]':
        df_grupos_usuario[col] = pd.NaT
    else:
        df_grupos_usuario[col].fillna('', inplace=True)

if os.path.exists(CAMINHO_PARAMETETROS_USUARIO):
    try:
        df_loaded = pd.read_excel(CAMINHO_PARAMETETROS_USUARIO)
        # Log para verificar o DataFrame do usuário após a leitura inicial
        logger.info(f"DataFrame df_grupos_usuario lido inicialmente. Formato: {df_loaded.shape}, Colunas: {df_loaded.columns.tolist()}")

        for col, dtype in colunas_base_parametros.items():
            if col in df_loaded.columns:
                if dtype == 'datetime64[ns]':
                    df_loaded[col] = pd.to_datetime(df_loaded[col], errors='coerce')
                    if df_loaded[col].isnull().any():
                        logger.warning(f"Coluna '{col}' em parametros_usuario.xlsx contém valores que não puderam ser convertidos para data (NaN/NaT).")
                else:
                    df_loaded[col] = df_loaded[col].astype(dtype).fillna('')
            else:
                logger.warning(f"Coluna '{col}' esperada mas não encontrada em '{CAMINHO_PARAMETETROS_USUARIO}'. Adicionando-a com valores vazios.")
                df_loaded[col] = pd.Series(dtype=dtype, index=df_loaded.index)
                if dtype == 'datetime64[ns]':
                    df_loaded[col] = pd.NaT
                else:
                    df_loaded[col].fillna('', inplace=True)
        df_grupos_usuario = df_loaded
        st.info(f"Parâmetros do usuário carregados de '{CAMINHO_PARAMETETROS_USUARIO}'. {df_grupos_usuario.shape[0]} regras carregadas.")
        logger.info(f"Parâmetros do usuário carregados de '{CAMINHO_PARAMETETROS_USUARIO}'. {df_grupos_usuario.shape[0]} regras carregadas.")
    except Exception as e:
        st.warning(f"Erro ao carregar o arquivo de parâmetros do usuário: {e}. Criando base vazia.")
        logger.warning(f"Erro ao carregar o arquivo de parâmetros do usuário: {e}. Criando base vazia.")
else:
    st.info("Arquivo de parâmetros do usuário não encontrado. Comece a adicionar suas regras abaixo.")
    logger.info("Arquivo de parâmetros do usuário não encontrado. Criando base vazia.")

# Configuração das colunas para o data_editor, incluindo os dropdowns de filial e a nova coluna 'Data'
column_configuration = {
    "Substituta": st.column_config.SelectboxColumn(
        "Substituta",
        help="Filial que irá substituir a coleta",
        options=lista_filiais,
        required=True,
    ),
    "Inicial": st.column_config.SelectboxColumn(
        "Inicial",
        help="Filial de origem da coleta a ser substituída (opcional)",
        options=lista_filiais,
        required=False,
    ),
    "Recebe": st.column_config.SelectboxColumn(
        "Recebe",
        help="Indica se esta regra define quem recebe a coleta (S/N)",
        options=['', 'S', 'N'],
        required=True,
    ),
    "UF": st.column_config.TextColumn(
        "UF",
        help="Estado da coleta (Ex: SP, MG)",
        width="small",
        required=True,
        max_chars=2,
    ),
    "Grupo Economico": st.column_config.TextColumn(
        "Grupo Econômico",
        help="Código ou nome do grupo econômico (opcional)",
    ),
    "Modalidade": st.column_config.SelectboxColumn(
        "Modalidade",
        help="Modalidade de transporte (FCA/EXW) (opcional)",
        options=['', 'FCA', 'EXW'],
        required=False,
    ),
    "Tipo de carga": st.column_config.SelectboxColumn(
        "Tipo de carga",
        help="Tipo de carga (Fracionado/Lotação) (opcional)",
        options=['', 'Fracionado', 'Lotação'],
        required=False,
    ),
    "Data": st.column_config.DateColumn( # Adicionada a coluna de data
        "Data",
        help="Data da última atualização desta regra. Preenchida automaticamente ao salvar se vazia.",
        format="DD/MM/YYYY",
        required=False,
    ),
}


df_grupos_usuario_editado = st.data_editor(
    df_grupos_usuario,
    num_rows="dynamic",
    use_container_width=True,
    key="regras_usuario_editadas",
    column_config=column_configuration,
    hide_index=True
)

# === BOTÃO PARA SALVAR AS ALTERAÇÕES DO USUÁRIO ===
if st.button("💾 Salvar minhas Regras (Usuário)"):
    try:
        df_to_save = df_grupos_usuario_editado.copy()
        
        for col, dtype in colunas_base_parametros.items():
            if col in df_to_save.columns:
                if dtype == 'datetime64[ns]':
                    df_to_save[col] = pd.to_datetime(df_to_save[col], errors='coerce')
                    # Preenche a coluna 'Data' com a data atual se for NaT
                    df_to_save[col] = df_to_save[col].apply(lambda x: pd.Timestamp.now().normalize() if pd.isna(x) else x)
                else:
                    df_to_save[col] = df_to_save[col].astype(str).replace(r'^\s*$', '', regex=True)
            else:
                # Caso uma coluna não exista no DataFrame editado (ex: o usuário a removeu acidentalmente)
                # Garante que ela seja adicionada com o tipo correto para evitar erros de concatenação
                df_to_save[col] = pd.Series(dtype=dtype, index=df_to_save.index)
                if dtype == 'datetime64[ns]':
                    df_to_save[col] = pd.NaT
                else:
                    df_to_save[col].fillna('', inplace=True)


        # Remove linhas que são completamente vazias
        df_to_save.replace('', pd.NA, inplace=True) # Temporariamente para dropna
        df_to_save = df_to_save.dropna(how='all')
        
        # Preenche os NAs de volta com strings vazias para as colunas de texto
        # e deixa NaT para as datas (excel exporta como vazio)
        for col, dtype in colunas_base_parametros.items():
            if col in df_to_save.columns and dtype == str:
                df_to_save[col].fillna('', inplace=True)

        # Filtra por linhas com campos obrigatórios preenchidos
        df_to_save = df_to_save[
            (df_to_save['Substituta'].astype(bool)) &
            (df_to_save['Recebe'].astype(bool)) &
            (df_to_save['UF'].astype(bool))
        ]

        if not df_to_save.empty:
            df_to_save.to_excel(CAMINHO_PARAMETETROS_USUARIO, index=False)
            st.success("✅ Suas regras foram salvas com sucesso!")
            st.warning("⚠️ Em ambientes de nuvem (como Streamlit Community Cloud), as alterações podem ser perdidas após o reinício do aplicativo.")
            logger.info("Regras do usuário salvas com sucesso.")
        else:
            st.error("❌ Nenhuma regra válida para salvar. As colunas 'Substituta', 'Recebe' e 'UF' são obrigatórias e não podem estar vazias.")

    except Exception as e:
        st.error(f"❌ Erro ao salvar suas regras: {e}")
        logger.error(f"Erro ao salvar regras do usuário: {e}")

st.divider()

# === BOTÃO PARA PROCESSAR (AGORA COM AS DUAS BASES CONCATENADAS) ===
if st.button("🚀 Rodar Roteirização"):
    log_stream.seek(0)
    log_stream.truncate(0)

    # Processar o DataFrame do usuário para garantir tipos corretos e dados limpos
    df_usuario_processed = df_grupos_usuario_editado.copy()
    for col, dtype in colunas_base_parametros.items():
        if col in df_usuario_processed.columns:
            if dtype == 'datetime64[ns]':
                df_usuario_processed[col] = pd.to_datetime(df_usuario_processed[col], errors='coerce')
            else:
                df_usuario_processed[col] = df_usuario_processed[col].astype(str).replace(r'^\s*$', '', regex=True)
        else:
            df_usuario_processed[col] = pd.Series(dtype=dtype, index=df_usuario_processed.index)
            if dtype == 'datetime64[ns]':
                df_usuario_processed[col] = pd.NaT
            else:
                df_usuario_processed[col].fillna('', inplace=True)

    # Concatene as bases padrão e do usuário.
    df_grupos_final = pd.concat([df_padrao_parametros, df_usuario_processed.dropna(how='all')], ignore_index=True)

    logger.info("Bases de parâmetros padrão e do usuário concatenadas para processamento.")

    # Remova linhas que não tenham os campos obrigatórios para uma regra
    # e garanta que os tipos estejam corretos após concatenação e manipulação
    df_grupos_final_validado = df_grupos_final.copy()
    for col, dtype in colunas_base_parametros.items():
        if col in df_grupos_final_validado.columns:
            if dtype == 'datetime64[ns]':
                df_grupos_final_validado[col] = pd.to_datetime(df_grupos_final_validado[col], errors='coerce')
            else:
                df_grupos_final_validado[col] = df_grupos_final_validado[col].astype(str) # Garante que sejam strings para o replace
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

        def buscar_melhor_regra_substituicao(df_regras_concat, uf, modalidade, tipo_carga, grupo_economico_origem=None):
            # Filtra regras que correspondem aos critérios básicos (UF, Recebe='S', Modalidade e Tipo de Carga)
            # Regras com Modalidade/Tipo de Carga vazios se aplicam a todas
            regras_aplicaveis = df_regras_concat[
                (df_regras_concat['UF'] == uf) &
                (df_regras_concat['Recebe'] == 'S') &
                (df_regras_concat['Modalidade'].isin(['', modalidade])) &
                (df_regras_concat['Tipo de carga'].isin(['', tipo_carga]))
            ].copy() # Use .copy() para evitar SettingWithCopyWarning

            # Se houver um grupo econômico de origem, filtre também por ele.
            # Regras com 'Grupo Economico' vazio se aplicam a qualquer grupo.
            if grupo_economico_origem and str(grupo_economico_origem).strip() != '':
                regras_aplicaveis = regras_aplicaveis[
                    (regras_aplicaveis['Grupo Economico'] == '') |
                    (regras_aplicaveis['Grupo Economico'] == grupo_economico_origem)
                ]
            else: # Se não houver grupo econômico de origem (ou for vazio), só considere regras sem grupo econômico especificado
                regras_aplicaveis = regras_aplicaveis[regras_aplicaveis['Grupo Economico'] == '']


            if regras_aplicaveis.empty:
                return pd.Series() # Retorna uma Series vazia se nenhuma regra for encontrada

            # Lógica para priorizar:
            # 1. Mais específica (mais critérios preenchidos: Grupo Economico > Tipo de Carga > Modalidade > Inicial)
            # 2. Mais recente (coluna 'Data')
            # 3. Se tudo empatar, usar o índice original (para consistência)

            # Pontuação de especificidade (ajuste os pesos conforme sua necessidade)
            regras_aplicaveis['specificity_score'] = 0
            regras_aplicaveis['specificity_score'] += regras_aplicaveis['Grupo Economico'].apply(lambda x: 4 if str(x).strip() != '' else 0)
            regras_aplicaveis['specificity_score'] += regras_aplicaveis['Tipo de carga'].apply(lambda x: 3 if str(x).strip() != '' else 0)
            regras_aplicaveis['specificity_score'] += regras_aplicaveis['Modalidade'].apply(lambda x: 2 if str(x).strip() != '' else 0)
            regras_aplicaveis['specificity_score'] += regras_aplicaveis['Inicial'].apply(lambda x: 1 if str(x).strip() != '' else 0)

            # Ordena por especificidade (descendente) e depois por Data (descendente)
            # Regras sem data (NaT) virão por último em sort descendente.
            # Adiciona o índice original para desempate final e consistência.
            regras_aplicaveis = regras_aplicaveis.sort_values(
                by=['specificity_score', 'Data', regras_aplicaveis.index.name if regras_aplicaveis.index.name else regras_aplicaveis.index],
                ascending=[False, False, True] # Mais específico, mais recente, menor índice
            )

            return regras_aplicaveis.iloc[0] # Retorna a regra de maior prioridade

        municipios = df_dist['MunicipioOrigem'].unique()
        total_municipios = len(municipios)
        resultados = []

        for i, municipio in enumerate(municipios):
            uf_municipio = municipio.split('-')[-1].strip()
            
            # ATENÇÃO: Se o Grupo Econômico for uma característica do município de origem,
            # você precisaria carregá-lo aqui (ex: de df_dist ou outra base).
            # Por exemplo:
            # grupo_economico_municipio = df_dist[df_dist['MunicipioOrigem'] == municipio]['GrupoEconomico'].iloc[0]
            grupo_economico_municipio = None # Mantenha None se você não tem essa informação por município

            for incoterm, tipo_carga, coluna_param in modalidades:
                try:
                    filial_encontrada = False

                    # Busca a melhor regra aplicando a nova função
                    regra = buscar_melhor_regra_substituicao(df_grupos_final_validado, uf_municipio, incoterm, tipo_carga, grupo_economico_municipio)

                    if not regra.empty:
                        # Certifica que Grupo Economico é uma string ou None
                        grupo_economico_str = str(regra['Grupo Economico']) if pd.notna(regra['Grupo Economico']) and str(regra['Grupo Economico']).strip() != '' else None

                        try:
                            cod_filial_subs = df_filiais[df_filiais['Filial'].astype(str) == str(regra['Substituta'])]['Codigo'].iloc[0]
                            logger.info(f"Regra de substituição aplicada para {municipio} ({incoterm}/{tipo_carga}): Filial {regra['Substituta']} (Código: {int(cod_filial_subs):04}).")
                        except IndexError:
                            logger.warning(f"Código não encontrado para filial substituta {regra['Substituta']} para {municipio} ({incoterm}/{tipo_carga}). Usando '0000'.")
                            cod_filial_subs = '0000'

                        descricao_regra = (
                            f"Regra de Substituição: {regra['Substituta']} recebe coletas de {grupo_economico_str if grupo_economico_str else 'qualquer grupo'} "
                            f"({regra['Modalidade'] if pd.notna(regra['Modalidade']) and str(regra['Modalidade']).strip() != '' else 'Todas modalidades'}, "
                            f"{regra['Tipo de carga'] if pd.notna(regra['Tipo de carga']) and str(regra['Tipo de carga']).strip() != '' else 'Todos os tipos de carga'})"
                        )
                        if pd.notna(regra.get('Inicial')) and str(regra['Inicial']).strip() != '':
                            descricao_regra += f" ao invés de {regra['Inicial']}"
                        
                        # Adicionar a data da regra à descrição, se existir
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
                            # A coluna 'GRUPO ECONOMICO' no resultado é o grupo da REGRA, não o do município.
                            # Se você precisar do grupo do município aqui, ele deve ser passado como `grupo_economico_municipio`
                            'GRUPO ECONOMICO': f"{int(float(grupo_economico_str)):04}" if grupo_economico_str and str(grupo_economico_str).replace('.', '', 1).isdigit() else None
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
                                logger.info(f"Filial compatível com modalidade encontrada para {municipio} ({incoterm}/{tipo_carga}): {mais_proxima['Filial']} (KM_ID: {mais_proxima['KM_ID']}).")


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
                                logger.info(f"Filial única no estado encontrada para {municipio} ({incoterm}/{tipo_carga}): {filial_unica['Filial']} (KM_ID: {mais_proxima['KM_ID']}).")


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
                            logger.info(f"Filial mais próxima (sem restrição) encontrada para {municipio} ({incoterm}/{tipo_carga}): {mais_proxima['Filial']} (KM_ID: {mais_proxima['KM_ID']}).")

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
                        logger.warning(f"Nenhuma filial encontrada para {municipio} ({incoterm}/{tipo_carga}).")

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
