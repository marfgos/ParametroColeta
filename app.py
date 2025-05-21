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

# Colunas esperadas para as bases de parâmetros
# Definindo os tipos para melhor compatibilidade com st.data_editor
colunas_base_parametros = {
    'Substituta': str,
    'Inicial': str,
    'Recebe': str,
    'UF': str,
    'Grupo Economico': str, # Pode ser string para códigos que começam com zero ou mistos
    'Modalidade': str,
    'Tipo de carga': str
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

# --- Extrair lista de filiais para os Selectbox ---
if 'Filial' in df_filiais.columns:
    # Garante que a lista contenha strings e ordena
    lista_filiais = [''] + sorted(df_filiais['Filial'].astype(str).unique().tolist())
else:
    lista_filiais = ['']
    st.warning("Coluna 'Filial' não encontrada em 'filiais_geocodificadas.xlsx'. Dropdowns de Filial podem não funcionar.")


# === CARREGAMENTO E EXIBIÇÃO DA BASE DE PARÂMETROS PADRÃO ===
st.header("📄 Parâmetros Contratuais Padrão")
try:
    df_padrao_parametros = pd.read_excel(CAMINHO_PARAMETROS_PADRAO)
    # Garante que as colunas do df_padrao_parametros tenham os tipos esperados
    for col, dtype in colunas_base_parametros.items():
        if col in df_padrao_parametros.columns:
            df_padrao_parametros[col] = df_padrao_parametros[col].astype(dtype).fillna('') # Preenche NaN com string vazia
        else: # Adiciona coluna se não existir, com tipo correto e valores vazios
            df_padrao_parametros[col] = pd.Series(dtype=dtype)
            df_padrao_parametros[col].fillna('', inplace=True)

    st.info(f"Parâmetros contratuais padrão carregados de '{CAMINHO_PARAMETROS_PADRAO}'.")
    logger.info(f"Parâmetros contratuais padrão carregados de '{CAMINHO_PARAMETROS_PADRAO}'.")
except FileNotFoundError:
    st.warning(f"Arquivo '{CAMINHO_PARAMETROS_PADRAO}' não encontrado. Criando DataFrame padrão vazio.")
    logger.warning(f"Arquivo '{CAMINHO_PARAMETROS_PADRÃO}' não encontrado. Criando DataFrame padrão vazio.")
    # Cria DataFrame vazio com os dtypes corretos
    df_padrao_parametros = pd.DataFrame(columns=list(colunas_base_parametros.keys()))
    for col, dtype in colunas_base_parametros.items():
        df_padrao_parametros[col] = pd.Series(dtype=dtype)
        df_padrao_parametros[col].fillna('', inplace=True) # Preenche NaN com string vazia
except Exception as e:
    st.error(f"Erro ao carregar o arquivo de parâmetros padrão '{CAMINHO_PARAMETROS_PADRAO}': {e}")
    logger.error(f"Erro ao carregar o arquivo de parâmetros padrão '{CAMINHO_PARAMETROS_PADRAO}': {e}")
    # Cria DataFrame vazio com os dtypes corretos em caso de erro
    df_padrao_parametros = pd.DataFrame(columns=list(colunas_base_parametros.keys()))
    for col, dtype in colunas_base_parametros.items():
        df_padrao_parametros[col] = pd.Series(dtype=dtype)
        df_padrao_parametros[col].fillna('', inplace=True) # Preenche NaN com string vazia

st.dataframe(df_padrao_parametros, use_container_width=True, height=200)
st.divider()

# === CARREGAMENTO E EDIÇÃO DA BASE DE PARÂMETROS DO USUÁRIO ===
st.header("✏️ Parâmetros Contratuais do Usuário (Editável)")

# Inicializa o DataFrame do usuário com os dtypes corretos
df_grupos_usuario = pd.DataFrame(columns=list(colunas_base_parametros.keys()))
for col, dtype in colunas_base_parametros.items():
    df_grupos_usuario[col] = pd.Series(dtype=dtype)
    df_grupos_usuario[col].fillna('', inplace=True) # Preenche NaN com string vazia

if os.path.exists(CAMINHO_PARAMETETROS_USUARIO):
    try:
        df_loaded = pd.read_excel(CAMINHO_PARAMETETROS_USUARIO)
        # Aplica os tipos corretos ao DataFrame carregado
        for col, dtype in colunas_base_parametros.items():
            if col in df_loaded.columns:
                df_loaded[col] = df_loaded[col].astype(dtype).fillna('')
            else: # Adiciona coluna se não existir, com tipo correto e valores vazios
                df_loaded[col] = pd.Series(dtype=dtype)
                df_loaded[col].fillna('', inplace=True)
        df_grupos_usuario = df_loaded # Atribui o DataFrame carregado e tipado
        st.info(f"Parâmetros do usuário carregados de '{CAMINHO_PARAMETETROS_USUARIO}'.")
        logger.info(f"Parâmetros do usuário carregados de '{CAMINHO_PARAMETETROS_USUARIO}'.")
    except Exception as e:
        st.warning(f"Erro ao carregar o arquivo de parâmetros do usuário: {e}. Criando base vazia.")
        logger.warning(f"Erro ao carregar o arquivo de parâmetros do usuário: {e}. Criando base vazia.")
        # df_grupos_usuario já está inicializado como vazio e tipado corretamente
else:
    st.info("Arquivo de parâmetros do usuário não encontrado. Comece a adicionar suas regras abaixo.")
    logger.info("Arquivo de parâmetros do usuário não encontrado. Criando base vazia.")

# Configuração das colunas para o data_editor, incluindo os dropdowns de filial
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
    "UF": st.column_config.TextColumn( # Permite entrada de texto livre para UF
        "UF",
        help="Estado da coleta (Ex: SP, MG)",
        width="small",
        required=True,
        max_chars=2, # Geralmente UF tem 2 caracteres
    ),
    "Grupo Economico": st.column_config.TextColumn( # Mantido como texto, pois pode ser código ou nome
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
        # Remover linhas completamente vazias que o data_editor pode criar
        # E também remover linhas onde 'Substituta' está vazio ou 'Recebe'/'UF' estão vazios
        df_to_save = df_grupos_usuario_editado.copy()
        
        # Converte para string para garantir que fillna('') funcione corretamente em todas as colunas
        for col in colunas_base_parametros.keys():
            if col in df_to_save.columns:
                df_to_save[col] = df_to_save[col].astype(str).replace(r'^\s*$', '', regex=True) # Trata espaços em branco como vazio

        # Remove linhas onde todas as colunas estão vazias
        df_to_save.replace('', pd.NA, inplace=True) # Temporariamente substitui '' por NA para dropna
        df_to_save = df_to_save.dropna(how='all')
        df_to_save.replace(pd.NA, '', inplace=True) # Volta NA para '' para salvar

        # Valida colunas obrigatórias após limpeza
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
            st.experimental_rerun()
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

    # Concatena as bases padrão e do usuário para o processamento
    # Garante que as colunas do df_grupos_usuario_editado sejam do tipo string para concatenação segura
    df_usuario_processed = df_grupos_usuario_editado.copy()
    for col in colunas_base_parametros.keys():
        if col in df_usuario_processed.columns:
            df_usuario_processed[col] = df_usuario_processed[col].astype(str).replace(r'^\s*$', '', regex=True)
            
    df_grupos_final = pd.concat([df_padrao_parametros, df_usuario_processed.dropna(how='all')], ignore_index=True)
    logger.info("Bases de parâmetros padrão e do usuário concatenadas para processamento.")

    if df_grupos_final.empty:
        st.error("Por favor, preencha os parâmetros contratuais (padrão ou do usuário) antes de processar.")
        logger.error("Tentativa de roteirização com base de parâmetros vazia.")
        st.stop()
    
    # Validação adicional das colunas obrigatórias nas regras finais
    # Garante que 'Substituta', 'Recebe', 'UF' não sejam vazios nas regras que serão usadas
    df_grupos_final_validado = df_grupos_final.dropna(subset=['Substituta', 'Recebe', 'UF'])
    
    # IMPORTANTE: replace de strings vazias para NaN para que dropna funcione para SelectboxColumn
    df_grupos_final_validado = df_grupos_final_validado.replace('', pd.NA)
    df_grupos_final_validado = df_grupos_final_validado.dropna(subset=['Substituta', 'Recebe', 'UF']) # Repete dropna após replace
    df_grupos_final_validado = df_grupos_final_validado.fillna('') # Volta NaN para string vazia para o restante do código

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

        def buscar_regras_substituicao(df_regras, uf, modalidade, tipo_carga):
            # Garante que UF, Modalidade e Tipo de carga são tratados como string para comparação
            regras = df_regras[(df_regras['UF'] == uf) & (df_regras['Recebe'] == 'S')]
            regras = regras[
                (regras['Modalidade'].isin(['', modalidade, None])) & # Inclui vazio, None e o valor da modalidade
                (regras['Tipo de carga'].isin(['', tipo_carga, None]))
            ]
            return regras.copy() # Retorna uma cópia para evitar SettingWithCopyWarning

        municipios = df_dist['MunicipioOrigem'].unique()
        total_municipios = len(municipios)
        resultados = []

        for i, municipio in enumerate(municipios):
            uf_municipio = municipio.split('-')[-1].strip()

            for incoterm, tipo_carga, coluna_param in modalidades:
                try:
                    filial_encontrada = False

                    # Prioridade 1: Regras de Substituição (base final concatenada)
                    regras_subs = buscar_regras_substituicao(df_grupos_final_validado, uf_municipio, incoterm, tipo_carga)
                    if not regras_subs.empty:
                        regra = regras_subs.iloc[0]
                        # Converte 'Grupo Economico' para str antes de verificar pd.notna
                        grupo_economico_str = str(regra['Grupo Economico']) if pd.notna(regra['Grupo Economico']) and str(regra['Grupo Economico']).strip() != '' else None
                        
                        try:
                            # Garante que a filial seja string antes de comparar
                            cod_filial_subs = df_filiais[df_filiais['Filial'].astype(str) == str(regra['Substituta'])]['Codigo'].iloc[0]
                            logger.info(f"Regra de substituição aplicada para {municipio} ({incoterm}/{tipo_carga}): Filial {regra['Substituta']} (Código: {int(cod_filial_subs):04}).")
                        except IndexError:
                            logger.warning(f"Código não encontrado para filial substituta {regra['Substituta']} para {municipio} ({incoterm}/{tipo_carga}). Usando '0000'.")
                            cod_filial_subs = '0000'

                        descricao_regra = (
                            f"Regra de Substituição: {regra['Substituta']} recebe coletas de {grupo_economico_str if grupo_economico_str else 'N/A'} "
                            f"({regra['Modalidade'] if pd.notna(regra['Modalidade']) and str(regra['Modalidade']).strip() != '' else 'Todas'}, "
                            f"{regra['Tipo de carga'] if pd.notna(regra['Tipo de carga']) and str(regra['Tipo de carga']).strip() != '' else 'Todos'})"
                        )
                        if pd.notna(regra.get('Inicial')) and str(regra['Inicial']).strip() != '':
                            descricao_regra += f" ao invés de {regra['Inicial']}"

                        resultados.append({
                            'Origem': municipio,
                            'Incoterm': incoterm,
                            'Tipo_Carga': tipo_carga,
                            'Filial': regra['Substituta'],
                            'Codigo_Filial': f"{int(cod_filial_subs):04}",
                            'KM_ID': None,
                            'Condicao_Atribuicao': descricao_regra,
                            'GRUPO ECONOMICO': f"{int(float(grupo_economico_str)):04}" if grupo_economico_str and str(grupo_economico_str).replace('.', '', 1).isdigit() else None
                        })
                        filial_encontrada = True

                    if not filial_encontrada:
                        # Prioridade 2: Filial compatível com a modalidade
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
                        # Prioridade 3: Única filial no estado
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
                        # Prioridade 4: Filial mais próxima (sem restrição)
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
