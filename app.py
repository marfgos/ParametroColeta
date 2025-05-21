import streamlit as st
import pandas as pd
import os
import logging
from io import BytesIO, StringIO # StringIO para capturar o log em memória
import zipfile # Para criar o arquivo ZIP

# === CONFIGURAÇÕES INICIAIS ===
st.set_page_config(page_title="Roteirização com Substituição", layout="wide")
st.title("📦 Roteirização com Regras de Substituição")

# === CONFIGURAÇÃO DE LOG ===
# Criar um buffer de string para capturar o log
log_stream = StringIO()
# Remover handlers existentes para evitar duplicação de logs, especialmente importante em reruns
for handler in logging.root.handlers[:]:
    logging.root.removeHandler(handler)
logging.basicConfig(stream=log_stream, level=logging.INFO, format='%(asctime)s - %(message)s')
logger = logging.getLogger() # Obter o logger raiz

# === CAMINHOS DOS ARQUIVOS DE PARÂMETROS ===
CAMINHO_PARAMETROS_PADRAO = "parametros_contrato.xlsx" # Base do repositório
CAMINHO_PARAMETROS_USUARIO = "parametros_usuario.xlsx" # Base editável pelo usuário

# Colunas esperadas para as bases de parâmetros
colunas_base_parametros = ['Substituta', 'Inicial', 'Recebe', 'UF', 'Grupo Economico', 'Modalidade', 'Tipo de carga']

# === CARREGAMENTO DAS BASES FIXAS INTERNAS (Distâncias e Filiais) ===
try:
    df_dist = pd.read_excel("municipios_distanciasreais.xlsx")
    df_filiais = pd.read_excel("filiais_geocodificadas.xlsx")
    st.success("📍 Bases internas de distâncias e filiais carregadas com sucesso.")
    logger.info("Bases internas de distâncias e filiais carregadas com sucesso.")
except Exception as e:
    st.error(f"❌ Erro ao carregar arquivos internos (distâncias/filiais): {e}")
    logger.error(f"Erro ao carregar arquivos internos (distâncias/filiais): {e}")
    st.stop() # Interrompe a execução se as bases principais não carregarem

# === CARREGAMENTO E EXIBIÇÃO DA BASE DE PARÂMETROS PADRÃO ===
st.header("📄 Parâmetros Contratuais Padrão")
try:
    df_padrao_parametros = pd.read_excel(CAMINHO_PARAMETROS_PADRAO)
    st.info(f"Parâmetros contratuais padrão carregados de '{CAMINHO_PARAMETROS_PADRAO}'.")
    logger.info(f"Parâmetros contratuais padrão carregados de '{CAMINHO_PARAMETROS_PADRAO}'.")
except FileNotFoundError:
    st.warning(f"Arquivo '{CAMINHO_PARAMETROS_PADRAO}' não encontrado. Criando DataFrame padrão vazio.")
    logger.warning(f"Arquivo '{CAMINHO_PARAMETROS_PADRAO}' não encontrado. Criando DataFrame padrão vazio.")
    df_padrao_parametros = pd.DataFrame(columns=colunas_base_parametros)
except Exception as e:
    st.error(f"Erro ao carregar o arquivo de parâmetros padrão '{CAMINHO_PARAMETROS_PADRAO}': {e}")
    logger.error(f"Erro ao carregar o arquivo de parâmetros padrão '{CAMINHO_PARAMETROS_PADRAO}': {e}")
    df_padrao_parametros = pd.DataFrame(columns=colunas_base_parametros)

st.dataframe(df_padrao_parametros, use_container_width=True, height=200) # Exibe a base padrão
st.divider()

# === CARREGAMENTO E EDIÇÃO DA BASE DE PARÂMETROS DO USUÁRIO ===
st.header("✏️ Parâmetros Contratuais do Usuário (Editável)")

df_grupos_usuario = pd.DataFrame(columns=colunas_base_parametros)
if os.path.exists(CAMINHO_PARAMETROS_USUARIO):
    try:
        df_grupos_usuario = pd.read_excel(CAMINHO_PARAMETROS_USUARIO)
        st.info(f"Parâmetros do usuário carregados de '{CAMINHO_PARAMETROS_USUARIO}'.")
        logger.info(f"Parâmetros do usuário carregados de '{CAMINHO_PARAMETROS_USUARIO}'.")
    except Exception as e:
        st.warning(f"Erro ao carregar o arquivo de parâmetros do usuário: {e}. Criando base vazia.")
        logger.warning(f"Erro ao carregar o arquivo de parâmetros do usuário: {e}. Criando base vazia.")
        df_grupos_usuario = pd.DataFrame(columns=colunas_base_parametros)
else:
    st.info("Arquivo de parâmetros do usuário não encontrado. Comece a adicionar suas regras abaixo.")
    logger.info("Arquivo de parâmetros do usuário não encontrado. Criando base vazia.")

# Permite ao usuário editar, adicionar e remover linhas
df_grupos_usuario_editado = st.data_editor(
    df_grupos_usuario,
    num_rows="dynamic", # Permite adicionar e remover linhas
    use_container_width=True,
    key="regras_usuario_editadas" # Chave única para o widget
)

# === BOTÃO PARA SALVAR AS ALTERAÇÕES DO USUÁRIO ===
if st.button("💾 Salvar minhas Regras (Usuário)"):
    try:
        # Removendo linhas completamente vazias que o data_editor pode criar
        df_to_save = df_grupos_usuario_editado.dropna(how='all')
        df_to_save.to_excel(CAMINHO_PARAMETROS_USUARIO, index=False)
        st.success("✅ Suas regras foram salvas com sucesso!")
        st.warning("⚠️ Em ambientes de nuvem (como Streamlit Community Cloud), as alterações podem ser perdidas após o reinício do aplicativo.")
        logger.info("Regras do usuário salvas com sucesso.")
        st.experimental_rerun() # Recarrega a página para refletir as alterações salvas
    except Exception as e:
        st.error(f"❌ Erro ao salvar suas regras: {e}")
        logger.error(f"Erro ao salvar regras do usuário: {e}")

st.divider()

# === BOTÃO PARA PROCESSAR (AGORA COM AS DUAS BASES CONCATENADAS) ===
if st.button("🚀 Rodar Roteirização"):
    # Limpa o stream de log para uma nova execução
    log_stream.seek(0)
    log_stream.truncate(0)

    # Concatena as bases padrão e do usuário para o processamento
    df_grupos_final = pd.concat([df_padrao_parametros, df_grupos_usuario_editado.dropna(how='all')], ignore_index=True)
    logger.info("Bases de parâmetros padrão e do usuário concatenadas para processamento.")

    if df_grupos_final.empty:
        st.error("Por favor, preencha os parâmetros contratuais (padrão ou do usuário) antes de processar.")
        logger.error("Tentativa de roteirização com base de parâmetros vazia.")
    else:
        # Barra de progresso
        progress_text = "Processando roteirização... Aguarde."
        my_bar = st.progress(0, text=progress_text)

        total_municipios = len(df_dist['MunicipioOrigem'].unique())
        municipios_processados = 0

        with st.spinner("Preparando roteirização..."):
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

            for i, municipio in enumerate(municipios):
                uf_municipio = municipio.split('-')[-1].strip()

                for incoterm, tipo_carga, coluna_param in modalidades:
                    try:
                        filial_encontrada = False

                        # Prioridade 1: Regras de Substituição (base final concatenada)
                        regras_subs = buscar_regras_substituicao(df_grupos_final, uf_municipio, incoterm, tipo_carga)
                        if not regras_subs.empty:
                            regra = regras_subs.iloc[0] # Pega a primeira regra aplicável
                            try:
                                cod_filial_subs = df_filiais[df_filiais['Filial'] == regra['Substituta']]['Codigo'].iloc[0]
                                logger.info(f"Regra de substituição aplicada para {municipio} ({incoterm}/{tipo_carga}): Filial {regra['Substituta']} (Código: {int(cod_filial_subs):04}).")
                            except IndexError:
                                logger.warning(f"Código não encontrado para filial substituta {regra['Substituta']} para {municipio} ({incoterm}/{tipo_carga}). Usando '0000'.")
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
                
                # Atualizar barra de progresso
                municipios_processados += 1
                progress = min(100, int(municipios_processados / total_municipios * 100))
                my_bar.progress(progress, text=f"{progress_text} {progress}% Concluído.")
                logger.debug(f"Progresso: {progress}% para {municipio}")

            # Finaliza a barra de progresso em 100%
            my_bar.progress(100, text=f"{progress_text} 100% Concluído.")

            df_resultado = pd.DataFrame(resultados)
            
            # === EXPORTAR RESULTADO E LOG EM ARQUIVO ZIP ===
            zip_buffer = BytesIO()
            with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zf:
                # Adicionar o Excel ao ZIP
                excel_buffer = BytesIO()
                df_resultado.to_excel(excel_buffer, index=False)
                excel_buffer.seek(0)
                zf.writestr("resultado_roteirizacao.xlsx", excel_buffer.getvalue())
                logger.info("Resultado da roteirização adicionado ao ZIP.")

                # Adicionar o Log ao ZIP
                log_content = log_stream.getvalue()
                zf.writestr("log_roteirizacao.log", log_content)
                logger.info("Log da roteirização adicionado ao ZIP.")
            
            zip_buffer.seek(0) # Retorna ao início do buffer para download

            st.success("✅ Processamento concluído!")
            st.dataframe(df_resultado)

            st.download_button(
                label="📥 Baixar Resultado (Excel + Log .zip)",
                data=zip_buffer,
                file_name="roteirizacao_completa.zip",
                mime="application/zip"
            )

            st.info("Log de processamento:")
            st.text_area("Visualizar Log", log_content, height=200) # Exibe o log na tela também
            logger.info("Aplicação concluída.")import streamlit as st
import pandas as pd
import os
import logging
from io import BytesIO, StringIO # StringIO para capturar o log em memória
import zipfile # Para criar o arquivo ZIP

# === CONFIGURAÇÕES INICIAIS ===
st.set_page_config(page_title="Roteirização com Substituição", layout="wide")
st.title("📦 Roteirização com Regras de Substituição")

# === CONFIGURAÇÃO DE LOG ===
# Criar um buffer de string para capturar o log
log_stream = StringIO()
# Remover handlers existentes para evitar duplicação de logs, especialmente importante em reruns
for handler in logging.root.handlers[:]:
    logging.root.removeHandler(handler)
logging.basicConfig(stream=log_stream, level=logging.INFO, format='%(asctime)s - %(message)s')
logger = logging.getLogger() # Obter o logger raiz

# === CAMINHOS DOS ARQUIVOS DE PARÂMETROS ===
CAMINHO_PARAMETROS_PADRAO = "parametros_contrato.xlsx" # Base do repositório
CAMINHO_PARAMETROS_USUARIO = "parametros_usuario.xlsx" # Base editável pelo usuário

# Colunas esperadas para as bases de parâmetros
colunas_base_parametros = ['Substituta', 'Inicial', 'Recebe', 'UF', 'Grupo Economico', 'Modalidade', 'Tipo de carga']

# === CARREGAMENTO DAS BASES FIXAS INTERNAS (Distâncias e Filiais) ===
try:
    df_dist = pd.read_excel("municipios_distanciasreais.xlsx")
    df_filiais = pd.read_excel("filiais_geocodificadas.xlsx")
    st.success("📍 Bases internas de distâncias e filiais carregadas com sucesso.")
    logger.info("Bases internas de distâncias e filiais carregadas com sucesso.")
except Exception as e:
    st.error(f"❌ Erro ao carregar arquivos internos (distâncias/filiais): {e}")
    logger.error(f"Erro ao carregar arquivos internos (distâncias/filiais): {e}")
    st.stop() # Interrompe a execução se as bases principais não carregarem

# === CARREGAMENTO E EXIBIÇÃO DA BASE DE PARÂMETROS PADRÃO ===
st.header("📄 Parâmetros Contratuais Padrão")
try:
    df_padrao_parametros = pd.read_excel(CAMINHO_PARAMETROS_PADRAO)
    st.info(f"Parâmetros contratuais padrão carregados de '{CAMINHO_PARAMETROS_PADRAO}'.")
    logger.info(f"Parâmetros contratuais padrão carregados de '{CAMINHO_PARAMETROS_PADRAO}'.")
except FileNotFoundError:
    st.warning(f"Arquivo '{CAMINHO_PARAMETROS_PADRAO}' não encontrado. Criando DataFrame padrão vazio.")
    logger.warning(f"Arquivo '{CAMINHO_PARAMETROS_PADRAO}' não encontrado. Criando DataFrame padrão vazio.")
    df_padrao_parametros = pd.DataFrame(columns=colunas_base_parametros)
except Exception as e:
    st.error(f"Erro ao carregar o arquivo de parâmetros padrão '{CAMINHO_PARAMETROS_PADRAO}': {e}")
    logger.error(f"Erro ao carregar o arquivo de parâmetros padrão '{CAMINHO_PARAMETROS_PADRAO}': {e}")
    df_padrao_parametros = pd.DataFrame(columns=colunas_base_parametros)

st.dataframe(df_padrao_parametros, use_container_width=True, height=200) # Exibe a base padrão
st.divider()

# === CARREGAMENTO E EDIÇÃO DA BASE DE PARÂMETROS DO USUÁRIO ===
st.header("✏️ Parâmetros Contratuais do Usuário (Editável)")

df_grupos_usuario = pd.DataFrame(columns=colunas_base_parametros)
if os.path.exists(CAMINHO_PARAMETROS_USUARIO):
    try:
        df_grupos_usuario = pd.read_excel(CAMINHO_PARAMETROS_USUARIO)
        st.info(f"Parâmetros do usuário carregados de '{CAMINHO_PARAMETROS_USUARIO}'.")
        logger.info(f"Parâmetros do usuário carregados de '{CAMINHO_PARAMETROS_USUARIO}'.")
    except Exception as e:
        st.warning(f"Erro ao carregar o arquivo de parâmetros do usuário: {e}. Criando base vazia.")
        logger.warning(f"Erro ao carregar o arquivo de parâmetros do usuário: {e}. Criando base vazia.")
        df_grupos_usuario = pd.DataFrame(columns=colunas_base_parametros)
else:
    st.info("Arquivo de parâmetros do usuário não encontrado. Comece a adicionar suas regras abaixo.")
    logger.info("Arquivo de parâmetros do usuário não encontrado. Criando base vazia.")

# Permite ao usuário editar, adicionar e remover linhas
df_grupos_usuario_editado = st.data_editor(
    df_grupos_usuario,
    num_rows="dynamic", # Permite adicionar e remover linhas
    use_container_width=True,
    key="regras_usuario_editadas" # Chave única para o widget
)

# === BOTÃO PARA SALVAR AS ALTERAÇÕES DO USUÁRIO ===
if st.button("💾 Salvar minhas Regras (Usuário)"):
    try:
        # Removendo linhas completamente vazias que o data_editor pode criar
        df_to_save = df_grupos_usuario_editado.dropna(how='all')
        df_to_save.to_excel(CAMINHO_PARAMETROS_USUARIO, index=False)
        st.success("✅ Suas regras foram salvas com sucesso!")
        st.warning("⚠️ Em ambientes de nuvem (como Streamlit Community Cloud), as alterações podem ser perdidas após o reinício do aplicativo.")
        logger.info("Regras do usuário salvas com sucesso.")
        st.experimental_rerun() # Recarrega a página para refletir as alterações salvas
    except Exception as e:
        st.error(f"❌ Erro ao salvar suas regras: {e}")
        logger.error(f"Erro ao salvar regras do usuário: {e}")

st.divider()

# === BOTÃO PARA PROCESSAR (AGORA COM AS DUAS BASES CONCATENADAS) ===
if st.button("🚀 Rodar Roteirização"):
    # Limpa o stream de log para uma nova execução
    log_stream.seek(0)
    log_stream.truncate(0)

    # Concatena as bases padrão e do usuário para o processamento
    df_grupos_final = pd.concat([df_padrao_parametros, df_grupos_usuario_editado.dropna(how='all')], ignore_index=True)
    logger.info("Bases de parâmetros padrão e do usuário concatenadas para processamento.")

    if df_grupos_final.empty:
        st.error("Por favor, preencha os parâmetros contratuais (padrão ou do usuário) antes de processar.")
        logger.error("Tentativa de roteirização com base de parâmetros vazia.")
    else:
        # Barra de progresso
        progress_text = "Processando roteirização... Aguarde."
        my_bar = st.progress(0, text=progress_text)

        total_municipios = len(df_dist['MunicipioOrigem'].unique())
        municipios_processados = 0

        with st.spinner("Preparando roteirização..."):
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

            for i, municipio in enumerate(municipios):
                uf_municipio = municipio.split('-')[-1].strip()

                for incoterm, tipo_carga, coluna_param in modalidades:
                    try:
                        filial_encontrada = False

                        # Prioridade 1: Regras de Substituição (base final concatenada)
                        regras_subs = buscar_regras_substituicao(df_grupos_final, uf_municipio, incoterm, tipo_carga)
                        if not regras_subs.empty:
                            regra = regras_subs.iloc[0] # Pega a primeira regra aplicável
                            try:
                                cod_filial_subs = df_filiais[df_filiais['Filial'] == regra['Substituta']]['Codigo'].iloc[0]
                                logger.info(f"Regra de substituição aplicada para {municipio} ({incoterm}/{tipo_carga}): Filial {regra['Substituta']} (Código: {int(cod_filial_subs):04}).")
                            except IndexError:
                                logger.warning(f"Código não encontrado para filial substituta {regra['Substituta']} para {municipio} ({incoterm}/{tipo_carga}). Usando '0000'.")
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
                
                # Atualizar barra de progresso
                municipios_processados += 1
                progress = min(100, int(municipios_processados / total_municipios * 100))
                my_bar.progress(progress, text=f"{progress_text} {progress}% Concluído.")
                logger.debug(f"Progresso: {progress}% para {municipio}")

            # Finaliza a barra de progresso em 100%
            my_bar.progress(100, text=f"{progress_text} 100% Concluído.")

            df_resultado = pd.DataFrame(resultados)
            
            # === EXPORTAR RESULTADO E LOG EM ARQUIVO ZIP ===
            zip_buffer = BytesIO()
            with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zf:
                # Adicionar o Excel ao ZIP
                excel_buffer = BytesIO()
                df_resultado.to_excel(excel_buffer, index=False)
                excel_buffer.seek(0)
                zf.writestr("resultado_roteirizacao.xlsx", excel_buffer.getvalue())
                logger.info("Resultado da roteirização adicionado ao ZIP.")

                # Adicionar o Log ao ZIP
                log_content = log_stream.getvalue()
                zf.writestr("log_roteirizacao.log", log_content)
                logger.info("Log da roteirização adicionado ao ZIP.")
            
            zip_buffer.seek(0) # Retorna ao início do buffer para download

            st.success("✅ Processamento concluído!")
            st.dataframe(df_resultado)

            st.download_button(
                label="📥 Baixar Resultado (Excel + Log .zip)",
                data=zip_buffer,
                file_name="roteirizacao_completa.zip",
                mime="application/zip"
            )

            st.info("Log de processamento:")
            st.text_area("Visualizar Log", log_content, height=200) # Exibe o log na tela também
            logger.info("Aplicação concluída.")
