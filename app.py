import streamlit as st
import pandas as pd
import os
import logging

# Configuração de logging
log_file = 'log_filiais_proximas.log'
logging.basicConfig(filename=log_file, level=logging.INFO, format='%(asctime)s - %(message)s')

st.title("Processamento de Filiais Mais Próximas com Regras de Substituição")

# Upload dos arquivos
dist_file = st.file_uploader("Upload da base de distâncias (municipios_distanciasreais.xlsx)", type=['xlsx'])
filial_file = st.file_uploader("Upload da base de filiais (filiais_geocodificadas.xlsx)", type=['xlsx'])
param_file = st.file_uploader("Upload da base de parâmetros contratuais (parametros_contrato.xlsx)", type=['xlsx'])

nome_arquivo = st.text_input("Nome do arquivo Excel para salvar o resultado (sem extensão)", "resultado_final")

if st.button("Processar"):

    if dist_file and filial_file and param_file:

        # Carregar dados
        df_dist = pd.read_excel(dist_file)
        df_filiais = pd.read_excel(filial_file)
        df_grupos = pd.read_excel(param_file)

        modalidades = [
            ("FCA", "Fracionado", "FCA/Fracionado"),
            ("FCA", "Lotação", "FCA/Lotação"),
            ("EXW", "Fracionado", "EXW/Fracionado"),
            ("EXW", "Lotação", "EXW/Lotação")
        ]

        resultados = []

        def buscar_regras_substituicao(df_regras, uf, modalidade, tipo_carga):
            regras = df_regras[
                (df_regras['UF'] == uf) & 
                (df_regras['Recebe'] == 'S')
            ]
            regras = regras[
                (regras['Modalidade'].isna() | (regras['Modalidade'] == modalidade)) &
                (regras['Tipo de carga'].isna() | (regras['Tipo de carga'] == tipo_carga))
            ]
            return regras

        municipios = df_dist['MunicipioOrigem'].unique()

        progress = st.progress(0)
        total = len(municipios)

        for i, municipio in enumerate(municipios):
            uf_municipio = municipio.split('-')[-1].strip()

            for incoterm, tipo_carga, coluna_param in modalidades:
                try:
                    filial_encontrada = False

                    # 1. Filial compatível
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

                    # 2. Filial única no estado
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

                    # 3. Filial mais próxima sem restrição
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

                        # Regras de substituição
                        regras_subs = buscar_regras_substituicao(
                            df_grupos, uf_municipio, incoterm, tipo_carga
                        )
                        for _, regra in regras_subs.iterrows():
                            try:
                                cod_filial_subs = df_filiais[df_filiais['Filial'] == regra['Substituta']]['Codigo'].iloc[0]
                            except:
                                logging.warning(f"Código não encontrado para filial {regra['Substituta']}")
                                cod_filial_subs = '0000'

                            descricao_regra = (
                                f"Regra de Substituição: {regra['Substituta']} recebe coletas de "
                                f"{regra['Grupo Economico']} "
                                f"({regra['Modalidade'] if pd.notna(regra['Modalidade']) else 'Todas as modalidades'}, "
                                f"{regra['Tipo de carga'] if pd.notna(regra['Tipo de carga']) else 'Todos os tipos de carga'})"
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

            progress.progress((i + 1) / total)

        df_resultado = pd.DataFrame(resultados)

        # Caminho para salvar
        nome_completo = f"{nome_arquivo}.xlsx"
        df_resultado.to_excel(nome_completo, index=False)

        st.success("Processo concluído!")

        # Botão para download direto
        with open(nome_completo, "rb") as f:
            dados = f.read()
            st.download_button(
                label="Download do arquivo Excel",
                data=dados,
                file_name=nome_completo,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

        st.write(f"Log de erros salvo em: {os.path.abspath(log_file)}")

    else:
        st.error("Por favor, faça o upload de todos os arquivos necessários.")
