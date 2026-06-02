'''
Gera a tabela Planilha_ONVC_ECR_CAR.xlsx que é usada na geração dos laudos
Anteriormente era feita manualmente a partir de um Join, aqui combinamos os resultados do intersect com a camada 
para a partir de um merge gerar a tabela
'''

import configparser
import os
import logging
from logging.handlers import RotatingFileHandler
from arcgis.gis import GIS
from arcgis.features import FeatureLayer
import pandas as pd
import urllib3
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

def gerar_tabela():
    """
    Função que compila a execução de todas as funções
    """
    print("\n⏳   Iniciando o processo de mesclagem dos dados...")
    # Lê os arquivos no formato DataFrame
    try:
        # Estes DataFrames não são usados no código original, mas mantidos por contexto
        car_df = pd.read_excel(r"Input\CSVs\tabela_car.xlsx")
        intersecoes_car = pd.read_csv(r"Input\CSVs\tabela_Dissolve_nova_camada_inter_GPL_08_Codigo_CAR.csv", delimiter=';')
        cicatrizes_df = pd.read_excel(r"Input\CSVs\cicatrizes_em_uso.xlsx")
        cicatrizes_df.rename(columns={'cdg_car': 'cod_imovel'})

    except FileNotFoundError as e:
        logging.error(f"❌ Erro ao ler arquivos: Certifique-se de que os arquivos foram baixados corretamente em /CSVs. Detalhes: {e}")
        return # Interrompe a execução se os arquivos não forem encontrados
    
    CHAVE_DE_LIGACAO = 'cod_imovel'

    # Mesclar os DataFrames
    df_mesclado = pd.merge(
        intersecoes_car,  # Tabela base (à esquerda) -> Sufixo _x
        car_df,       # Tabela a ser unida (à direita) -> Sufixo _y
        on=CHAVE_DE_LIGACAO, # Coluna comum para a mesclagem
        how='left'           # Tipo de junção: 'left' join
    )

    COLUNAS_CORRIGIDAS = [
            'OID_', 'idtxt', 
            'data_refer', 'data_ocorr', 'area_m2', 'area_ha',
            'centro_x', 'centro_y', 'link_kml', 'ant_dep', 
            'cod_imovel', 'mod_fiscal', 'area_prop', 'class_fund', 'fmp',
            'fase_processo', 'Shape_Length', 'Shape_Area', 'area_inter', 'porc',
            'nome_imovel', 'nome', 'cpf'
        ]

    # Renomeia as colunas selecionadas para o nome original, removendo o sufixo _x
    df_meio = df_mesclado[COLUNAS_CORRIGIDAS].drop_duplicates()
    df_final = pd.merge(
        df_meio,  # Tabela base (à esquerda)
        cicatrizes_df[['idtxt', 'municipio']],       # Tabela a ser unida (à direita)
        on='idtxt',
        how='left'           # Tipo de junção: 'left' join
    )

    COLUNAS_FINAIS_CORRIGIDAS = [
            'OID_', 'idtxt', 
            'data_refer', 'data_atual', 'area_m2', 'area_ha',
            'centro_x', 'centro_y', 'link_kml', 'ant_dep', 'superinten', 'status',
            'cod_imovel', 'mod_fiscal', 'area_prop', 'class_fund', 'fmp', 'fase_processo',
            'Shape_Length', 'Shape_Area', 'area_inter', 'porc',
            'nome_imovel', 'nome', 'cpf', 'municipio'
        ]
    # Salvar o resultado 
    NOME_ARQUIVO_FINAL = 'Planilha_ONVC_ECR_CAR.xlsx'
    CAMINHO = r'Output\Planilha_ONVC_CAR.xlsx'
    df_final.to_excel(CAMINHO, index=False) # index=False para não incluir a coluna de índice do pandas
    print("\n" + "✨ "*32)
    print(f"✨  Processo concluído. O resultado foi salvo em '{NOME_ARQUIVO_FINAL}' com {len(df_final)} linhas.  ✨")
    print("✨ "*32)

if __name__ == "__main__":
    try:
        gerar_tabela()
    except Exception as e:
        logging.error(f"❌ Erro inesperado durante a execução: {e}")

