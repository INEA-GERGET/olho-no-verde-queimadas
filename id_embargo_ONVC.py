import pandas as pd
import numpy as np
from pathlib import Path
from io import StringIO

# --- 1. Configuração de Caminho e Constantes ---
NOME_ARQUIVO = "df_id_embargo_ONVC.xlsx"
PASTA_CSV = Path(r"\\Bp-1hd57t3-inea\e\COGET\INPUTS_SCRIPTS")
CAMINHO_COMPLETO_MAPA = PASTA_CSV / NOME_ARQUIVO

# Colunas que, combinadas, formam a chave única do mapeamento
COLUNAS_CHAVE = ['idtxt', 'cod_imovel'] 
COLUNA_ID_LAUDO = 'id_embargo_onvc'
PREFIXO_ID = 'EMQMD'

# --- 2. Função Auxiliar para Encontrar o Próximo ID ---
def get_proximo_contador(df_id_embargo_ONVC, coluna_id=COLUNA_ID_LAUDO, prefixo=PREFIXO_ID):
    """
    Determina o próximo número sequencial a ser usado.
    """
    if df_id_embargo_ONVC is None or df_id_embargo_ONVC.empty or coluna_id not in df_id_embargo_ONVC.columns:
        return 1

    # Extrai a parte numérica de todos os IDs existentes
    ids_existentes = df_id_embargo_ONVC[coluna_id].astype(str).str.extract(f'{prefixo}(\d+)', expand=False)
    
    # Converte para inteiro, ignora NaNs e encontra o máximo
    numeros_existentes = ids_existentes.apply(pd.to_numeric, errors='coerce').fillna(0).astype(int)

    if numeros_existentes.empty:
        return 1
    
    return numeros_existentes.max() + 1


# --- 3. Função para CARREGAR o DataFrame do arquivo ---
def carregar_df_id_embargo_ONVC(colunas_chave=COLUNAS_CHAVE, id_coluna=COLUNA_ID_LAUDO):
    """
    Tenta carregar o DataFrame de mapeamento do arquivo XLSX.
    Retorna o DataFrame ou None se o arquivo não existir/for inválido.
    """
    if CAMINHO_COMPLETO_MAPA.is_file():
        print(f"Arquivo '{NOME_ARQUIVO}' encontrado. Carregando dados...")
        try:
            # Garante que as colunas chave e o ID sejam lidas como string
            dtype_dict = {col: str for col in colunas_chave + [id_coluna]}
            return pd.read_excel(CAMINHO_COMPLETO_MAPA, dtype=dtype_dict)
        except Exception as e:
            print(f"Erro ao ler o arquivo XLSX: {e}")
            return None
    else:
        print(f"Arquivo '{NOME_ARQUIVO}' não encontrado.")
        return None

# --- 4. Função para SALVAR/CRIAR o DataFrame no arquivo ---
def salvar_df_id_embargo_ONVC(df_para_salvar):
    """
    Cria a pasta CSVs (se não existir) e salva o DataFrame no arquivo XLSX.
    """
    try:
        PASTA_CSV.mkdir(parents=True, exist_ok=True)
        df_para_salvar.to_excel(CAMINHO_COMPLETO_MAPA, index=False, sheet_name='Mapeamento LDLRT')
        print(f"df_id_embargo_ONVC salvo com sucesso em: {CAMINHO_COMPLETO_MAPA}")
    except Exception as e:
        print(f"ERRO ao salvar o arquivo XLSX: {e}")


# --- 5. Função Principal de Atualização (Integrada com Load/Save) ---
def atualizar_id_embargo_onvc(df_inicial, colunas_chave=COLUNAS_CHAVE, id_coluna=COLUNA_ID_LAUDO, prefixo=PREFIXO_ID):
    
    # 1. Carrega o mapeamento existente
    df_id_embargo_ONVC = carregar_df_id_embargo_ONVC(colunas_chave=colunas_chave, id_coluna=id_coluna)
    
    # ** PASSO CHAVE: Inicializa df_id_embargo_ONVC se ele não foi carregado **
    if df_id_embargo_ONVC is None or df_id_embargo_ONVC.empty:
        print("Inicializando novo df_id_embargo_ONVC.")
        # O novo DF deve ter as colunas chave e a coluna de ID
        df_id_embargo_ONVC = pd.DataFrame(columns=colunas_chave + [id_coluna])

    # Garante que as colunas chave estejam no formato string para o merge
    for col in colunas_chave:
        df_inicial[col] = df_inicial[col].astype(str)
        if col in df_id_embargo_ONVC.columns:
            df_id_embargo_ONVC[col] = df_id_embargo_ONVC[col].astype(str)
            
    # 2. Junção (Merge) para trazer os IDs existentes para o df_inicial
    merge_cols = colunas_chave + [id_coluna]
    
    df_temp = pd.merge(
        left=df_inicial, 
        right=df_id_embargo_ONVC[merge_cols], 
        on=colunas_chave, # Merge nas colunas 'id' E 'cod_imovel'
        how='left'
    )
    
    # Remove duplicatas da entrada (o arquivo inicial pode ter linhas repetidas)
    df_temp.drop_duplicates(subset=colunas_chave, inplace=True) 

    # 3. Identificar as novas entradas (onde 'id_embargo_onvc' é NaN)
    novas_entradas = df_temp[df_temp[id_coluna].isna()].copy()
    
    # Se não houver novas entradas, retorna e não salva
    if novas_entradas.empty:
        print("Nenhuma nova entrada. df_id_embargo_ONVC não foi modificado.")
        # Retorna o mapeamento existente
        return df_id_embargo_ONVC

    # 4. Determinar o próximo contador (Baseado no df_id_embargo_ONVC carregado/criado)
    proximo_contador = get_proximo_contador(df_id_embargo_ONVC, id_coluna, prefixo)
    
    # 5. Gerar IDs sequenciais para as novas entradas
    novos_ids = []
    
    for _ in range(len(novas_entradas)):
        # Formata com 6 dígitos (ex: LDLRT000001)
        novo_id = f"{prefixo}{proximo_contador:06d}" 
        novos_ids.append(novo_id)
        proximo_contador += 1
        
    novas_entradas.loc[:, id_coluna] = novos_ids # Uso de .loc para evitar SettingWithCopyWarning
    
    # 6. Atualizar o DataFrame mestre (df_id_embargo_ONVC)
    novo_mapeamento = novas_entradas[colunas_chave + [id_coluna]].copy()
    
    # Concatena o mapeamento antigo com o novo
    df_id_embargo_ONVC = pd.concat([df_id_embargo_ONVC, novo_mapeamento], ignore_index=True)
    
    # 7. Salvar o arquivo XLSX
    salvar_df_id_embargo_ONVC(df_id_embargo_ONVC)
    
    return df_id_embargo_ONVC.reset_index(drop=True)

# ----------------------------------------------------------------------
## 🔑 Bloco de Execução de Teste
# ----------------------------------------------------------------------

if __name__ == "__main__":
    print("--- Execução de Teste do id_embargo_onvc ---")
        
    # Caminho do arquivo conforme solicitado
    alerta_file = r'Output\Planilha_ONVC_CAR.xlsx' 
    
    # Carrega o arquivo `alerta`
    if alerta_file:
        try:
            # Tenta ler o arquivo como XLSX (o formato nativo)
            alerta_df = pd.read_excel(alerta_file)
        except FileNotFoundError:
            raise FileNotFoundError(f"Arquivo do alerta ONV não encontrado em: {alerta_file}")
        except Exception as e:
            # Em caso de erro, tenta ler como CSV (com base no histórico de arquivos do usuário)
            print(f"Erro ao ler XLSX ({e}). Tentando ler o CSV correspondente...")
            try:
                # Assume que o CSV correspondente foi gerado
                csv_path = alerta_file.replace('.xlsx', '.csv')
                alerta_df = pd.read_csv(csv_path, sep=';', encoding='latin1', dtype={'id': str, 'cod_imovel': str})
            except Exception as csv_e:
                raise IOError(f"Falha ao ler o arquivo {alerta_file} como XLSX e CSV: {csv_e}")
    else:
        raise FileNotFoundError("Caminho do arquivo do alerta ONV não definido.")

    
    # Colunas que serão usadas como chave única
    colunas_chave_teste = COLUNAS_CHAVE
    
    # Verifica se as colunas necessárias existem no DF lido
    if not all(col in alerta_df.columns for col in colunas_chave_teste):
        raise ValueError(f"O DataFrame de entrada deve conter as colunas: {colunas_chave_teste}.")

    # Prepara o DataFrame para o mapeamento: apenas as colunas chave e remove duplicatas
    df_alerta_mapeamento = alerta_df[colunas_chave_teste].drop_duplicates(subset=colunas_chave_teste).copy()

    # Define o tipo como string para garantir que o merge funcione corretamente
    for col in colunas_chave_teste:
        df_alerta_mapeamento[col] = df_alerta_mapeamento[col].astype(str)
    
    print(f"Preparando DataFrame com combinações únicas de '{colunas_chave_teste[0]}' e '{colunas_chave_teste[1]}' para mapeamento.")

    print("\n--- RODADA 1: Gerando IDs iniciais ---")
    
    # Chama a função principal passando as colunas chave
    df_mapeamento_final_run1 = atualizar_id_embargo_onvc(
        df_alerta_mapeamento, 
        colunas_chave=colunas_chave_teste, 
        id_coluna=COLUNA_ID_LAUDO
    )
    print("\nIDs Gerados (Rodada 1):")
    # Exibe as colunas chave e o ID gerado
    print(df_mapeamento_final_run1[colunas_chave_teste + [COLUNA_ID_LAUDO]])
    
    print("\n" + "="*50)
    
    
    # Simulação da Rodada 2 para teste de persistência
    if CAMINHO_COMPLETO_MAPA.is_file():
        print(f"\n--- RODADA 2: Tentando adicionar novas entradas (usará os IDs já existentes) ---")
        
        # Se o DF de mapeamento não tiver sido alterado, a função não deve criar novos IDs
        df_mapeamento_final_run2 = atualizar_id_embargo_onvc(
            df_alerta_mapeamento, 
            colunas_chave=colunas_chave_teste, 
            id_coluna=COLUNA_ID_LAUDO
        )
        print("\nIDs Gerados (Rodada 2 - Sem novas entradas):")
        print(df_mapeamento_final_run2[colunas_chave_teste + [COLUNA_ID_LAUDO]])
        
    print("\n" + "="*50)
