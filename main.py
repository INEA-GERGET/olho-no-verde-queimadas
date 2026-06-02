import os
import time
import pandas as pd
import logging
import sys

# --- CONFIGURAÇÕES GLOBAIS ---
CSV_FOLDER = 'CSVs'
PNG_FOLDER = 'antes_depois'
ALERTA_FILE = 'cicatrizes_em_uso.xlsx'
TABELA_CPF = r'Output\Planilha_ONVC_CAR.xlsx'
INPUT_CAMADA = 'cicatrizes_em_uso.shp'
ID_COLUMN = 'id'

def setup_logging():
    """Configura o log de forma limpa e ú nica."""
    root = logging.getLogger()
    if root.handlers:
        for handler in root.handlers[:]:
            root.removeHandler(handler)
            
    logging.basicConfig(
        level=logging.INFO,
        format='%(levelname)s: %(message)s',
        handlers=[logging.StreamHandler(sys.stdout)]
    )
    sys.stdout.flush()

def gerar_merge():
    setup_logging()
    
    print("\n" + "="*100)
    print("="*20 + " "*15 + "GERAÇÃO DA TABELA DE ONVC_CAR" + " "*16 + "="*20)
    print("="*100)
    print("\nPara gerar a tabela usada nos laudos de embargo, certifique-se que os arquivos necessários estejam")
    print(" na pasta 'CSVs' ->")
    print(" "*20 + "'tabela_car.xlsx'")
    print(" "*20 + "'cicatrizes_em_uso.xlsx' e")
    print(" "*20 + "'tabela_Dissolve_nova_camada_inter_GPL_08_Codigo_CAR.csv'\n")
    print("="*100)

    entrada = input("Deseja gerar a tabela de merge? (s/n): ").strip().lower()
    if entrada != 's':
        logging.info("Operação de geração de tabela de merge cancelada pelo usuário.")
        return

    # 1. MENSAGEM INICIAL E MESCLAGEM
    sys.stdout.write("Gerando a tabela usada nos laudos de embargo...\n")
    sys.stdout.flush()

    try:
        # Importação tardia da função de merge
        from gerar_tabela_merge_onvc import gerar_tabela
        gerar_tabela()
        
        # Pausa para o terminal processar as mensagens da mesclagem
        time.sleep(1.5) 
    except Exception as e:
        logging.error(f'❌ Erro ao gerar a tabela de merge: {e}')


def gerar_ids_embargo():
    """Lógica para garantir que o mapeamento de IDs de laudo esteja atualizado."""
    try:
        from id_embargo_ONVC import atualizar_id_embargo_onvc, COLUNAS_CHAVE
        
        logging.info("Verificando e atualizando mapeamento de IDs de Laudo...")
        
        # Carrega a tabela que contém ID do Alerta e Código do CAR
        if not os.path.exists(TABELA_CPF):
            logging.error(f"Arquivo {TABELA_CPF} não encontrado para gerar IDs.")
            return False

        df_entrada = pd.read_excel(TABELA_CPF)
        
        # Chama a função do script id_embargo_ONV.py
        # Ela carrega o arquivo existente, adiciona novos se houver e salva.
        atualizar_id_embargo_onvc(df_entrada)
        
        logging.info("Mapeamento de IDs (LDEMBGR) concluído.")
        return True
    except Exception as e:
        logging.error(f"Erro ao processar IDs de laudo: {e}")
        return False

def gerar_laudos_embargo():
    try:
        
        # Importações de processamento (tardias para evitar conflitos)
        import geopandas as gpd
        from gerar_embargo_onvc import plot_map_for_idtxt_satellite, create_pdf_for_idtxt

    except Exception as e:
        logging.error(f"❌ Ocorreu um erro inesperado: {e}")
    
    finally:
        # A MENSAGEM FINAL
        logging.info("\n" + "="*65)
        logging.info("✨✨ Processamento concluído com sucesso. ✨✨")
        logging.info("="*65)
        sys.stdout.flush()
        
        # MATAR O PROCESSO: Isso impede que o Python tente rodar o script de novo
        # por causa de algum import circular ou limpeza de cache.
        os._exit(0) 

def gerar_laudos_alerta():
    try:
        
        # Importações de processamento (tardias para evitar conflitos)
        import geopandas as gpd
        from gerar_alerta_onvc import plot_map_for_idtxt_satellite, create_pdf_for_idtxt

    except Exception as e:
        logging.error(f"❌ Ocorreu um erro inesperado: {e}")
    
    finally:
        # A MENSAGEM FINAL
        logging.info("\n" + "="*65)
        logging.info("✨✨ Processamento concluído com sucesso. ✨✨")
        logging.info("="*65)
        sys.stdout.flush()
        
        # MATAR O PROCESSO: Isso impede que o Python tente rodar o script de novo
        # por causa de algum import circular ou limpeza de cache.
        os._exit(0) 

if __name__ == "__main__":
    # Garante que o main só rode se o arquivo for executado diretamente

    print("\n" + "="*100)
    print("="*20 + " "*9 + "GERAÇÃO DOS LAUDOS - ONVC QUEIMADAS 🔥🔥🔥" + " "*10 + "="*20)
    print("="*100)
    print("\nPara gerar os laudos, certifique-se que os arquivos necessários estejam na pasta correta:")
    print(" na pasta 'CSVs' ->")
    print(" "*20 + "'tabela_car.xlsx' (APENAS PARA EMBARGOS),")
    print(" "*20 + "'alertas_em_uso.xlsx',")
    print(" "*20 + "'todas as tabelas dissolves'")
    print(" na pasta 'camada' ->")
    print(" "*20 + "'alertas_em_uso.shp'\n")
    print(" na pasta 'antes_depois' ->")
    print(" "*20 + "imagens de antes e depois nomeadas com o id\n") 
    print("Esteja conectado ao COGET (APENAS PARA EMBARGOS) ->")
    print(" "*20 + "'df_id_embargo_ONV.xlsx' é atualizado automaticamente no COGET\n") 
    print("="*100)

    while True:
        entrada = input("Deseja gerar laudos de alerta ou laudos de embargo? (alerta/embargo): ").strip().lower()
        if entrada == 'embargo':
            gerar_merge()
            gerar_ids_embargo()
            gerar_laudos_embargo()
            break
        elif entrada == 'alerta':
            gerar_laudos_alerta()
            break
        else:
            print("Entrada inválida. Por favor, digite 'alerta' ou 'embargo' para continuar.")
