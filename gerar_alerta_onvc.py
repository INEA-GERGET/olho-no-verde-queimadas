'''
Script para gerar os laudos das cicatrizes. Ele usa a tabela gerada pelo código "gerar_tabela_usada_no_layout" como um dos inputs
Além de um shape das cicatrizes, tabela de referencias contendo os links das camadas, imagens do papel timbrado e as logos do INEA
Também utiliza uma tabela da camada em uso que também provém do código "gerar_tabela_usada_no_layout" e os arquivos do dissolve que devem ter as 
tabelas exportadas usando o script de conversão CSV.
As tabelas excel (xlsx) devem ser salvas como "Pasta de Trabalho do Excel" e os arquivos em CSV como "CSV-UTF-8" (padrão nos scripts).
Este script deve estar em uma pasta dedicadas com todos os arquivos necessários. e pastas de output organizados da seguinte maneira:
Pastas principal:
- este script
- pasta output
- pasta contendo os csvs e xlsx
- pasta contendo as imagens
- pasta contendo o shp da camada
- pasta dedicada para logs
- demais scripts
'''

import geopandas as gpd
import matplotlib.pyplot as plt
import numpy as np
import os
import contextily as ctx
import pandas as pd
import logging
from logging.handlers import RotatingFileHandler
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.units import inch, cm
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib.colors import HexColor
from reportlab.lib.styles import ParagraphStyle
from PIL import Image, ImageDraw
from reportlab.platypus import PageBreak
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, Image as RLImage
from PIL import Image as PILImage
import warnings
from reportlab.lib.enums import TA_CENTER, TA_JUSTIFY, TA_LEFT
from gerar_mapa_onvc import gerar_mapa_alerta
import string

pd.set_option('future.no_silent_downcasting', True)

# Ignora o Warning de geometria:
warnings.filterwarnings('ignore', message="Measured (M) geometry types are not supported. Original type 'Measured 3D Polygon' is converted to 'Polygon Z'", category=UserWarning)

# ----------------------------------------------------------------- Configurações -----------------------------------------------------------------

# Pastas dos prints de antes e depois:
pasta_AD = r"Input\antes_depois"

# Pasta onde se encontram os CSVs e o XLSX
csv_folder = r"Input\CSVs"
id_embargo_folder = r"Input"
ref_folder = r"arquivos"

# Planilhas de input, a gerada pelo script e a de referências
referencias = os.path.join(ref_folder, 'referencias.csv')

# Camada dos alertas (Usado para gerar os mapas)
input_camada = r"Input\camada\cicatrizes_em_uso.shp" 

# Excel da camada dos alertas
cicatriz_file = r'Input\CSVs\cicatrizes_em_uso.xlsx'

# Imagens do papel timbrado:
imagem_cbc1 = r"arquivos\cbc1.jpg"
imagem_cbc2 = r"arquivos\cbc2.jpg"

# Logos:
logo_olho_no_verde_caminho = r"arquivos\Logo_Olho_no_Verde_(P)_Vertical.png"

# Identificações dos arquivos de interseção ou dissolve, o nome padrão que tem no começo de todos os arquivos
inter_prefix = 'tabela_Dissolve_nova_camada_inter'

# Coluna onde se encontra o ID nos arquivos
id_column = 'idtxt'

# Criando as pastas de saída (Não precisa configurar)
output_folder = f"Output"
if not os.path.exists(output_folder):
    os.makedirs(output_folder)
png_folder = os.path.join(output_folder, "mapas")
if not os.path.exists(png_folder):
    os.makedirs(png_folder)
layout_folder = os.path.join(output_folder, "laudos_alerta_onvc")
if not os.path.exists(layout_folder):
    os.makedirs(layout_folder)
mapas_folder = os.path.join(output_folder, "mapa")
if not os.path.exists(mapas_folder):
    os.makedirs(mapas_folder)

# Inicia um contador
cntd = 1

# Função para gerar o PDF para cada `id`
styles = getSampleStyleSheet()

# 1. Adicionando o estilo CertificadoTitulo, conforme solicitado
styles.add(ParagraphStyle(name='Titulo',
                            parent=styles['Normal'],
                            fontName='Helvetica-Bold',
                            fontSize=20,
                            alignment=TA_CENTER, 
                            spaceAfter=0.5 * cm,
                            textColor=HexColor("#000000")))

styles.add(ParagraphStyle(name='Titulo2',
                            parent=styles['Normal'],
                            fontName='Helvetica-Bold',
                            fontSize=16,
                            alignment=TA_LEFT, 
                            spaceAfter=1 * cm,
                            textColor=HexColor("#000000")))

styles.add(ParagraphStyle(name='Subtitulo',
                            parent=styles['Normal'],
                            fontName='Helvetica',
                            fontSize=10,
                            alignment=TA_CENTER, 
                            spaceAfter=0.5 * cm,
                            textColor=HexColor("#000000")))

styles.add(ParagraphStyle(name='CorpoTexto',
                            parent=styles['Normal'],
                            fontName='Helvetica',
                            fontSize=12,
                            alignment=TA_JUSTIFY, 
                            spaceAfter=0.5 * cm,
                            textColor=HexColor("#000000")))

styles.add(ParagraphStyle(name='CorpoTexto1',
                            parent=styles['Normal'],
                            fontName='Helvetica',
                            fontSize=12,
                            alignment=TA_CENTER, 
                            spaceAfter=0.5 * cm,
                            textColor=HexColor("#000000")))

styles.add(ParagraphStyle(name='HeaderTable',
                            parent=styles['Normal'],
                            fontName='Helvetica-Bold',
                            fontSize=12,
                            alignment=TA_CENTER, 
                            spaceAfter=0.5 * cm,
                            textColor=HexColor("#000000")))


styles.add(ParagraphStyle(name='Legenda',
                            parent=styles['Normal'],
                            fontName='Helvetica',
                            fontSize=12,
                            alignment=TA_CENTER, 
                            spaceAfter=0.5 * cm,
                            textColor=HexColor("#000000")))


# Estilo para células de dados: centralizado
centered_style = ParagraphStyle(name='CenteredStyle', parent=styles['Normal'], alignment=TA_CENTER, fontSize=10, leading=12, fontName='Helvetica')
# Estilo para cabeçalhos de tabela: negrito e centralizado
header_style = ParagraphStyle(name='HeaderStyle', parent=centered_style, fontName='Helvetica-Bold')


# ----------------------------------------------------------------- Definição das funções -----------------------------------------------------------------

def setup_logging():
    log_format = '%(asctime)s - %(name)s - %(levelname)s - %(filename)s:%(lineno)d - %(message)s'
    log_file = os.path.join("logs", "execucao.log") # Diretório dedicado para logs
    handlers = [RotatingFileHandler(log_file, maxBytes=5*1024*1024, backupCount=3, encoding="utf-8"), logging.StreamHandler()]
    logging.basicConfig(level=logging.INFO, format=log_format, handlers=handlers)

def addPageNumber(canvas, doc, current_idtxt):
    '''
    Adiciona a numeração de página no canto inferior direito e o ID do Laudo
    no canto inferior esquerdo.
    '''
    canvas.saveState()
    # Define a fonte e o tamanho
    canvas.setFont('Helvetica', 10)
    
    # --- LÓGICA EXISTENTE: PÁGINA (Canto Inferior Direito) ---
    page_num = canvas.getPageNumber()
    text_page = f"Página {page_num}"
    
    # Posição: 1 cm da margem direita (A4[0] - cm) e 1.5 cm do fundo
    x_position_right = A4[0] - cm 
    y_position = 1.5 * cm 
    
    # Desenha a string no canto inferior direito
    canvas.drawRightString(x_position_right, y_position, text_page)
    
    # --- NOVA LÓGICA: RELATÓRIO DO LAUDO (Canto Inferior Esquerdo) ---
    report_text = f"Laudo da Cicatriz {current_idtxt}"
    # Posição: 1.5 cm da margem esquerda e 1.5 cm do fundo (mesma altura)
    x_position_left = cm 
    
    # Desenha a string no canto inferior esquerdo
    canvas.drawString(x_position_left, y_position, report_text)
    
    canvas.restoreState()

# Função para gerar o mapa em PNG para cada `idtxt`
def plot_map_for_idtxt_satellite(gdf, idtxt, png_folder, buffer_distance=19, color='none', opacity=0.9):
    ''''
    função que gera a imagem do mapa que é usada no layout com imagem de satélite
    '''
    gdf_filtered = gdf[gdf['idtxt'] == str(idtxt)]
    if gdf_filtered.empty:
        logging.error(f"ID {idtxt} não encontrado no shapefile.")
        return
    
    # Reprojetar para Web Mercator
    gdf_filtered = gdf_filtered.to_crs(epsg=3857)
    bounds = gdf_filtered.total_bounds
    xmin, ymin, xmax, ymax = bounds
    margin = buffer_distance
    xlim = (xmin - margin, xmax + margin)
    ylim = (ymin - margin, ymax + margin)
    
    # Criar figura e eixo
    fig, ax = plt.subplots(figsize=(10, 10))
    gdf_filtered.plot(ax=ax, color=color, edgecolor='red', alpha=opacity, linewidth=3.5)
    
    # Adicionar basemap de imagens de satélite
    basemap_source="https://services.arcgisonline.com/ArcGIS/rest/services/World_Imagery/MapServer/tile/{z}/{y}/{x}"
    ctx.add_basemap(ax, source=basemap_source, zoom=18)
    
    # Definir os limites
    ax.set_xlim(xlim)
    ax.set_ylim(ylim)
    
    # Salvar o mapa
    output_path = os.path.join(png_folder, f"{idtxt}.png")
    plt.savefig(output_path, dpi=300, bbox_inches='tight')
    plt.close()


# Função para gerar o PDF para cada `idtxt`
def create_pdf_for_idtxt(idtxt, cicatriz_data, inter_data, png_folder, cntd, link_poligono, link_imagemAD, current_idtxt):
    '''
    Função que gera o PDF para cada ID
    Entradas: id (em loop), DF da camada de cicatriz, DF da interseção, pasta onde estão os mapas, contador (variável), link kml, nome do CAR, nome da pessoa
    '''
    pdf_file = os.path.join(layout_folder, f"laudo_cicatriz_{idtxt}.pdf")

    
    # Definindo a função wrapper para passar o ID do laudo ao rodapé
    def add_page_num_wrapper(canvas, doc):
        addPageNumber(canvas, doc, current_idtxt)
    
    doc = SimpleDocTemplate(pdf_file, pagesize=A4, leftMargin=1.5 * cm, rightMargin=1.5 * cm,
                            topMargin=1.5 * cm, bottomMargin=2.5 * cm)

    # Formatando as colunas para mostrar apenas dois dígitos após a vírgula
    if 'Área (m²)' in cicatriz_data.columns:
        cicatriz_data['Área (m²)'] = cicatriz_data['Área (m²)'].fillna(0).astype(str).str.replace(',', '.').apply(lambda x: f"{float(x):.2f}")
    if 'Centróide X' in cicatriz_data.columns:
        cicatriz_data['Centróide X'] = cicatriz_data['Centróide X'].fillna(0).astype(str).str.replace(',', '.').apply(lambda x: f"{float(x):.4f}")
    if 'Centróide Y' in cicatriz_data.columns:
        cicatriz_data['Centróide Y'] = cicatriz_data['Centróide Y'].fillna(0).astype(str).str.replace(',', '.').apply(lambda x: f"{float(x):.4f}")
    if 'Área (ha)' in cicatriz_data.columns:
        cicatriz_data['Área (ha)'] = cicatriz_data['Área (ha)'].fillna(0).astype(str).str.replace(',', '.').apply(lambda x: f"{float(x):.4f}")
        
    # Para inter_data
    if 'Porcentagem (%)' in inter_data.columns:
        inter_data['Porcentagem (%)'] = inter_data['Porcentagem (%)'].fillna('0').astype(str).str.replace(',', '.').apply(lambda x: f"{float(x):.3f}")
    if 'Área da interseção (m²)' in inter_data.columns:
        inter_data['Área da interseção (m²)'] = inter_data['Área da interseção (m²)'].fillna('0').astype(str).str.replace(',', '.').apply(lambda x: f"{float(x):.2f}")    
    if 'area_ha' in inter_data.columns:
        inter_data['area_ha'] = inter_data['area_ha'].fillna('0').astype(str).str.replace(',', '.').apply(lambda x: f"{float(x):.2f}")
    if 'area_m2' in inter_data.columns:
        inter_data['area_m2'] = inter_data['area_m2'].fillna('0').astype(str).str.replace(',', '.').apply(lambda x: f"{float(x):.2f}")

    # Estilo para células de dados: centralizado
    centered_style = ParagraphStyle(name='CenteredStyle', parent=styles['Normal'], alignment=TA_CENTER, fontSize=10, leading=12, fontName='Helvetica')
    # Estilo para cabeçalhos de tabela: negrito e centralizado
    header_style = ParagraphStyle(name='HeaderStyle', parent=centered_style, fontName='Helvetica-Bold', fontSize=10, leading=12)

    def clean_data(data):
        ''' Pequena função para limpar os DataFrames removendo dados vazios '''
        data = data.dropna(how='all')  # Remove linhas totalmente vazias
        data = data.dropna(axis=1, how='all')  # Remove colunas totalmente vazias
        data = data.fillna('-')  # Preenche valores NaN restantes com marcador
        data = data[~(data == '-').all(axis=1)]  # Remove linhas onde todos os valores são '-'
        return data

    def add_table_with_split(data, title, qtd_colunas, col_widths_percent=None):
        data = pd.DataFrame(data)
        try:
            data = clean_data(data)
        except Exception:
            return

        if data.empty:
            return

        elements.append(Paragraph(title, styles['Titulo2']))
        columns = data.columns.tolist()
        num_tables = (len(columns) + qtd_colunas - 1) // qtd_colunas

        for i in range(num_tables):
            start_col = i * qtd_colunas
            end_col = min(start_col + qtd_colunas, len(columns))
            sub_data = data.iloc[:, start_col:end_col]
            sub_data = clean_data(sub_data)
            if sub_data.empty:
                continue
            table_data = [sub_data.columns.tolist()]
            custom_style = ParagraphStyle(name='CustomStyle', parent=styles['CorpoTexto'], fontSize=12, alignment=TA_CENTER)
            
            for row in sub_data.values.tolist():
                new_row = []
                for cell in row:
                    cell_text = str(cell)
                    new_row.append(Paragraph(cell_text, custom_style))
                table_data.append(new_row)
            
            if col_widths_percent and len(col_widths_percent) == len(sub_data.columns):
                largura_total = A4[0] - 2 * cm
                colWidths = [largura_total * (p/100) for p in col_widths_percent]
            else:
                colWidths = (A4[0] - 2 * cm) / len(sub_data.columns)
            
            sub_table = Table(table_data, colWidths=colWidths)

            style = [
                ('BACKGROUND', (0, 0), (-1, 0), HexColor("#cb7f38")),
                ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                ('FONTNAME', (0, 1), (-1, -1), 'Helvetica'),
                ('GRID', (0, 0), (-1, -1), 0.5, HexColor("#9b5f26")),
                ('FONTSIZE', (0, 0), (-1, -1), 12),
                ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
                ('vAlign', (0, 0), (-1, -1), 'MIDDLE'),
            ]

            sub_table.setStyle(TableStyle(style))
            elements.append(sub_table)


    def add_table_without_space(data, qtd_colunas, col_widths_percent=None):
        data = pd.DataFrame(data)
        try:
            data = clean_data(data)
        except Exception:
            return

        if data.empty:
            return

        columns = data.columns.tolist()
        num_tables = (len(columns) + qtd_colunas - 1) // qtd_colunas

        for i in range(num_tables):
            start_col = i * qtd_colunas
            end_col = min(start_col + qtd_colunas, len(columns))
            sub_data = data.iloc[:, start_col:end_col]
            sub_data = clean_data(sub_data)
            if sub_data.empty:
                continue
            table_data = [sub_data.columns.tolist()]
            custom_style = ParagraphStyle(name='CustomStyle', parent=styles['CorpoTexto'], fontSize=12, alignment=TA_CENTER)
            
            for row in sub_data.values.tolist():
                new_row = []
                for cell in row:
                    cell_text = str(cell)
                    new_row.append(Paragraph(cell_text, custom_style))
                table_data.append(new_row)
            
            if col_widths_percent and len(col_widths_percent) == len(sub_data.columns):
                largura_total = A4[0] - 2 * cm
                colWidths = [largura_total * (p/100) for p in col_widths_percent]
            else:
                colWidths = (A4[0] - 2 * cm) / len(sub_data.columns)
            
            sub_table = Table(table_data, colWidths=colWidths)

            style = [
                ('BACKGROUND', (0, 0), (-1, 0), HexColor("#cb7f38")),
                ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                ('FONTNAME', (0, 1), (-1, -1), 'Helvetica'),
                ('GRID', (0, 0), (-1, -1), 0.5, HexColor("#9b5f26")),
                ('FONTSIZE', (0, 0), (-1, -1), 12),
                ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
                ('vAlign', (0, 0), (-1, -1), 'MIDDLE'),
            ]

            sub_table.setStyle(TableStyle(style))
            elements.append(sub_table)

    
    def add_table_REF(data, title, qtd_colunas, col_widths_percent=None):
        data = pd.DataFrame(data)
        try:
            data = clean_data(data)
        except Exception:
            return

        if data.empty:
            return

        elements.append(Paragraph(title, styles['Titulo2']))
        columns = data.columns.tolist()
        num_tables = (len(columns) + qtd_colunas - 1) // qtd_colunas

        for i in range(num_tables):
            start_col = i * qtd_colunas
            end_col = min(start_col + qtd_colunas, len(columns))
            sub_data = data.iloc[:, start_col:end_col]
            sub_data = clean_data(sub_data)
            if sub_data.empty:
                continue
            table_data = [sub_data.columns.tolist()]
            custom_style = ParagraphStyle(name='CustomStyle', parent=styles['CorpoTexto'], fontSize=11, alignment=TA_CENTER)
            
            for row in sub_data.values.tolist():
                new_row = []
                for cell in row:
                    cell_text = str(cell)
                    new_row.append(Paragraph(cell_text, custom_style))
                table_data.append(new_row)
            
            if col_widths_percent and len(col_widths_percent) == len(sub_data.columns):
                largura_total = A4[0] - 2 * cm
                colWidths = [largura_total * (p/100) for p in col_widths_percent]
            else:
                colWidths = (A4[0] - 2 * cm) / len(sub_data.columns)
            
            sub_table = Table(table_data, colWidths=colWidths)

            style = [
                ('BACKGROUND', (0, 0), (-1, 0), HexColor("#cb7f38")),
                ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                ('FONTNAME', (0, 1), (-1, -1), 'Helvetica'),
                ('GRID', (0, 0), (-1, -1), 0.5, HexColor("#9b5f26")),
                ('FONTSIZE', (0, 0), (-1, -1), 12),
                ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
                ('vAlign', (0, 0), (-1, -1), 'MIDDLE'),
            ]

            sub_table.setStyle(TableStyle(style))
            elements.append(sub_table)
            #elements.append(Spacer(1, 0.5 * cm))

    def add_tables_by_columns(data, title_prefix, qtd_colunas, remover_id):
        ''' Função para adicionar especificamente as tabelas do intersect,
        Funciona semelhante a comum, mas lida melhor com as diferente interseções'''

        campos_a_manter = ['zona', 'label', 'cod_imovel', 'tp_app9', 'Tipo_8', 'Tipo_7', 'Tipo_6', 'tipo_5', 'Tipo_4', 'Tipo_3', 'tipo_ap,p2', 'Tipo1', 
                        'nome_ucm', 'nome_uce', 'nome_ucf', 'municn', 'rotulo', 'grau', 'classe_05', 'vegetacao2024', 'classe_01','classe_02', 'classe_03', 'classe_04',
                        'Nome_RPPN', 'empreend', 'nprocess', 'numero_pro', 'instrument', 'descriptio', 'name_2', 'name_3'] # Atualizado (cecilia)
        
        # Dicionário dos titulos de acordo com o nome das colunas
        titulos_colunas = {'zona': 'Zoneamento UCs Estaduais', 'label': 'Zoneamento UCs Estaduais', 'cod_imovel': 'Imóvel do CAR', 'tp_app9': 'Área de Preservação Permanente', 
            'Tipo_8': 'Faixa Marginal de Proteção', 'Tipo_7': 'Área de Preservação Permanente', 'Tipo_6': 'Área de Preservação Permanente', 'tipo_5': 'Área de Preservação Permanente', 
            'Tipo_4': 'Área de Preservação Permanente', 'Tipo_3': 'Área de Preservação Permanente', 'tipo_app2': 'Área de Preservação Permanente', 'Tipo1': 'Área de Preservação Permanente', 
            'nome_ucm': 'Unidades de Conservação Municipais',
            'nome_uce': 'Unidades de Conservação Estaduais',
            'nome_ucf': 'Unidades de Conservação Federais', 'municn': 'Município', 
            'rotulo': 'Zona de Amortecimento Estadual', 'grau': 'Área de Uso Restrito',
            'classe_05': 'Uso e Cobertura do INEA (2018)', 'vegetacao2024': 'Cobertura Vegetal (INEA - 2024)', 
            'classe_01': 'Corrida de Massa','classe_02': 'Enxurrada', 'classe_03': 'Inundação', 'classe_04': 'Movimento de Massa', 'Nome_RPPN': 'Nome da RPPN',
            'empreend': 'Autorização de Supressão Vegetal', 'Nome_FDA': 'Limites Restauração (Florestas do Amanhã)',
            'numero_pro': 'Licenças Estaduais', 'descriptio': 'Licenças Municipais'} # Atualizado (cecilia)

        dicionario_dissolve = {'classe_05': 'Classe', 'class': 'Classe', 'apps': 'APPs', 
            'municipio': 'Municipio', 'vegetacao2': 'Vegetação', 'NOME': 'Restinga', 'instrument': 'Instrumento Emitido',
            'NomeOficia': 'Nome UCM', 'Tipo_16': 'Nome UCF', 'nome_uce': 'Nome UCE', 'rotulo': 'ZA UCE', 'Geocode': 'Código FDA',
            'Nome_RPPN': 'Nome da RPPN', 'descriptio': 'Número do Processo', 'name_2': 'Tipo da Licença', 'name_3': 'Instrumento Emitido',
            'numero_pro': 'Número do Processo', 'empreend': 'Empreendimento', 'nprocess': 'Número do Processo', 'tp_app9': 'Tipo',
            'tipo_app2': 'Tipo', 'nome_ucm': 'Nome', 'nome_uce': 'Nome', 'cod_imovel': 'Código do CAR',
            'nome_ucf': 'Nome', 'municn': 'Nome', 'rotulo': 'Rótulo UC Estadual', 'vegetacao2024': 'Tipo',
            'Tipo_8': 'Tipo', 'Tipo_7': 'Tipo', 'tipo_5': 'Tipo', 'Tipo_3': 'Tipo', 'Tipo1': 'Tipo', 
            'nm_municip': 'Nome município', 'zona': 'Zona', 'grau': 'Grau de Declividade', 'Tipo_6': 'Tipo',
            'Nome_FDA': 'Nome FDA', 'Tipo_4': 'Tipo', 'Nome': 'Declividade', 'label': 'Rótulo UC Estadual',
            'classe_01': 'Classe','classe_02': 'Classe', 'classe_03': 'Classe', 'classe_04': 'Classe', 'descriptio': 'Número do Processo', 
            'name_2': 'Tipo da Licença', 'name_3': 'Instrumento Emitido'}
        
        # Iterar sobre cada coluna em `manter_colunas`
        for coluna in campos_a_manter:
            if coluna in data.columns:
                # Processar cada valor único na coluna
                for valor in data[coluna].unique():
                    sub_data = data[data[coluna] == valor].copy()
                    sub_data = clean_data(sub_data)
                    # Pular conjuntos de dados vazios
                    if sub_data.empty:
                        continue
                    if remover_id == 1 and 'idtxt' in sub_data.columns:
                        sub_data = sub_data.drop(columns=['idtxt'])
                    # Continua com o processamento normal das demais colunas
                    columns = sub_data.columns.tolist()
                    num_tables = (len(columns) + qtd_colunas - 1) // qtd_colunas

                    # Loop para dividir a tabela em sub-tabelas
                    for i in range(num_tables):
                        start_col = i * qtd_colunas
                        end_col = min(start_col + qtd_colunas, len(columns))
                        sub_data = sub_data.iloc[:, start_col:end_col]
                        # Limpa a sub-tabela após a divisão
                        sub_data = clean_data(sub_data)
                        # Se a sub-tabela estiver vazia após a limpeza, pula
                        if sub_data.empty:
                            continue
                        elements.append(Paragraph(f"{title_prefix} - {titulos_colunas[coluna]}: ", styles['Titulo2']))
                        sub_data = sub_data.rename(columns=dicionario_dissolve)
                        # Formatação dos dados com 2 casas decimais para valores numéricos
                        # Transforma cabeçalhos e células em Parágrafos para permitir quebra de linha
                        header_row = [Paragraph(str(col), styles['HeaderTable']) for col in sub_data.columns]
                        table_data = [header_row]
                        for row in sub_data.values.tolist():
                            new_row = []
                            for cell in row:
                                cell_text = str(cell)
                                new_row.append(Paragraph(cell_text, styles['CorpoTexto1']))
                            table_data.append(new_row)
                        # Criação e estilo da sub-tabela
                        sub_table = Table(table_data, colWidths=(A4[0] - 0.8 * inch) / len(sub_data.columns))
                        sub_table.setStyle(TableStyle([
                            ('BACKGROUND', (0, 0), (-1, 0), HexColor("#cb7f38")),  # Fundo para os títulos
                            ('GRID', (0, 0), (-1, -1), 0.5, HexColor("#9b5f26")),  # Linhas da grade
                            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE')                 # Alinhamento vertical
                        ]))
                        elements.append(sub_table)
                        elements.append(Spacer(1, 0.2 * inch))

    # Lista que armazena todos os elementos que serão inseridos no PDF
    elements = []

    # Seção de Cabeçalho (CBC1 e Olho No Verde)
    if os.path.exists(imagem_cbc1):
        img_cbc1 = RLImage(imagem_cbc1)
        img_cbc1.drawHeight = 4 * cm
        img_cbc1.drawWidth = 7.72 * cm
        elements.append(img_cbc1)

    if os.path.exists(logo_olho_no_verde_caminho):
        img_olho_no_verde = RLImage(logo_olho_no_verde_caminho)
        img_olho_no_verde.drawHeight = 3.9 * cm
        img_olho_no_verde.drawWidth = 7 * cm
        elements.append(img_olho_no_verde)
        elements.append(Spacer(1, 0.2 * cm))

    # Adiciona o título
    elements.append(Paragraph(f"Laudo da Cicatriz", styles['Titulo']))
    elements.append(Spacer(0, 0 * cm))

    # Adiciona o subtítulo
    subtitulo = "Documento automatizado gerado a partir do sistema de detecção de alertas de cicatrizes de queimadas"
    centered_style_subtitle = styles['Subtitulo'].clone('CenteredStyle')
    centered_style_subtitle.alignment = TA_CENTER
    elements.append(Paragraph(subtitulo, centered_style_subtitle))
    elements.append(Spacer(0, 0 * cm))

    # Introdução
    introducao = """O Programa Olho no Verde - Queimadas realiza o monitoramento por intermédio de disponibilização sistemática e contínua de 
        produtos espectrais, fruto de uma constelação de satélites, que geram imagens de alta resolução espacial. 
        O método de aquisição das informações se dá por meio do processamento automático e semiautomático utilizando técnicas de 
        sensoriamento remoto e aprendizagem de máquina. Detectada a mudança na vegetação, a partir da comparação de imagens de diferentes datas, 
        é materializado o polígono resultante do processamento. """
    elements.append(Paragraph(introducao, styles['CorpoTexto']))
    elements.append(Spacer(1, 0.5 * cm))

    # Ajustando dados mapa e informações

    info_df = pd.read_excel(cicatriz_file)
    mapa_info = info_df.rename(columns={'idtxt': 'ID da Cicatriz', 'centro_x': 'Latitude (centro)',
                                       'centro_y': 'Longitude (centro)', 'municipio': 'Município'})
    cidade = mapa_info.loc[mapa_info['ID da Cicatriz'] == current_idtxt, 'Município'].values[0]
    # Capitaliza a string
    cidade_normalizada = string.capwords(cidade)
    
    # Mapeamento de correção de nomes
    lista_cidades_renomear = {'Aperibe':'Aperibé', 
                                'Armação Dos Buzios':'Armação Dos Búzios',
                                'Barra Do Pirai':'Barra Do Piraí',
                                'Conceicao De Macabu':'Conceição De Macabu',
                                'Itaborai':'Itaboraí',
                                'Itaguai':'Itaguaí',
                                'Laje Do Muriae':'Laje Do Muriaé',
                                'Macae':'Macaé',
                                'Mage':'Magé',
                                'Marica':'Maricá',
                                'Nilopolis':'Nilópolis',
                                'Niteroi':'Niterói',,
                                'Nova Iguacu':'Nova Iguaçu',
                                'Paraiba Do Sul':'Paraíba Do Sul',
                                'Petropolis':'Petrópolis',
                                'Pirai':'Piraí',
                                'Porcincula':'Porciúncula',
                                'Quissama':'Quissamã',
                                'Santo Antonio De Padua':'Santo Antônio De Pádua',
                                'Sao Francisco De Itabapoana':'São Francisco De Itabapoana',
                                'Sao Fidelis':'São Fidélis',
                                'Sao Goncalo':'São Gonçalo',
                                'Sao Joao Da Barra':'São João Da Barra',
                                'Sao Joao De Meriti':'São João De Meriti',
                                'Sao Jose De Uba':'São José De Ubá',
                                'Sao Jose Do Vale Do Rio Preto':'São José Do Vale Do Rio Preto',
                                'Sao Pedro Da Aldeia':'São Pedro Da Aldeia',
                                'Sao Sebastiao Do Alto':'São Sebastião Do Alto',
                                'Seropedica':'Seropédica',
                                'Tangua':'Tanguá',
                                'Teresopolis':'Teresópolis',
                                'Tres Rios':'Três Rios',
                                'Valenca':'Valença',
                                'Varre Sai':'Varre-Sai'}
    
    # Obtém o nome final. Se 'cidade_normalizada' estiver no dicionário, usa o valor
    # Se não estiver, usa 'cidade_normalizada' como fallback.
    cidade_final_str = lista_cidades_renomear.get(cidade_normalizada, cidade_normalizada)
    cidade_df = pd.DataFrame({'Município da Cicatriz': [cidade_final_str]})


    # Adiciona tabela de cicatriz com divisão, se necessário
    alerta_info = pd.DataFrame({'ID da Cicatriz' : [current_idtxt]})
    add_table_with_split(alerta_info, "Informações da Cicatriz", 1, [100])
    add_table_without_space(cidade_df, 1, [100])
    alerta_info2 = cicatriz_data[['Área (ha)', 'Área (m²)', 'Data antes', 'Data depois', 'Centróide X', 'Centróide Y']].copy()
    add_table_without_space(alerta_info2, 2, [50, 50])
    elements.append(Spacer(1, 0.5 * cm))

    # Adiciona um salto de página
    elements.append(PageBreak())

    # Adiciona os mapas e informações do local

    # Informações da Fiscalização
    elements.append(Paragraph("Informações do local", styles['Titulo']))
    elements.append(Spacer(1, 1 * cm))

    # Cria o mapa Estado - cidade - alerta
 
    mapa_info['Longitude (centro)'] = mapa_info['Longitude (centro)'].astype(str).str.replace(',', '.', regex=True).astype(float)
    mapa_info['Latitude (centro)'] = mapa_info['Latitude (centro)'].astype(str).str.replace(',', '.', regex=True).astype(float)

    # 1. Obter valores da linha atual
    longitude = mapa_info.loc[mapa_info['ID da Cicatriz'] == current_idtxt, 'Longitude (centro)'].values[0]  
    latitude = mapa_info.loc[mapa_info['ID da Cicatriz'] == current_idtxt, 'Latitude (centro)'].values[0] 
    id_alerta = mapa_info.loc[mapa_info['ID da Cicatriz'] == current_idtxt, 'ID da Cicatriz'].values[0]

    # Usa cidade_final_str para criar o caminho e para a lógica
    mapa_localizacao = f"Output\\mapa\\mapa_{cidade_final_str.replace(' ', '_')}_{id_alerta}.png"
    
    if os.path.exists(mapa_localizacao):
            print(f"😎 🗺️  Mapa {cidade_final_str} (ID {id_alerta}).png já existe")
    else: 
        # Agora o .lower() é chamado em uma string
        if cidade_final_str.lower() in ['nan', 'none']:
            print(f"⚠️ 🗺️  AVISO: Pulando registro devido a nome de município ausente (ID {id_alerta}).")     
        try:
            # Passa a string correta para a função
            gerar_mapa_alerta(cidade_final_str, id_alerta, longitude, latitude)
        except Exception as e:
            print(f"❌ 🗺️  ERRO ao gerar mapa para {cidade_final_str} (ID {id_alerta}): {e}")
        

    # Imagem do Mapa Estado - cidade
    png_ilustrativo1 = f"Output\\mapa\\mapa_{cidade_final_str.replace(' ', '_')}_{id_alerta}.png"
    if os.path.exists(png_ilustrativo1):
        img = RLImage(png_ilustrativo1)
        img.drawHeight = 12 * cm
        img.drawWidth = 22 * cm
        elements.append(img)
        elements.append(Spacer(1.5, 0.5 * cm))


    # Verifica se a imagem existe e adiciona ao PDF - MAPA
    png_file = f"{idtxt}.png"
    png_path = os.path.join(png_folder, png_file)
    if os.path.exists(png_path):
        img = RLImage(png_path)  
        img.drawHeight = 9 *cm 
        img.drawWidth = 9 *cm 
        elements.append(img)
        elements.append(Spacer(1, 0.5 * cm))
    
    # Adiciona as imagens de antes, durante e depois
    elements.append(PageBreak())
    antes_file = next((f for f in os.listdir(pasta_AD) if f.endswith(f"{idtxt}.jpg")), None)
    antes_path = os.path.join(pasta_AD, antes_file) if antes_file else None
    elements.append(Paragraph("Antes, durante e depois:", styles['Titulo2']))
    # Verifica se a imagem existe e adiciona ao PDF - Antes e depois
    if antes_path and os.path.exists(antes_path):
        img_A = RLImage(antes_path)  
        img_A.drawHeight = 8 * inch
        img_A.drawWidth = 3 * inch
        elements.append(Spacer(1, 0.2 * inch))
        elements.append(img_A)
        elements.append(Spacer(0.5, 0.2 * inch))
    else:
        logging.warning(f"Imagens AD não encontradas para ID: {idtxt}")
        elements.append(Paragraph('Imagens de antes, durante e depois não disponíveis.', styles['CorpoTexto1']))

    # Adiciona a string 'link_imagem' como um parágrafo ao documento
    link_imagem = f"Link para imagens de antes, durante e depois: {link_imagemAD}"
    elements.append(Paragraph(link_imagem, styles['CorpoTexto']))
    elements.append(Spacer(1, 0.2 * inch))
    
    # Adiciona a string 'link_poligono' como um parágrafo ao documento
    link_poligono = f"Link para o polígono georeferenciado: {link_poligono}"
    elements.append(Paragraph(link_poligono, styles['CorpoTexto']))
    elements.append(Spacer(1, 0.2 * inch))

    #elements.append(PageBreak())
    if not inter_data.empty: # Adiciona as tabelas de interseções
        elements.append(Paragraph("Interseções da Cicatriz com outras camadas", styles['Titulo']))
        elements.append(Spacer(1, 0.2 * inch))
        add_tables_by_columns(inter_data, " ", 3, 1)
    else:
        logging.warning("Tabela de interseções vazia")

    # Adiciona a string 'conclusao' como um parágrafo ao documento
    conclusao = """A supressão de vegetação detectada no Alerta em epígrafe configura queimada e/ou incêndio florestal 
        realizado em desconformidade com a Lei Federal n° 12.651/2012, impondo a lavratura de Embargo Remoto Cautelar, 
        nos termos do art. 11 §3º, c/c art. 29 c/c art. 2º, inciso VII, ambos da Lei Estadual nº 3.467/2000 
        e Decreto Estadual nº 48.691 de 14 de setembro de 2023."""
    elements.append(Paragraph(conclusao, styles['CorpoTexto']))
    elements.append(Spacer(1, 0.2 * inch))
    
    # Adiciona a string 'observacao' como um parágrafo ao documento
    observacao = """Observação: Todas as informações estão disponíveis em acesso aberto para consulta a processos 
        no Portal do SEI (https://portalsei.rj.gov.br/)."""
    elements.append(Paragraph(observacao, styles['CorpoTexto']))
    elements.append(Spacer(1, 0.2 * inch))


    # Adiciona a tabela de referencia
    referencia_DF = pd.read_csv(referencias, delimiter=';', encoding='latin1')
    if not referencia_DF.empty:
        add_table_REF(referencia_DF, "Referências:", 2, [20, 80])
        elements.append(Spacer(1, 0.4 * cm))

    # Adiciona a obs2    
    obs_2 = "Observação: As camadas aqui listadas que não foram apresentadas no ítem de interseções do alerta com outras camadas não apresentam interseção espacial com o alerta em questão."
    elements.append(Paragraph(obs_2, styles['CorpoTexto']))

    # Rodapé
    if os.path.exists(imagem_cbc2):
        img_cbc2 = RLImage(imagem_cbc2)
        img_cbc2.drawHeight = 3.81 * cm
        img_cbc2.drawWidth = 15.7 * cm
        elements.append(img_cbc2)

    # Cria o PDF após a adição dos elementos
    doc.build(elements, 
              onFirstPage=add_page_num_wrapper, 
              onLaterPages=add_page_num_wrapper)
    logging.info(f"✅  {cntd} PDFs gerados, Alerta: {idtxt}\n")

# ----------------------------------------------------------------- Execução e tratamento dos arquivos -----------------------------------------------------------------

# Carrega o arquivo `cicatriz`
if cicatriz_file:
    cicatriz_df = pd.read_excel(cicatriz_file)
else:
    raise FileNotFoundError("Arquivo de cicatriz não encontrado.")

# Carrega todos os arquivos `inter` que começam com `tabela_Intersecao`
inter_files = [f for f in os.listdir(csv_folder) if f.startswith(inter_prefix) and f.endswith('.csv')]
inter_dfs = [pd.read_csv(os.path.join(csv_folder, f), sep=';', on_bad_lines='skip') for f in inter_files]

# Processamento de interseção
valid_inter_dfs = []
for df in inter_dfs:
    if id_column in df.columns:
        valid_inter_dfs.append(df)
    else:
        logging.error(f"Coluna {id_column} não encontrada em um dos arquivos de interseção.")

# Colunas que serão mantidas no DataFrame
manter_colunas =['idtxt', 'zona', 'label', 'cod_imovel', 'tp_app9', 'Tipo_8', 'Tipo_7', 'Tipo_6', 'tipo_5', 'Tipo_4', 'Tipo_3', 'tipo_ap,p2', 'Tipo1', 'nm_ucmp,i', 
            'nm_ucmu,s', 'nm_ucepi', 'nm_uceus', 'nm_ucfpi', 'nm_ucfus', 'nm_upam', 'superinten', 'municn', 'rotulo', 'grau', 'classe', 'vegetacao2024', 'area_inter', 'porc'] 


# Processar a tabela de nomes
nome_table = cicatriz_df.rename(columns={'municipio': 'Município'})[['idtxt', 'Município']]
nome_table = nome_table.drop_duplicates(subset=['idtxt'], keep='first')

try:
    setup_logging() # Chama a função que começa o logging
    # Lê o arquivo da camada em formato GeoDataFrame para gerar os mapas
    gdf = gpd.read_file(input_camada)
    
    # IDs que serão processados
    # Solicita ao usuário uma lista de IDs (separada por vírgulas) ou ENTER para processar todos.
    try:
        prompt = (
            "====================================================================================================\n"
            "============ 🔥🔥🔥 Bem-vindo ao gerador de laudos de Alerta - ONVC QUEIMADAS  🔥🔥🔥 ==============\n"
            "====================================================================================================\n"
            "====================================================================================================\n"
            "Digite IDs separados por vírgula para processar, ou pressione ENTER para processar todos\n"
            "(Exemplo: 202501002, 202401006, 2022010808)\n"
            "====================================================================================================\n"
            "Seus ids: "
        )

        user_input = input(prompt).strip()
    except KeyboardInterrupt:
        logging.info("Entrada de usuário cancelada. Encerrando.")
        raise SystemExit(0)

    all_ids = cicatriz_df[id_column].astype(str).tolist()

    if not user_input or user_input.strip().upper() == 'ALL':
        lista_ids = all_ids
        logging.info(f"\n🔎 Nenhum ID específico fornecido; processando todos os IDs.")
    else:
        requested_ids = [s.strip() for s in user_input.split(',') if s.strip()]
        missing_ids = [i for i in requested_ids if i not in all_ids]
        if missing_ids:
            logging.warning(f"⚠️ IDs não encontrados nos dados de alerta e serão ignorados: {missing_ids}")
        lista_ids = [i for i in requested_ids if i in all_ids]
        if not lista_ids:
            logging.error("Nenhum dos IDs fornecidos foi encontrado nos dados de alerta. Encerrando.")
            raise SystemExit(1)
        logging.info(f"\n🔎 Processando IDs fornecidos: {lista_ids}")



    for _, row in nome_table.iterrows():
        # Captura a variável necessária:
        current_idtxt = row['idtxt']
        idtxt = str(current_idtxt)
        if idtxt not in lista_ids: 
            logging.info(f"ID {idtxt} não está em lista_ids; pulando.")
            continue

        try: # Gerar mapa para o idtxt (usando geometria do shapefile)
            plot_map_for_idtxt_satellite(gdf, current_idtxt, png_folder) 
            print("⏳🗺️  Criando os mapas da cicatriz...")
        except Exception as e:
            logging.error(f"Erro ao gerar a imagem no ID {current_idtxt}: {str(e)}")

        # Filtrando os dados de cicatriz e interseção
        # ... (o bloco de filtragem da cicatriz e verificação de empty é mantido e está correto)
        cicatriz_df[id_column] = cicatriz_df[id_column].astype(str)
        id_form = str(idtxt)[:4]
        cicatriz_filtered = cicatriz_df[cicatriz_df[id_column].astype(str).str.strip() == str(idtxt).strip()] 
        if cicatriz_filtered.empty:
            logging.warning(f"⚠️ Alerta/ID {idtxt} encontrado na tabela de nomes, mas NÃO encontrado na tabela de cicatriz. Pulando geração de PDF.")
            continue # Volta para o próximo item do loop
            
        # Extração dos links (usando .iloc[0] para pegar a string única, não a lista de 1 elemento)
        link_AD = cicatriz_filtered['ant_dep']
        link_imagemAD = link_AD.replace(r'^\s*$', np.nan, regex=True)
        link_imagemAD = link_imagemAD.fillna('Link não disponível').astype(str).tolist()
        link_poligono = cicatriz_filtered['link_kml']
        link_poligono = link_poligono.replace(r'^\s*$', np.nan, regex=True)
        link_poligono = link_poligono.fillna('Link não disponível').astype(str).tolist()
        cic_colunas = ['idtxt', 'area_ha', 'area_m2', 'data_refer', 'data_ocorr', 'centro_x', 'centro_y', 'municipio']
        cicatriz_filtered = cicatriz_filtered.drop(columns=[column for column in cicatriz_filtered.columns if column not in cic_colunas], axis=1)
        columns = ['idtxt'] + [col for col in cicatriz_filtered.columns if col != 'idtxt'] # Reordenando para idtxt ficar em primeiro
        cicatriz_filtered = cicatriz_filtered[columns]
        cicatriz_filtered = cicatriz_filtered.rename(columns={'idtxt': 'ID da Cicatriz', 'area_ha': 'Área (ha)', 'area_m2': 'Área (m²)', 'data_refer': 'Data antes',
                                                            'data_ocorr': 'Data depois', 'centro_x': 'Centróide X', 'centro_y': 'Centróide Y', 'municipio': 'Município'}) # Renomeia colunas
        colunas_data = ['Data antes', 'Data depois']

        for col in colunas_data:
            if col in cicatriz_filtered.columns:
                cicatriz_filtered[col] = pd.to_datetime(cicatriz_filtered[col], unit='ms')
                
                cicatriz_filtered[col] = cicatriz_filtered[col].dt.strftime('%d/%m/%Y')

        cicatriz_filtered = cicatriz_filtered.reset_index(drop=True)

        # PARA OS DADOS DO INTERSECT:
        if valid_inter_dfs:
            # Filtra os dataframes de interseção e concatena apenas os que não estão vazios para evitar FutureWarning.
            filtered_parts = [df[df[id_column].astype(str) == str(idtxt)] for df in valid_inter_dfs]
            non_empty_parts = [part for part in filtered_parts if not part.empty]

            if non_empty_parts: # Concatena os DFs se eles não estiverem vazios
                inter_filtered = pd.concat(non_empty_parts, ignore_index=True)
            else: # se estiverem vazios, cria um DF vazio
                inter_filtered = pd.DataFrame()
            try:
                inter_filtered = inter_filtered.drop(columns=[column for column in inter_filtered.columns if column not in manter_colunas], axis=1) # Remove colunas
                inter_filtered = inter_filtered.rename(columns={'area_inter': 'Área da interseção (m²)', 'porc': 'Porcentagem (%)'}) # Renomeia colunas
            except Exception as e:
                logging.exception(f"Ocorreu um erro durante tratar intersect table: {e}")        
            try: # Geração do PDF
                    logging.info(f"⏳  Gerando layout para: ID {current_idtxt}")
                    create_pdf_for_idtxt(
                        idtxt, 
                        cicatriz_filtered, 
                        inter_filtered, 
                        png_folder, 
                        cntd, 
                        link_poligono,
                        link_imagemAD, # <-- Variável de link de imagem (Antes/Depois)
                        current_idtxt
                    )
            except Exception as e:
                logging.error(f"Ocorreu um erro ao gerar PDF para {idtxt}: {e}")
            finally:
                cntd += 1
except Exception as e:
    logging.exception(f"Ocorreu um erro: {e}")

def delete_files(directory, condition):
    # Apaga todas as imagens dos mapas para economizar memória
    files = os.listdir(directory)
    for file in files:
        file_path = os.path.join(directory, file)
        if os.path.isfile(file_path) and not file.startswith(condition):
            os.remove(file_path)
    logging.info(f"Arquivos excluídos: {directory}")
delete_files(png_folder, "APAGAR")


logging.info("✨✨✨  Fim do processamento  ✨✨✨")
