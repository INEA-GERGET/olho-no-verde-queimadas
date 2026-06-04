# Gerar Laudos - De Olho No Verde Queimadas 

## Sumário
1. [Descrição](#Descrição)
2. [Uso](#Uso)
   - [Sobre os arquivos](#Sobre-os-arquivos)
4. [Instalação](#Instalação)


## Descrição
Este repositório contém todos os arquivos necessárrios para gerar Relatórios Laudo do Alerta e Laudo de Emabrgo de forma automática.

## Uso
Esse script possui uma interface interativa no terminal, onde o usuário pode escolher qual tipo de laudo quer gerar, se vai gerar ou não a tabela de merge e selecionar os IDs dos alertas a serem processados.
O primeiro passo ao executar o [main.py](main.py) deverá ser essa tela:
```
====================================================================================================
====================         GERAÇÃO DOS LAUDOS - ONVC QUEIMADAS 🔥🔥🔥          ====================
====================================================================================================

Para gerar os laudos, certifique-se que os arquivos necessários estejam na pasta correta:
 na pasta 'CSVs' ->
                    'tabela_car.xlsx' (APENAS PARA EMBARGOS),
                    'alertas_em_uso.xlsx',
                    'todas as tabelas dissolves'
 na pasta 'camada' ->
                    'alertas_em_uso.shp'

 na pasta 'antes_depois' ->
                    imagens de antes e depois nomeadas com o id

Esteja conectado ao COGET (APENAS PARA EMBARGOS) ->
                    'df_id_embargo_ONV.xlsx' é atualizado automaticamente no COGET

====================================================================================================
Deseja gerar laudos de alerta ou laudos de embargo? (alerta/embargo):
```

Se você escolher a opção "alerta" para gerar os laudos de alerta, a próxima tela irá pedir a entrada da lista de IDs ou pressionar o ENTER para gerar os laudos para todos os IDs disponíveis.

```
====================================================================================================
============ 🔥🔥🔥 Bem-vindo ao gerador de laudos de Alerta - ONVC QUEIMADAS  🔥🔥🔥 ==============
====================================================================================================

Digite IDs separados por vírgula para processar, ou pressione ENTER para processar todos
(Exemplo: 202501002, 202401006, 2022010808)
====================================================================================================
Seus ids:
```

Se você selecionar a opção "embargo" para gerar os laudos de embargo, na primeira tela no terminal, a próxima tela irá perguntar se você gostaria de gerar a tabela merge, necessária para a gerção de embargos.
```
====================================================================================================
====================               GERAÇÃO DA TABELA DE ONVC_CAR                ====================
====================================================================================================

Para gerar a tabela usada nos laudos de embargo, certifique-se que os arquivos necessários estejam
 na pasta 'CSVs' ->
                    'tabela_car.xlsx'
                    'cicatrizes_em_uso.xlsx' e
                    'tabela_Dissolve_nova_camada_inter_GPL_08_Codigo_CAR.csv'

====================================================================================================
Deseja gerar a tabela de merge? (s/n):
```

Independente da sua decisão sobre a geração da tabela merge, a próxima tela mostrar a atualização dos IDs de Laudo e irá pedir como entrada os IDs a serem processados. 

```
====================================================================================================
============ 🔥🔥🔥 Bem-vindo ao gerador de laudos de Embargo - ONVC QUEIMADAS  🔥🔥🔥 =============
====================================================================================================
====================================================================================================
Digite IDs separados por vírgula para processar, ou pressione ENTER para processar todos
(Exemplo: 202501002, 202401006, 2022010808)
====================================================================================================
Seus ids:
```
### Sobre os arquivos 
Baixe todos os arquivos e pastas e salve em um único diretório. 

Alguns arquivos necessitam de mudança no caminho para os documentos, atente-se à isso. 

Sobre os arquivos:
1. **Input**: Nesta pasta todos os arquivos mutáveis do programa estarão aqui. São os inputs necessários para gerar os laudos. Nela deverão conter os seguintes arquivos:
   1. **CSVs**: Esta pasta deverá armazenar a `tabela_car.xlsx` (para embargos), `cicatrizes_em_uso.xlsx` e todas as tabelas provenientes do Dissolve. 
   2. **antes_depois**: Nesta pasta você irá colocar as imagens de antes e depois em png pu jpg. O nome da imagem deve ser o nome do ID do Alerta.
   3. **camada**: Esta pasta deverá conter arquivos em formato GeoDataFrame, `cicatrizes_em_uso.shp` para gerar mapas no Laudo.
2. **Output**: Nesta pasta estarão os resultados do script: a planilha merge, os relatórios de alerta e de embargo e duas pastas de mapas, que no final do script deverão ficar vazias.
3. **arquivos**: Nesta pasta estão os arquivos de imagem necessários para o layout do documento, tais como: papel timbrado, logos, cabeçalho, rodapé e referências.
4. **config**: Nesta pasta está o acesso para o Portal GEOINEA.
5. **logs**: Pasta com o aquivo de execução.


[**ONVC_intersect_dissolve.txt**](ONV_intersect_dissolve.txt): Este arquivo de texto é para ser colocado no terminal Python dentro do ArcGIS, onde irá juntar as camadas e fazer os cálculos das áreas de interseção. É o código mais demorado para rodar, pois os arquivos iniciais são mais pesados. **Há a necessidade de mudar o caminho do diretório para os arquivos de input e os outputs**, mas mantenha esses arquivos dentro da pasta com todos os arquivos.

    ``` python
    # Caminhos 
    feature_class_principal = r"C:\Users\Nome-de-usuario\Desktop\Laudos_ONV\camada\operações_diss.shp"  # TROCAR
    gdb_directory = r"C:\Users\Nome-de-usuario\Desktop\Laudos_ONV\Cruzamentos_ONV_.gdb"  # TROCAR

    output_intersect_dir = r"C:\Users\Nome-de-usuario\Desktop\Laudos_ONV\Output_intersect.gdb"  # TROCAR
    output_dissolve_dir = r"C:\Users\Nome-de-usuario\Desktop\Laudos_ONV\Output_dissolve.gdb"  # TROCAR
    ```
    
* [**ONVC_table_to_excel.txt**](ONV_table_to_excel.txt): Este arquivo de texto é para ser colocado no terminal Python dentro do ArcGIS, onde irá juntar as camadas e fazer os cálculos das áreas de interseção. É o código mais demorado para rodar, pois os arquivos iniciais são mais pesados. **Há a necessidade de mudar o caminho do diretório para os arquivos de input e os outputs**, mas mantenha esses arquivos dentro da pasta com todos os arquivos. Ele vai gerar os CSVs das camadas. 

    ``` python
    # Caminho 
    gdb_directory = r"C:\Users\Nome-de-usuario\Desktop\Laudos_ONV\Output_dissolve.gdb"  # TROCAR
    arcpy.env.workspace = gdb_directory
    gdb_layers = arcpy.ListFeatureClasses()

    Saida = r"C:\Users\Nome-de-usuario\Desktop\Laudos_ONV\CSVs"  # TROCAR
    ```
    
* [**main.py**](main.py): Gera os Relatórios de Laudos de Alerta e Embargo e a Planinha de merge. 
## Instalação
Os arquivos .txt são para serem utilizados dentro do [ArcGIS](https://www.arcgis.com/index.html) e os .py em um compilador Python, como explicado no [mapa mental](#Uso).
É recomendado criar um ambiente virtual a partir do ArcGIS no seu compilador para fazer as intealações das bibliotecas necessárias.

Para utilizar os scripts é necessária a instalação das bibliotecas Python:

* **configparser**: Para ler arquivos de configuração.
* **os**: Para interagir com o sistema operacional, como manipular caminhos de arquivos.
* **logging**: Para registrar eventos e mensagens do sistema.
* **arcgis**: Para interagir com o ArcGIS.
* **pandas**: Para manipulação e análise de dados tabulares.
* **geopandas**: Para trabalhar com dados geoespaciais (camadas de feições).
* **matplotlib.pyplot**: Para criar gráficos e visualizações.
* **contextily**: Para adicionar mapas base (mapas de fundo) a gráficos criados com Matplotlib.
* **PIL (Pillow)**: Para manipulação de imagens. A biblioteca é importada como `Image` e `ImageDraw`.
* **reportlab**: Para gerar documentos PDF de forma programática.
* **warnings**: Para controlar avisos (warnings).
* **arcpy**: Para a automação de tarefas do ArcGIS Pro e ArcMap.
