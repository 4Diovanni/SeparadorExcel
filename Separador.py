# separador de casos por área -----------------------------------------------------

# -*- coding: utf-8 -*-
"""
Separador de casos por área
Automação desenvolvida por: Giovanni Moreira
Assistente financeiro da empresa Allcare Gestora de Saúde
Data: 09/05/2025
version: 1.0.2
Descrição: Este código tem como objetivo automatizar a separação de arquivos por área, garantindo que cada arquivo seja movido para a pasta correta, com o mesmo nome da respectiva área.
"""

import os
import shutil
def separador_pasta():
    """
    Função para separar arquivos em pastas específicas com base no sufixo do nome do arquivo.
    """
    # Mapeamento das áreas para nomes de pastas numeradas
    mapa_areas = {# Nome da separação do arquivo: Nome da pasta do arquivo em respectivo
        "FABRICA": "1 - FABRICA",
        "LOJA": "2 - LOJA",
        "JOALHERIA": "3 - JOALHERIA",
        # Em exemplo, arquivo: Produtos_stockA_FABRICA.xlsx, logo ele colocará na pasta "1 - FABRICA"
    }

    # Caminho da pasta de origem
    pasta_origem = r"Caminho/da/pasta/origem"

    for arquivo in os.listdir(pasta_origem):
        caminho_arquivo = os.path.join(pasta_origem, arquivo)
        if os.path.isfile(caminho_arquivo):
            nome_base = os.path.splitext(arquivo)[0]
            if "_" in nome_base: # captura a divisão do arquivo / a determinante do grupo pelo nome
                sufixo = nome_base.split("_")[-1].upper().strip()
                # Redireciona (Um mapeamento) para (Uma pasta especifica citada no mapeamento)
                if "nome_map" in sufixo or "nome_map2" in sufixo: # capturar variações
                    pasta_destino_nome = mapa_areas["nome_pasta_map"] # assim ele coleta o nome do arquivo e redireciona para a pasta "nome_pasta_map"
                else:
                    pasta_destino_nome = mapa_areas.get(sufixo, sufixo)
                pasta_destino = os.path.join(pasta_origem, pasta_destino_nome)
                if not os.path.exists(pasta_destino):
                    os.makedirs(pasta_destino)
                shutil.move(caminho_arquivo, os.path.join(pasta_destino, arquivo))
                print(f"Arquivo '{arquivo}' movido para '{pasta_destino_nome}'.")

    print("Processo concluído.")

separador_pasta()