# -*- coding: utf-8 -*-
"""
Created on Thu Dec  19 10:01:32 2024

@Author: Anderson Silva
"""


import os
import pandas as pd
from pyxlsb import open_workbook

# Diretório base onde estão as pastas dos estados
base_dir = r'\\acswpj2\OP_PROJETOS_NOVA\APFJ\Apuração\Indiretos\ICMS IPI\Privado\2024\01 - Apuração'

# Diretório onde os arquivos filtrados serão salvos
output_dir = r'\\acswpj2\OP_PROJETOS_NOVA\OPERACOES\ADV\Condições Comerciais\Compartilhado\6. BI ONI\1. Apuração - Synchro\2024.v1'

# Lista com as siglas dos estados
estados = ["RJ"]
           #  ["AC", "AL", "AM", "BA", "CE", "DF", "ES", "GO", "MA", "MG", "MS", "MT",
             #"PA", "PB", "PE", "PI", "PR", "RJ", "RN", "RO", "RR", "RS", "SC", "SE", "SP", "TO"]
             
             #"PR"

# Extensões de arquivo a serem verificadas
file_extensions = ['xlsx', 'xls', 'csv', 'xlsb']

# Função para listar abas do arquivo .xlsb
def listar_abas_xlsb(file_path):
    """
    Função para listar abas do arquivo .xlsb

    Recebe o caminho do arquivo .xlsb e retorna uma lista com os nomes das abas
    existentes no arquivo.

    Parameters
    ----------
    file_path : str
        Caminho do arquivo .xlsb

    Returns
    -------
    list
        Lista com os nomes das abas do arquivo .xlsb
    """
    
    try:
        with open_workbook(file_path) as wb:
            # Corrigindo a maneira de acessar as abas
            sheet_names = []
            for sheet in wb.sheets:
                sheet_names.append(sheet)
        return sheet_names
    except Exception as e:
        print(f"Erro ao listar abas do arquivo .xlsb {file_path}: {e}")
        return []

# Função para ler arquivos .xlsb
# Recebe o caminho do arquivo e o nome da aba que se deseja ler
def ler_xlsb(file_path, sheet_name):
    """
    Função para ler arquivos .xlsb

    Recebe o caminho do arquivo .xlsb e o nome da aba que se deseja ler e
    retorna um DataFrame com os dados da aba.

    Parameters
    ----------
    file_path : str
        Caminho do arquivo .xlsb
    sheet_name : str
        Nome da aba que se deseja ler

    Returns
    -------
    pd.DataFrame or None
        Se o arquivo for lido com sucesso, retorna um DataFrame com os dados
        da aba. Caso contrário, retorna None.
    """

    try:
        with open_workbook(file_path) as wb:
            with wb.get_sheet(sheet_name) as sheet:
                data = []
                for row in sheet.rows():
                    data.append([item.v for item in row])
                df = pd.DataFrame(data[1:], columns=data[0])
        return df
    except Exception as e:
        print(f"Erro ao ler o arquivo .xlsb {file_path}: {e}")
        return None

# Função para buscar arquivos dentro de uma pasta de estado
# Recebe o diretório do estado e verifica as extensões desejadas
def buscar_arquivos(estado_dir, file_extensions):
    """
    Função para buscar arquivos dentro de uma pasta de estado

    Recebe o diretório do estado e verifica as extensões desejadas

    Parameters
    ----------
    estado_dir : str
        Diretório do estado
    file_extensions : list
        Lista com as extensões de arquivo a serem verificadas

    Returns
    -------
    list
        Lista com os caminhos dos arquivos encontrados
    """
    arquivos_encontrados = []
    
    # Loop pelos meses dentro do estado
    for mes in os.listdir(estado_dir):
        mes_path = os.path.join(estado_dir, mes)
        if os.path.isdir(mes_path):
            print(f"  Processando mês: {mes}")
            # Loop pelas unidades dentro do mês
            for unidade in os.listdir(mes_path):
                unidade_path = os.path.join(mes_path, unidade)
                if os.path.isdir(unidade_path):
                    print(f"    Processando unidade: {unidade}")
                    # Verificar se existe uma pasta de APURAÇÃO
                    apuracao_path = os.path.join(unidade_path, 'APURAÇÃO ICMS_ICMS ST_IPI')
                    if os.path.isdir(apuracao_path):
                        print(f"      Processando pasta de APURAÇÃO: {apuracao_path}")
                        # Verificar arquivos dentro da pasta de APURAÇÃO
                        for file in os.listdir(apuracao_path):
                            if file.startswith("19.") and file.split('.')[-1].lower() in file_extensions:
                                file_path = os.path.join(apuracao_path, file)
                                arquivos_encontrados.append((file_path, mes))
                                print(f"        Arquivo encontrado: {file_path}")
                    else:
                        # Se não existir pasta de APURAÇÃO, procurar diretamente na pasta da unidade
                        print(f"      Sem pasta de APURAÇÃO, verificando diretamente na unidade: {unidade_path}")
                        for file in os.listdir(unidade_path):
                            if file.startswith("19.") and file.split('.')[-1].lower() in file_extensions:
                                file_path = os.path.join(unidade_path, file)
                                arquivos_encontrados.append((file_path, mes))
                                print(f"        Arquivo encontrado: {file_path}")
    
    return arquivos_encontrados

# Função para processar os arquivos encontrados
# Carrega os dados da aba e salva em um novo arquivo Excel
def processar_arquivos(arquivos, output_dir):
    """
    Processa os arquivos encontrados e salva os dados da aba 'ENTR E SAIDAS MASTER' em um novo arquivo Excel
    no diretório de saída.

    Parameters
    ----------
    arquivos : list
        Lista com os caminhos dos arquivos encontrados e o mês correspondente
    output_dir : str
        Diretório de saída onde os arquivos serão salvos

    Returns
    -------
    None
    """
    for file_path, mes in arquivos:
        try:
            print(f"  Processando arquivo: {file_path}")
            
            # Diretório de saída correspondente ao mês
            output_mes_dir = os.path.join(output_dir, mes)
            os.makedirs(output_mes_dir, exist_ok=True)
            
            # Verificar as abas disponíveis
            abas = listar_abas_xlsb(file_path)
            print(f"    Abas encontradas no arquivo {file_path}: {abas}")

            # Verifica se a aba 'ENTR E SAIDAS MASTER' está disponível
            if 'ENTR E SAIDAS MASTER' in abas:
                # Carregar a planilha específica
                df = ler_xlsb(file_path, sheet_name='ENTR E SAIDAS MASTER')
            else:
                print(f"    Aba 'ENTR E SAIDAS MASTER' não encontrada no arquivo {file_path}.")
                continue

            if df is not None:
                # Caminho para salvar o novo arquivo
                output_file_path = os.path.join(output_mes_dir, os.path.basename(file_path).replace('.xlsb', '.xlsx'))

                # Salvar a planilha em um novo arquivo
                try:
                    df.to_excel(output_file_path, index=False, engine='openpyxl')
                    print(f"    Arquivo processado e salvo em: {output_file_path}")
                except Exception as e:
                    print(f"    Erro ao salvar o arquivo {output_file_path}: {e}")
            else:
                print(f"    Falha ao processar o arquivo {file_path}")

        except Exception as e:
            print(f"    Erro ao processar o arquivo {file_path}: {e}")

# Função principal para executar o processamento de todos os estados
def main():
    """
    Função principal para executar o processamento de todos os estados
    Presente na lista estados.

    Essa função percorre a lista de estados e para cada estado:
        1. Verifica se a pasta do estado existe;
        2. Busca arquivos com as extens es especificadas em file_extensions;
        3. Se encontrar arquivos, processa-os;
        4. Se n o encontrar arquivos, imprime uma mensagem de aviso.

    Parameters
    ----------
    None

    Returns
    -------
    None
    """
    for estado in estados:
        print(f"Processando estado: {estado}")
        estado_dir = os.path.join(base_dir, estado)
        if os.path.exists(estado_dir):
            arquivos = buscar_arquivos(estado_dir, file_extensions)
            if arquivos:
                print(f"Arquivos encontrados para o estado {estado}, iniciando processamento...")
                processar_arquivos(arquivos, output_dir)
            else:
                print(f"Nenhum arquivo encontrado para o estado {estado}.")
        else:
            print(f"Pasta do estado {estado} não encontrada.")
    print("Processo concluído.")

if __name__ == "__main__":
    main()
