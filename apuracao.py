import os
import shutil
import pandas as pd

# Diretório principal onde estão as pastas dos estados
base_dir = r'\\acswpj2\OP_PROJETOS_NOVA\APFJ\Apuração\Indiretos\ICMS IPI\Privado\2024\01 - Apuração'

# Diretório onde os arquivos filtrados serão salvos
output_dir = r'\\acswpj2\OP_PROJETOS_NOVA\OPERACOES\ADV\Condições Comerciais\Compartilhado\6. BI ONI\1. Apuração - Synchro\2024.v1'

# Extensões de arquivo a serem verificadas
file_extensions = ['xlsx', 'xls', 'csv', 'xlsb']

# Loop pelos estados
for estado in os.listdir(base_dir):
    estado_path = os.path.join(base_dir, estado)
    if os.path.isdir(estado_path):
        # Loop pelos meses
        for mes in os.listdir(estado_path):
            mes_path = os.path.join(estado_path, mes)
            if os.path.isdir(mes_path):
                # Verificar arquivos dentro da pasta do mês
                for file in os.listdir(mes_path):
                    if file.startswith("19.") and file.split('.')[-1] in file_extensions:
                        file_path = os.path.join(mes_path, file)
                        try:
                            # Carregar a planilha específica
                            if file.endswith('.csv'):
                                df = pd.read_csv(file_path, sheet_name='ENTR E SAIDAS MASTER')
                            else:
                                df = pd.read_excel(file_path, sheet_name='ENTR E SAIDAS MASTER')

                            # Diretório de saída correspondente ao mês
                            output_mes_dir = os.path.join(output_dir, mes)
                            os.makedirs(output_mes_dir, exist_ok=True)

                            # Caminho para salvar o novo arquivo
                            output_file_path = os.path.join(output_mes_dir, file)

                            # Salvar a sheet "ENTR E SAIDAS MASTER" em um novo arquivo
                            if file.endswith('.csv'):
                                df.to_csv(output_file_path, index=False)
                            else:
                                df.to_excel(output_file_path, index=False, engine='openpyxl')

                            print(f"Arquivo processado e salvo em: {output_file_path}")

                        except Exception as e:
                            print(f"Erro ao processar o arquivo {file_path}: {e}")

print("Processo concluído.")
