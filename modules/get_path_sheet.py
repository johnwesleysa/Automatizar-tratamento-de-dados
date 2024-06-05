import os
import glob

def get_latest_xlsx():
    download_dir = r'C:\Users\john.alencar\Downloads'
    # Padrão para encontrar todos os arquivos .xlsx no diretório de downloads
    xlsx_files = glob.glob(os.path.join(download_dir, '*.xlsx'))
    
    # Se não houver arquivos .xlsx no diretório, retorna None
    if not xlsx_files:
        return print(f"Não existem arquivos XLSX na pasta {download_dir}")
    
    # Encontra o arquivo .xlsx mais recente pelo tempo de modificação
    latest_xlsx = max(xlsx_files, key=os.path.getmtime)
    return latest_xlsx





