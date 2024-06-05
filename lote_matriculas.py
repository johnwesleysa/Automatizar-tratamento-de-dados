import re
import openpyxl
from modules.organizar import linha, cls
from modules.tratar_data_set import tratar_data_set, gerar_arquivo
from datetime import datetime



def main():
    cls()
    #PEGA A DATA ATUAL
    # Obtém a data atual do sistema no formato dd-mm-aaaa
    linha()
    
    data = datetime.now().strftime("%d-%m-%Y")
    print("Data atual:", data)

    linha()

    retorno_tratar = tratar_data_set()
    print(retorno_tratar[1])

    while True:
        try:
            gerar = int(input("Deseja gerar os arquivos? 1 - Sim | 2 - Não: "))
            if gerar not in [1, 2]:
                raise ValueError("Escolha inválida. Digite 1 para Sim ou 2 para Não.")
            break
        
        except ValueError as e:
            print(f"Erro: {e}")
            continue
    
    if gerar == 1:
        gerar_arquivo(data, retorno_tratar[0])
        print("Arquivos Gerados.")
    else:
        print("Saindo...")

if __name__ == "__main__":
    main()
