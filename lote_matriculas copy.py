import re
import openpyxl
from modules.organizar import linha, cls
from modules.tratar_data_set import tratar_data_set, gerar_arquivo
from datetime import datetime



def main():
    #cidades = [
    #    ["Recife", "Recife"], 
    #    ["Brasília", "Brasília"], 
    #    ["São Paulo", "São Paulo"], 
    #    ["Barra", "Rio de Janeiro"], 
    #    ["Botafogo", "Rio de Janeiro"], 
    #    ["Urca", "Rio de Janeiro"]
    #]
    
    cls()
    """while True:
        linha()
        try:  
            #print("Escolha a cidade:\n0 - Recife\n1 - Brasília\n2 - São Paulo\n3 - Barra\n4 - Botafogo\n5 - Urca\n")
            #cidade = int(input("Digite o número correspondente à cidade: "))

            #if cidade < 0 or cidade >= len(cidades):
            #    raise ValueError("Cidade inválida. Escolha um número de 0 a 5.")

            #data = input("Data (dd-mm-aaaa): ")
            #if not re.match(r'^\d{2}-\d{2}-\d{4}$', data):
            #    raise ValueError("Data inválida. Use o formato dd-mm-aaaa.")
            
            # Se chegou até aqui, significa que o usuário digitou valores válidos
            break
            

        
        except ValueError as e:
            cls()
            print(f"Erro: {e}")
            continue"""
        
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
