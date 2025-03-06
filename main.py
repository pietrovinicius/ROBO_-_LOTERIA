#06/03/2025
#@PLima

import random
import tkinter as tk
from openpyxl import Workbook, load_workbook
import os

# Nome do arquivo da planilha
arquivo_excel = "apostas.xlsx"

def gerar_aposta():
    """Gera uma aposta válida seguindo as regras."""
    while True:
        numeros = random.sample(range(1, 61), 6)
        if (
            not tem_sequencia_consecutiva(numeros) and
            not muitos_multiplos_de_5(numeros) and
            not padrao_visual_obvio(numeros)
        ):
            return sorted(numeros)

def tem_sequencia_consecutiva(numeros):
    """Verifica se há 4 ou mais números consecutivos na aposta."""
    numeros.sort()
    contador = 1
    for i in range(len(numeros) - 1):
        if numeros[i] + 1 == numeros[i + 1]:
            contador += 1
            if contador >= 4:
                return True
        else:
            contador = 1
    return False

def muitos_multiplos_de_5(numeros):
    """Evita que mais da metade dos números sejam múltiplos de 5."""
    multiplos = [n for n in numeros if n % 5 == 0]
    return len(multiplos) >= 3  # Se houver 3 ou mais múltiplos de 5, rejeita

def padrao_visual_obvio(numeros):
    """Evita padrões como todos números da mesma dezena."""
    dezenas = [n // 10 for n in numeros]
    return len(set(dezenas)) <= 2  # Se forem apenas 2 dezenas diferentes, rejeita

def salvar_aposta_excel(numeros):
    """Salva a aposta na planilha Excel no modo append."""
    if os.path.exists(arquivo_excel):
        # Se o arquivo já existe, abre e carrega a planilha
        wb = load_workbook(arquivo_excel)
        ws = wb.active
    else:
        # Se o arquivo não existe, cria uma nova planilha
        wb = Workbook()
        ws = wb.active
        ws.append(["N1", "N2", "N3", "N4", "N5", "N6"])  # Cabeçalho

    # Adiciona a nova aposta
    ws.append(numeros)
    wb.save(arquivo_excel)

def exibir_aposta():
    """Gera uma nova aposta, exibe na interface e salva no Excel."""
    aposta = gerar_aposta()
    label_resultado.config(text=f"Sua aposta: {aposta}")
    salvar_aposta_excel(aposta)


if __name__ == "__main__":
    print('\n\ninicio.py - __main__ - Inicio')

    # Criando a interface gráfica
    janela = tk.Tk()
    janela.title("Gerador de Aposta Mega-Sena")
    janela.geometry("300x200")

    # Botão para gerar aposta
    botao_gerar = tk.Button(janela, text="Gerar Aposta", command=exibir_aposta, font=("Arial", 12))
    botao_gerar.pack(pady=20)

    # Label para exibir o resultado
    label_resultado = tk.Label(janela, text="Clique no botão para gerar uma aposta", font=("Arial", 12))
    label_resultado.pack()

    # Iniciar a interface gráfica
    janela.mainloop()

    print('\n\ninicio.py - __main__ - Fim')
