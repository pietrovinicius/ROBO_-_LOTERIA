#06/03/2025
#@PLima

import random
import tkinter as tk
from tkinter import messagebox, ttk
from openpyxl import Workbook, load_workbook
import os
from datetime import datetime

# Constantes
ARQUIVO_EXCEL = "apostas.xlsx"
NUMERO_DE_NUMEROS_POR_APOSTA = 6
INTERVALO_NUMEROS = range(1, 61)
NUMERO_MAXIMO_TENTATIVAS_GERAR_APOSTA = 10
ATUALIZACAO_PLANILHA_INTERVALO = 5000  # Milissegundos (5 segundos)

def gerar_aposta() -> list[int] | None:
    """Gera uma aposta válida seguindo as regras."""
    for _ in range(NUMERO_MAXIMO_TENTATIVAS_GERAR_APOSTA):
        numeros = random.sample(INTERVALO_NUMEROS, NUMERO_DE_NUMEROS_POR_APOSTA)
        if (
            not tem_sequencia_consecutiva(numeros) and
            not muitos_multiplos_de_5(numeros) and
            not padrao_visual_obvio(numeros)
        ):
            return sorted(numeros)
    return None

def tem_sequencia_consecutiva(numeros: list[int]) -> bool:
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

def muitos_multiplos_de_5(numeros: list[int]) -> bool:
    """Evita que mais da metade dos números sejam múltiplos de 5."""
    multiplos = [n for n in numeros if n % 5 == 0]
    return len(multiplos) >= 3

def padrao_visual_obvio(numeros: list[int]) -> bool:
    """Evita padrões como todos números da mesma dezena."""
    dezenas = {n // 10 for n in numeros}
    return len(dezenas) <= 2

def salvar_aposta_excel(numeros: list[int]) -> None:
    """Salva a aposta na planilha Excel no modo append, incluindo a data."""
    try:
        data_atual = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

        if os.path.exists(ARQUIVO_EXCEL):
            wb = load_workbook(ARQUIVO_EXCEL)
            ws = wb.active
        else:
            wb = Workbook()
            ws = wb.active
            ws.append(["Data", "N1", "N2", "N3", "N4", "N5", "N6"])

        ws.append([data_atual] + numeros)
        wb.save(ARQUIVO_EXCEL)
    except Exception as e:
        print(f"Erro ao salvar a aposta no Excel: {e}")

def exibir_aposta() -> None:
    """Gera uma nova aposta, exibe na interface, salva no Excel e atualiza a janela da planilha."""
    aposta = gerar_aposta()
    if aposta:
        label_resultado.config(text=f"Sua aposta: {aposta}")
        salvar_aposta_excel(aposta)
        atualizar_total_sorteios()
        atualizar_janela_planilha()  # Atualiza a janela da planilha após gerar a aposta
    else:
        label_resultado.config(text="Não foi possível gerar uma aposta válida.")

def ao_fechar():
    """Confirma o fechamento do aplicativo."""
    resultado = messagebox.askyesno("Confirmação", "Tem certeza de que deseja fechar o aplicativo?")
    if resultado:
        print("Fechando aplicativo...")
        janela.destroy()

def contar_sorteios_excel() -> int:
    """Conta o número de sorteios na planilha Excel."""
    try:
        if os.path.exists(ARQUIVO_EXCEL):
            wb = load_workbook(ARQUIVO_EXCEL)
            ws = wb.active
            return ws.max_row - 1
        else:
            return 0
    except Exception as e:
        print(f"Erro ao contar os sorteios no Excel: {e}")
        return 0

def atualizar_total_sorteios() -> None:
    """Atualiza o label com o total de sorteios."""
    total_sorteios = contar_sorteios_excel()
    label_total_sorteios.config(text=f"Total de sorteios: {total_sorteios}")

def exibir_janela_planilha() -> None:
    """Cria e exibe a janela com a visualização da planilha."""
    global janela_planilha, tree
    janela_planilha = tk.Toplevel(janela)  # Janela filha da principal
    janela_planilha.title("Visualização da Planilha")
    janela_planilha.geometry("600x400")

    # Treeview para exibir a planilha
    tree = ttk.Treeview(janela_planilha, columns=("Data", "N1", "N2", "N3", "N4", "N5", "N6"), show="headings")

    # Definir a largura das colunas
    tree.column("Data", width=150)  # Largura da coluna "Data"
    tree.column("N1", width=50)    # Largura da coluna "N1"
    tree.column("N2", width=50)    # Largura da coluna "N2"
    tree.column("N3", width=50)    # Largura da coluna "N3"
    tree.column("N4", width=50)    # Largura da coluna "N4"
    tree.column("N5", width=50)    # Largura da coluna "N5"
    tree.column("N6", width=50)    # Largura da coluna "N6"

    tree.heading("Data", text="Data")
    tree.heading("N1", text="N1")
    tree.heading("N2", text="N2")
    tree.heading("N3", text="N3")
    tree.heading("N4", text="N4")
    tree.heading("N5", text="N5")
    tree.heading("N6", text="N6")
    tree.pack(expand=True, fill="both")

    # Agendar a atualização periódica da janela
    atualizar_janela_planilha()

def atualizar_janela_planilha() -> None:
    """Atualiza os dados exibidos na janela da planilha."""
    try:
        if os.path.exists(ARQUIVO_EXCEL):
            wb = load_workbook(ARQUIVO_EXCEL)
            ws = wb.active
            dados = list(ws.values)  # Obtém todos os dados da planilha

            # Limpa a Treeview antes de atualizar
            for item in tree.get_children():
                tree.delete(item)

            # Exibe os dados na Treeview em ordem decrescente (excluindo o cabeçalho)
            for row in reversed(dados[1:]):
                tree.insert("", "end", values=row)

        # Agendar a próxima atualização
        janela_planilha.after(ATUALIZACAO_PLANILHA_INTERVALO, atualizar_janela_planilha)
    except Exception as e:
        print(f"Erro ao atualizar a janela da planilha: {e}")

if __name__ == "__main__":
    print('\ninicio.py - __main__ - Inicio')

    # Criando a interface gráfica
    janela = tk.Tk()
    janela.title("Gerador de Apostas Mega-Sena")
    janela.geometry("500x250")
    janela.maxsize(500, 250)
    janela.protocol("WM_DELETE_WINDOW", ao_fechar)

    # Label para exibir o resultado
    label_resultado = tk.Label(janela, text="Clique no botão para gerar uma aposta", font=("Arial", 12))
    label_resultado.place(relx=0.5, y=75, anchor='center')

    # Botão para gerar aposta
    botao_gerar = tk.Button(janela, text="Gerar Aposta", command=exibir_aposta, font=("Arial", 12))
    botao_gerar.place(relx=0.5, y=175, anchor='center')

    # Label para exibir o total de sorteios (inicialmente vazio)
    label_total_sorteios = tk.Label(janela, text="", font=("Arial", 10))
    label_total_sorteios.place(relx=1.0, rely=1.0, anchor='se')

     # Botão para exibir a planilha
    botao_exibir_planilha = tk.Button(janela, text="Exibir Planilha", command=exibir_janela_planilha, font=("Arial", 12))
    botao_exibir_planilha.place(relx=0.5, y=220, anchor='center') # Ajuste a posição conforme necessário

    # Inicializa o total de sorteios ao iniciar o app
    atualizar_total_sorteios()

    # Iniciar a interface gráfica
    janela.mainloop()

    print('\ninicio.py - __main__ - Fim')