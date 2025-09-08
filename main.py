# 08/09/2025
# @PLima
# Refatorado por Gemini

import logging
import os
import random
import tkinter as tk
from collections import defaultdict
from datetime import datetime
from tkinter import messagebox

import pandas as pd
import ttkbootstrap as ttk
from ttkbootstrap.constants import *

# --- Constantes ---
ARQUIVO_EXCEL_MEGA_SENA = "Mega-Sena.xlsx"
ARQUIVO_EXCEL_APOSTAS = "apostas.xlsx"
ARQUIVO_LOG = "log.txt"
NUMERO_DE_NUMEROS_POR_APOSTA = 6
INTERVALO_NUMEROS = range(1, 61)
MAX_TENTATIVAS_VALIDACAO = 100

# --- Configuração do Logging ---
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s - %(levelname)s - %(message)s",
    handlers=[
        logging.FileHandler(ARQUIVO_LOG, mode='w'),
        logging.StreamHandler()
    ]
)

class LotteryLogic:
    """Lida com a lógica de negócio para geração de apostas."""

    def __init__(self, mega_sena_file: str):
        self.mega_sena_file = mega_sena_file
        self.frequencias = self._calcular_frequencia_numeros()

    def _calcular_frequencia_numeros(self) -> dict[int, int] | None:
        """Calcula a frequência dos números usando Pandas."""
        try:
            logging.info(f"Lendo o arquivo de dados: {self.mega_sena_file}")
            df = pd.read_excel(self.mega_sena_file)
            numeros_cols = df.columns[2:8]
            frequencias = defaultdict(int)
            for col in numeros_cols:
                for numero in df[col].dropna():
                    if 1 <= int(numero) <= 60:
                        frequencias[int(numero)] += 1
            logging.info(f"Frequências calculadas para {len(frequencias)} números.")
            return frequencias
        except FileNotFoundError:
            logging.error(f"Arquivo de dados não encontrado: {self.mega_sena_file}")
            messagebox.showerror("Erro", f"Arquivo de dados não encontrado: {self.mega_sena_file}")
            return None
        except Exception as e:
            logging.error(f"Erro ao calcular a frequência dos números: {e}")
            messagebox.showerror("Erro", f"Erro ao processar o arquivo de dados: {e}")
            return None

    def gerar_aposta_analisada(self) -> list[int] | None:
        """Gera uma aposta com análise e probabilidade ponderada."""
        if not self.frequencias:
            logging.warning("Não foi possível gerar aposta pois não há dados de frequência.")
            return None

        tentativas = 0
        while tentativas < MAX_TENTATIVAS_VALIDACAO:
            aposta_temp = self._gerar_aposta_ponderada()
            if aposta_temp and not self._validar_regras(aposta_temp):
                logging.info(f"Aposta gerada e validada: {aposta_temp}")
                return aposta_temp
            tentativas += 1
        
        logging.warning(f"Não foi possível gerar uma aposta válida após {MAX_TENTATIVAS_VALIDACAO} tentativas.")
        return None

    def _gerar_aposta_ponderada(self) -> list[int] | None:
        """Gera uma aposta com base na frequência ponderada dos números de forma eficiente."""
        numeros = list(self.frequencias.keys())
        pesos = list(self.frequencias.values())

        if len(numeros) < NUMERO_DE_NUMEROS_POR_APOSTA:
            logging.warning("Não há números únicos suficientes para gerar uma aposta.")
            return None

        aposta = set()
        while len(aposta) < NUMERO_DE_NUMEROS_POR_APOSTA:
            numero_escolhido = random.choices(numeros, weights=pesos, k=1)[0]
            aposta.add(numero_escolhido)
        
        return sorted(list(aposta))

    def _validar_regras(self, numeros: list[int]) -> bool:
        if self._tem_sequencia_consecutiva(numeros):
            logging.debug(f"Aposta {numeros} reprovada por sequência consecutiva.")
            return True
        if self._muitos_multiplos_de_5(numeros):
            logging.debug(f"Aposta {numeros} reprovada por múltiplos de 5.")
            return True
        if self._padrao_visual_obvio(numeros):
            logging.debug(f"Aposta {numeros} reprovada por padrão visual.")
            return True
        return False

    def _tem_sequencia_consecutiva(self, numeros: list[int]) -> bool:
        for i in range(len(numeros) - 3):
            if numeros[i+1] == numeros[i]+1 and numeros[i+2] == numeros[i]+2 and numeros[i+3] == numeros[i]+3:
                return True
        return False

    def _muitos_multiplos_de_5(self, numeros: list[int]) -> bool:
        return sum(1 for n in numeros if n % 5 == 0) >= 3

    def _padrao_visual_obvio(self, numeros: list[int]) -> bool:
        dezenas = {n // 10 for n in numeros}
        return len(dezenas) <= 2

    @staticmethod
    def salvar_aposta_excel(numeros: list[int]) -> None:
        """Salva a aposta na planilha "apostas.xlsx"."""
        try:
            data_atual = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            nova_linha = [data_atual] + numeros
            
            if not os.path.exists(ARQUIVO_EXCEL_APOSTAS):
                df = pd.DataFrame(columns=["Data", "N1", "N2", "N3", "N4", "N5", "N6"])
                df.loc[0] = nova_linha
            else:
                df = pd.read_excel(ARQUIVO_EXCEL_APOSTAS)
                df.loc[len(df)] = nova_linha

            df.to_excel(ARQUIVO_EXCEL_APOSTAS, index=False)
            logging.info(f"Aposta {numeros} salva em {ARQUIVO_EXCEL_APOSTAS}")

        except Exception as e:
            logging.error(f"Erro ao salvar a aposta no Excel: {e}")
            messagebox.showerror("Erro de Gravação", f"Não foi possível salvar a aposta: {e}")


class App(ttk.Window):
    """Classe principal da aplicação GUI com ttkbootstrap."""

    def __init__(self, logic: LotteryLogic, themename: str = "superhero"):
        super().__init__(themename=themename)
        self.logic = logic
        
        logging.info("Iniciando a aplicação com interface moderna.")

        self.title("Gerador de Apostas Mega-Sena")
        self.geometry("700x600")
        self.protocol("WM_DELETE_WINDOW", self._ao_fechar)

        self._configurar_layout()
        self._configurar_widgets()
        self.atualizar_janela_planilha()

    def _configurar_layout(self):
        """Configura os frames principais usando o layout grid."""
        self.grid_rowconfigure(1, weight=1)
        self.grid_columnconfigure(0, weight=1)

        self.control_frame = ttk.Frame(self, padding=(20, 10))
        self.control_frame.grid(row=0, column=0, sticky="ew")
        self.control_frame.grid_columnconfigure(0, weight=1)

        self.history_frame = ttk.Frame(self, padding=(20, 10))
        self.history_frame.grid(row=1, column=0, sticky="nsew")
        self.history_frame.grid_rowconfigure(0, weight=1)
        self.history_frame.grid_columnconfigure(0, weight=1)

    def _configurar_widgets(self):
        """Cria e posiciona os widgets na janela."""
        # --- Frame de Controle ---
        title_label = ttk.Label(self.control_frame, text="Gerador de Apostas", font=("-size 20 -weight bold"))
        title_label.grid(row=0, column=0, pady=(0, 10))

        self.balls_frame = ttk.Frame(self.control_frame)
        self.balls_frame.grid(row=1, column=0, pady=(10, 20))

        self.ball_labels = []
        for i in range(NUMERO_DE_NUMEROS_POR_APOSTA):
            label = ttk.Label(self.balls_frame, text="--", font=("-size 16 -weight bold"), anchor=CENTER, bootstyle=(INVERSE, PRIMARY), padding=10, width=3)
            label.grid(row=0, column=i, padx=5)
            self.ball_labels.append(label)

        self.botao_gerar = ttk.Button(self.control_frame, text="Gerar Nova Aposta", command=self.exibir_aposta, bootstyle=SUCCESS, padding=10)
        self.botao_gerar.grid(row=2, column=0, pady=(0, 10))

        # --- Frame de Histórico ---
        history_title = ttk.Label(self.history_frame, text="Histórico de Apostas", font=("-size 14"))
        history_title.grid(row=0, column=0, sticky="w", pady=(0, 5))
        
        cols = ("Data", "N1", "N2", "N3", "N4", "N5", "N6")
        self.tree = ttk.Treeview(self.history_frame, columns=cols, show="headings", bootstyle=PRIMARY)
        
        for col in cols:
            self.tree.heading(col, text=col)
            self.tree.column(col, width=80, anchor=CENTER)
        self.tree.column("Data", width=160)

        # Adiciona scrollbars
        vsb = ttk.Scrollbar(self.history_frame, orient="vertical", command=self.tree.yview)
        hsb = ttk.Scrollbar(self.history_frame, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)

        self.tree.grid(row=1, column=0, sticky="nsew")
        vsb.grid(row=1, column=1, sticky="ns")
        hsb.grid(row=2, column=0, sticky="ew")

        self.history_frame.grid_rowconfigure(1, weight=1)
        self.history_frame.grid_columnconfigure(0, weight=1)
        
        # Estilo para linhas alternadas
        self.tree.tag_configure('oddrow', background=self.style.colors.get('light'))
        self.tree.tag_configure('evenrow', background=self.style.colors.get('dark'))


    def exibir_aposta(self):
        """Gera, valida, exibe e salva a aposta."""
        logging.info("Botão 'Gerar Aposta' clicado.")
        self.botao_gerar.config(state="disabled")
        self.update_idletasks() # Força a atualização da UI

        aposta = self.logic.gerar_aposta_analisada()

        if aposta:
            for i, numero in enumerate(aposta):
                self.ball_labels[i].config(text=f"{numero:02}")
            self.logic.salvar_aposta_excel(aposta)
            self.atualizar_janela_planilha(highlight_new=True)
        else:
            msg = "Não foi possível gerar uma aposta válida."
            messagebox.showwarning("Aviso", f"{msg} Verifique o log para mais detalhes.")
        
        self.botao_gerar.config(state="normal")

    def atualizar_janela_planilha(self, highlight_new=False):
        """Atualiza os dados exibidos na Treeview."""
        try:
            # Limpa a treeview
            for item in self.tree.get_children():
                self.tree.delete(item)

            if os.path.exists(ARQUIVO_EXCEL_APOSTAS):
                df = pd.read_excel(ARQUIVO_EXCEL_APOSTAS)
                
                # Insere os novos dados em ordem reversa (mais recente primeiro)
                for i, row in df.iloc[::-1].iterrows():
                    tag = 'evenrow' if i % 2 == 0 else 'oddrow'

                    # --- Lógica de Formatação de Data ---
                    try:
                        date_obj = datetime.strptime(str(row['Data']), "%Y-%m-%d %H:%M:%S")
                        formatted_date = date_obj.strftime("%H:%M %d/%m/%y")
                    except (ValueError, TypeError):
                        formatted_date = row['Data'] # Usa o valor original em caso de erro
                    
                    display_values = [formatted_date] + list(row[1:])
                    # --- Fim da Lógica ---

                    if highlight_new and i == len(df) - 1:
                        tag = 'success' # Estilo do ttkbootstrap para destaque
                    self.tree.insert("", "end", values=display_values, tags=(tag,))
        except Exception as e:
            logging.error(f"Erro ao atualizar a janela da planilha: {e}")

    def _ao_fechar(self):
        """Confirma o fechamento do aplicativo."""
        if messagebox.askyesno("Confirmação", "Tem certeza de que deseja fechar o aplicativo?"):
            logging.info("Aplicação fechada pelo usuário.")
            self.destroy()

def main():
    """Função principal para iniciar a aplicação."""
    try:
        logic = LotteryLogic(ARQUIVO_EXCEL_MEGA_SENA)
        if logic.frequencias is not None:
            app = App(logic)
            app.mainloop()
        else:
            logging.critical("A aplicação não pode iniciar pois os dados de frequência não foram carregados.")
    except Exception as e:
        logging.critical(f"Ocorreu um erro fatal na aplicação: {e}")
        messagebox.showerror("Erro Fatal", f"A aplicação encontrou um erro e precisa fechar: {e}")

if __name__ == "__main__":
    main()