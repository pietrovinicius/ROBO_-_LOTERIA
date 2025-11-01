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
        self.pares_coocorrencia = self._calcular_coocorrencia_pares()
        self.trincas_coocorrencia = self._calcular_coocorrencia_trincas()
        # Frequências suavizadas (Bayes) para uso auxiliar
        self.freq_suavizadas = self._calcular_frequencias_suavizadas(alpha=80)

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

    def _calcular_coocorrencia_pares(self) -> dict[tuple[int, int], int] | None:
        """Calcula a co-ocorrência de pares de números a partir da base histórica."""
        try:
            logging.info("Calculando co-ocorrência de pares de números.")
            df = pd.read_excel(self.mega_sena_file)
            numeros_cols = df.columns[2:8]
            cooc = defaultdict(int)
            for _, row in df[numeros_cols].dropna().iterrows():
                try:
                    nums = [int(x) for x in row.values if 1 <= int(x) <= 60]
                except Exception:
                    nums = []
                if len(nums) == 6:
                    nums_sorted = sorted(set(nums))
                    # Atualiza contagem para todos os pares únicos
                    for i in range(len(nums_sorted)):
                        for j in range(i + 1, len(nums_sorted)):
                            a, b = nums_sorted[i], nums_sorted[j]
                            cooc[(a, b)] += 1
            logging.info(f"Co-ocorrência calculada para {len(cooc)} pares.")
            return cooc
        except FileNotFoundError:
            logging.error(f"Arquivo de dados não encontrado: {self.mega_sena_file}")
            return None
        except Exception as e:
            logging.error(f"Erro ao calcular co-ocorrência de pares: {e}")
            return None

    def _calcular_coocorrencia_trincas(self) -> dict[tuple[int, int, int], int] | None:
        """Calcula a co-ocorrência de trincas de números a partir da base histórica."""
        try:
            logging.info("Calculando co-ocorrência de trincas de números.")
            df = pd.read_excel(self.mega_sena_file)
            numeros_cols = df.columns[2:8]
            cooc3 = defaultdict(int)
            for _, row in df[numeros_cols].dropna().iterrows():
                try:
                    nums = [int(x) for x in row.values if 1 <= int(x) <= 60]
                except Exception:
                    nums = []
                if len(nums) == 6:
                    nums_sorted = sorted(set(nums))
                    if len(nums_sorted) == 6:
                        # 20 trincas por bilhete
                        for i in range(0, 4):
                            for j in range(i + 1, 5):
                                for k in range(j + 1, 6):
                                    a, b, c = nums_sorted[i], nums_sorted[j], nums_sorted[k]
                                    cooc3[(a, b, c)] += 1
            logging.info(f"Co-ocorrência calculada para {len(cooc3)} trincas.")
            return cooc3
        except FileNotFoundError:
            logging.error(f"Arquivo de dados não encontrado: {self.mega_sena_file}")
            return None
        except Exception as e:
            logging.error(f"Erro ao calcular co-ocorrência de trincas: {e}")
            return None

    def _calcular_frequencias_suavizadas(self, alpha: int = 80) -> dict[int, float] | None:
        """Aplica suavização Bayesiana às frequências individuais.
        Retorna probabilidade normalizada aproximada por número.
        """
        if not self.frequencias:
            return None
        total_obs = sum(self.frequencias.values())
        if total_obs == 0:
            return {n: 1.0 / 60.0 for n in range(1, 61)}
        # prior simétrico: alpha para cada número, total 60*alpha
        suav = {}
        prior_total = 60 * alpha
        for n in range(1, 61):
            f = self.frequencias.get(n, 0)
            suav[n] = (f + alpha) / (total_obs + prior_total)
        return suav

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

    # --- Estratégia 1: Aleatória uniforme com equilíbrio ---
    def gerar_aposta_estrategia1(self) -> list[int] | None:
        """Gera aposta por amostragem uniforme aplicando restrições leves de equilíbrio.
        Critérios:
        - 2 a 4 números pares (preferência por 3 pares/3 ímpares)
        - Cobertura de ao menos 4 décadas distintas
        - Soma entre 150 e 210
        - Evita regras já existentes (sequências de 4+, muitos múltiplos de 5)
        """
        tentativas = 0
        while tentativas < MAX_TENTATIVAS_VALIDACAO:
            aposta = sorted(random.sample(list(INTERVALO_NUMEROS), NUMERO_DE_NUMEROS_POR_APOSTA))
            if not self._validar_regras_equilibrio(aposta):
                return aposta
            tentativas += 1
        return None

    # --- Estratégia 2: Maximiza raridade de pares (cobertura) ---
    def gerar_aposta_estrategia2(self) -> list[int] | None:
        """Gera aposta priorizando pares historicamente raros (baixa co-ocorrência),
        mantendo restrições de equilíbrio semelhantes à estratégia 1.
        """
        if not self.pares_coocorrencia:
            logging.warning("Co-ocorrência de pares indisponível; utilizando estratégia 1 como fallback.")
            return self.gerar_aposta_estrategia1()

        # Pesos de raridade dos pares: w = 1/(freq+1)
        wpar = defaultdict(lambda: 1.0)
        for (a, b), c in self.pares_coocorrencia.items():
            wpar[(a, b)] = 1.0 / (c + 1)

        # Peso auxiliar por número baseado em suavização (quanto menor prob., maior incentivo)
        invfreq = defaultdict(lambda: 1.0)
        if self.freq_suavizadas:
            for n, p in self.freq_suavizadas.items():
                invfreq[n] = 1.0 / (p + 1e-6)

        melhor_aposta = None
        melhor_score = -1.0
        for _ in range(50):  # múltiplas tentativas para escapar de ótimos locais
            escolhidos = []
            candidatos = set(INTERVALO_NUMEROS)
            # semente aleatória
            seed = random.choice(list(candidatos))
            escolhidos.append(seed)
            candidatos.remove(seed)

            while len(escolhidos) < NUMERO_DE_NUMEROS_POR_APOSTA:
                melhor_c = None
                melhor_s = -1.0
                for c in candidatos:
                    # soma de raridade com os já escolhidos
                    s = 0.0
                    for e in escolhidos:
                        a, b = (e, c) if e < c else (c, e)
                        s += wpar[(a, b)]
                    # contribuição de trincas raras, aproximada: combine c com dois dos escolhidos
                    if self.trincas_coocorrencia:
                        le = len(escolhidos)
                        if le >= 2:
                            for i in range(le):
                                for j in range(i + 1, le):
                                    tri = tuple(sorted([escolhidos[i], escolhidos[j], c]))
                                    freq_tri = self.trincas_coocorrencia.get(tri, 0)
                                    s += 0.5 * (1.0 / (freq_tri + 1))
                    # pequeno incentivo a números com menor frequência individual
                    s += 0.1 * invfreq[c]
                    if s > melhor_s:
                        melhor_s = s
                        melhor_c = c
                escolhidos.append(melhor_c)
                candidatos.remove(melhor_c)

            aposta = sorted(escolhidos)
            if not self._validar_regras_equilibrio(aposta):
                # score final: raridade de pares + trincas + entropia por décadas
                score_total = 0.0
                # pares
                for i in range(len(aposta)):
                    for j in range(i + 1, len(aposta)):
                        a, b = aposta[i], aposta[j]
                        score_total += wpar[(a, b)]
                # trincas
                if self.trincas_coocorrencia:
                    for i in range(0, 4):
                        for j in range(i + 1, 5):
                            for k in range(j + 1, 6):
                                tri = tuple(sorted([aposta[i], aposta[j], aposta[k]]))
                                freq_tri = self.trincas_coocorrencia.get(tri, 0)
                                score_total += 0.5 * (1.0 / (freq_tri + 1))
                # entropia por décadas
                score_total += 0.3 * self._entropia_decadas(aposta)
                if score_total > melhor_score:
                    melhor_score = score_total
                    melhor_aposta = aposta

        return melhor_aposta

    def _validar_regras_equilibrio(self, numeros: list[int]) -> bool:
        """Retorna True se a aposta deve ser rejeitada (inválida)."""
        numeros_sorted = sorted(numeros)
        # Regras originais
        if self._tem_sequencia_consecutiva(numeros_sorted):
            return True
        if self._muitos_multiplos_de_5(numeros_sorted):
            return True
        # Regras de equilíbrio
        pares = sum(1 for n in numeros_sorted if n % 2 == 0)
        impares = NUMERO_DE_NUMEROS_POR_APOSTA - pares
        if not (2 <= pares <= 4):
            return True
        dezenas = {n // 10 for n in numeros_sorted}
        if len(dezenas) < 4:  # cobertura mínima de 4 décadas
            return True
        soma = sum(numeros_sorted)
        if not (150 <= soma <= 210):
            return True
        return False

    @staticmethod
    def _entropia_decadas(numeros: list[int]) -> float:
        """Entropia de Shannon aproximada das décadas presentes na aposta."""
        from math import log2
        decs = [n // 10 for n in numeros]
        total = len(numeros)
        counts = defaultdict(int)
        for d in decs:
            counts[d] += 1
        probs = [c / total for c in counts.values()]
        ent = -sum(p * log2(p) for p in probs)
        # normalizar por log2(total) ~ máximo quando distribuição uniforme
        ent_max = log2(total)
        return ent / ent_max if ent_max > 0 else ent

    # --- Geração de portfólio (Estratégia 2) ---
    def gerar_portfolio_estrategia2(self, quantidade: int) -> list[list[int]]:
        """Gera N apostas maximizando cobertura de pares/trincas e baixa sobreposição.
        Simples heurística: gerar candidatas via gerar_aposta_estrategia2 e aceitar
        aquelas que adicionam mais pares/trincas novos ao conjunto atual.
        """
        portfolio: list[list[int]] = []
        pares_cobertos = set()
        trincas_cobertas = set()

        tentativas = 0
        while len(portfolio) < quantidade and tentativas < quantidade * 10:
            cand = self.gerar_aposta_estrategia2()
            tentativas += 1
            if not cand:
                continue
            # medir ganho de cobertura
            novos_pares = set()
            for i in range(6):
                for j in range(i + 1, 6):
                    a, b = cand[i], cand[j]
                    novos_pares.add((a, b))
            ganho_pares = len(novos_pares - pares_cobertos)

            novos_trincas = set()
            for i in range(0, 4):
                for j in range(i + 1, 5):
                    for k in range(j + 1, 6):
                        tri = tuple(sorted([cand[i], cand[j], cand[k]]))
                        novos_trincas.add(tri)
            ganho_trincas = len(novos_trincas - trincas_cobertas)

            # aceitador simples: exige algum ganho de cobertura
            if ganho_pares + ganho_trincas > 0:
                portfolio.append(cand)
                pares_cobertos |= novos_pares
                trincas_cobertas |= novos_trincas

        # se insuficiente, completa com estratégia 1
        while len(portfolio) < quantidade:
            alt = self.gerar_aposta_estrategia1()
            if alt:
                portfolio.append(alt)
            else:
                break
        return portfolio

    # --- Relatório estatístico (qui-quadrado) ---
    def gerar_relatorio_estatistico(self) -> str:
        """Gera um resumo com teste de uniformidade por números (qui-quadrado)."""
        try:
            if not self.frequencias:
                return "Sem dados de frequência."
            import math
            total = sum(self.frequencias.values())
            if total == 0:
                return "Sem ocorrências nos dados."
            esperado = total / 60.0
            chi2 = 0.0
            for n in range(1, 61):
                obs = self.frequencias.get(n, 0)
                chi2 += (obs - esperado) ** 2 / esperado
            # gl = 59
            resumo = (
                f"Total observações: {total}\n"
                f"Esperado por número: {esperado:.2f}\n"
                f"Qui-quadrado (gl=59): {chi2:.2f}\n"
                "Este resultado é apenas indicativo; use com cautela."
            )
            return resumo
        except Exception as e:
            logging.error(f"Erro ao gerar relatório estatístico: {e}")
            return f"Erro ao gerar relatório: {e}"

    @staticmethod
    def salvar_aposta_excel(numeros: list[int], estrategia: str) -> None:
        """Salva a aposta na planilha "apostas.xlsx" com a coluna da estratégia."""
        try:
            data_atual = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            # Monta a linha como dict para evitar erro de 'mismatched columns'
            nova_linha_dict = {
                "Estrategia": estrategia,
                "Data": data_atual,
                "N1": numeros[0],
                "N2": numeros[1],
                "N3": numeros[2],
                "N4": numeros[3],
                "N5": numeros[4],
                "N6": numeros[5],
            }
            
            if not os.path.exists(ARQUIVO_EXCEL_APOSTAS):
                df = pd.DataFrame(columns=["Estrategia", "Data", "N1", "N2", "N3", "N4", "N5", "N6"])
                df.loc[0] = nova_linha_dict
            else:
                df = pd.read_excel(ARQUIVO_EXCEL_APOSTAS)
                # Garantir coluna 'Estrategia'
                if 'Estrategia' not in df.columns:
                    df.insert(0, 'Estrategia', '-')
                # Garantir coluna 'Data'
                if 'Data' not in df.columns:
                    df.insert(1, 'Data', pd.NA)
                # Garantir ordem das colunas
                colunas_alvo = ["Estrategia", "Data", "N1", "N2", "N3", "N4", "N5", "N6"]
                # Se existir colunas extras, manter após as principais
                # Garante colunas N1..N6
                for c in ["N1","N2","N3","N4","N5","N6"]:
                    if c not in df.columns:
                        df[c] = pd.NA
                # Reordena exatamente para as colunas alvo (descarta extras para evitar mismatch)
                df = df[[c for c in colunas_alvo]]
                # Adiciona a nova linha por dict alinhando por nome de coluna
                df.loc[len(df)] = nova_linha_dict

            df.to_excel(ARQUIVO_EXCEL_APOSTAS, index=False)
            logging.info(f"Aposta {numeros} (Estratégia: {estrategia}) salva em {ARQUIVO_EXCEL_APOSTAS}")

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
        # Aumenta largura para acomodar todos os 6 números visíveis
        self.geometry("960x640")
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

        # Botões de geração por estratégia
        btns_frame = ttk.Frame(self.control_frame)
        btns_frame.grid(row=2, column=0, pady=(0, 10))

        self.botao_e1 = ttk.Button(btns_frame, text="Gerar (Estratégia 1)", command=self.exibir_aposta_e1, bootstyle=SUCCESS, padding=10)
        self.botao_e1.grid(row=0, column=0, padx=5)

        self.botao_e2 = ttk.Button(btns_frame, text="Gerar (Estratégia 2)", command=self.exibir_aposta_e2, bootstyle=INFO, padding=10)
        self.botao_e2.grid(row=0, column=1, padx=5)

        # Portfólio (Estratégia 2 avançada)
        ttk.Label(btns_frame, text="Qtd:").grid(row=0, column=2, padx=(15, 5))
        self.qtd_portfolio = ttk.Spinbox(btns_frame, from_=2, to=20, width=5)
        self.qtd_portfolio.set(5)
        self.qtd_portfolio.grid(row=0, column=3)
        self.botao_portfolio = ttk.Button(btns_frame, text="Gerar Portfólio (E2)", command=self.exibir_portfolio_e2, bootstyle=WARNING, padding=10)
        self.botao_portfolio.grid(row=0, column=4, padx=5)

        # Relatório estatístico
        self.botao_relatorio = ttk.Button(btns_frame, text="Relatório Estatístico", command=self.exibir_relatorio_estatistico, bootstyle=SECONDARY, padding=10)
        self.botao_relatorio.grid(row=0, column=5, padx=5)

        # --- Frame de Histórico ---
        history_title = ttk.Label(self.history_frame, text="Histórico de Apostas", font=("-size 14"))
        history_title.grid(row=0, column=0, sticky="w", pady=(0, 5))
        
        cols = ("Estratégia", "Data", "N1", "N2", "N3", "N4", "N5", "N6")
        self.tree = ttk.Treeview(self.history_frame, columns=cols, show="headings", bootstyle=PRIMARY)
        
        for col in cols:
            self.tree.heading(col, text=col)
            self.tree.column(col, width=80, anchor=CENTER)
        self.tree.column("Estratégia", width=110)
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


    def exibir_aposta_e1(self):
        """Gera, valida, exibe e salva a aposta pela Estratégia 1."""
        logging.info("Botão 'Gerar (Estratégia 1)' clicado.")
        self._toggle_botoes(False)
        self.update_idletasks()

        aposta = self.logic.gerar_aposta_estrategia1()

        if aposta:
            for i, numero in enumerate(aposta):
                self.ball_labels[i].config(text=f"{numero:02}")
            self.logic.salvar_aposta_excel(aposta, estrategia="E1")
            self.atualizar_janela_planilha(highlight_new=True)
        else:
            msg = "Não foi possível gerar uma aposta válida pela Estratégia 1."
            messagebox.showwarning("Aviso", f"{msg} Verifique o log para mais detalhes.")

        self._toggle_botoes(True)

    def exibir_aposta_e2(self):
        """Gera, valida, exibe e salva a aposta pela Estratégia 2."""
        logging.info("Botão 'Gerar (Estratégia 2)' clicado.")
        self._toggle_botoes(False)
        self.update_idletasks()

        aposta = self.logic.gerar_aposta_estrategia2()

        if aposta:
            for i, numero in enumerate(aposta):
                self.ball_labels[i].config(text=f"{numero:02}")
            self.logic.salvar_aposta_excel(aposta, estrategia="E2")
            self.atualizar_janela_planilha(highlight_new=True)
        else:
            msg = "Não foi possível gerar uma aposta válida pela Estratégia 2."
            messagebox.showwarning("Aviso", f"{msg} Verifique o log para mais detalhes.")

        self._toggle_botoes(True)

    def exibir_portfolio_e2(self):
        """Gera múltiplas apostas pela Estratégia 2 e salva todas."""
        logging.info("Botão 'Gerar Portfólio (Estratégia 2)' clicado.")
        self._toggle_botoes(False)
        self.update_idletasks()

        try:
            qtd = int(self.qtd_portfolio.get())
        except Exception:
            qtd = 5
        qtd = max(1, min(50, qtd))

        apostas = self.logic.gerar_portfolio_estrategia2(qtd)
        if apostas and len(apostas) > 0:
            for aposta in apostas:
                for i, numero in enumerate(aposta):
                    self.ball_labels[i].config(text=f"{numero:02}")
                self.logic.salvar_aposta_excel(aposta, estrategia="E2-PORT")
            self.atualizar_janela_planilha(highlight_new=True)
        else:
            messagebox.showwarning("Aviso", "Não foi possível gerar o portfólio solicitado.")

        self._toggle_botoes(True)

    def exibir_relatorio_estatistico(self):
        """Exibe relatório estatístico básico (qui-quadrado)."""
        rel = self.logic.gerar_relatorio_estatistico()
        messagebox.showinfo("Relatório Estatístico", rel)

    def _toggle_botoes(self, habilitar: bool):
        """Habilita/Desabilita os botões de geração para evitar cliques simultâneos."""
        state = "normal" if habilitar else "disabled"
        self.botao_e1.config(state=state)
        self.botao_e2.config(state=state)
        self.botao_portfolio.config(state=state)
        self.botao_relatorio.config(state=state)

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

                    estrategia_val = row['Estrategia'] if 'Estrategia' in df.columns else "-"
                    numeros_vals = list(row[2:]) if 'Estrategia' in df.columns else list(row[1:])
                    display_values = [estrategia_val, formatted_date] + numeros_vals
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