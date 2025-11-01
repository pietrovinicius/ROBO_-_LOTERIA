# ROBO_-_LOTERIA

Gerador de apostas para a Mega-Sena com interface moderna (ttkbootstrap) e estratégias baseadas em princípios estatísticos.

Sumário
- Visão geral
- Instalação
- Execução
- Estratégias de geração
- Portfólio (Estratégia 2)
- Relatório estatístico
- Estrutura dos dados e arquivos
- Logs e solução de problemas
 - Guia rápido (passo a passo)
 - Screenshots da interface

Visão geral
Este projeto lê o histórico de sorteios (Mega-Sena.xlsx), calcula estatísticas básicas (frequência individual, co-ocorrência de pares e trincas) e oferece duas estratégias para gerar apostas:
- Estratégia 1 (E1): Seleção uniforme com restrições de equilíbrio (paridade 2–4 pares; ao menos 4 décadas; soma entre 150–210; evita sequências longas e excesso de múltiplos de 5).
- Estratégia 2 (E2): Geração que prioriza pares e trincas raros (baixa co-ocorrência histórica), com suavização Bayesiana das frequências e termo de entropia por décadas para espalhar os números.

Instalação
1) Requisitos: Python 3.12+ (Windows), pip.
2) Criar o ambiente virtual e instalar dependências:
   - py -m venv venv
   - .\venv\Scripts\python.exe -m pip install --upgrade pip
   - .\venv\Scripts\python.exe -m pip install pandas openpyxl ttkbootstrap pillow
3) Coloque o arquivo Mega-Sena.xlsx na raiz do projeto.

Execução
- .\venv\Scripts\python.exe main.py
- A janela “Gerador de Apostas Mega-Sena” será aberta.
- Botões disponíveis:
  - Gerar (Estratégia 1)
  - Gerar (Estratégia 2)
  - Qtd (Spinbox): quantidade para portfólio
  - Gerar Portfólio (E2)
  - Relatório Estatístico

Estratégias de geração
- E1 (uniforme com equilíbrio):
  - Paridade: 2–4 pares (preferência natural por 3 pares/3 ímpares)
  - Décadas: pelo menos 4 décadas distintas entre os 6 números
  - Soma: entre 150 e 210
  - Evita: 4+ números consecutivos e 3+ múltiplos de 5
- E2 (raridade de pares/trincas + suavização + entropia):
  - Priorização de pares/trincas pouco co-ocorrentes (pesos 1/(freq+1))
  - Suavização Bayesiana para frequência individual (alpha=80)
  - Termo de entropia por décadas para favorecer distribuição espalhada
  - Mantém as mesmas restrições de equilíbrio da E1

Portfólio (Estratégia 2)
- Gera N apostas maximizando a cobertura de pares/trincas e reduzindo sobreposição entre bilhetes.
- Salva cada aposta com estratégia “E2-PORT”.
- Use quantidades moderadas (ex.: 5 a 20) para bom desempenho.

Relatório estatístico
- Botão “Relatório Estatístico” calcula um qui-quadrado simples sobre frequências individuais.
- Útil como diagnóstico: ver se o histórico aparenta desvio relevante da uniformidade.

Estrutura dos dados e arquivos
- Mega-Sena.xlsx: base histórica usada para cálculos.
- apostas.xlsx: registro das apostas geradas com colunas [Estrategia, Data, N1..N6].
- log.txt: logs de execução e eventos.
- main.py: código da aplicação.

Logs e solução de problemas
- Se houver erro ao importar numpy/pandas/pillow, reinstale-os sem cache:
  - .\venv\Scripts\python.exe -m pip install --force-reinstall --no-cache-dir numpy pandas pillow
- Se houver erro ao salvar no Excel (“mismatched columns” ou “['Data'] not in index”), o código já tenta migrar automaticamente para o formato correto; delete o arquivo apostas.xlsx se necessário e gere novamente.
- A interface foi ampliada (960x640) para exibir claramente o N6.

Aviso estatístico
- As estratégias não garantem aumento de probabilidade de acerto de um bilhete isolado; o objetivo é equilibrar combinações, evitar padrões populares e aumentar diversidade quando se gera várias apostas.
- Guia rápido (passo a passo)
1) Prepare o ambiente:
   - py -m venv venv
   - .\\venv\\Scripts\\python.exe -m pip install pandas openpyxl ttkbootstrap pillow
2) Coloque Mega-Sena.xlsx na raiz do projeto.
3) Execute a aplicação:
   - .\\venv\\Scripts\\python.exe main.py
4) Use os botões:
   - Gerar (Estratégia 1): aposta uniforme e equilibrada.
   - Gerar (Estratégia 2): aposta priorizando pares/trincas raros.
   - Qtd + Gerar Portfólio (E2): gera N apostas focando cobertura.
   - Relatório Estatístico: mostra um resumo de qui-quadrado.
5) Verifique resultados:
   - Histórico na UI com coluna "Estratégia".
   - Apostas salvas em apostas.xlsx com [Estrategia, Data, N1..N6].

Screenshots da interface
Coloque imagens em docs/screenshots e elas aparecerão corretamente no GitHub.

Exemplos (placeholders):
![Tela principal](docs/screenshots/tela-principal.png)
![E1 gerada](docs/screenshots/e1-gerada.png)
![E2 gerada](docs/screenshots/e2-gerada.png)
![Portfólio E2](docs/screenshots/portfolio-e2.png)
![Relatório](docs/screenshots/relatorio-estatistico.png)