# Changelog

Todas as mudanças notáveis deste projeto serão documentadas aqui.

## [1.1.0] - 2025-11-01
### Adicionado
- Estratégia 2 avançada: raridade de pares e trincas (co-ocorrência histórica), suavização Bayesiana e termo de entropia por décadas.
- Geração de Portfólio (E2) com Spinbox de quantidade e botão dedicado.
- Relatório estatístico simples (qui-quadrado) acessível via botão.
- Coluna "Estratégia" no histórico (UI) e no arquivo apostas.xlsx.

### Alterado
- UI ampliada para 960x640 para exibir o N6 com mais conforto.
- Histórico reorganizado para mostrar "Estratégia" à esquerda de "Data".
- Ajustes de layout: novos botões de geração (E1/E2), portfólio e relatório.

### Corrigido
- Salvamento em Excel: 
  - Inserção de linhas via dicionário para evitar "mismatched columns".
  - Migração automática adicionando colunas ausentes (Estrategia, Data, N1..N6) e reordenação segura.
- Problemas de importação de dependências (numpy/pandas/pillow) resolvidos com reinstalação sem cache.

## [1.0.0] - 2025-09-08
### Adicionado
- Interface moderna com ttkbootstrap.
- Estratégia 1: aleatória uniforme com restrições de equilíbrio (paridade, décadas, soma) e filtros básicos.
- Leitura de Mega-Sena.xlsx e cálculo de frequências históricas por número.
- Salvamento de apostas em apostas.xlsx e exibição de histórico na UI.