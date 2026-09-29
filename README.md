# vazoes-sin-consistencia
# Análise e consistência de séries de vazões naturais do SIN

Ferramenta desenvolvida na Iniciação Científica "Análise e consistência de séries de
vazões afluentes nas usinas hidrelétricas brasileiras" (Escola Politécnica da USP, 2026),
sob orientação do Prof. Renato Carlos Zambon.

## O que faz
- Calcula as vazões incrementais (Qinc = Qi − ΣQm) das 154 usinas do SIN e classifica as inconsistências
- Testes de tendência e estacionariedade (Mann-Kendall com correção de Yue, Sen, Pettitt)
- Sazonalidade, curvas de permanência, evaporação líquida e usos consuntivos (diagnóstico)
- Correção por conservação de massa, gerando um deck corrigido no formato do NEWAVE
- Interface interativa em HTML com todos os resultados

## Requisitos
Python 3.10+ com pandas, numpy, scipy e openpyxl:
    pip install pandas numpy scipy openpyxl

## Dados de entrada
Incluídos no repositório (versão usada no relatório final):
- `vazoes`: deck de vazões naturais do NEWAVE, jan/1931 a ago/2026 (CCEE, www.ccee.org.br)
- `usinas atualizadas.xlsx`: cadastro das usinas
- `dados_hidroterm_completo.xlsx`: base do modelo HIDROTERM (cadastro, polinômios área-cota,
  evaporação)

Não incluída (arquivo grande, pública):
- Séries de usos consuntivos da ANA (Resolução nº 92/2021), disponível no portal da ANA.
  Sem ela, o programa roda normalmente, apenas sem a aba Usos Consuntivos.

Para reproduzir os resultados do relatório, basta executar na pasta do repositório:
    python gerar_interface.py

## Como executar
    python gerar_interface.py
    python gerar_interface.py --ana-usos [arquivo_da_ANA].xlsx

Saídas: interface HTML, `saidas/diagnostico.txt`, CSVs e `saidas/vazoes_corrigidas_massa.txt`.

## Citação
NAKAMURA, E. H. Análise e consistência de séries de vazões afluentes nas usinas
hidrelétricas brasileiras. Relatório final de Iniciação Científica – Escola Politécnica
da USP, São Paulo, 2026.
