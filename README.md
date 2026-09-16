# Conta Fácil — Web

Aplicação web para processamento e conferência das faturas de água da EMBASA na UFRB.

## Publicação
Este repositório está preparado para o Streamlit Community Cloud.

Arquivo principal: `app.py`

## Funcionalidades
- upload do PDF consolidado da EMBASA;
- planilha de controle;
- histórico em Excel;
- pesquisa por local, centro e matrícula;
- gráficos analíticos;
- exportação Excel;
- geração do Relatório Mensal de Ateste em Word.

## Observação sobre histórico
No Streamlit Community Cloud, o disco local pode ser reiniciado. O sistema aceita o Excel exportado pelo próprio Conta Fácil para recuperar o histórico. Persistência em banco (ex.: Supabase) pode ser adicionada depois.
