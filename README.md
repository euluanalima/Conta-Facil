# Conta Fácil — Web

Aplicação web para processamento, conferência e acompanhamento das faturas de água da EMBASA na UFRB — NUMAM.

## Acesso e publicação
O projeto está preparado para publicação no Streamlit Community Cloud.

- Arquivo principal: `app.py`
- Núcleo de regras: `conta_facil_core.py`
- Dependências: `requirements.txt`
- Configuração visual do Streamlit: `.streamlit/config.toml`

## Funcionalidades
- upload do PDF consolidado da EMBASA;
- extração das faturas por matrícula;
- consolidação da planilha de controle;
- histórico mensal;
- pesquisa por local, centro e matrícula;
- gráficos analíticos;
- exportação Excel;
- geração do Relatório Mensal de Ateste em Word;
- comparação com o mês calendário imediatamente anterior.

## Documentação

A documentação completa está na pasta `docs/`:

- [Guia do usuário](docs/GUIA_DO_USUARIO.md)
- [Documentação técnica](docs/DOCUMENTACAO_TECNICA.md)
- [Regras de cálculo](docs/REGRAS_DE_CALCULO.md)
- [Arquitetura](docs/ARQUITETURA.md)
- [Deploy no Streamlit](docs/DEPLOY_STREAMLIT.md)
- [Manutenção](docs/MANUTENCAO.md)
- [Histórico de alterações](CHANGELOG.md)

## Regra de precisão monetária
Os valores são mantidos internamente em **centavos inteiros**. A conversão para reais ocorre apenas na apresentação e nas exportações.

## Histórico
O sistema aceita o Excel exportado pelo próprio Conta Fácil como histórico. No Streamlit Community Cloud, o armazenamento local pode ser reiniciado, portanto o histórico persistente em nuvem deve ser tratado separadamente caso seja necessário.

## Observação
A documentação descreve o comportamento atual do código. Alterações de interface devem ser feitas sem modificar as regras de cálculo sem revisão específica.
