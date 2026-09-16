# Documentação Técnica — Conta Fácil

## 1. Visão geral
O projeto web está dividido em duas partes principais:

- `app.py`: camada de interface web em Streamlit;
- `conta_facil_core.py`: núcleo de extração, consolidação, cálculos, histórico, gráficos e exportações.

A separação permite modificar a interface sem reescrever as regras de negócio.

## 2. Principais dependências
O projeto utiliza:
- Streamlit;
- pandas;
- pypdf;
- matplotlib;
- python-docx;
- xlsxwriter;
- openpyxl.

Consulte `requirements.txt` para a lista usada na implantação.

## 3. Estrutura de dados

### 3.1 Faturas
A função `extract_pdf()` produz um DataFrame com campos como:
- `Página`;
- `Matrícula`;
- `Mês/Ano`;
- `Data Emissão`;
- `Data Vencimento`;
- `Valor_Centavos`;
- `Nome na Fatura`.

### 3.2 Controle
A função `build_control()` combina o histórico com os dados do mês atual.

Estrutura principal:
- `Nº G`;
- `LOCAL`;
- `CEN`;
- `MAT`;
- colunas mensais no padrão `Jan/2026`, `Fev/2026`, etc.;
- `Aum / Red`;
- `% Centro`;
- `% Total`.

### 3.3 Resumo por centro
O resumo é calculado com agrupamento pandas sobre o DataFrame de controle.

Campos principais:
- centro;
- valor atual em centavos;
- valor anterior em centavos;
- variação percentual;
- participação no total institucional.

## 4. Precisão monetária
Os valores monetários são convertidos para **centavos inteiros** no momento da extração.

Exemplo:

`R$ 1.234,56 -> 123456`

Isso evita erro de ponto flutuante durante somas e agrupamentos.

## 5. Histórico
O sistema possui:
- base histórica incorporada;
- carregamento de histórico via Excel;
- histórico automático local em JSON;
- integração do histórico com o mês atual por matrícula.

A competência anterior usada para comparação é o **mês calendário imediatamente anterior**, não simplesmente a coluna anterior disponível.

## 6. Exportações

### Excel
`export_excel()` gera:
- `Controle`;
- `Faturas`;
- `Resumo Centros`;
- `Memória de Cálculo`.

### Gráficos
`make_chart_images()` gera:
- distribuição por centro;
- evolução mensal por centro.

### Word
`generate_word_report()` gera o Relatório Mensal de Ateste, incluindo narrativa analítica, tabelas, gráficos e seção de conformidade.

## 7. Interface web
O `app.py` é responsável por:
- upload de PDF;
- upload de histórico;
- acionamento do processamento;
- apresentação dos dados;
- pesquisa;
- downloads;
- formulários do ateste;
- responsividade visual.

A interface chama as funções do núcleo e não deve duplicar regras financeiras.

## 8. Erros e validações relevantes
O sistema interrompe o processamento quando:
- nenhuma fatura é reconhecida;
- não encontra matrícula;
- não encontra valor da fatura;
- não consegue identificar a competência;
- encontra competências diferentes no mesmo PDF.

Matrículas sem cadastro institucional são mantidas para revisão, em vez de serem descartadas.
