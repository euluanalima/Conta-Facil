# Guia do Usuário — Conta Fácil

## 1. Objetivo
O Conta Fácil foi desenvolvido para auxiliar a conferência mensal das contas de água da EMBASA vinculadas à UFRB, com consolidação por matrícula e centro, geração de planilha, gráficos e Relatório Mensal de Ateste.

## 2. Fluxo de uso

### 2.1 Enviar o PDF
Na tela inicial, use **Arquivo de faturamento** para selecionar o PDF consolidado da EMBASA.

O sistema espera que as páginas permitam identificar, entre outros dados:
- matrícula;
- mês/ano de referência;
- valor da fatura;
- data de vencimento;
- data de emissão, quando disponível.

### 2.2 Carregar histórico
O campo **Carregar histórico Excel (opcional)** permite selecionar um arquivo histórico.

O sistema reconhece:
- planilhas simples com cabeçalho na primeira linha;
- planilhas exportadas pelo próprio Conta Fácil, mesmo quando possuem título antes do cabeçalho.

As colunas cadastrais esperadas são:
- `LOCAL`;
- `CEN` ou `Centro`;
- `MAT` ou `Matrícula`.

### 2.3 Processar PDF
Clique em **Processar PDF**.

Após o processamento, o sistema:
1. extrai as faturas;
2. identifica a competência;
3. consolida os valores por matrícula;
4. combina o mês atual com o histórico;
5. calcula as variações;
6. monta o resumo por centro.

### 2.4 Planilha de Controle
Na aba **Planilha de Controle**, é possível:
- visualizar os registros;
- pesquisar por local, centro ou matrícula;
- conferir os valores mensais;
- conferir `Aum / Red`, `% Centro` e `% Total`;
- baixar a planilha Excel.

### 2.5 Gráficos Analíticos
A aba **Gráficos Analíticos** apresenta:
- distribuição do faturamento do mês por centro;
- evolução mensal do faturamento por centro.

Quando não há dado para determinado mês, o gráfico deve representar ausência de dado, e não faturamento zero.

### 2.6 Relatório de Ateste
Na aba **Relatório de Ateste**, preencha:
- contrato;
- nome do gestor;
- SIAPE;
- data do ateste;
- ocorrências/justificativas, quando existirem.

Depois, clique em **Gerar relatório Word**.

O relatório inclui:
- identificação do período;
- panorama geral;
- resumo por centro;
- gráficos;
- análise por centro;
- conclusão;
- planilha detalhada;
- verificação de conformidade.

## 3. Interpretação das principais colunas
- **Nº G**: ordem do cadastro.
- **LOCAL**: identificação do local da conta.
- **CEN**: centro/unidade responsável.
- **MAT**: matrícula da EMBASA.
- **Mês/Ano**: valor faturado no período.
- **Aum / Red**: variação percentual da matrícula contra o mês calendário imediatamente anterior.
- **% Centro**: participação da matrícula no total do seu centro.
- **% Total**: participação da matrícula no total institucional do mês.

## 4. Matrícula nova
Quando uma matrícula não existe no histórico, o sistema pode incluí-la e marcar o local como **MATRÍCULA NOVA - REVISAR LOCAL** e o centro como **NÃO MAPEADO** até revisão cadastral.

## 5. Conferência recomendada antes do ateste
Antes de emitir o relatório definitivo:
1. confirme a competência do PDF;
2. confira o total geral;
3. confira matrículas novas ou não mapeadas;
4. verifique grandes variações;
5. revise as ocorrências informadas pelo gestor;
6. confirme se o histórico do mês anterior está presente.
