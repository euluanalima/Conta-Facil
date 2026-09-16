# Regras de Cálculo — Conta Fácil

## 1. Princípio geral
Todos os cálculos financeiros são realizados com valores em **centavos inteiros**.

## 2. Valor total do mês
O total institucional do mês é a soma dos valores atuais de todas as matrículas:

```text
Total do mês = Σ valor atual de cada matrícula
```

## 3. Aum / Red por matrícula
A comparação é feita contra o **mês calendário imediatamente anterior**.

```text
Aum / Red (%) = ((Atual - Anterior) / Anterior) × 100
```

Casos especiais:
- anterior = 0 e atual = 0: sem variação calculável;
- anterior = 0 e atual > 0: 100%;
- mês anterior sem dados: variação não calculada.

## 4. Percentual dentro do centro

```text
% Centro = (Valor da matrícula / Total do centro) × 100
```

O total do centro é obtido pelo agrupamento das matrículas daquele centro.

## 5. Percentual no total institucional

```text
% Total = (Valor da matrícula / Total institucional) × 100
```

## 6. Resumo por centro
Para cada centro:

```text
Valor atual do centro = Σ matrículas do centro no mês atual
Valor anterior do centro = Σ matrículas do centro no mês anterior
Variação do centro (%) = ((Atual - Anterior) / Anterior) × 100
```

## 7. Variação geral usada no relatório
O texto do relatório compara o total institucional atual com o total institucional do mês anterior:

```text
Variação geral (%) = ((Total atual - Total anterior) / Total anterior) × 100
```

O percentual é exibido com uma casa decimal.

## 8. Ausência de histórico
Se não houver dados para o mês imediatamente anterior:
- o relatório não deve afirmar aumento ou redução;
- o gráfico deve representar ausência de dado em vez de criar um zero artificial.

## 9. Fonte dos totais
O núcleo do sistema calcula os totais diretamente a partir das linhas das matrículas agrupadas pelo pandas.

Não devem ser usadas faixas fixas de linhas como regra de negócio para subtotal por centro, porque inclusões ou mudanças de posição podem deixar uma matrícula fora da faixa.

## 10. Exemplo de auditoria
Para verificar uma variação geral:
1. some todas as matrículas do mês anterior;
2. some todas as matrículas do mês atual;
3. aplique a fórmula de variação geral;
4. compare com o panorama e a conclusão do relatório.

Uma divergência entre um subtotal manual de Excel e o relatório deve ser investigada pela composição das linhas antes de se alterar a fórmula do sistema.
