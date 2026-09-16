# Manutenção — Conta Fácil

## 1. Regra principal
Antes de qualquer alteração, identifique se ela é:

### Visual
Exemplos:
- logo;
- cores;
- espaçamentos;
- responsividade;
- cabeçalho;
- largura de componentes.

Essas mudanças devem ficar preferencialmente em `app.py`.

### Regra de negócio
Exemplos:
- fórmula de variação;
- agrupamento por centro;
- tratamento de matrícula;
- leitura de valores;
- geração dos totais.

Essas mudanças afetam `conta_facil_core.py` e devem ser revisadas com teste de regressão.

## 2. Checklist antes de publicar
- [ ] o PDF de teste ainda é lido;
- [ ] a quantidade de faturas é a mesma;
- [ ] o total geral é o mesmo;
- [ ] os subtotais por centro estão corretos;
- [ ] a competência foi identificada;
- [ ] a comparação usa o mês calendário anterior;
- [ ] o Excel abre normalmente;
- [ ] o Word é gerado;
- [ ] os gráficos não transformam ausência de dado em zero;
- [ ] o app continua responsivo.

## 3. Teste de regressão financeira
Para um PDF já conferido, registre:
- competência;
- quantidade de faturas;
- total geral;
- totais por centro;
- principais variações.

Depois de qualquer alteração no núcleo, os mesmos dados devem produzir os mesmos resultados, salvo quando a alteração tiver sido feita justamente para corrigir uma regra documentada.

## 4. Histórico
Ao alterar o tratamento de histórico, testar:
- Excel do próprio sistema;
- planilha externa simples;
- mês anterior presente;
- mês anterior ausente;
- matrícula nova;
- matrícula sem centro.

## 5. Excel
A aba **Memória de Cálculo** deve permanecer coerente com as regras reais do sistema.

Evite criar fórmulas dependentes de faixas fixas de linha para subtotais de centros. O núcleo usa agrupamento por centro sobre as matrículas.

## 6. Word
Ao alterar o relatório:
- não mudar valores ou percentuais apenas por alteração de texto;
- manter o panorama e a conclusão usando a mesma variação geral;
- conferir se a tabela detalhada corresponde ao controle;
- revisar renderização no Word/Google Docs.

## 7. Git
Recomendações:
- usar commits pequenos e descritivos;
- separar alteração visual de alteração de lógica;
- registrar correções relevantes no `CHANGELOG.md`;
- evitar subir PDFs, planilhas reais e relatórios com dados administrativos se não forem necessários ao código.

## 8. Recuperação
Se uma mudança causar erro:
1. identifique o commit anterior funcional;
2. compare `app.py` e `conta_facil_core.py`;
3. reverta apenas o arquivo afetado;
4. execute novamente o teste de regressão.
