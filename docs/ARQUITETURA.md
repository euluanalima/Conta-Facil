# Arquitetura — Conta Fácil

## 1. Visão lógica

```text
Navegador
   │
   ▼
Streamlit (app.py)
   │
   ▼
Núcleo Python (conta_facil_core.py)
   ├── Extração PDF
   ├── Histórico
   ├── Consolidação
   ├── Cálculos
   ├── Excel
   ├── Gráficos
   └── Word
```

## 2. Camada de interface
Arquivo: `app.py`

Responsabilidades:
- renderizar a aplicação;
- receber arquivos;
- manter estado da sessão;
- apresentar resultados;
- oferecer downloads;
- coletar dados administrativos para o ateste.

Não deve ser usada para definir novas fórmulas financeiras.

## 3. Camada de domínio
Arquivo: `conta_facil_core.py`

Responsabilidades:
- converter valores monetários;
- extrair faturas;
- normalizar histórico;
- consolidar controle;
- calcular totais e percentuais;
- gerar narrativa analítica;
- exportar arquivos.

## 4. Fluxo principal

```text
PDF EMBASA
  ↓
extract_pdf()
  ↓
DataFrame de faturas
  ↓
build_control()
  ↓
Controle + Resumo
  ├── Interface
  ├── Excel
  ├── Gráficos
  └── Relatório Word
```

## 5. Histórico

```text
Base incorporada / Excel histórico / JSON local
                 ↓
             Matrícula
                 ↓
            build_control()
                 ↓
       histórico + mês atual
```

## 6. Critérios de projeto
- valores monetários em centavos inteiros;
- associação principal por matrícula;
- separação entre interface e regra de negócio;
- ausência de dado não deve virar zero visual;
- exportações devem reutilizar os dados consolidados do núcleo.

## 7. Implantação atual
O frontend e o backend Python são executados juntos pelo Streamlit Community Cloud.

Essa arquitetura é adequada para o volume atual do sistema. Caso seja necessário histórico permanente multiusuário, a persistência pode ser migrada para banco de dados sem reescrever as fórmulas do núcleo.
