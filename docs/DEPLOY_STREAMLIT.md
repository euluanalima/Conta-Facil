# Deploy no Streamlit Community Cloud

## 1. Pré-requisitos
- repositório no GitHub;
- conta no Streamlit Community Cloud;
- repositório autorizado no Streamlit.

## 2. Repositório
Repositório atual:

`euluanalima/Conta-Facil`

Branch principal:

`main`

Arquivo de entrada:

`app.py`

## 3. Publicação
No Streamlit Community Cloud:

1. clique em **Create app**;
2. selecione o repositório;
3. selecione a branch `main`;
4. informe `app.py` como Main file path;
5. escolha o endereço;
6. clique em **Deploy**.

## 4. Dependências
O Streamlit instala os pacotes presentes em `requirements.txt`.

## 5. Atualizações
Quando um commit é enviado para a branch vinculada, o Streamlit normalmente reconstrói ou atualiza a aplicação automaticamente.

Se a nova versão ainda não aparecer:
- aguarde a atualização;
- recarregue a página;
- use Ctrl + F5 no desktop quando necessário;
- consulte os logs do app em caso de falha de build.

## 6. Armazenamento local
O armazenamento local do ambiente de nuvem não deve ser tratado como banco permanente.

O arquivo de histórico JSON pode funcionar durante a execução, mas pode ser perdido em reinicializações do ambiente.

Para uso com histórico permanente entre sessões e dispositivos, recomenda-se futuramente usar um banco persistente.

## 7. Segurança operacional
Não armazene credenciais diretamente no código ou no repositório.

Caso futuramente sejam adicionadas chaves de banco ou API, use os mecanismos de secrets do provedor de hospedagem.

## 8. Teste após deploy
Após uma publicação:
1. abra o app;
2. envie um PDF conhecido;
3. confira quantidade de faturas;
4. confira total geral;
5. confira competência;
6. gere o Excel;
7. gere o Word;
8. confira os gráficos;
9. teste em tela desktop e mobile.
