# Ocorrências do Plantão – NIR

## Como o plantão ativo é identificado
- O sistema considera **plantão ativo** quando existe um ID preenchido em `CONFIG_PLANTAO!A2`.
- A criação de um novo plantão (botão "Abrir Plantão" + salvar equipe) grava o ID em `A2` e o horário de abertura em `J2`.
- Ao encerrar o plantão (botão "Encerrar"), os dados da linha 2 são movidos para `HIST_PLANTOES` e `A2:K2` são limpos.
- Não há checagem automática de data/hora para expirar o plantão; se `A2` estiver vazio, a interface mostra "Nenhum plantão ativo".
- Se houver "Nenhum plantão ativo" mesmo com a linha preenchida, o código agora preenche `A2` com um ID gerado e registra o evento. Se o aviso persistir, confira se a planilha está sendo lida pelo mesmo arquivo ativo do Apps Script.
