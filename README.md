# NIR

Aplicação em Google Apps Script e HTML para registro e consulta das ocorrências do NIR.

## Abas usadas na planilha

- Ocorrências do fluxo principal são lançadas nas abas existentes apontadas em `NIR_SHEETS` (`RESERVA CONFIRMADA`, `PROCEDIMENTO CONFIRMADO`, `RESERVA NEGADA`, `PLANTÃO ANTERIOR`).
- Os relatórios de ocorrências por turno são gravados automaticamente nas abas **`REL ENF`** (enfermagem) e **`REL MED`** (médica) quando a função `saveRelatorio` é acionada. Se essas abas não existirem na planilha, o código as cria com o cabeçalho padrão antes de gravar o primeiro registro.

## Observação

Caso precise localizar os registros recentes de enfermagem ou médica diretamente na planilha, procure pelas abas `REL ENF` e `REL MED`; se ainda não houver nenhum relatório salvo pelo WebApp, elas não aparecerão até o primeiro uso.