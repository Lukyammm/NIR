# NIR - Passagem de Plantao e Regulacao

Aplicativo para o Nucleo Interno de Regulacao (NIR), desenhado a partir da planilha operacional e dos rascunhos de fluxo recebidos em 16/06/2026.

O foco principal e gerar relatorios diarios de troca de plantao a partir de filtros, mantendo tambem o registro estruturado de regulacoes, leitos/UIB, procedimentos externos, altas de UTI, bloqueios, pendencias e auditoria.

## Decisoes do projeto

- **MVP em Google Apps Script + Google Sheets**: aproveita a rotina atual baseada em planilhas, reduz infraestrutura inicial e permite implantacao rapida.
- **Modelo de dados normalizado**: em vez de duplicar uma aba por dia/turno, o app grava eventos em tabelas permanentes e gera relatorios por filtro.
- **Relatorio como produto central**: o relatorio diario e montado por periodo, turno, status, especialidade, origem, unidade e modulo.
- **Historico e auditoria por padrao**: toda edicao relevante recebe timestamp, usuario e snapshot quando o plantao e encerrado.
- **Repositorio sem dados reais**: este repo e publico. Nao versionar nomes de pacientes, prontuarios, telefones, planilhas reais ou exports operacionais.

## Estrutura

- `Code.gs`: backend Apps Script, criacao de abas, CRUD, filtros e geracao de relatorio.
- `index.html`: interface web responsiva, com modo demo local e integracao com `google.script.run`.
- `docs/ANALISE_PLANILHA.md`: leitura critica da planilha e dos rascunhos.
- `docs/MODELO_DADOS.md`: entidades, campos, status e regras.
- `docs/RELATORIO_DIARIO.md`: formato do relatorio de passagem de plantao.
- `docs/IMPLANTACAO.md`: como publicar e operar o MVP.

## Funcionalidades

- **Painel** com KPIs operacionais (regulacoes, reservados, internados, bloqueios, altas UTI) e indicadores de tempo medio (reserva, chegada e permanencia).
- **Fluxo visual** com sala de situacao, mapa de UIB/leitos, agenda de exames/procedimentos e fila protegida para pendencias que atravessam turno.
- **Relatorio diario** filtravel por periodo, turno, busca e modulo, com copia para area de transferencia e impressao/PDF.
- **Modulos CRUD**: regulacao, UIB/leitos, procedimentos, UTI/SRPA e bloqueios, com arquivamento protegido (soft delete) e exportacao CSV.
- **Auditoria**: aba com historico real de alteracoes (NIR_LOG), alem da governanca de privacidade.
- **Modo demo local**: roda sem backend usando `localStorage`, util para testes e demonstracoes.

## Ajuste para o modelo operacional solicitado

O app evita a logica de "varias abas" como ponto de trabalho. A interface concentra os registros em uma visao unica:

- UIB/leitos aparecem como mapa visual por unidade, com alerta de especialidade potencialmente incompativel.
- Procedimentos e exames aparecem em calendario operacional a partir da data programada.
- Registros em acompanhamento ficam em fila protegida ate receberem desfecho claro.
- `RESERVA_CONFIRMADA` continua sendo pendencia operacional, nao status final.
- Arquivar registro aberto e bloqueado para reduzir erro de copia/cola entre plantoes.

## Abas criadas no Google Sheets

- `NIR_CONFIG`
- `NIR_PLANTOES`
- `NIR_REGULACOES`
- `NIR_LEITOS_UIB`
- `NIR_PROCEDIMENTOS`
- `NIR_UTI_SRPA`
- `NIR_BLOQUEIOS`
- `NIR_CONTATOS`
- `NIR_RELATORIOS`
- `NIR_LOG`




## Guia visual do sistema

> Telas demonstrativas baseadas nos componentes, textos, cores e fluxos encontrados no código deste repositório. Os dados exibidos são fictícios e não representam pacientes, profissionais ou instituições reais.

### NIR - sala de situação
![NIR - sala de situação](./docs/screenshots/situation-room.svg)

### NIR - mapa de leitos
![NIR - mapa de leitos](./docs/screenshots/bed-map.svg)

### NIR - relatório diário
![NIR - relatório diário](./docs/screenshots/daily-report.svg)

## Como rodar

1. Crie uma planilha Google privada para o NIR.
2. Abra **Extensoes > Apps Script**.
3. Copie `Code.gs` e `index.html` para o projeto Apps Script.
4. Em **Propriedades do script**, defina `PLANILHA_ID` com o ID da planilha.
5. Execute `criarEstruturaNIR()` uma vez para criar as abas.
6. Publique como Web App.

Para testar a interface localmente, abra `index.html` no navegador. Ela usa dados sinteticos em `localStorage` quando `google.script.run` nao existe.

## Regras de privacidade

Dados reais de pacientes e contatos devem ficar apenas no ambiente privado do hospital. Se for necessario compartilhar prints ou exports, remover identificadores diretos antes.
