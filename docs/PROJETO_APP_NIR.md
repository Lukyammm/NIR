# Projeto App NIR

## Síntese executiva

O NIR precisa de um aplicativo operacional para transformar a planilha de plantão em um sistema rastreável, filtrável e capaz de gerar relatórios diários de troca de plantão sem retrabalho.

A função principal é o **relatório diário por filtros**, mas para esse relatório ser confiável o app precisa estruturar os dados que o alimentam: regulação, UIB/leitos, procedimentos, UTI/SRPA, bloqueios, contatos e auditoria.

## Problema atual

- A operação depende de múltiplas abas e regiões da planilha.
- O relatório de plantão exige copiar informações de lugares diferentes.
- Status e desfechos não são totalmente padronizados.
- A análise por período, turno, ADS, município, especialidade e origem fica difícil.
- Tempos importantes não são capturados automaticamente.
- O histórico por plantão existe, mas é caro de consolidar.

## Proposta

Criar um Web App do NIR com:

- abertura e encerramento de plantão;
- registros estruturados por módulo;
- filtros globais;
- painel com indicadores;
- geração de relatório diário;
- snapshot do relatório no encerramento;
- log de alterações;
- modo de implantação rápido em Google Apps Script + Google Sheets.

## Módulos

### Painel

Visão imediata do plantão:

- regulações totais;
- reservados;
- internados;
- negados;
- cancelados;
- pendências;
- bloqueios;
- altas de UTI aguardando leito;
- gráficos por status e especialidade.

### Relatório diário

Gera o texto de passagem de plantão a partir de filtros:

- período;
- turno;
- status;
- especialidade;
- origem de regulação;
- unidade;
- busca livre;
- somente pendentes.

Saídas:

- prévia em tela;
- copiar texto;
- imprimir/salvar em PDF pelo navegador;
- snapshot no encerramento do plantão.

### Regulação

Unifica CRL, SAMU, local, SIREG e fila Dona Libânia.

Campos críticos:

- Fastmedic;
- unidade de origem;
- ADS/município;
- especialidade;
- status;
- avaliador;
- justificativa;
- observação;
- marcos de tempo.

### UIB / Leitos

Controla ocupação, reserva, falta chegar, transferência entre UIBs, permanência e desfecho.

Inclui regra futura para sinalizar paciente em UIB incompatível com especialidade.

### Procedimentos

Consolida imagem, endoscopia, vascular, radiointervenção, urologia, hemodiálise, transferência externa e SAD/EMAD/CCC.

### UTI / SRPA

Controla:

- pedidos de UTI;
- SRPA aguardando leito;
- altas de UTI aguardando leito;
- altas com leito.

### Bloqueios

Controla:

- manutenção;
- isolamento;
- leito ectópico;
- reserva técnica;
- previsão e encerramento.

## Indicadores

- regulações por status;
- regulações por especialidade;
- vagas por turno;
- vagas remanescentes;
- pacientes em tela;
- tempo médio vaga dada -> reserva;
- tempo médio reserva -> chegada;
- tempo médio chegada -> internação/cancelamento;
- permanência média em UIB;
- bloqueios abertos;
- procedimentos pendentes.

## Arquitetura do MVP

### Camada de dados

Google Sheets privado com abas normalizadas.

### Backend

Google Apps Script (`Code.gs`):

- cria estrutura das abas;
- registra plantões;
- salva/atualiza/arquiva registros;
- gera relatórios;
- salva snapshots;
- mantém log.

### Frontend

HTML/CSS/JS (`index.html`):

- painel operacional;
- filtros;
- tabelas;
- formulário lateral;
- relatório;
- modo demo local sem Google Apps Script.

## Privacidade

O repositório é público e não deve conter dados reais.

Regras:

- não versionar planilhas reais;
- não versionar nomes de pacientes;
- não versionar prontuários reais;
- não versionar telefones reais;
- usar `paciente_alias` em demonstrações;
- manter dados reais no ambiente privado do hospital.

## Fases sugeridas

### Fase 1 - MVP funcional

- estrutura de abas;
- frontend operacional;
- abertura/encerramento de plantão;
- CRUD dos módulos principais;
- relatório por filtros;
- log básico.

### Fase 2 - Migração assistida

- importador da planilha atual;
- de/para de status;
- padronização de especialidades;
- validação de campos obrigatórios.

### Fase 3 - Inteligência operacional

- gráficos por dia/mês;
- indicadores por ADS e município;
- alerta de UIB incompatível;
- alertas de tempo parado;
- painel de gargalos.

### Fase 4 - Plataforma institucional

- banco relacional;
- autenticação por perfil;
- integração com sistemas externos;
- geração PDF server-side;
- trilha de auditoria avançada.

## Critérios de sucesso

- gerar relatório diário em poucos cliques;
- reduzir retrabalho de copiar/colar;
- permitir filtro por período, turno, status, especialidade, origem, ADS e unidade;
- acompanhar pendências entre plantões;
- manter histórico confiável;
- não expor dados sensíveis no repositório.
