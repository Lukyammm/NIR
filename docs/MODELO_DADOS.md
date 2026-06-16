# Modelo de dados proposto

## Princípios

- Um registro operacional deve existir uma vez e ser filtrável por plantão, período, turno, status, especialidade e origem.
- O relatório é uma visão gerada a partir dos registros, não uma aba digitada do zero.
- Nenhuma exclusão física no fluxo operacional: registros são arquivados com `deleted_at`.
- Toda alteração crítica deve ter log com antes/depois.

## Entidades

### Plantão

Tabela: `NIR_PLANTOES`

Campos principais:

- `id`
- `data`
- `dia_semana`
- `turno` (`MT` ou `SN`)
- `status` (`ABERTO`, `ENCERRADO`)
- médicos por ilha
- enfermeiros por ilha
- administrativos por ilha
- coordenação
- `aberto_em`, `aberto_por`
- `encerrado_em`, `encerrado_por`
- `observacao`

### Regulação

Tabela: `NIR_REGULACOES`

Representa a fila Dona Libânia, SAMU, CRL, local/SIREG e solicitações acompanhadas na tela.

Campos principais:

- `origem_regulacao`
- `fastmedic`
- `prontuario`
- `paciente_nome` ou, para uso público, `paciente_alias`
- `ads`
- `municipio`
- `unidade_origem`
- `especialidade`
- `status`
- `avaliado_por`
- `justificativa`
- `observacao`
- `sincronizacao`
- marcos de tempo: `vaga_dada_em`, `reservado_em`, `chegou_em`, `internado_em`, `cancelado_em`
- `prioridade`
- `flags`

### UIB / Leitos

Tabela: `NIR_LEITOS_UIB`

Representa ocupação, reserva, falta chegar, permanência e desfecho.

Campos principais:

- `unidade`
- `leito`
- `status_leito`
- paciente
- `data_internacao`
- `especialidade_origem`
- `especialidade_destino`
- `isolamento`
- `patologia`
- `pendencia`
- `permanencia_dias`
- `desfecho`
- `transferir_para`
- `observacao`

Regra recomendada: alertar quando `especialidade_destino` não combinar com a unidade do leito.

### Procedimentos / Rede externa

Tabela: `NIR_PROCEDIMENTOS`

Unifica imagem, endoscopia, vascular, radiointervenção, urologia, hemodiálise, transferência externa e SAD/EMAD/CCC.

Campos principais:

- `tipo`
- paciente
- `fastmedic`
- `origem`
- `apto_medico`
- `judicial`
- `avaliado_por`
- `agendado_por`
- `procedimento`
- `data_programacao`
- `hora_programacao`
- `status`
- `ultimo_contato`
- `observacao`

### UTI / SRPA

Tabela: `NIR_UTI_SRPA`

Campos principais:

- `tipo`: pedido de UTI, SRPA aguardando leito, alta da UTI aguardando leito, alta da UTI com leito.
- `prontuario`
- paciente
- `leito_atual`
- `especialidade`
- `status`
- `destino`
- `comunicado_hora`
- `observacao`

### Bloqueios

Tabela: `NIR_BLOQUEIOS`

Campos principais:

- `tipo_bloqueio`: manutenção, isolamento, ectópico, reserva técnica.
- `unidade`
- `leito`
- paciente, quando aplicável
- `motivo`
- `data_inicio`
- `previsao`
- `status`
- `observacao`
- `encerrado_em`

### Relatórios e log

`NIR_RELATORIOS` guarda snapshots gerados.

`NIR_LOG` guarda eventos de abertura/encerramento de plantão, criação, atualização, arquivamento e geração de relatório.

## Status padronizados

Regulação:

- `NOVO`
- `EM_AVALIACAO`
- `AGUARDANDO_REGULACAO`
- `RESERVADO`
- `FALTA_CHEGAR`
- `INTERNADO`
- `NEGADO`
- `CANCELADO`

Procedimentos:

- `PENDENTE`
- `AGENDADO`
- `REAGENDADO`
- `REALIZADO`
- `CANCELADO`
- `NEGADO`

Leitos:

- `LIVRE`
- `RESERVADO`
- `FALTA_CHEGAR`
- `INTERNADO`
- `BLOQUEADO`

Desfechos:

- `ALTA`
- `OBITO`
- `TRANSFERIDO`
- `PERMANECE`
- `CANCELADO`

## Indicadores derivados

- total de regulações por período/turno
- reservados, internados, negados, cancelados
- pacientes ainda pendentes
- vagas por especialidade
- pacientes em tela
- tempo médio vaga dada -> reserva
- tempo médio reserva -> chegada
- tempo médio chegada -> internação/cancelamento
- permanência média em UIB
- bloqueios abertos
- altas de UTI aguardando leito
- procedimentos pendentes por tipo
