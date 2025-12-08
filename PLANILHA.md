# Guia das abas da planilha NIR

Este arquivo descreve cada aba usada pelo WebApp, os nomes exatos das colunas e para que servem. Sempre mantenha esses cabeçalhos, pois o código valida e cria automaticamente qualquer coluna ausente ao salvar um registro.

## Abas de ocorrências

### RESERVA CONFIRMADA e PROCEDIMENTO CONFIRMADO
Cabeçalho padrão de ocorrências confirmadas. As colunas `Dia` e `Turno` (posições 21 e 22) permitem filtragem rápida na interface.

| Coluna | Função |
| --- | --- |
| TIPO | Categoria operacional do registro (ex.: Reserva, Procedimento). |
| Fastmedic | Código ou referência do chamado. |
| Nome do Paciente | Identificação do paciente. |
| Leito Reservado | Número do leito reservado. |
| Especialidade | Especialidade responsável. |
| Origem | Local de origem do paciente. |
| Status | Situação atual (ex.: Em andamento/Concluído). |
| Data da Reserva | Data em que a reserva foi cadastrada. |
| Hora da Reserva | Horário do cadastro da reserva. |
| Data da Confirmação da Reserva | Data em que a reserva foi confirmada. |
| Hora da Confirmação da Reserva | Horário da confirmação. |
| Tempo entre Alocação e Aceite | Tempo decorrido entre a reserva e a confirmação (formatado como duração). |
| CHEGOU | Indica se o paciente já chegou (texto livre, ex.: "Sim"/"Não"). |
| Observação | Observações complementares. |
| Dia | Dia consolidado para filtros (mesma data da ocorrência). |
| Turno | Turno associado (ex.: Manhã/Tarde/Noite). |
| Registro ID | Identificador único gerado pelo sistema. |
| Criado por | E-mail do usuário autenticado que registrou a linha. |
| Criado em | Data/hora em que o sistema gravou a linha. |

### RESERVA NEGADA
Estrutura para reservas recusadas.

| Coluna | Função |
| --- | --- |
| TIPO | Categoria do pedido recusado. |
| FASTMEDIC | Código ou referência do chamado. |
| NOME DO PACIENTE | Identificação do paciente. |
| ORIGEM | Local de origem do paciente. |
| ESPECIALIDADE | Especialidade responsável pelo pedido. |
| JUSTIFICATIVA | Motivo principal da negativa. |
| Data da Reserva | Data originalmente solicitada. |
| Hora da Reserva | Horário originalmente solicitado. |
| Data do Cancelamento da Reserva | Data em que a reserva foi negada/cancelada. |
| Hora do Cancelamento da Reserva | Horário da negativa/cancelamento. |
| Tempo entre Alocação e Cancelamento | Tempo decorrido até a negativa (duração). |
| JUSTIFICATIVA COMPLEMENTAR | Detalhamento adicional do motivo. |
| dia | Dia consolidado para filtros. |
| turno | Turno consolidado para filtros. |
| Registro ID | Identificador único gerado pelo sistema. |
| Criado por | E-mail do usuário autenticado que registrou a linha. |
| Criado em | Data/hora em que o sistema gravou a linha. |

### PLANTÃO ANTERIOR
Usado para rastrear ocorrências herdadas do plantão anterior.

| Coluna | Função |
| --- | --- |
| TIPO | Categoria operacional do registro herdado. |
| Fastmedic | Código ou referência do chamado. |
| Nome do Paciente | Identificação do paciente. |
| Leito Reservado | Leito associado no plantão anterior. |
| Especialidade | Especialidade responsável. |
| Origem | Local de origem do paciente. |
| Status | Situação atual do caso herdado. |
| Data da Reserva | Data original do registro. |
| Hora da Reserva | Horário original do registro. |
| Data de Admissão | Data em que o paciente foi admitido. |
| Hora da Admissão | Horário de admissão. |
| Tempo Decorrido | Tempo entre a reserva e a admissão (duração). |
| CHEGOU | Indica se o paciente chegou. |
| Observação | Observações complementares. |
| dia | Dia consolidado para filtros. |
| turno | Turno consolidado para filtros. |
| Registro ID | Identificador único gerado pelo sistema. |
| Criado por | E-mail do usuário autenticado que registrou a linha. |
| Criado em | Data/hora em que o sistema gravou a linha. |

## Abas auxiliares

### PLANTAO ATUAL
Criada automaticamente caso não exista. Guarda os responsáveis pelo plantão em curso.

| Coluna | Função |
| --- | --- |
| Data do Plantão | Data de referência do plantão. |
| Dia da Semana | Dia da semana correspondente. |
| Turno | Turno do plantão (ex.: Manhã/Tarde/Noite). |
| Médico(a) | Profissional médico responsável. |
| Enfermeiro(a) | Enfermeiro(a) principal. |
| Enfermeiro(a) Sombra | Enfermeiro(a) de apoio. |
| Auxiliar Administrativo | Responsável administrativo. |
| Registrado em | Data/hora em que o plantão foi lançado. |

### REL ENF e REL MED
Relatórios de ocorrências por turno. Se as abas não existirem, o sistema as cria com o cabeçalho abaixo antes do primeiro lançamento.

| Coluna | Função |
| --- | --- |
| Data Registro | Data/hora do lançamento do relatório. |
| Dia | Dia de referência informado no formulário. |
| Turno | Turno relacionado ao relatório. |
| Tipo | "ENF" para enfermagem ou "MED" para médica. |
| Texto | Conteúdo do relatório. |
| Criado em | Data/hora gravada pelo sistema. |

### OBS FIXAS
Linhas fixas exibidas nos relatórios, gerenciadas no WebApp.

| Coluna | Função |
| --- | --- |
| Seção | Bloco ao qual o texto pertence (`enfermagem`, `medica` ou `exames`). |
| Texto | Conteúdo padrão exibido em cada relatório. |
| Atualizado em | Data/hora da última edição. |
| Atualizado por | Usuário que fez a edição (ou "Sistema" se automática). |
