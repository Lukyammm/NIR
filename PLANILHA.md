# Estrutura da Planilha – Ocorrências do Plantão (NIR)

Este sistema de Apps Script + WebApp controla os registros de plantões do NIR e sincroniza cada ação com abas específicas da planilha. A macro `Criar Estrutura Ocorrências` cria ou atualiza automaticamente as abas e cabeçalhos descritos abaixo; os formulários do WebApp gravam sempre na aba "viva" correspondente e, quando há um plantão ativo (`CONFIG_PLANTAO!A2` preenchido), também duplicam o registro nas abas históricas associadas.

## Abas operacionais (edição em tempo real)

### CONFIG_PLANTAO
Linha 2 mantém o plantão ativo. O botão "Abrir Plantão" preenche essa linha; "Encerrar Plantão" move os dados para `HIST_PLANTOES` e limpa `A2:K2`.
- **ID_PLANTAO** – Identificador único do plantão ativo.
- **Data do Plantão** – Data da escala corrente (formato data).
- **Dia da Semana** – Nome do dia (ex.: segunda-feira).
- **Turno** – Manhã/Tarde/Noite ou similar.
- **Médico(a) 1 / Médico(a) 2** – Profissionais médicos escalados.
- **Enfermeiro(a) 1 / Enfermeiro(a) 2** – Enfermeiros de referência.
- **Auxiliar Administrativo** – Responsável administrativo.
- **Abertura (Timestamp)** – Data/hora em que o plantão foi aberto.
- **Encerramento (Timestamp)** – Data/hora de encerramento (preenchida ao encerrar).

### RESERVA_CONFIRMADA
Reservas confirmadas e em andamento.
- **ID** – Identificador do registro.
- **Tipo** – Categoria do atendimento (ex.: internação, transferência).
- **Fastmedic** – Número ou código do chamado Fastmedic.
- **Nome do Paciente** – Paciente vinculado à reserva.
- **Leito Reservado** – Leito destinado ao paciente.
- **Especialidade** – Especialidade responsável.
- **Origem** – Origem da solicitação (ex.: setor, unidade externa).
- **Status** – Situação atual da reserva.
- **Data da Reserva / Hora da Reserva** – Momento da solicitação.
- **Data da Confirmação / Hora da Confirmação** – Momento do aceite.
- **Tempo entre Alocação e Aceite** – Duração entre abertura e confirmação.
- **Chegou** – Indica se o paciente já chegou.
- **Observação** – Comentários adicionais.

### PROCEDIMENTO_VASCULAR
Estrutura idêntica a `RESERVA_CONFIRMADA`, usada para procedimentos vasculares.
- Mesmos campos: **ID**, **Tipo**, **Fastmedic**, **Nome do Paciente**, **Leito Reservado**, **Especialidade**, **Origem**, **Status**, **Data/Hora da Reserva**, **Data/Hora da Confirmação**, **Tempo entre Alocação e Aceite**, **Chegou**, **Observação**.

### RESERVA_NEGADA
Registros de solicitações recusadas.
- **ID** – Identificador do registro.
- **Tipo** – Categoria do atendimento.
- **Fastmedic** – Número/código do chamado.
- **Nome do Paciente** – Nome informado na solicitação.
- **Origem** – Local de origem.
- **Especialidade** – Especialidade solicitada.
- **Justificativa** – Motivo da negativa.
- **Data da Reserva / Hora da Reserva** – Momento da abertura.
- **Data do Cancelamento / Hora do Cancelamento** – Quando a negativa foi registrada.
- **Tempo entre Alocação e Cancelamento** – Intervalo entre abertura e cancelamento.

### PLANTAO_ANTERIOR
Fila de registros herdados de plantões anteriores que ainda demandam acompanhamento.
- **ID** – Identificador do registro.
- **Tipo** – Tipo de atendimento.
- **Fastmedic** – Número/código do chamado.
- **Nome do Paciente** – Paciente associado.
- **Leito Reservado** – Leito em uso ou previsto.
- **Especialidade** – Especialidade responsável.
- **Origem** – Origem do caso.
- **Status** – Situação atual.
- **Data/Hora da Reserva** – Momento da reserva original.
- **Data/Hora de Admissão** – Momento de entrada/admissão.
- **Tempo Decorrido** – Tempo total desde a reserva.
- **Chegou** – Indica chegada do paciente.
- **Observação** – Observações gerais.

### BLOQUEADOS_MANUTENCAO
Leitos bloqueados por manutenção.
- **ID** – Identificador do bloqueio.
- **Leito** – Número do leito.
- **Unidade** – Unidade/setor do leito.
- **Manutenção Acionada** – Breve descrição ou protocolo de acionamento.
- **Data de Início** – Quando o bloqueio começou.
- **Previsão de Reparo** – Data prevista de liberação.
- **Observação** – Observações sobre o reparo.

### BLOQUEADOS_ISOLAMENTO
Leitos bloqueados por isolamento.
- **ID** – Identificador do bloqueio.
- **Unidade** – Unidade/setor.
- **Leito** – Número do leito.
- **Paciente** – Nome do paciente em isolamento.
- **Tipo de Isolamento** – Tipo (ex.: respiratório, de contato).
- **Patologia** – Diagnóstico associado.
- **Início do Isolamento** – Data de início.
- **Tempo Previsto** – Duração prevista.
- **Observação** – Observações clínicas ou logísticas.

## Abas históricas (preenchimento automático)
Ao encerrar um plantão ou inserir registros com um plantão ativo, os dados são duplicados aqui com o ID do plantão para rastreabilidade.

### HIST_PLANTOES
- **ID_PLANTAO** – ID do plantão encerrado.
- **Data** – Data do plantão.
- **Dia** – Dia da semana.
- **Turno** – Turno do plantão.
- **Medico1 / Medico2** – Médicos escalados.
- **Enf1 / Enf2** – Enfermeiros.
- **Aux** – Auxiliar administrativo.
- **Hora Abertura / Hora Encerramento** – Carimbos de horário.
- **Usuario_Encerramento** – Usuário que executou o encerramento.

### HIST_RESERVA_CONFIRMADA e HIST_PROCEDIMENTO_VASCULAR
Mesmas colunas das abas operacionais correspondentes, antecedidas por **PLANTAO_ID** para vincular ao plantão ativo no momento do registro.

### HIST_RESERVA_NEGADA
- **PLANTAO_ID**, seguido das colunas: **ID**, **Tipo**, **Fastmedic**, **Nome**, **Origem**, **Especialidade**, **Justificativa**, **Data Reserva**, **Hora Reserva**, **Data Cancelamento**, **Hora Cancelamento**, **Tempo Cancelamento**.

### HIST_MANUTENCAO
- **PLANTAO_ID**, **ID**, **Leito**, **Unidade**, **Manutenção Acionada**, **Data Início**, **Previsão**, **Observação**, **Encerrado em** (data/hora de liberação).

### HIST_ISOLAMENTO
- **PLANTAO_ID**, **ID**, **Unidade**, **Leito**, **Paciente**, **Isolamento**, **Patologia**, **Início**, **Tempo Previsto**, **Observação**, **Encerrado em**.

### LOG_NIR
- **ID_EVENTO** – ID único do evento no log.
- **ID_REGISTRO** – ID do registro afetado.
- **MODULO** – Módulo origem (plantão, reserva, manutenção etc.).
- **TIPO_EVENTO** – Tipo de evento (inserção, encerramento, recuperação de ID).
- **USUARIO** – E-mail do usuário executando a ação.
- **DATA_HORA** – Timestamp do evento.
- **OBSERVACAO** – Observações sobre a ação.

## Observações operacionais
1. Sempre mantenha `CONFIG_PLANTAO!A2` preenchido enquanto o plantão estiver ativo; os formulários impedem novo plantão se já existir ID na linha 2.
2. O WebApp "Ocorrências do Plantão – NIR" lê essas abas e permite inserir, editar ou excluir registros; exclusões removem apenas das abas vivas e não retroagem às abas históricas já registradas.
3. Se o plantão estiver preenchido mas sem ID, o script gera um novo ID automaticamente ao carregar a interface para evitar perda de rastreio nos históricos.
4. Use a ação "Encerrar Plantão Atual" para mover a linha ativa para `HIST_PLANTOES` e limpar a área de trabalho antes do próximo plantão.
