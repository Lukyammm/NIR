function criarEstruturaNIR() {
  assertAuthorized_();
  const ss = getSpreadsheet_();

  const abasVivas = {
    "CONFIG_PLANTAO": [
      "ID_PLANTAO","Data do Plantão","Dia da Semana","Turno",
      "Médico(a) 1","Médico(a) 2",
      "Enfermeiro(a) 1","Enfermeiro(a) 2",
      "Auxiliar Administrativo",
      "Abertura (Timestamp)","Encerramento (Timestamp)"
    ],

    "RESERVA_CONFIRMADA": [
      "ID","Tipo","Fastmedic","Nome do Paciente","Leito Reservado",
      "Especialidade","Origem","Status",
      "Data da Reserva","Hora da Reserva",
      "Data da Confirmação","Hora da Confirmação",
      "Tempo entre Alocação e Aceite","Chegou","Observação"
    ],

    "PROCEDIMENTO_VASCULAR": [
      "ID","Tipo","Fastmedic","Nome do Paciente","Leito Reservado",
      "Especialidade","Origem","Status",
      "Data da Reserva","Hora da Reserva",
      "Data da Confirmação","Hora da Confirmação",
      "Tempo entre Alocação e Aceite","Chegou","Observação"
    ],

    "RESERVA_NEGADA": [
      "ID","Tipo","Fastmedic","Nome do Paciente","Origem",
      "Especialidade","Justificativa",
      "Data da Reserva","Hora da Reserva",
      "Data do Cancelamento","Hora do Cancelamento",
      "Tempo entre Alocação e Cancelamento"
    ],

    "PLANTAO_ANTERIOR": [
      "ID","Tipo","Fastmedic","Nome do Paciente","Leito Reservado",
      "Especialidade","Origem","Status",
      "Data da Reserva","Hora da Reserva",
      "Data de Admissão","Hora de Admissão",
      "Tempo Decorrido","Chegou","Observação"
    ],

    "BLOQUEADOS_MANUTENCAO": [
      "ID","Leito","Unidade","Manutenção Acionada","Data de Início","Previsão de Reparo","Observação"
    ],

    "BLOQUEADOS_ISOLAMENTO": [
      "ID","Unidade","Leito","Paciente",
      "Tipo de Isolamento","Patologia",
      "Início do Isolamento","Tempo Previsto","Observação"
    ]
  };

  const abasHistoricos = {
    "HIST_PLANTOES": [
      "ID_PLANTAO","Data","Dia","Turno",
      "Medico1","Medico2",
      "Enf1","Enf2",
      "Aux","Hora Abertura","Hora Encerramento","Usuario_Encerramento"
    ],

    "HIST_RESERVA_CONFIRMADA": [
      "PLANTAO_ID","ID","Tipo","Fastmedic","Nome do Paciente","Leito Reservado",
      "Especialidade","Origem","Status",
      "Data Reserva","Hora Reserva",
      "Data Confirmação","Hora Confirmação",
      "Tempo Aceite","Chegou","Observação"
    ],

    "HIST_PROCEDIMENTO_VASCULAR": [
      "PLANTAO_ID","ID","Tipo","Fastmedic","Nome do Paciente","Leito Reservado",
      "Especialidade","Origem","Status",
      "Data Reserva","Hora Reserva",
      "Data Confirmação","Hora Confirmação",
      "Tempo Aceite","Chegou","Observação"
    ],

    "HIST_RESERVA_NEGADA": [
      "PLANTAO_ID","ID","Tipo","Fastmedic","Nome","Origem","Especialidade",
      "Justificativa","Data Reserva","Hora Reserva",
      "Data Cancelamento","Hora Cancelamento","Tempo Cancelamento"
    ],

    "HIST_MANUTENCAO": [
      "PLANTAO_ID","ID","Leito","Unidade","Manutenção Acionada","Data Início","Previsão","Observação","Encerrado em"
    ],

    "HIST_ISOLAMENTO": [
      "PLANTAO_ID","ID","Unidade","Leito","Paciente",
      "Isolamento","Patologia","Início","Tempo Previsto","Observação","Encerrado em"
    ],

    "LOG_NIR": [
      "ID_EVENTO","ID_REGISTRO","MODULO","TIPO_EVENTO",
      "USUARIO","DATA_HORA","OBSERVACAO"
    ]
  };

  criarAbas_(ss, abasVivas);
  criarAbas_(ss, abasHistoricos);

  SpreadsheetApp.getUi().alert("Estrutura da aba Ocorrências NIR criada/atualizada.");
}

function criarAbas_(ss, estrutura) {
  for (let nome in estrutura) {
    let aba = ss.getSheetByName(nome);
    const header = estrutura[nome];
    if (!aba) {
      aba = ss.insertSheet(nome);
      aba.appendRow(header);
    } else {
      const range = aba.getRange(1, 1, 1, header.length);
      range.setValues([header]);
    }
  }
}
