/*********** CONFIGURAÇÃO BÁSICA ***********/

// Nome exato das abas que já existem na planilha
const NIR_SHEETS = {
  'RESERVA CONFIRMADA': 'RESERVA CONFIRMADA',
  'PROCEDIMENTO CONFIRMADO': 'PROCEDIMENTO CONFIRMADO',
  'RESERVA NEGADA': 'RESERVA NEGADA',
  'PLANTÃO ANTERIOR': 'PLANTÃO ANTERIOR'
};

// Abas (serão criadas se não existirem) para relatórios por turno
const REL_ENF_SHEET = 'REL ENF';
const REL_MED_SHEET = 'REL MED';


/**
 * Abre o WebApp (usa o arquivo index.html)
 */
function doGet(e) {
  return HtmlService
    .createHtmlOutputFromFile('index')
    .setTitle('NIR – Ocorrências');
}

/**
 * Utilitário: pega planilha ativa
 */
function getSS() {
  return SpreadsheetApp.getActiveSpreadsheet();
}

/**
 * Normaliza texto (para comparar cabeçalhos sem se importar com maiúscula/minúscula)
 */
function normalize_(value) {
  return String(value || '').trim().toLowerCase();
}

/**
 * Pega/Cria aba
 */
function getOrCreateSheet_(name) {
  var ss = getSS();
  var sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
  }
  return sheet;
}

/**
 * Retorna a linha de cabeçalhos já normalizada
 */
function getHeaderNormalized_(sheet) {
  var lastCol = sheet.getLastColumn();
  if (lastCol === 0) return [];
  var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  return headers.map(function (h) { return normalize_(h); });
}

/**
 * Mapeamento de campos do formulário → cabeçalhos de cada aba
 * (usa o nome exato das colunas da sua planilha)
 */
function getFieldMappings_() {
  return {
    'RESERVA CONFIRMADA': {
      tipo: 'TIPO',
      fastmedic: 'Fastmedic',
      nomePaciente: 'Nome do Paciente',
      leitoReservado: 'Leito Reservado',
      especialidade: 'Especialidade',
      origem: 'Origem',
      status: 'Status',
      dataReserva: 'Data da Reserva',
      horaReserva: 'Hora da Reserva',
      dataConfirmacao: 'Data da Confirmação da Reserva',
      horaConfirmacao: 'Hora da Confirmação da Reserva',
      tempoEntre: 'Tempo entre Alocação e Aceite',
      chegou: 'CHEGOU',
      observacao: 'Observação',
      dia: 'Dia',
      turno: 'Turno'
    },
    'PROCEDIMENTO CONFIRMADO': {
      tipo: 'TIPO',
      fastmedic: 'Fastmedic',
      nomePaciente: 'Nome do Paciente',
      leitoReservado: 'Leito Reservado',
      especialidade: 'Especialidade',
      origem: 'Origem',
      status: 'Status',
      dataReserva: 'Data da Reserva',
      horaReserva: 'Hora da Reserva',
      dataConfirmacao: 'Data da Confirmação da Reserva',
      horaConfirmacao: 'Hora da Confirmação da Reserva',
      tempoEntre: 'Tempo entre Alocação e Aceite',
      observacao: 'Observação',
      dia: 'dia',
      turno: 'turno'
    },
    'RESERVA NEGADA': {
      tipo: 'TIPO',
      fastmedic: 'FASTMEDIC',
      nomePaciente: 'NOME DO PACIENTE',
      origem: 'ORIGEM',
      especialidade: 'ESPECIALIDADE',
      justificativa: 'JUSTIFICATIVA',
      dataReserva: 'Data da Reserva',
      horaReserva: 'Hora da Reserva',
      dataCancelamento: 'Data do Cancelamento da Reserva',
      horaCancelamento: 'Hora do Cancelamento da Reserva',
      tempoEntre: 'Tempo entre Alocação e Cancelamento',
      justificativaComplementar: 'JUSTIFICATIVA COMPLEMENTAR',
      dia: 'dia',
      turno: 'turno'
    },
    'PLANTÃO ANTERIOR': {
      tipo: 'TIPO',
      fastmedic: 'Fastmedic',
      nomePaciente: 'Nome do Paciente',
      leitoReservado: 'Leito Reservado',
      especialidade: 'Especialidade',
      origem: 'Origem',
      status: 'Status',
      dataReserva: 'Data da Reserva',
      horaReserva: 'Hora da Reserva',
      dataAdmissao: 'Data de Admissão',
      horaAdmissao: 'Hora da Admissão',
      tempoDecorrido: 'Tempo Decorrido',
      chegou: 'CHEGOU',
      observacao: 'Observação',
      dia: 'dia',
      turno: 'turno'
    }
  };
}

/**
 * Lança uma nova linha na aba correta, baseado na categoria e nos campos
 */
function appendOccurrence_(category, data) {
  var sheetName = NIR_SHEETS[category];
  if (!sheetName) {
    throw new Error('Categoria inválida: ' + category);
  }

  var sheet = getOrCreateSheet_(sheetName);
  var headerNorm = getHeaderNormalized_(sheet);

  if (!headerNorm.length) {
    throw new Error('A aba "' + sheetName + '" precisa ter o cabeçalho preenchido na Linha 1.');
  }

  var fieldMappings = getFieldMappings_();
  var mapping = fieldMappings[category] || {};

  var lastCol = headerNorm.length;
  var newRow = new Array(lastCol).fill('');

  Object.keys(mapping).forEach(function (fieldKey) {
    var headerLabel = mapping[fieldKey];
    var idx = headerNorm.indexOf(normalize_(headerLabel));
    var value = data[fieldKey];

    if (idx > -1 && value !== null && value !== undefined && value !== '') {
      newRow[idx] = value;
    }
  });

  sheet.appendRow(newRow);
}

/**
 * Função chamada pelo front-end para salvar uma ocorrência
 */
function saveOccurrence(payload) {
  if (!payload) {
    throw new Error('Dados não recebidos.');
  }

  var category = payload.category;
  var fields = payload.fields || {};

  if (!category || !NIR_SHEETS[category]) {
    throw new Error('Tipo de ocorrência inválido ou ausente.');
  }

  appendOccurrence_(category, fields);

  return {
    ok: true,
    message: 'Ocorrência salva em "' + category + '".'
  };
}

/**
 * Monta o "dash" com os últimos 3 turnos encontrados
 * (baseado nas colunas Dia/dia e Turno/turno das 4 abas)
 */
function getDashboardData() {
  var ss = getSS();
  var shiftMap = {};

  Object.keys(NIR_SHEETS).forEach(function (categoryName) {
    var sheetName = NIR_SHEETS[categoryName];
    var sh = ss.getSheetByName(sheetName);
    if (!sh) return;

    var lastRow = sh.getLastRow();
    var lastCol = sh.getLastColumn();
    if (lastRow <= 1 || lastCol === 0) return;

    var headersNorm = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(function (h) {
      return normalize_(h);
    });

    var diaIdx = headersNorm.indexOf('dia');
    var turnoIdx = headersNorm.indexOf('turno');

    if (diaIdx === -1 && turnoIdx === -1) {
      return;
    }

    var dataValues = sh.getRange(2, 1, lastRow - 1, lastCol).getValues();

    dataValues.forEach(function (row, i) {
      var dia = diaIdx > -1 ? row[diaIdx] : '';
      var turno = turnoIdx > -1 ? row[turnoIdx] : '';

      if (!dia && !turno) return;

      var key = String(dia) + ' | ' + String(turno);

      if (!shiftMap[key]) {
        shiftMap[key] = {
          key: key,
          dia: dia,
          turno: turno,
          counts: {
            'RESERVA CONFIRMADA': 0,
            'PROCEDIMENTO CONFIRMADO': 0,
            'RESERVA NEGADA': 0,
            'PLANTÃO ANTERIOR': 0
          },
          lastRowIndex: 0
        };
      }

      shiftMap[key].counts[categoryName] =
        (shiftMap[key].counts[categoryName] || 0) + 1;

      var globalRowIndex = i + 2; // linha real na aba
      if (globalRowIndex > shiftMap[key].lastRowIndex) {
        shiftMap[key].lastRowIndex = globalRowIndex;
      }
    });
  });

  var shifts = Object.keys(shiftMap).map(function (k) {
    return shiftMap[k];
  });

  // Ordena pelos "últimos registros" (mais recentes primeiro)
  shifts.sort(function (a, b) {
    return b.lastRowIndex - a.lastRowIndex;
  });

  var top3 = shifts.slice(0, 3).map(function (s) {
    var total = Object.keys(s.counts).reduce(function (sum, k) {
      return sum + (s.counts[k] || 0);
    }, 0);

    return {
      key: s.key,
      dia: s.dia,
      turno: s.turno,
      counts: s.counts,
      total: total
    };
  });

  return {
    shifts: top3
  };
}

/**
 * Salva Relatório de Ocorrências de Enfermagem / Médicas por turno
 * tipo: 'ENF' ou 'MED'
 */
function saveRelatorio(payload) {
  if (!payload) {
    throw new Error('Dados do relatório não recebidos.');
  }

  var tipo = payload.tipo;   // ENF ou MED
  var texto = payload.texto;
  var dia = payload.dia || '';
  var turno = payload.turno || '';

  if (!tipo || !texto) {
    throw new Error('Tipo de relatório e texto são obrigatórios.');
  }

  var sheetName = tipo === 'ENF' ? REL_ENF_SHEET : REL_MED_SHEET;
  var sheet = getOrCreateSheet_(sheetName);

  // Se a aba estiver vazia, cria cabeçalho padrão
  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, 6).setValues([[
      'Data Registro',
      'Dia',
      'Turno',
      'Tipo',
      'Texto',
      'Criado em'
    ]]);
  }

  var now = new Date();

  sheet.appendRow([
    now,
    dia,
    turno,
    tipo,
    texto,
    now
  ]);

  return {
    ok: true,
    message: 'Relatório salvo com sucesso.'
  };
}

/**
 * Retorna últimos relatórios (misturando enfermagem e médico) para mostrar no painel
 */
function getRelatoriosRecentes(limit) {
  limit = limit || 10;
  var ss = getSS();
  var out = [];

  [
    { tipo: 'ENF', name: REL_ENF_SHEET },
    { tipo: 'MED', name: REL_MED_SHEET }
  ].forEach(function (conf) {
    var sh = ss.getSheetByName(conf.name);
    if (!sh) return;

    var lastRow = sh.getLastRow();
    var lastCol = sh.getLastColumn();
    if (lastRow <= 1 || lastCol === 0) return;

    var numRows = Math.min(lastRow - 1, limit * 2); // pega um pouco mais para misturar
    var values = sh.getRange(lastRow - numRows + 1, 1, numRows, lastCol).getValues();

    values.forEach(function (r) {
      out.push({
        tipo: conf.tipo,
        data: r[0],
        dia: r[1],
        turno: r[2],
        texto: r[4]
      });
    });
  });

  out.sort(function (a, b) {
    var ta = a.data instanceof Date ? a.data.getTime() : 0;
    var tb = b.data instanceof Date ? b.data.getTime() : 0;
    return tb - ta;
  });

  return out.slice(0, limit);
}
