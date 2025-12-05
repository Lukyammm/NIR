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
const SHIFT_SHEET = 'PLANTAO ATUAL';
const FIXED_NOTES_SHEET = 'OBS FIXAS';
const FIXED_NOTE_SECTIONS = ['enfermagem', 'medica', 'exames'];


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

function getShiftSheet_() {
  var sheet = getOrCreateSheet_(SHIFT_SHEET);
  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, 8).setValues([[
      'Data do Plantão',
      'Dia da Semana',
      'Turno',
      'Médico(a)',
      'Enfermeiro(a)',
      'Enfermeiro(a) Sombra',
      'Auxiliar Administrativo',
      'Registrado em'
    ]]);
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

function sanitizeValue_(value) {
  if (value === null || value === undefined) return '';
  if (value instanceof Date) return value;

  var str = String(value).trim();
  if (!str) return '';

  if (/^[=+\-@]/.test(str)) {
    return "'" + str; // evita fórmulas
  }

  return str;
}

function getDefaultFixedNotes_() {
  return {
    enfermagem: [
      'OBSERVAÇÕES:',
      '📍 ORTOPEDIA: 715 A 721 -  ENFERMARIA LIBERADA PARA LEITOS ECTÓPICOS',
      '📍VASCULAR:  LEITOS DE ISOLAMENTO 600.1 E 600.2 SÃO LEITOS DE ISOLAMENTO DA VASCULAR, NÃO SÃO ECTÓPICOS.',
      '📍BARIÁTRICA:  LEITOS 611.01 AO 612.02 SÃO PARA PACIENTES DA BARIÁTRICA',
      '📍 CENTRO DE IMAGEM: HEMODINAMICA - HORARIOS DE FUNCIONAMENTO: SEGUNDA A SEXTA-FEIRA / NOS FINAIS DE SEMANA NÃO FUNCIONA.',
      '📍 CATETERISMOS MARCADOS NO HM - TODOS ÀS 07 HORAS',
      '📍ATENÇÃO!',
      ' Perfis que não são nosso (VASCULAR) sinalizado por Dra Grazi: Doença carotídea, Isquemia Mesentérica, Aneurismas de aorta, Trauma vascular e Embolizações'
    ].join('\n'),
    medica: [
      'OBSERVAÇÕES:',
      '📍 ORTOPEDIA: 715 A 721 -  ENFERMARIA LIBERADA PARA LEITOS ECTÓPICOS',
      '📍VASCULAR:  LEITOS DE ISOLAMENTO 600.1 E 600.2 SÃO LEITOS DE ISOLAMENTO DA VASCULAR, NÃO SÃO ECTÓPICOS.',
      '📍BARIÁTRICA:  LEITOS 611.01 AO 612.02 SÃO PARA PACIENTES DA BARIÁTRICA',
      '📍 CENTRO DE IMAGEM:HEMODINAMICA - HORARIOS DE FUNCIONAMENTO: SEGUNDA A SEXTA-FEIRA / NOS FINAIS DE SEMANA NÃO FUNCIONA.',
      '📍ATENÇÃO!',
      ' Perfis que não são nosso (VASCULAR) sinalizado por Dra Grazi: Doença carotídea, Isquemia Mesentérica, Aneurismas de aorta, Trauma vascular e Embolizações',
      'FOI AUTORIZADO PELA DIREÇÃO (TARDE DE 06/09/2025), A PEDIDO DA CRL SESA, A REALIZAÇÃO DE TCs DA REDE, VISTO QUE TC DO HSJ ESTÁ QUEBRADA.',
      'DR. JURANDIR SOLICITA QUE OS PACIENTES PARA ARTERIOGRAFIA SEJAM ALOCADOS PARA O IJF - CRL CIENTE'
    ].join('\n'),
    exames: [
      '📍EXAMES:',
      'VER PLANILHA  AGENDAMENTO EXAMES EXTERNOS (POR FAVOR, ATUALIZAR A CADA  AGENDAMENTO)',
      '  Agendamento Exames H. Externos (RNM e TC);',
      '🚨  AGENDADOS - 18/10/2025 -7H00',
      '3687301        KALEBE BARBOSA TELES - HGWA - RNM CERVICAL E CRÂNIO (CRIANÇA 7 ANOS)',
      '3685751        JOAO PAULO ALVES DE SOUZA - MULUNGU - TC DE CRÂNIO S/C',
      '3685356        MARIA ALEUDA DE SOUSA - UPA HORIZONTE - TC DE CRÂNIO  S/C',
      '3686927        FRANCISCA ANA DE SOUSA MOURA - ITAITINGA - TC DE CRÂNIO  S/C',
      '3683819        MANOEL DOMINGOS DE OLIVEIRA - UPA AUTRAN NUNES - TC DE CRÂNIO S/C'
    ].join('\n')
  };
}

function getRequiredFieldsConfig_() {
  return {
    'RESERVA CONFIRMADA': ['tipo', 'fastmedic', 'nomePaciente', 'especialidade', 'origem', 'status', 'dataReserva', 'horaReserva', 'dia', 'turno'],
    'PROCEDIMENTO CONFIRMADO': ['tipo', 'fastmedic', 'nomePaciente', 'especialidade', 'origem', 'status', 'dataReserva', 'horaReserva', 'dia', 'turno'],
    'RESERVA NEGADA': ['tipo', 'fastmedic', 'nomePaciente', 'origem', 'especialidade', 'justificativa', 'dataReserva', 'horaReserva', 'dia', 'turno'],
    'PLANTÃO ANTERIOR': ['tipo', 'fastmedic', 'nomePaciente', 'especialidade', 'origem', 'status', 'dataReserva', 'horaReserva', 'dia', 'turno']
  };
}

function validateRequiredFields_(category, fields, mapping) {
  var required = getRequiredFieldsConfig_()[category] || [];
  var missing = [];

  required.forEach(function (key) {
    var value = sanitizeValue_(fields[key]);
    if (value === '') {
      missing.push(mapping[key] || key);
    }
  });

  return missing;
}

function buildOccurrenceFingerprint_(category, data, mapping) {
  var keyParts = [normalize_(category)];
  Object.keys(mapping).forEach(function (fieldKey) {
    keyParts.push(normalize_(sanitizeValue_(data[fieldKey])));
  });
  var digest = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, keyParts.join('||'));
  return Utilities.base64Encode(digest);
}

function getAuditInfo_(category, data, mapping) {
  var createdBy = '';
  try {
    createdBy = Session.getActiveUser().getEmail() || '';
  } catch (err) {
    createdBy = '';
  }

  var createdAt = new Date();
  var recordId = buildOccurrenceFingerprint_(category, data, mapping);

  return {
    createdBy: createdBy || 'Anônimo',
    createdAt: createdAt,
    recordId: recordId
  };
}

function clearDashboardCache_() {
  try {
    CacheService.getScriptCache().remove('dashboardData');
  } catch (err) {
    // ignora erros de cache
  }
}

function ensureAuditColumns_(sheet) {
  var headersNorm = getHeaderNormalized_(sheet);
  if (!headersNorm.length) return headersNorm;

  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var additions = [];

  if (headersNorm.indexOf('registro id') === -1) {
    additions.push('Registro ID');
  }
  if (headersNorm.indexOf('criado por') === -1) {
    additions.push('Criado por');
  }
  if (headersNorm.indexOf('criado em') === -1) {
    additions.push('Criado em');
  }

  if (additions.length) {
    sheet.insertColumnsAfter(sheet.getLastColumn() || 1, additions.length);
    sheet.getRange(1, headers.length + 1, 1, additions.length).setValues([additions]);
  }

  return getHeaderNormalized_(sheet);
}

function getFixedNotesSheet_() {
  var sheet = getOrCreateSheet_(FIXED_NOTES_SHEET);
  var lastRow = sheet.getLastRow();

  if (lastRow === 0) {
    sheet.getRange(1, 1, 1, 4).setValues([[
      'Seção',
      'Texto',
      'Atualizado em',
      'Atualizado por'
    ]]);
    lastRow = 1;
  }

  var defaults = getDefaultFixedNotes_();

  var rows = Math.max(lastRow - 1, 0);
  var existing = {};
  if (rows > 0) {
    var values = sheet.getRange(2, 1, rows, 2).getValues();
    values.forEach(function (r) {
      var sec = normalize_(r[0]);
      if (sec) existing[sec] = true;
    });
  }

  var toAppend = [];
  var now = new Date();
  var user = '';
  try {
    user = Session.getActiveUser().getEmail() || '';
  } catch (err) {
    user = '';
  }

  FIXED_NOTE_SECTIONS.forEach(function (sec) {
    if (!existing[normalize_(sec)]) {
      toAppend.push([sec, defaults[sec] || '', now, user || 'Sistema']);
    }
  });

  if (toAppend.length) {
    sheet.getRange(lastRow + 1, 1, toAppend.length, 4).setValues(toAppend);
  }

  return sheet;
}

function getFixedNotes() {
  var sheet = getFixedNotesSheet_();
  var defaults = getDefaultFixedNotes_();
  var notes = Object.assign({}, defaults);

  var lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    var values = sheet.getRange(2, 1, lastRow - 1, 2).getValues();
    values.forEach(function (r) {
      var sec = normalize_(r[0]);
      if (sec && notes.hasOwnProperty(sec)) {
        notes[sec] = String(r[1] || '').trim();
      }
    });
  }

  return { notes: notes };
}

function saveFixedNotes(payload) {
  if (!payload) {
    throw new Error('Nenhum bloco fixo recebido para salvar.');
  }

  var sheet = getFixedNotesSheet_();

  var incomingKeys = Object.keys(payload).filter(function (k) {
    return FIXED_NOTE_SECTIONS.indexOf(k) !== -1;
  });

  if (!incomingKeys.length) {
    throw new Error('Nenhuma seção válida enviada.');
  }

  var lastRow = sheet.getLastRow();
  var existingRows = {};
  if (lastRow > 1) {
    var secValues = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
    secValues.forEach(function (r, idx) {
      var sec = normalize_(r[0]);
      if (sec) existingRows[sec] = idx + 2; // linha real na planilha
    });
  }

  var now = new Date();
  var user = '';
  try {
    user = Session.getActiveUser().getEmail() || '';
  } catch (err) {
    user = '';
  }

  incomingKeys.forEach(function (key) {
    var value = sanitizeValue_(payload[key]);
    var normalized = normalize_(key);
    var rowIndex = existingRows[normalized];
    var rowValues = [key, value || '', now, user || 'Sistema'];

    if (rowIndex) {
      sheet.getRange(rowIndex, 1, 1, 4).setValues([rowValues]);
    } else {
      sheet.appendRow(rowValues);
    }
  });

  return getFixedNotes();
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
  var headerNorm = ensureAuditColumns_(sheet);

  if (!headerNorm.length) {
    throw new Error('A aba "' + sheetName + '" precisa ter o cabeçalho preenchido na Linha 1.');
  }

  var fieldMappings = getFieldMappings_();
  var mapping = fieldMappings[category] || {};

  var lastCol = headerNorm.length;
  var newRow = new Array(lastCol).fill('');

  var auditInfo = getAuditInfo_(category, data, mapping);
  var recordIdIdx = headerNorm.indexOf('registro id');
  if (recordIdIdx > -1 && sheet.getLastRow() > 1) {
    var existing = sheet.getRange(2, recordIdIdx + 1, sheet.getLastRow() - 1, 1).getValues();
    var hasDuplicate = existing.some(function (row) { return row[0] === auditInfo.recordId; });
    if (hasDuplicate) {
      throw new Error('Ocorrência duplicada detectada.');
    }
  }

  Object.keys(mapping).forEach(function (fieldKey) {
    var headerLabel = mapping[fieldKey];
    var idx = headerNorm.indexOf(normalize_(headerLabel));
    var value = sanitizeValue_(data[fieldKey]);

    if (idx > -1 && value !== null && value !== undefined && value !== '') {
      newRow[idx] = value;
    }
  });

  var createdByIdx = headerNorm.indexOf('criado por');
  if (createdByIdx > -1) {
    newRow[createdByIdx] = auditInfo.createdBy;
  }

  var createdAtIdx = headerNorm.indexOf('criado em');
  if (createdAtIdx > -1) {
    newRow[createdAtIdx] = auditInfo.createdAt;
  }

  if (recordIdIdx > -1) {
    newRow[recordIdIdx] = auditInfo.recordId;
  }

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

  var fieldMappings = getFieldMappings_();
  var mapping = fieldMappings[category] || {};

  var missing = validateRequiredFields_(category, fields, mapping);
  if (missing.length) {
    throw new Error('Campos obrigatórios ausentes: ' + missing.join(', '));
  }

  appendOccurrence_(category, fields);

  clearDashboardCache_();

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
  var cache = CacheService.getScriptCache();
  var cached = cache.get('dashboardData');
  if (cached) {
    try {
      return JSON.parse(cached);
    } catch (err) {
      // segue para recomputar
    }
  }

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

  var result = {
    shifts: top3
  };

  try {
    cache.put('dashboardData', JSON.stringify(result), 300);
  } catch (err) {
    // ignora cache falho
  }

  return result;
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

/**
 * Retorna informações do plantão atual (última linha da aba de plantão)
 */
function getShiftInfo() {
  var sheet = getShiftSheet_();
  var lastRow = sheet.getLastRow();
  if (lastRow <= 1) return {};

  var values = sheet.getRange(lastRow, 1, 1, 8).getValues()[0];
  return {
    dataPlantao: values[0],
    diaSemana: values[1],
    turno: values[2],
    medico: values[3],
    enfermeiro: values[4],
    enfermeiroSombra: values[5],
    auxiliar: values[6],
    registradoEm: values[7]
  };
}

/**
 * Salva/atualiza o registro do plantão atual
 */
function saveShift(payload) {
  if (!payload) {
    throw new Error('Dados do plantão não recebidos.');
  }

  if (!payload.dataPlantao || !payload.turno) {
    throw new Error('Data do plantão e turno são obrigatórios.');
  }

  var dayName = payload.diaSemana;
  if (!dayName) {
    var parsedDate = payload.dataPlantao instanceof Date
      ? payload.dataPlantao
      : new Date(payload.dataPlantao);
    if (!isNaN(parsedDate.getTime())) {
      var weekdays = ['Domingo', 'Segunda-feira', 'Terça-feira', 'Quarta-feira', 'Quinta-feira', 'Sexta-feira', 'Sábado'];
      dayName = weekdays[parsedDate.getDay()] || '';
    }
  }

  var sheet = getShiftSheet_();
  var now = new Date();
  sheet.appendRow([
    payload.dataPlantao,
    dayName || payload.diaSemana || '',
    payload.turno || '',
    payload.medico || '',
    payload.enfermeiro || '',
    payload.enfermeiroSombra || '',
    payload.auxiliar || '',
    now
  ]);

  return {
    ok: true,
    message: 'Plantão salvo/atualizado com sucesso.'
  };
}

/**
 * Retorna últimos 2 dias com as 2 ocorrências mais recentes de cada
 */
function getLatestOccurrencesDays(limitDays, perDayLimit) {
  limitDays = limitDays || 2;
  perDayLimit = perDayLimit || 2;

  var occurrences = collectOccurrences_();
  if (!occurrences.length) {
    return { days: [] };
  }

  var grouped = {};
  occurrences.forEach(function (occ) {
    var diaKey = String(occ.dia || 'Sem dia');
    if (!grouped[diaKey]) {
      grouped[diaKey] = [];
    }
    grouped[diaKey].push(occ);
  });

  var days = Object.keys(grouped).map(function (key) {
    var list = grouped[key];
    list.sort(function (a, b) {
      return (b.order || 0) - (a.order || 0);
    });

    var turnoResumo = list.find(function (i) { return i.turno; });

    return {
      dia: key,
      turnoResumo: turnoResumo ? turnoResumo.turno : '',
      occurrences: list.slice(0, perDayLimit).map(mapOccurrenceForClient_)
    };
  });

  days.sort(function (a, b) {
    var da = parseDateFlexible_(a.dia);
    var db = parseDateFlexible_(b.dia);
    if (da && db) {
      return db.getTime() - da.getTime();
    }
    return (grouped[b.dia][0].order || 0) - (grouped[a.dia][0].order || 0);
  });

  return { days: days.slice(0, limitDays) };
}

/**
 * Busca ocorrências por texto livre
 */
function searchOccurrences(query) {
  if (!query || String(query).trim().length < 2) {
    return [];
  }

  var term = String(query).toLowerCase();
  var occurrences = collectOccurrences_();

  var filtered = occurrences.filter(function (occ) {
    var haystack = [
      occ.category,
      occ.paciente,
      occ.especialidade,
      occ.origem,
      occ.status,
      occ.dia,
      occ.turno,
      occ.observacao
    ].join(' ').toLowerCase();
    return haystack.indexOf(term) > -1;
  });

  filtered.sort(function (a, b) {
    return (b.order || 0) - (a.order || 0);
  });

  return filtered.slice(0, 30).map(mapOccurrenceForClient_);
}

/**
 * Coleta todas as ocorrências das abas NIR
 */
function collectOccurrences_() {
  var ss = getSS();
  var out = [];

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

    var values = sh.getRange(2, 1, lastRow - 1, lastCol).getValues();

    values.forEach(function (row, idx) {
      var dia = getValueByHeader_(headersNorm, row, ['dia']);
      var turno = getValueByHeader_(headersNorm, row, ['turno']);
      var paciente = getValueByHeader_(headersNorm, row, ['nome do paciente', 'paciente']);
      var especialidade = getValueByHeader_(headersNorm, row, ['especialidade']);
      var origem = getValueByHeader_(headersNorm, row, ['origem']);
      var status = getValueByHeader_(headersNorm, row, ['status']);
      var observacao = getValueByHeader_(headersNorm, row, ['observação', 'observacao']);
      var dataReserva = getValueByHeader_(headersNorm, row, [
        'data da reserva',
        'data da confirmação da reserva',
        'data do cancelamento da reserva',
        'data de admissão'
      ]);
      var horaReserva = getValueByHeader_(headersNorm, row, [
        'hora da reserva',
        'hora da confirmação da reserva',
        'hora do cancelamento da reserva',
        'hora da admissão'
      ]);

      var timestamp = buildTimestamp_(dataReserva, horaReserva);
      var order = timestamp || (lastRow - idx); // fallback usando posição

      out.push({
        category: categoryName,
        dia: dia,
        turno: turno,
        paciente: paciente,
        especialidade: especialidade,
        origem: origem,
        status: status,
        observacao: observacao,
        dataHora: timestamp,
        order: order
      });
    });
  });

  out.sort(function (a, b) {
    return (b.order || 0) - (a.order || 0);
  });

  return out;
}

function getValueByHeader_(headersNorm, row, keys) {
  for (var i = 0; i < keys.length; i++) {
    var idx = headersNorm.indexOf(normalize_(keys[i]));
    if (idx > -1) {
      var val = row[idx];
      if (val !== undefined && val !== null && String(val).trim() !== '') {
        return val;
      }
    }
  }
  return '';
}

function buildTimestamp_(dateValue, timeValue) {
  var dateObj = parseDateFlexible_(dateValue);
  if (!dateObj) return 0;

  if (timeValue) {
    var time = String(timeValue);
    var parts = time.split(':');
    if (parts.length >= 2) {
      dateObj.setHours(Number(parts[0]) || 0, Number(parts[1]) || 0, 0, 0);
    }
  }
  return dateObj.getTime();
}

function parseDateFlexible_(value) {
  if (!value) return null;
  if (value instanceof Date && !isNaN(value.getTime())) return value;

  if (typeof value === 'string') {
    var str = value.trim();
    if (!str) return null;

    var isoMatch = str.match(/^([0-9]{4})-([0-9]{2})-([0-9]{2})$/);
    if (isoMatch) {
      var y = Number(isoMatch[1]);
      var m = Number(isoMatch[2]) - 1;
      var d = Number(isoMatch[3]);
      var dt = new Date(y, m, d);
      return isNaN(dt.getTime()) ? null : dt;
    }

    var parts = str.split('/');
    if (parts.length === 3) {
      var dd = Number(parts[0]);
      var mm = Number(parts[1]) - 1;
      var yyyy = Number(parts[2]);
      var dt2 = new Date(yyyy, mm, dd);
      return isNaN(dt2.getTime()) ? null : dt2;
    }
  }
  return null;
}

/**
 * Retorna visão estilo aba por dia/turno/categoria
 */
function getSheetView(options) {
  options = options || {};
  var category = options.category || '';
  var dia = normalize_(options.dia || '');
  var turno = normalize_(options.turno || '');

  if (category && category !== 'TODOS') {
    return buildSheetViewByCategory_(category, dia, turno);
  }

  return buildSheetViewCombined_(dia, turno);
}

function buildSheetViewCombined_(dia, turno) {
  var occurrences = collectOccurrences_();

  var filtered = occurrences.filter(function (occ) {
    if (dia && normalize_(occ.dia) !== dia) {
      return false;
    }

    if (turno && normalize_(occ.turno) !== turno) {
      return false;
    }

    return true;
  });

  var mapped = filtered.slice(0, 120).map(function (occ) {
    var base = mapOccurrenceForClient_(occ);
    base.dia = occ.dia || '';
    base.turno = occ.turno || '';
    return base;
  });

  var headers = [
    'Categoria',
    'Paciente',
    'Especialidade',
    'Origem',
    'Status',
    'Turno',
    'Data/Hora',
    'Observação'
  ];

  var rows = mapped.map(function (item) {
    return [
      item.category,
      item.paciente || 'Paciente não informado',
      item.especialidade || '—',
      item.origem || '—',
      item.status || '—',
      item.turno || '—',
      item.dataHora || '—',
      item.observacao || '—'
    ];
  });

  return {
    total: filtered.length,
    headers: headers,
    rows: rows
  };
}

function buildSheetViewByCategory_(category, dia, turno) {
  var sheetName = NIR_SHEETS[category];
  if (!sheetName) {
    return buildSheetViewCombined_(dia, turno);
  }

  var sh = getSS().getSheetByName(sheetName);
  if (!sh) {
    return buildSheetViewCombined_(dia, turno);
  }

  var lastRow = sh.getLastRow();
  var lastCol = sh.getLastColumn();
  if (lastRow <= 1 || lastCol === 0) {
    return { total: 0, headers: [], rows: [] };
  }

  var headers = sh.getRange(1, 1, 1, lastCol).getValues()[0];
  var headersNorm = headers.map(function (h) { return normalize_(h); });
  var values = sh.getRange(2, 1, lastRow - 1, lastCol).getValues();

  var diaIdx = headersNorm.indexOf('dia');
  var turnoIdx = headersNorm.indexOf('turno');

  var filteredRows = values.filter(function (row) {
    if (dia && diaIdx > -1) {
      if (normalize_(row[diaIdx]) !== dia) return false;
    }

    if (turno && turnoIdx > -1) {
      if (normalize_(row[turnoIdx]) !== turno) return false;
    }

    return true;
  });

  var rows = filteredRows.slice(0, 120).map(function (row) {
    return row.map(formatSheetCellForDisplay_);
  });

  return {
    total: filteredRows.length,
    headers: headers,
    rows: rows
  };
}

function formatSheetCellForDisplay_(value) {
  if (value instanceof Date && !isNaN(value.getTime())) {
    var tz = Session.getScriptTimeZone() || 'America/Sao_Paulo';
    return Utilities.formatDate(value, tz, 'dd/MM/yyyy HH:mm');
  }
  return sanitizeValue_(value);
}

function mapOccurrenceForClient_(occ) {
  return {
    category: occ.category,
    dia: occ.dia,
    turno: occ.turno,
    paciente: occ.paciente,
    especialidade: occ.especialidade,
    origem: occ.origem,
    status: occ.status,
    observacao: occ.observacao,
    dataHora: occ.dataHora ? formatDateTime_(occ.dataHora) : ''
  };
}

function formatDateTime_(timestamp) {
  if (!timestamp) return '';
  var d = new Date(timestamp);
  if (isNaN(d.getTime())) return '';
  var dd = ('0' + d.getDate()).slice(-2);
  var mm = ('0' + (d.getMonth() + 1)).slice(-2);
  var yyyy = d.getFullYear();
  var hh = ('0' + d.getHours()).slice(-2);
  var min = ('0' + d.getMinutes()).slice(-2);
  return dd + '/' + mm + '/' + yyyy + ' ' + hh + ':' + min;
}
