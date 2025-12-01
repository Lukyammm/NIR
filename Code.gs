
/*********** CONFIGURAÇÃO BÁSICA ***********/

const NIR_SHEETS = {
  'RESERVA CONFIRMADA': 'RESERVA CONFIRMADA',
  'PROCEDIMENTO CONFIRMADO': 'PROCEDIMENTO CONFIRMADO',
  'RESERVA NEGADA': 'RESERVA NEGADA',
  'PLANTÃO ANTERIOR': 'PLANTÃO ANTERIOR'
};

/**
 * Abre o WebApp (usa o arquivo Index.html)
 */
function doGet(e) {
  return HtmlService
    .createHtmlOutputFromFile('Index')
    .setTitle('NIR – Ocorrências');
}

/**
 * Utilitário: pega planilha ativa
 */
function getSS() {
  return SpreadsheetApp.getActiveSpreadsheet();
}

/**
 * Utilitário: retorna aba pelo nome ou lança erro amigável
 */
function getSheetByName_(name) {
  const ss = getSS();
  const sheet = ss.getSheetByName(name);
  if (!sheet) {
    throw new Error('Aba não encontrada: ' + name);
  }
  return sheet;
}

/**
 * Normaliza string (minúscula, sem acento, trim) para comparar
 */
function normalize_(str) {
  return String(str || '')
    .toLowerCase()
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .trim();
}

/**
 * Localiza índice (0-based) de uma coluna pelo texto do cabeçalho
 */
function findColIndexByHeader_(headers, target) {
  const normalizedTarget = normalize_(target);
  for (let i = 0; i < headers.length; i++) {
    if (normalize_(headers[i]) === normalizedTarget) {
      return i;
    }
  }
  return -1;
}

/*********** API PRINCIPAL – GRAVAR OCORRÊNCIA ***********/

/**
 * Salva uma ocorrência em uma das 4 abas do NIR.
 *
 * Espera receber do front-end um objeto payload:
 * {
 *   tipoGrupo: 'RESERVA CONFIRMADA' | 'PROCEDIMENTO CONFIRMADO' | 'RESERVA NEGADA' | 'PLANTÃO ANTERIOR',
 *   // opção 1 – linha como array, NA ORDEM das colunas da planilha:
 *   linha: [...],
 *
 *   // opção 2 – linha como objeto, com chaves que batem com os cabeçalhos
 *   // (ex.: "TIPO", "Fastmedic", "Nome do Paciente", "Dia", "Turno" etc.)
 *   // linha: { "TIPO": "RESERVA CONFIRMADA", "Fastmedic": "1234", ... }
 * }
 */
function salvarOcorrencia(payload) {
  if (!payload || typeof payload !== 'object') {
    throw new Error('Payload inválido recebido em salvarOcorrencia.');
  }

  const tipoGrupo = payload.tipoGrupo;
  const linha = payload.linha;

  if (!tipoGrupo || !NIR_SHEETS[tipoGrupo]) {
    throw new Error('Tipo de grupo inválido: ' + tipoGrupo);
  }
  if (!linha) {
    throw new Error('Nenhuma linha enviada para salvar.');
  }

  const sheetName = NIR_SHEETS[tipoGrupo];
  const sheet = getSheetByName_(sheetName);

  const lastColumn = sheet.getLastColumn();
  const headers = sheet.getRange(1, 1, 1, lastColumn).getValues()[0]; // linha 1 – cabeçalhos
  const numCols = headers.length;

  let rowToAppend = new Array(numCols).fill('');

  // Caso 1: linha é array → cola direto na ordem
  if (Array.isArray(linha)) {
    for (let i = 0; i < Math.min(numCols, linha.length); i++) {
      rowToAppend[i] = linha[i];
    }
  }
  // Caso 2: linha é objeto → mapeia pelas chaves x cabeçalhos
  else if (typeof linha === 'object') {
    const normalizedData = {};
    Object.keys(linha).forEach(function (key) {
      normalizedData[normalize_(key)] = linha[key];
    });

    for (let col = 0; col < numCols; col++) {
      const header = headers[col];
      const normHeader = normalize_(header);
      if (normalizedData.hasOwnProperty(normHeader)) {
        rowToAppend[col] = normalizedData[normHeader];
      } else {
        // Se não tiver no objeto, fica em branco (ou você pode colocar algum default aqui)
        rowToAppend[col] = '';
      }
    }
  } else {
    throw new Error('Formato de linha não suportado em salvarOcorrencia.');
  }

  sheet.appendRow(rowToAppend);

  return {
    ok: true,
    message: 'Ocorrência salva com sucesso em "' + sheetName + '".',
    sheet: sheetName
  };
}

/*********** API PARA DASHBOARD – RESUMO POR DIA/TURNO ***********/

/**
 * Retorna um resumo do total de ocorrências por aba para um dia/turno específico.
 *
 * Espera receber do front-end:
 *   dia   → string exatamente como está na coluna "dia" das abas (ex.: "01/12/2025")
 *   turno → string exatamente como está na coluna "turno" (ex.: "MATUTINO", "VESPERTINO", "NOTURNO")
 *
 * Retorno:
 * {
 *   filtros: { dia, turno },
 *   totais: [
 *     { tipoGrupo: 'RESERVA CONFIRMADA', sheet: 'RESERVA CONFIRMADA', total: 5 },
 *     { tipoGrupo: 'PROCEDIMENTO CONFIRMADO', sheet: 'PROCEDIMENTO CONFIRMADO', total: 2 },
 *     ...
 *   ]
 * }
 */
function getResumoPorDiaTurno(dia, turno) {
  if (!dia || !turno) {
    throw new Error('Informe "dia" e "turno" para obter o resumo.');
  }

  const ss = getSS();
  const result = [];

  Object.keys(NIR_SHEETS).forEach(function (tipoGrupo) {
    const sheetName = NIR_SHEETS[tipoGrupo];
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) return;

    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();
    if (lastRow < 2) {
      // só cabeçalho, sem dados
      result.push({
        tipoGrupo: tipoGrupo,
        sheet: sheetName,
        total: 0
      });
      return;
    }

    const range = sheet.getRange(1, 1, lastRow, lastCol);
    const values = range.getValues();
    const headers = values[0];

    const diaColIndex = findColIndexByHeader_(headers, 'dia');
    const turnoColIndex = findColIndexByHeader_(headers, 'turno');

    if (diaColIndex === -1 || turnoColIndex === -1) {
      // se não achar dia/turno nessa aba, ignora ou loga
      result.push({
        tipoGrupo: tipoGrupo,
        sheet: sheetName,
        total: 0
      });
      return;
    }

    let count = 0;
    for (let r = 1; r < values.length; r++) {
      const row = values[r];
      const diaVal = String(row[diaColIndex] || '').trim();
      const turnoVal = String(row[turnoColIndex] || '').trim();

      if (diaVal === String(dia).trim() && turnoVal === String(turno).trim()) {
        count++;
      }
    }

    result.push({
      tipoGrupo: tipoGrupo,
      sheet: sheetName,
      total: count
    });
  });

  return {
    filtros: { dia: dia, turno: turno },
    totais: result
  };
}

/*********** OPCIONAL – EXEMPLO DE HELPER PARA O FRONT ***********/

/**
 * Retorna configuração básica para o front montar selects, etc.
 */
function getConfigNIR() {
  return {
    tiposGrupo: Object.keys(NIR_SHEETS),
    sheets: NIR_SHEETS
  };
}
