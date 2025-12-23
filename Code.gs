const DEFAULT_TIMEZONE = "America/Sao_Paulo";
const SPREADSHEET_ID_PROPERTY = "PLANILHA_ID";
const PLANILHA_ID =
  PropertiesService.getScriptProperties().getProperty(SPREADSHEET_ID_PROPERTY) ||
  "1vnnGQEkAjP9eTRLWWSb2lSngwGTRYa_rtk6DsG8HGqc";

const AUTHORIZED_USERS = [];
const ACTIVE_SHIFT_KEY = "ACTIVE_SHIFT_STATE";
const REPORT_SHEET_CURRENT = "Atual";
const REPORT_SHEET_HISTORY = "Histórico";
const REPORT_COLUMNS = ["ID_LINHA", "COLUNA", "ORDEM", "TEXTO", "TRAVADO", "ATUALIZADO_EM"];
const REPORT_HISTORY_COLUMNS = [
  "ID_REGISTRO",
  "DATA_HORA_ENCERRAMENTO",
  "PLANTAO",
  "COLUNA_ENFERMAGEM",
  "COLUNA_MEDICA",
  "USUARIO"
];

function gerarID_() {
  return Utilities.getUuid().split("-")[0].toUpperCase();
}

function nowIso_() {
  return new Date().toISOString();
}

function safeJsonParse_(value, fallback) {
  try {
    return JSON.parse(value);
  } catch (err) {
    return fallback;
  }
}

function safeJsonStringify_(value) {
  try {
    return JSON.stringify(value);
  } catch (err) {
    return "";
  }
}

function normalizarDataHora_(valor, tz) {
  try {
    if (valor === null || valor === undefined || valor === "") return null;
    const timezone = tz || Session.getScriptTimeZone() || DEFAULT_TIMEZONE;

    if (Object.prototype.toString.call(valor) === "[object Date]") {
      if (isNaN(valor)) return null;
      return ajustarParaTimezone_(valor, timezone);
    }

    if (typeof valor === "number") {
      const base = new Date(Date.UTC(1899, 11, 30));
      const millis = valor * 24 * 60 * 60 * 1000;
      const dtSerial = new Date(base.getTime() + millis);
      return ajustarParaTimezone_(dtSerial, timezone);
    }

    const str = String(valor).trim();
    if (!str) return null;

    const normalizado = str.replace(/\s+/g, " ").replace(/-/g, "/");
    const m = normalizado.match(
      /^(\d{1,2})[\/\.](\d{1,2})[\/\.](\d{2,4})(?:[ T](\d{1,2})[:hH]?(\d{1,2}))?/
    );
    if (m) {
      const dia = Number(m[1]);
      const mes = Number(m[2]) - 1;
      let ano = Number(m[3]);
      if (ano < 100) ano += 2000;
      const hora = Number(m[4] || 0);
      const minuto = Number(m[5] || 0);
      const candidato = new Date(ano, mes, dia, hora, minuto);
      return isNaN(candidato) ? null : ajustarParaTimezone_(candidato, timezone);
    }

    const isoLike = new Date(str.replace(/ /g, "T"));
    if (!isNaN(isoLike)) {
      return ajustarParaTimezone_(isoLike, timezone);
    }

    const dt = new Date(str);
    return isNaN(dt) ? null : ajustarParaTimezone_(dt, timezone);
  } catch (e) {
    Logger.log("normalizarDataHora_ falhou para valor: " + valor + " -> " + e);
    return null;
  }
}

function normalizarDataSimples_(valor, tz) {
  const dt = normalizarDataHora_(valor, tz);
  if (!dt) return "";
  const timezone = tz || Session.getScriptTimeZone() || DEFAULT_TIMEZONE;
  return Utilities.formatDate(dt, timezone, "yyyy-MM-dd");
}

function getDayOfWeekFromDate_(valor, tz) {
  const dt = normalizarDataHora_(valor, tz);
  if (!dt) return "";
  const dias = [
    "Domingo",
    "Segunda-feira",
    "Terça-feira",
    "Quarta-feira",
    "Quinta-feira",
    "Sexta-feira",
    "Sábado"
  ];
  return dias[dt.getDay()] || "";
}

function normalizeTurno_(valor) {
  const turno = String(valor || "").trim().toUpperCase();
  return turno === "MT" || turno === "SN" ? turno : "";
}

function ajustarParaTimezone_(dateObj, tz) {
  if (!dateObj) return null;
  const timezone = tz || Session.getScriptTimeZone() || DEFAULT_TIMEZONE;
  const iso = Utilities.formatDate(dateObj, timezone, "yyyy-MM-dd'T'HH:mm:ss'Z'");
  const ajustado = new Date(iso);
  return isNaN(ajustado) ? null : ajustado;
}

function getSpreadsheet_() {
  if (!PLANILHA_ID) {
    throw new Error("PLANILHA_ID não configurado para o WebApp.");
  }
  return SpreadsheetApp.openById(PLANILHA_ID);
}

function isUserAuthorized_() {
  if (!AUTHORIZED_USERS || AUTHORIZED_USERS.length === 0) {
    return true;
  }
  const email = Session.getActiveUser().getEmail();
  return AUTHORIZED_USERS.indexOf(email) !== -1;
}

function assertAuthorized_() {
  if (!isUserAuthorized_()) {
    throw new Error("Acesso não autorizado.");
  }
}

function buildResponse_(data) {
  return { ok: true, data: data };
}

function buildError_(err, code, details) {
  const message = err && err.message ? err.message : String(err || "Erro desconhecido");
  return {
    ok: false,
    error: {
      message: message,
      code: code || "APP_ERROR",
      details: details || null
    }
  };
}

function runWithResult_(fn) {
  try {
    return buildResponse_(fn());
  } catch (err) {
    Logger.log("Erro: " + err);
    return buildError_(err, "APP_ERROR", err && err.stack ? String(err.stack) : null);
  }
}

function getActiveShiftFromProperties_() {
  const raw = PropertiesService.getScriptProperties().getProperty(ACTIVE_SHIFT_KEY);
  if (!raw) return null;
  const parsed = safeJsonParse_(raw, null);
  return parsed && parsed.status === "active" ? parsed : null;
}

function setActiveShift_(shift) {
  PropertiesService.getScriptProperties().setProperty(ACTIVE_SHIFT_KEY, safeJsonStringify_(shift));
}

function clearActiveShift_() {
  PropertiesService.getScriptProperties().deleteProperty(ACTIVE_SHIFT_KEY);
}

function getShiftFromSheet_() {
  const ss = getSpreadsheet_();
  const sheet = ss.getSheetByName("CONFIG_PLANTAO");
  if (!sheet) return null;

  const values = sheet.getRange("A2:K2").getValues()[0];
  const hasContent = values.some((v) => v !== "" && v !== null);
  if (!hasContent) return null;

  let id = values[0] ? String(values[0]) : "";
  if (!id) {
    id = gerarID_();
    sheet.getRange("A2").setValue(id);
  }

  const tz = Session.getScriptTimeZone() || DEFAULT_TIMEZONE;
  const startedAt = normalizarDataHora_(values[9], tz);
  const endedAt = normalizarDataHora_(values[10], tz);

  return {
    id: id,
    date: normalizarDataSimples_(values[1], tz) || "",
    dayOfWeek: values[2] || "",
    shift: values[3] || "",
    team: {
      medico1: values[4] || "",
      medico2: values[5] || "",
      enf1: values[6] || "",
      enf2: values[7] || "",
      aux: values[8] || ""
    },
    startedAt: startedAt ? startedAt.toISOString() : "",
    endedAt: endedAt ? endedAt.toISOString() : "",
    status: endedAt ? "closed" : "active",
    updatedAt: nowIso_()
  };
}

function writeShiftToSheet_(shift) {
  const ss = getSpreadsheet_();
  const sheet = ss.getSheetByName("CONFIG_PLANTAO");
  if (!sheet) return;

  const row = [
    shift.id,
    shift.date || "",
    shift.dayOfWeek || "",
    shift.shift || "",
    shift.team && shift.team.medico1 ? shift.team.medico1 : "",
    shift.team && shift.team.medico2 ? shift.team.medico2 : "",
    shift.team && shift.team.enf1 ? shift.team.enf1 : "",
    shift.team && shift.team.enf2 ? shift.team.enf2 : "",
    shift.team && shift.team.aux ? shift.team.aux : "",
    shift.startedAt || "",
    shift.endedAt || ""
  ];

  sheet.getRange("A2:K2").setValues([row]);
}

function getActiveShift_() {
  const shift = getActiveShiftFromProperties_();
  if (shift) return shift;

  const recovered = getShiftFromSheet_();
  if (recovered && recovered.status === "active") {
    setActiveShift_(recovered);
    writeAction_("SHIFT_RECOVERED", { shiftId: recovered.id });
    return recovered;
  }

  return null;
}

function writeAction_(type, payload) {
  const ss = getSpreadsheet_();
  const log = ss.getSheetByName("LOG_NIR");
  if (!log) return;

  const shift = getActiveShiftFromProperties_();
  const idRegistro = payload && payload.shiftId ? payload.shiftId : shift ? shift.id : "";
  const obs = payload ? safeJsonStringify_(payload) : "";

  log.appendRow([
    gerarID_(),
    idRegistro,
    "APP",
    type,
    Session.getActiveUser().getEmail(),
    new Date(),
    obs
  ]);
}

function getUserInfo_() {
  return {
    email: Session.getActiveUser().getEmail() || "",
    timezone: Session.getScriptTimeZone() || DEFAULT_TIMEZONE
  };
}

function getAppState() {
  return runWithResult_(function () {
    assertAuthorized_();
    const shift = getActiveShift_();
    return {
      activeShift: shift,
      user: getUserInfo_(),
      now: nowIso_()
    };
  });
}

function startShift(payload) {
  return runWithResult_(function () {
    assertAuthorized_();
    const lock = LockService.getScriptLock();
    lock.waitLock(10000);

    try {
      const existing = getActiveShiftFromProperties_();
      if (existing) {
        return { activeShift: existing };
      }

      const tz = Session.getScriptTimeZone() || DEFAULT_TIMEZONE;
      const data = payload || {};
      const shift = {
        id: gerarID_(),
        date: normalizarDataSimples_(data.data, tz) || "",
        dayOfWeek: getDayOfWeekFromDate_(data.data, tz),
        shift: normalizeTurno_(data.turno),
        team: {
          medico1: data.medico1 || "",
          medico2: data.medico2 || "",
          enf1: data.enf1 || "",
          enf2: data.enf2 || "",
          aux: data.aux || ""
        },
        startedAt: nowIso_(),
        endedAt: "",
        status: "active",
        updatedAt: nowIso_(),
        openedBy: Session.getActiveUser().getEmail() || ""
      };

      setActiveShift_(shift);
      writeShiftToSheet_(shift);
      writeAction_("SHIFT_STARTED", { shiftId: shift.id });

      return { activeShift: shift };
    } finally {
      lock.releaseLock();
    }
  });
}

function updateShift(payload) {
  return runWithResult_(function () {
    assertAuthorized_();
    const lock = LockService.getScriptLock();
    lock.waitLock(10000);

    try {
      const existing = getActiveShiftFromProperties_();
      if (!existing) {
        throw new Error("Nenhum plantão ativo para atualizar.");
      }

      const tz = Session.getScriptTimeZone() || DEFAULT_TIMEZONE;
      const data = payload || {};
      const baseDate = data.data || existing.date;
      const dayOfWeek = getDayOfWeekFromDate_(baseDate, tz) || existing.dayOfWeek || "";

      const updated = {
        id: existing.id,
        date: normalizarDataSimples_(data.data, tz) || existing.date || "",
        dayOfWeek: dayOfWeek,
        shift: normalizeTurno_(data.turno) || existing.shift || "",
        team: {
          medico1: data.medico1 || existing.team.medico1 || "",
          medico2: data.medico2 || existing.team.medico2 || "",
          enf1: data.enf1 || existing.team.enf1 || "",
          enf2: data.enf2 || existing.team.enf2 || "",
          aux: data.aux || existing.team.aux || ""
        },
        startedAt: existing.startedAt,
        endedAt: existing.endedAt || "",
        status: existing.status,
        updatedAt: nowIso_(),
        openedBy: existing.openedBy || ""
      };

      setActiveShift_(updated);
      writeShiftToSheet_(updated);
      writeAction_("SHIFT_UPDATED", { shiftId: updated.id });

      return { activeShift: updated };
    } finally {
      lock.releaseLock();
    }
  });
}

function endShift(shiftId) {
  return runWithResult_(function () {
    assertAuthorized_();
    const lock = LockService.getScriptLock();
    lock.waitLock(10000);

    try {
      const existing = getActiveShiftFromProperties_();
      if (!existing) {
        return { activeShift: null };
      }

      if (shiftId && String(shiftId) !== String(existing.id)) {
        throw new Error("O plantão informado não corresponde ao plantão ativo.");
      }

      const endedAt = nowIso_();
      const closed = Object.assign({}, existing, {
        endedAt: endedAt,
        status: "closed",
        updatedAt: endedAt
      });

      const ss = getSpreadsheet_();
      const hist = ss.getSheetByName("HIST_PLANTOES");
      if (hist) {
        const row = [
          closed.id,
          closed.date || "",
          closed.dayOfWeek || "",
          closed.shift || "",
          closed.team.medico1 || "",
          closed.team.medico2 || "",
          closed.team.enf1 || "",
          closed.team.enf2 || "",
          closed.team.aux || "",
          closed.startedAt || "",
          closed.endedAt || "",
          Session.getActiveUser().getEmail() || ""
        ];
        hist.appendRow(row);
      }

      snapshotReportState_(closed);

      const conf = ss.getSheetByName("CONFIG_PLANTAO");
      if (conf) {
        conf.getRange("A2:K2").clearContent();
      }

      clearActiveShift_();
      writeAction_("SHIFT_ENDED", { shiftId: closed.id });

      return { activeShift: null };
    } finally {
      lock.releaseLock();
    }
  });
}

function writeAction(type, payload) {
  return runWithResult_(function () {
    assertAuthorized_();
    writeAction_(type, payload);
    return { saved: true };
  });
}

function readDashboardData(payload) {
  return runWithResult_(function () {
    assertAuthorized_();
    const view = payload && payload.view ? payload.view : "confirmadas";
    const principalMap = {
      confirmadas: "RESERVA_CONFIRMADA",
      vascular: "PROCEDIMENTO_VASCULAR",
      negadas: "RESERVA_NEGADA",
      anterior: "PLANTAO_ANTERIOR"
    };
    const principalSheet = principalMap[view] || "RESERVA_CONFIRMADA";

    return {
      activeShift: getActiveShift_(),
      indicadores: getIndicadores_(),
      principal: getTabela_(principalSheet),
      manutencao: getTabela_("BLOQUEADOS_MANUTENCAO"),
      isolamento: getTabela_("BLOQUEADOS_ISOLAMENTO"),
      principalNome: principalSheet,
      now: nowIso_()
    };
  });
}

function addRecord(payload) {
  return runWithResult_(function () {
    assertAuthorized_();
    const modulo = payload && payload.modulo ? payload.modulo : "";
    const registro = payload && payload.registro ? payload.registro : {};
    const resultado = adicionarRegistro_(modulo, registro);
    return resultado || { saved: true };
  });
}

function deleteRecord(payload) {
  return runWithResult_(function () {
    assertAuthorized_();
    const modulo = payload && payload.modulo ? payload.modulo : "";
    const id = payload && payload.id ? payload.id : "";
    excluirRegistro_(modulo, id);
    return { deleted: true };
  });
}

function criarEstruturaNIR() {
  assertAuthorized_();
  const ss = getSpreadsheet_();

  const abasVivas = {
    "CONFIG_PLANTAO": [
      "ID_PLANTAO", "Data do Plantão", "Dia da Semana", "Turno",
      "Médico(a) 1", "Médico(a) 2",
      "Enfermeiro(a) 1", "Enfermeiro(a) 2",
      "Auxiliar Administrativo",
      "Abertura (Timestamp)", "Encerramento (Timestamp)"
    ],

    "RESERVA_CONFIRMADA": [
      "ID", "Tipo", "Fastmedic", "Nome do Paciente", "Leito Reservado",
      "Especialidade", "Origem", "Status",
      "Data da Reserva", "Hora da Reserva",
      "Data da Confirmação", "Hora da Confirmação",
      "Tempo entre Alocação e Aceite", "Chegou", "Observação"
    ],

    "PROCEDIMENTO_VASCULAR": [
      "ID", "Tipo", "Fastmedic", "Nome do Paciente", "Leito Reservado",
      "Especialidade", "Origem", "Status",
      "Data da Reserva", "Hora da Reserva",
      "Data da Confirmação", "Hora da Confirmação",
      "Tempo entre Alocação e Aceite", "Chegou", "Observação"
    ],

    "RESERVA_NEGADA": [
      "ID", "Tipo", "Fastmedic", "Nome do Paciente", "Origem",
      "Especialidade", "Justificativa",
      "Data da Reserva", "Hora da Reserva",
      "Data do Cancelamento", "Hora do Cancelamento",
      "Tempo entre Alocação e Cancelamento"
    ],

    "PLANTAO_ANTERIOR": [
      "ID", "Tipo", "Fastmedic", "Nome do Paciente", "Leito Reservado",
      "Especialidade", "Origem", "Status",
      "Data da Reserva", "Hora da Reserva",
      "Data de Admissão", "Hora de Admissão",
      "Tempo Decorrido", "Chegou", "Observação"
    ],

    "BLOQUEADOS_MANUTENCAO": [
      "ID", "Leito", "Unidade", "Manutenção Acionada", "Data de Início", "Previsão de Reparo", "Observação"
    ],

    "BLOQUEADOS_ISOLAMENTO": [
      "ID", "Unidade", "Leito", "Paciente",
      "Tipo de Isolamento", "Patologia",
      "Início do Isolamento", "Tempo Previsto", "Observação"
    ],

    "Atual": REPORT_COLUMNS
  };

  const abasHistoricos = {
    "HIST_PLANTOES": [
      "ID_PLANTAO", "Data", "Dia", "Turno",
      "Medico1", "Medico2",
      "Enf1", "Enf2",
      "Aux", "Hora Abertura", "Hora Encerramento", "Usuario_Encerramento"
    ],

    "HIST_RESERVA_CONFIRMADA": [
      "PLANTAO_ID", "ID", "Tipo", "Fastmedic", "Nome do Paciente", "Leito Reservado",
      "Especialidade", "Origem", "Status",
      "Data Reserva", "Hora Reserva",
      "Data Confirmação", "Hora Confirmação",
      "Tempo Aceite", "Chegou", "Observação"
    ],

    "HIST_PROCEDIMENTO_VASCULAR": [
      "PLANTAO_ID", "ID", "Tipo", "Fastmedic", "Nome do Paciente", "Leito Reservado",
      "Especialidade", "Origem", "Status",
      "Data Reserva", "Hora Reserva",
      "Data Confirmação", "Hora Confirmação",
      "Tempo Aceite", "Chegou", "Observação"
    ],

    "HIST_RESERVA_NEGADA": [
      "PLANTAO_ID", "ID", "Tipo", "Nome", "Origem", "Especialidade",
      "Justificativa", "Data Reserva", "Hora Reserva",
      "Data Cancelamento", "Hora Cancelamento", "Tempo Cancelamento"
    ],

    "HIST_MANUTENCAO": [
      "PLANTAO_ID", "ID", "Leito", "Unidade", "Manutenção Acionada", "Data Início", "Previsão", "Observação", "Encerrado em"
    ],

    "HIST_ISOLAMENTO": [
      "PLANTAO_ID", "ID", "Unidade", "Leito", "Paciente",
      "Isolamento", "Patologia", "Início", "Tempo Previsto", "Observação", "Encerrado em"
    ],

    "LOG_NIR": [
      "ID_EVENTO", "ID_REGISTRO", "MODULO", "TIPO_EVENTO",
      "USUARIO", "DATA_HORA", "OBSERVACAO"
    ],

    "Histórico": REPORT_HISTORY_COLUMNS
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

function getIndicadores_() {
  const cache = CacheService.getScriptCache();
  const cached = cache.get("indicadores");
  if (cached) {
    return safeJsonParse_(cached, null);
  }

  const ss = getSpreadsheet_();
  const confirmadas = ss.getSheetByName("RESERVA_CONFIRMADA");
  const vascular = ss.getSheetByName("PROCEDIMENTO_VASCULAR");
  const negadas = ss.getSheetByName("RESERVA_NEGADA");
  const anterior = ss.getSheetByName("PLANTAO_ANTERIOR");

  const countConfirmadas = confirmadas ? Math.max(confirmadas.getLastRow() - 1, 0) : 0;
  const countVascular = vascular ? Math.max(vascular.getLastRow() - 1, 0) : 0;
  const countNegadas = negadas ? Math.max(negadas.getLastRow() - 1, 0) : 0;
  const countAnterior = anterior ? Math.max(anterior.getLastRow() - 1, 0) : 0;

  const indicadores = {
    alocados: countConfirmadas + countVascular,
    reservasConfirmadas: countConfirmadas,
    reservasCanceladas: countNegadas,
    admitidosUIB: countAnterior
  };

  cache.put("indicadores", safeJsonStringify_(indicadores), 30);
  return indicadores;
}

function getTabela_(sheetName) {
  const cache = CacheService.getScriptCache();
  const cacheKey = "table:" + sheetName;
  const cached = cache.get(cacheKey);
  if (cached) {
    return safeJsonParse_(cached, { headers: [], rows: [] });
  }

  const ss = getSpreadsheet_();
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) return { headers: [], rows: [] };

  const lastCol = sheet.getLastColumn();
  const headers = lastCol > 0 ? sheet.getRange(1, 1, 1, lastCol).getValues()[0].map((h) => h || "") : [];

  if (headers.length === 0) {
    return { headers: [], rows: [] };
  }

  const dataRange = sheet.getDataRange();
  const values = dataRange.getValues();
  const linhas = values
    .slice(1)
    .filter((row) => row.some((cell) => cell !== "" && cell !== null))
    .map((row) => {
      if (row.length < headers.length) {
        return row.concat(Array(headers.length - row.length).fill(""));
      }
      return row.slice(0, headers.length);
    });

  const payload = { headers: headers, rows: linhas };
  cache.put(cacheKey, safeJsonStringify_(payload), 15);
  return payload;
}

function ensureReportSheets_() {
  const ss = getSpreadsheet_();
  let current = ss.getSheetByName(REPORT_SHEET_CURRENT);
  if (!current) {
    current = ss.insertSheet(REPORT_SHEET_CURRENT);
    current.appendRow(REPORT_COLUMNS);
  } else if (current.getLastRow() === 0) {
    current.appendRow(REPORT_COLUMNS);
  } else {
    current.getRange(1, 1, 1, REPORT_COLUMNS.length).setValues([REPORT_COLUMNS]);
  }

  let history = ss.getSheetByName(REPORT_SHEET_HISTORY);
  if (!history) {
    history = ss.insertSheet(REPORT_SHEET_HISTORY);
    history.appendRow(REPORT_HISTORY_COLUMNS);
  } else if (history.getLastRow() === 0) {
    history.appendRow(REPORT_HISTORY_COLUMNS);
  } else {
    history.getRange(1, 1, 1, REPORT_HISTORY_COLUMNS.length).setValues([REPORT_HISTORY_COLUMNS]);
  }

  return { current: current, history: history };
}

function normalizeReportLines_(lines) {
  return (lines || [])
    .filter((line) => line && typeof line === "object")
    .map((line, index) => ({
      id: line.id ? String(line.id) : gerarID_(),
      text: line.text ? String(line.text) : "",
      locked: Boolean(line.locked),
      order: typeof line.order === "number" ? line.order : index
    }));
}

function getReportState() {
  return runWithResult_(function () {
    assertAuthorized_();
    const sheets = ensureReportSheets_();
    const dataRange = sheets.current.getDataRange();
    const values = dataRange.getValues();

    const rows = values.slice(1).filter((row) => row.some((cell) => cell !== "" && cell !== null));
    const enfermagem = [];
    const medica = [];

    rows.forEach((row) => {
      const line = {
        id: row[0] ? String(row[0]) : gerarID_(),
        text: row[3] ? String(row[3]) : "",
        locked: String(row[4]).toLowerCase() === "true",
        order: Number(row[2]) || 0
      };

      const coluna = row[1] ? String(row[1]).toLowerCase() : "";
      if (coluna === "enfermagem") {
        enfermagem.push(line);
      } else if (coluna === "medica") {
        medica.push(line);
      }
    });

    enfermagem.sort((a, b) => a.order - b.order);
    medica.sort((a, b) => a.order - b.order);

    return {
      enfermagem: enfermagem.length ? enfermagem : [{ id: gerarID_(), text: "", locked: false, order: 0 }],
      medica: medica.length ? medica : [{ id: gerarID_(), text: "", locked: false, order: 0 }]
    };
  });
}

function saveReportState(payload) {
  return runWithResult_(function () {
    assertAuthorized_();
    const lock = LockService.getScriptLock();
    lock.waitLock(10000);

    try {
      const sheets = ensureReportSheets_();
      const data = payload || {};
      const enfermagem = normalizeReportLines_(data.enfermagem);
      const medica = normalizeReportLines_(data.medica);
      const updatedAt = new Date();

      const rows = [];
      enfermagem.forEach((line, index) => {
        rows.push([line.id, "enfermagem", index, line.text, line.locked ? "TRUE" : "FALSE", updatedAt]);
      });
      medica.forEach((line, index) => {
        rows.push([line.id, "medica", index, line.text, line.locked ? "TRUE" : "FALSE", updatedAt]);
      });

      const sheet = sheets.current;
      if (sheet.getLastRow() > 1) {
        sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).clearContent();
      }

      if (rows.length > 0) {
        sheet.getRange(2, 1, rows.length, REPORT_COLUMNS.length).setValues(rows);
      }

      return { saved: true };
    } finally {
      lock.releaseLock();
    }
  });
}

function snapshotReportState_(closedShift) {
  const sheets = ensureReportSheets_();
  const current = sheets.current;
  const dataRange = current.getDataRange();
  const values = dataRange.getValues();
  const rows = values.slice(1).filter((row) => row.some((cell) => cell !== "" && cell !== null));
  const enfermagem = [];
  const medica = [];

  rows.forEach((row) => {
    const line = {
      id: row[0] ? String(row[0]) : gerarID_(),
      text: row[3] ? String(row[3]) : "",
      locked: String(row[4]).toLowerCase() === "true",
      order: Number(row[2]) || 0
    };
    const coluna = row[1] ? String(row[1]).toLowerCase() : "";
    if (coluna === "enfermagem") {
      enfermagem.push(line);
    } else if (coluna === "medica") {
      medica.push(line);
    }
  });

  enfermagem.sort((a, b) => a.order - b.order);
  medica.sort((a, b) => a.order - b.order);

  const historyRow = [
    gerarID_(),
    closedShift && closedShift.endedAt ? closedShift.endedAt : nowIso_(),
    closedShift && closedShift.id ? closedShift.id : "",
    safeJsonStringify_(enfermagem),
    safeJsonStringify_(medica),
    Session.getActiveUser().getEmail() || ""
  ];

  sheets.history.appendRow(historyRow);
}

function adicionarRegistro_(modulo, registro) {
  const ss = getSpreadsheet_();
  const shift = getActiveShift_();
  const shiftId = shift && shift.id ? shift.id : "";

  const id = gerarID_();

  let abaNome = "";
  let histNome = null;
  let row = [];

  switch (modulo) {
    case "RESERVA_CONFIRMADA":
      abaNome = "RESERVA_CONFIRMADA";
      histNome = "HIST_RESERVA_CONFIRMADA";
      row = [
        id,
        registro.tipo || "",
        registro.fastmedic || "",
        registro.nome || "",
        registro.leito || "",
        registro.especialidade || "",
        registro.origem || "",
        registro.status || "",
        registro.dataReserva || "",
        registro.horaReserva || "",
        registro.dataConfirmacao || "",
        registro.horaConfirmacao || "",
        registro.tempoAceite || "",
        registro.chegou || "",
        registro.observacao || ""
      ];
      break;

    case "PROCEDIMENTO_VASCULAR":
      abaNome = "PROCEDIMENTO_VASCULAR";
      histNome = "HIST_PROCEDIMENTO_VASCULAR";
      row = [
        id,
        registro.tipo || "",
        registro.fastmedic || "",
        registro.nome || "",
        registro.leito || "",
        registro.especialidade || "",
        registro.origem || "",
        registro.status || "",
        registro.dataReserva || "",
        registro.horaReserva || "",
        registro.dataConfirmacao || "",
        registro.horaConfirmacao || "",
        registro.tempoAceite || "",
        registro.chegou || "",
        registro.observacao || ""
      ];
      break;

    case "RESERVA_NEGADA":
      abaNome = "RESERVA_NEGADA";
      histNome = "HIST_RESERVA_NEGADA";
      row = [
        id,
        registro.tipo || "",
        registro.fastmedic || "",
        registro.nome || "",
        registro.origem || "",
        registro.especialidade || "",
        registro.justificativa || "",
        registro.dataReserva || "",
        registro.horaReserva || "",
        registro.dataCancelamento || "",
        registro.horaCancelamento || "",
        registro.tempoCancelamento || ""
      ];
      break;

    case "PLANTAO_ANTERIOR":
      abaNome = "PLANTAO_ANTERIOR";
      row = [
        id,
        registro.tipo || "",
        registro.fastmedic || "",
        registro.nome || "",
        registro.leito || "",
        registro.especialidade || "",
        registro.origem || "",
        registro.status || "",
        registro.dataReserva || "",
        registro.horaReserva || "",
        registro.dataAdmissao || "",
        registro.horaAdmissao || "",
        registro.tempoDecorrido || "",
        registro.chegou || "",
        registro.observacao || ""
      ];
      break;

    case "BLOQUEADOS_MANUTENCAO":
      abaNome = "BLOQUEADOS_MANUTENCAO";
      histNome = "HIST_MANUTENCAO";
      row = [
        id,
        registro.leito || "",
        registro.unidade || "",
        registro.manutencaoAcionada || "",
        registro.dataInicio || "",
        registro.previsaoReparo || "",
        registro.observacao || ""
      ];
      break;

    case "BLOQUEADOS_ISOLAMENTO":
      abaNome = "BLOQUEADOS_ISOLAMENTO";
      histNome = "HIST_ISOLAMENTO";
      row = [
        id,
        registro.unidade || "",
        registro.leito || "",
        registro.paciente || "",
        registro.tipoIsolamento || "",
        registro.patologia || "",
        registro.inicioIsolamento || "",
        registro.tempoPrevisto || "",
        registro.observacao || ""
      ];
      break;

    default:
      throw new Error("Módulo inválido.");
  }

  const aba = ss.getSheetByName(abaNome);
  if (!aba) throw new Error("Aba não encontrada: " + abaNome);
  aba.appendRow(row);

  writeAction_("REGISTRO_INSERIDO", { modulo: modulo, registroId: id });

  if (histNome && shiftId) {
    const hist = ss.getSheetByName(histNome);
    if (hist) {
      const linhaHist = [shiftId, id].concat(row.slice(1));
      if (histNome === "HIST_MANUTENCAO" || histNome === "HIST_ISOLAMENTO") {
        linhaHist.push(new Date());
      }
      hist.appendRow(linhaHist);
    }
  }

  return { id: id };
}

function excluirRegistro_(modulo, id) {
  const mapa = {
    "RESERVA_CONFIRMADA": { aba: "RESERVA_CONFIRMADA", hist: "HIST_RESERVA_CONFIRMADA" },
    "PROCEDIMENTO_VASCULAR": { aba: "PROCEDIMENTO_VASCULAR", hist: "HIST_PROCEDIMENTO_VASCULAR" },
    "RESERVA_NEGADA": { aba: "RESERVA_NEGADA", hist: "HIST_RESERVA_NEGADA" },
    "PLANTAO_ANTERIOR": { aba: "PLANTAO_ANTERIOR", hist: null },
    "BLOQUEADOS_MANUTENCAO": { aba: "BLOQUEADOS_MANUTENCAO", hist: "HIST_MANUTENCAO" },
    "BLOQUEADOS_ISOLAMENTO": { aba: "BLOQUEADOS_ISOLAMENTO", hist: "HIST_ISOLAMENTO" }
  };

  const conf = mapa[modulo];
  if (!conf) throw new Error("Módulo inválido.");

  const ss = getSpreadsheet_();
  const sheet = ss.getSheetByName(conf.aba);
  if (!sheet) throw new Error("Aba não encontrada: " + conf.aba);

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  const values = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();
  const shift = getActiveShift_();
  const shiftId = shift && shift.id ? shift.id : "";

  for (let i = 0; i < values.length; i++) {
    if (String(values[i][0]) === String(id)) {
      const rowIndex = i + 2;
      const rowData = values[i];

      if (conf.hist && shiftId) {
        const histSheet = ss.getSheetByName(conf.hist);
        if (histSheet) {
          const registro = [shiftId, id].concat(rowData.slice(1));
          if (conf.hist === "HIST_MANUTENCAO" || conf.hist === "HIST_ISOLAMENTO") {
            registro.push(new Date());
          }
          histSheet.appendRow(registro);
        }
      }

      sheet.deleteRow(rowIndex);
      writeAction_("REGISTRO_EXCLUIDO", { modulo: modulo, registroId: id });
      break;
    }
  }
}

function onOpen() {
  const ui = getSpreadsheet_().getUi();
  ui.createMenu("Ocorrências NIR")
    .addItem("Criar Estrutura Ocorrências", "criarEstruturaNIR")
    .addItem("Abrir Ocorrências (Sidebar)", "abrirWebAppSidebar")
    .addItem("Encerrar Plantão Atual", "encerrarPlantaoLegacy")
    .addToUi();
}

function abrirWebAppSidebar() {
  assertAuthorized_();
  const html = HtmlService.createTemplateFromFile("index").evaluate()
    .setTitle("Ocorrências do Plantão – NIR")
    .setWidth(1200);
  SpreadsheetApp.getUi().showSidebar(html);
}

function doGet() {
  if (!isUserAuthorized_()) {
    return HtmlService.createHtmlOutput(
      "<p>Acesso não autorizado. Solicite permissão ao administrador.</p>"
    );
  }
  return HtmlService.createTemplateFromFile("index")
    .evaluate()
    .setTitle("Ocorrências do Plantão – NIR");
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function encerrarPlantaoLegacy() {
  endShift();
}
