const NIR_TIMEZONE = "America/Sao_Paulo";
const NIR_SPREADSHEET_PROPERTY = "PLANILHA_ID";
const NIR_ACTIVE_SHIFT_KEY = "ACTIVE_SHIFT_ID";

const NIR_SHEETS = {
  config: {
    name: "NIR_CONFIG",
    headers: ["CHAVE", "VALOR", "ATUALIZADO_EM", "ATUALIZADO_POR"]
  },
  shifts: {
    name: "NIR_PLANTOES",
    headers: [
      "id", "data", "dia_semana", "turno", "status",
      "medico_ilha_1", "medico_ilha_2", "medico_ilha_3",
      "enfermeiro_ilha_1", "enfermeiro_ilha_2", "enfermeiro_ilha_3",
      "adm_ilha_1", "adm_ilha_2", "adm_ilha_3",
      "coordenacao", "aberto_em", "aberto_por", "encerrado_em", "encerrado_por", "observacao"
    ]
  },
  regulations: {
    name: "NIR_REGULACOES",
    headers: [
      "id", "created_at", "updated_at", "plantao_id", "data_solicitacao", "turno",
      "origem_regulacao", "fastmedic", "prontuario", "paciente_nome", "paciente_alias",
      "ads", "municipio", "unidade_origem", "especialidade", "status",
      "avaliado_por", "justificativa", "observacao", "sincronizacao",
      "vaga_dada_em", "reservado_em", "chegou_em", "internado_em", "cancelado_em",
      "prioridade", "flags", "deleted_at"
    ]
  },
  beds: {
    name: "NIR_LEITOS_UIB",
    headers: [
      "id", "created_at", "updated_at", "plantao_id", "unidade", "leito", "status_leito",
      "paciente_nome", "paciente_alias", "prontuario", "fastmedic", "data_internacao",
      "especialidade_origem", "especialidade_destino", "isolamento", "patologia",
      "pendencia", "permanencia_dias", "desfecho", "transferir_para", "observacao", "deleted_at"
    ]
  },
  procedures: {
    name: "NIR_PROCEDIMENTOS",
    headers: [
      "id", "created_at", "updated_at", "plantao_id", "tipo", "paciente_nome", "paciente_alias",
      "fastmedic", "origem", "apto_medico", "judicial", "avaliado_por", "agendado_por",
      "procedimento", "data_programacao", "hora_programacao", "status", "ultimo_contato",
      "observacao", "deleted_at"
    ]
  },
  icu: {
    name: "NIR_UTI_SRPA",
    headers: [
      "id", "created_at", "updated_at", "plantao_id", "tipo", "prontuario", "paciente_nome",
      "paciente_alias", "leito_atual", "especialidade", "status", "destino", "comunicado_hora",
      "observacao", "deleted_at"
    ]
  },
  blocks: {
    name: "NIR_BLOQUEIOS",
    headers: [
      "id", "created_at", "updated_at", "plantao_id", "tipo_bloqueio", "unidade", "leito",
      "paciente_nome", "paciente_alias", "motivo", "data_inicio", "previsao", "status",
      "observacao", "encerrado_em", "deleted_at"
    ]
  },
  contacts: {
    name: "NIR_CONTATOS",
    headers: ["id", "setor", "ramal", "nome", "telefone", "grupo", "observacao", "deleted_at"]
  },
  reports: {
    name: "NIR_RELATORIOS",
    headers: [
      "id", "plantao_id", "gerado_em", "gerado_por", "filtros_json", "kpis_json",
      "relatorio_texto", "snapshot_json"
    ]
  },
  log: {
    name: "NIR_LOG",
    headers: [
      "id", "data_hora", "usuario", "plantao_id", "modulo", "acao",
      "registro_id", "antes_json", "depois_json", "observacao"
    ]
  }
};

const NIR_FINAL_STATUSES = [
  "INTERNADO", "NEGADO", "CANCELADO", "CONCLUIDO", "REALIZADO", "ALTA", "OBITO", "TRANSFERIDO"
];

const NIR_SPECIALTIES = [
  "Clinica Medica", "Geriatria", "Reumatologia", "Pneumologia", "Hematologia",
  "Oncologia Clinica", "Cirurgia Oncologica", "Cirurgia Geral", "Coloproctologia",
  "Cirurgia Digestiva", "Cirurgia Ginecologica", "Cabeca e Pescoco", "Cirurgia Vascular",
  "Urologia", "Ortopedia", "Obstetricia", "UCINCO", "UTI Neonatal", "UTI Adulto",
  "Cardiopediatria", "Cirurgia Toracica", "Exame", "Hemodialise", "SAD/EMAD/CCC"
];

function doGet() {
  return HtmlService.createTemplateFromFile("index")
    .evaluate()
    .setTitle("NIR - Passagem de Plantao")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("NIR")
    .addItem("Criar/atualizar estrutura", "criarEstruturaNIR")
    .addItem("Abrir app em sidebar", "abrirSidebarNIR")
    .addItem("Inserir dados demo", "seedDemoDataNIR")
    .addToUi();
}

function abrirSidebarNIR() {
  const html = HtmlService.createTemplateFromFile("index").evaluate()
    .setTitle("NIR - Passagem de Plantao")
    .setWidth(1200);
  SpreadsheetApp.getUi().showSidebar(html);
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function criarEstruturaNIR() {
  return runNir_(function () {
    ensureAllSheets_();
    return { created: true, sheets: Object.keys(NIR_SHEETS).map(function (key) { return NIR_SHEETS[key].name; }) };
  });
}

function getAppState(payload) {
  return runNir_(function () {
    // Caminho de abertura: nao garantimos as 10 abas aqui. Cada readObjects_
    // garante sua propria aba (lazy), evitando tocar abas nao usadas no painel
    // (contatos, relatorios, log) e cortando varias idas a planilha.
    const filters = payload && payload.filters ? payload.filters : {};
    const activeShift = getActiveShift_();
    const options = getOptions_();
    const dashboard = buildDashboard_(filters);
    return {
      activeShift: activeShift,
      options: options,
      dashboard: dashboard,
      user: getUser_(),
      generatedAt: nowIso_()
    };
  });
}

function startShift(payload) {
  return runNir_(function () {
    ensureAllSheets_();
    const existing = getActiveShift_();
    if (existing) return { activeShift: existing, reused: true };

    const data = payload || {};
    const id = generateId_("PLT");
    const openedAt = nowIso_();
    const shiftDate = normalizeDate_(data.data || openedAt);
    const row = {
      id: id,
      data: shiftDate,
      dia_semana: dayName_(shiftDate),
      turno: normalizeTurno_(data.turno),
      status: "ABERTO",
      medico_ilha_1: data.medico_ilha_1 || "",
      medico_ilha_2: data.medico_ilha_2 || "",
      medico_ilha_3: data.medico_ilha_3 || "",
      enfermeiro_ilha_1: data.enfermeiro_ilha_1 || "",
      enfermeiro_ilha_2: data.enfermeiro_ilha_2 || "",
      enfermeiro_ilha_3: data.enfermeiro_ilha_3 || "",
      adm_ilha_1: data.adm_ilha_1 || "",
      adm_ilha_2: data.adm_ilha_2 || "",
      adm_ilha_3: data.adm_ilha_3 || "",
      coordenacao: data.coordenacao || "",
      aberto_em: openedAt,
      aberto_por: getUser_().email,
      encerrado_em: "",
      encerrado_por: "",
      observacao: data.observacao || ""
    };
    appendObject_("shifts", row);
    setConfig_(NIR_ACTIVE_SHIFT_KEY, id);
    writeLog_("plantao", "PLANTAO_ABERTO", id, null, row, "");
    return { activeShift: row };
  });
}

function updateShift(payload) {
  return runNir_(function () {
    ensureAllSheets_();
    const active = getActiveShift_();
    if (!active) throw new Error("Nenhum plantao ativo.");
    const next = Object.assign({}, active, payload || {});
    next.data = normalizeDate_(next.data || active.data);
    next.dia_semana = dayName_(next.data);
    next.turno = normalizeTurno_(next.turno);
    updateObject_("shifts", active.id, next);
    writeLog_("plantao", "PLANTAO_ATUALIZADO", active.id, active, next, "");
    return { activeShift: next };
  });
}

function closeShift(payload) {
  return runNir_(function () {
    ensureAllSheets_();
    const active = getActiveShift_();
    if (!active) return { activeShift: null, alreadyClosed: true };

    const closed = Object.assign({}, active, {
      status: "ENCERRADO",
      encerrado_em: nowIso_(),
      encerrado_por: getUser_().email
    });
    updateObject_("shifts", active.id, closed);
    setConfig_(NIR_ACTIVE_SHIFT_KEY, "");
    const report = buildCloseShiftSnapshotReport_(closed, payload || {});
    appendObject_("reports", {
      id: report.id,
      plantao_id: report.plantao_id,
      gerado_em: report.generatedAt,
      gerado_por: report.generatedBy,
      filtros_json: jsonStringify_(report.filters),
      kpis_json: jsonStringify_(report.kpis),
      relatorio_texto: report.text,
      snapshot_json: jsonStringify_(report.sections)
    });
    writeLog_("plantao", "PLANTAO_ENCERRADO", active.id, active, closed, "Relatorio salvo: " + report.id);
    return { activeShift: null, report: report };
  });
}

function listRecords(payload) {
  return runNir_(function () {
    const moduleKey = payload && payload.module ? payload.module : "regulations";
    const filters = payload && payload.filters ? payload.filters : {};
    return { rows: filterRows_(readObjects_(moduleKey), filters), module: moduleKey };
  });
}

function saveRecord(payload) {
  return runNir_(function () {
    ensureAllSheets_();
    const moduleKey = payload && payload.module ? payload.module : "";
    const data = payload && payload.record ? payload.record : {};
    if (!NIR_SHEETS[moduleKey]) throw new Error("Modulo invalido: " + moduleKey);

    const active = getActiveShift_();
    const now = nowIso_();
    const before = data.id ? findObject_(moduleKey, data.id) : null;
    const record = normalizeRecord_(moduleKey, data, active, now);

    if (before) {
      record.created_at = before.created_at || record.created_at;
      updateObject_(moduleKey, record.id, record);
      writeLog_(moduleKey, "REGISTRO_ATUALIZADO", record.id, before, record, "");
    } else {
      appendObject_(moduleKey, record);
      writeLog_(moduleKey, "REGISTRO_CRIADO", record.id, null, record, "");
    }
    return { record: record };
  });
}

function deleteRecord(payload) {
  return runNir_(function () {
    ensureAllSheets_();
    const moduleKey = payload && payload.module ? payload.module : "";
    const id = payload && payload.id ? payload.id : "";
    const before = findObject_(moduleKey, id);
    if (!before) return { deleted: false };
    const after = Object.assign({}, before, { deleted_at: nowIso_(), updated_at: nowIso_() });
    updateObject_(moduleKey, id, after);
    writeLog_(moduleKey, "REGISTRO_ARQUIVADO", id, before, after, "Soft delete");
    return { deleted: true };
  });
}

function generateReport(payload) {
  return runNir_(function () {
    return generateReportInternal_(payload || {});
  });
}

function listLog(payload) {
  return runNir_(function () {
    ensureSheet_("log");
    const opts = payload || {};
    const limit = opts.limit && opts.limit > 0 ? opts.limit : 100;
    return readRecentLogRows_(opts, limit);
  });
}

function seedDemoDataNIR() {
  return runNir_(function () {
    ensureAllSheets_();
    const active = getActiveShift_() || startShift({
      data: normalizeDate_(nowIso_()),
      turno: "MT",
      medico_ilha_1: "Dra. Exemplo A",
      medico_ilha_2: "Dr. Exemplo B",
      enfermeiro_ilha_1: "Enf. Exemplo A",
      enfermeiro_ilha_2: "Enf. Exemplo B",
      adm_ilha_1: "Adm. Exemplo"
    }).data.activeShift;

    const demoRegulations = [
      {
        origem_regulacao: "CRL", fastmedic: "FM-DEMO-001", paciente_alias: "Paciente 001",
        data_solicitacao: normalizeDate_(nowIso_()), turno: "MT", ads: "ADS Fortaleza",
        municipio: "Fortaleza", unidade_origem: "Hospital de Origem A",
        especialidade: "Obstetricia", status: "RESERVADO", avaliado_por: "Dr. Exemplo B",
        justificativa: "Nao se aplica", observacao: "Aguardar chegada e confirmar leito.",
        reservado_em: nowIso_(), prioridade: "alta"
      },
      {
        origem_regulacao: "SAMU", fastmedic: "FM-DEMO-002", paciente_alias: "Paciente 002",
        data_solicitacao: normalizeDate_(nowIso_()), turno: "MT", ads: "ADS Sertao",
        municipio: "Municipio B", unidade_origem: "UPA Exemplo",
        especialidade: "Cirurgia Vascular", status: "NEGADO", avaliado_por: "Dra. Exemplo A",
        justificativa: "Dentro do perfil, mas sem vagas", observacao: "Manter em tela para reavaliacao.",
        prioridade: "media"
      }
    ];
    demoRegulations.forEach(function (item) { saveRecord({ module: "regulations", record: item }); });

    [
      {
        unidade: "UIB Vascular", leito: "L01", status_leito: "INTERNADO", paciente_alias: "Paciente 003",
        data_internacao: normalizeDate_(nowIso_()), especialidade_origem: "Cirurgia Vascular",
        especialidade_destino: "Cirurgia Vascular", pendencia: "Aguardar procedimento", permanencia_dias: 2
      },
      {
        unidade: "UIB Onco/Hemato", leito: "L04", status_leito: "RESERVADO", paciente_alias: "Paciente 004",
        especialidade_destino: "Oncologia Clinica", pendencia: "Falta chegar no leito", permanencia_dias: 0
      }
    ].forEach(function (item) { saveRecord({ module: "beds", record: item }); });

    [
      {
        tipo: "Imagem", paciente_alias: "Paciente 005", origem: "Hospital Externo",
        apto_medico: "SIM", avaliado_por: "Dra. Exemplo A", agendado_por: "Adm. Exemplo",
        procedimento: "TC de cranio", data_programacao: normalizeDate_(nowIso_()),
        hora_programacao: "08:00", status: "AGENDADO", observacao: "Com sedacao."
      }
    ].forEach(function (item) { saveRecord({ module: "procedures", record: item }); });

    [
      {
        tipo: "ALTA_UTI_AGUARDANDO_LEITO", prontuario: "PR-DEMO-001", paciente_alias: "Paciente 006",
        leito_atual: "UTI Adulto II", especialidade: "Cirurgia Geral", status: "ALTA_24H",
        observacao: "Aguardando leito de enfermaria."
      }
    ].forEach(function (item) { saveRecord({ module: "icu", record: item }); });

    [
      {
        tipo_bloqueio: "MANUTENCAO", unidade: "Ala B", leito: "B511.01",
        motivo: "Manutencao predial", data_inicio: normalizeDate_(nowIso_()),
        previsao: "", status: "ABERTO", observacao: "Acompanhar reparo."
      }
    ].forEach(function (item) { saveRecord({ module: "blocks", record: item }); });

    return { seeded: true, activeShift: active };
  });
}

function runNir_(fn) {
  try {
    return { ok: true, data: fn() };
  } catch (err) {
    return {
      ok: false,
      error: {
        message: err && err.message ? err.message : String(err),
        stack: err && err.stack ? String(err.stack) : ""
      }
    };
  }
}

// Caches por execucao: evitam reabrir a planilha e re-garantir abas a cada
// chamada interna. Em Apps Script o estado global vive durante uma execucao,
// entao isso elimina dezenas de openById/getSheetByName redundantes por request.
let NIR_SPREADSHEET_CACHE_ = null;
const NIR_SHEET_CACHE_ = {};
const NIR_ENSURED_SHEETS_ = {};

function getSpreadsheet_() {
  if (NIR_SPREADSHEET_CACHE_) return NIR_SPREADSHEET_CACHE_;
  const configuredId = PropertiesService.getScriptProperties().getProperty(NIR_SPREADSHEET_PROPERTY);
  if (configuredId) {
    NIR_SPREADSHEET_CACHE_ = SpreadsheetApp.openById(configuredId);
    return NIR_SPREADSHEET_CACHE_;
  }
  const active = SpreadsheetApp.getActiveSpreadsheet();
  if (!active) throw new Error("Configure PLANILHA_ID nas propriedades do script.");
  NIR_SPREADSHEET_CACHE_ = active;
  return NIR_SPREADSHEET_CACHE_;
}

function ensureAllSheets_() {
  Object.keys(NIR_SHEETS).forEach(function (key) {
    ensureSheet_(key);
  });
}

function ensureSheet_(key) {
  if (NIR_ENSURED_SHEETS_[key]) return NIR_SHEET_CACHE_[key];

  const spec = NIR_SHEETS[key];
  const ss = getSpreadsheet_();
  let sheet = ss.getSheetByName(spec.name);
  let isNew = false;
  if (!sheet) {
    sheet = ss.insertSheet(spec.name);
    isNew = true;
  }
  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, spec.headers.length).setValues([spec.headers]);
    sheet.setFrozenRows(1);
  } else {
    const current = sheet.getRange(1, 1, 1, Math.max(spec.headers.length, sheet.getLastColumn())).getValues()[0];
    const changed = spec.headers.some(function (h, i) { return current[i] !== h; });
    if (changed) sheet.getRange(1, 1, 1, spec.headers.length).setValues([spec.headers]);
    // Congelar linha e escrita/formatacao cara: so aplica quando necessario,
    // nunca em todo caminho de leitura como antes.
    if (isNew || sheet.getFrozenRows() < 1) sheet.setFrozenRows(1);
  }

  NIR_ENSURED_SHEETS_[key] = true;
  NIR_SHEET_CACHE_[key] = sheet;
  return sheet;
}

function getUser_() {
  let email = "";
  try {
    email = Session.getActiveUser().getEmail() || "";
  } catch (err) {
    email = "";
  }
  return { email: email, timezone: NIR_TIMEZONE };
}

function getConfig_(key) {
  const rows = readObjects_("config", { includeDeleted: true });
  const found = rows.find(function (row) { return row.CHAVE === key; });
  return found ? found.VALOR : "";
}

function setConfig_(key, value) {
  const sheet = ensureSheet_("config");
  const values = sheet.getDataRange().getValues();
  const now = nowIso_();
  const user = getUser_().email;
  for (let i = 1; i < values.length; i++) {
    if (values[i][0] === key) {
      sheet.getRange(i + 1, 2, 1, 3).setValues([[value, now, user]]);
      return;
    }
  }
  sheet.appendRow([key, value, now, user]);
}

function getActiveShift_() {
  const id = getConfig_(NIR_ACTIVE_SHIFT_KEY);
  if (!id) return null;
  const shift = findObject_("shifts", id);
  if (!shift || shift.status !== "ABERTO") return null;
  return shift;
}

function readObjects_(moduleKey, options) {
  const spec = NIR_SHEETS[moduleKey];
  if (!spec) throw new Error("Modulo invalido: " + moduleKey);
  const sheet = ensureSheet_(moduleKey);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  const values = sheet.getRange(1, 1, lastRow, spec.headers.length).getValues();
  const headers = values[0];
  return values.slice(1)
    .filter(function (row) { return row.some(function (cell) { return cell !== "" && cell !== null; }); })
    .map(function (row) {
      const obj = {};
      headers.forEach(function (header, index) {
        obj[header] = normalizeCell_(row[index], header);
      });
      return obj;
    })
    .filter(function (obj) {
      return options && options.includeDeleted ? true : !obj.deleted_at;
    });
}

function readRecentLogRows_(opts, limit) {
  const spec = NIR_SHEETS.log;
  const sheet = ensureSheet_("log");
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return { rows: [], total: 0 };

  const headers = sheet.getRange(1, 1, 1, spec.headers.length).getValues()[0];
  const rows = [];
  const total = lastRow - 1;
  const chunkSize = Math.min(Math.max(limit * 5, 200), 1000);

  for (let end = lastRow; end >= 2 && rows.length < limit; end -= chunkSize) {
    const start = Math.max(2, end - chunkSize + 1);
    const values = sheet.getRange(start, 1, end - start + 1, spec.headers.length).getValues();
    for (let i = values.length - 1; i >= 0 && rows.length < limit; i--) {
      const raw = values[i];
      if (!raw.some(function (cell) { return cell !== "" && cell !== null; })) continue;
      const obj = {};
      headers.forEach(function (header, index) {
        obj[header] = normalizeCell_(raw[index], header);
      });
      if (opts.module && obj.modulo !== opts.module) continue;
      if (opts.plantao_id && obj.plantao_id !== opts.plantao_id) continue;
      rows.push(obj);
    }
  }

  return { rows: rows, total: total };
}

function appendObject_(moduleKey, obj) {
  const spec = NIR_SHEETS[moduleKey];
  const sheet = ensureSheet_(moduleKey);
  const row = spec.headers.map(function (header) {
    return obj[header] === undefined || obj[header] === null ? "" : obj[header];
  });
  sheet.appendRow(row);
}

function updateObject_(moduleKey, id, obj) {
  const spec = NIR_SHEETS[moduleKey];
  const sheet = ensureSheet_(moduleKey);
  const values = sheet.getRange(1, 1, sheet.getLastRow(), spec.headers.length).getValues();
  const idCol = spec.headers.indexOf("id");
  if (idCol < 0) throw new Error("Modulo sem coluna id: " + moduleKey);
  for (let i = 1; i < values.length; i++) {
    if (String(values[i][idCol]) === String(id)) {
      const row = spec.headers.map(function (header) {
        return obj[header] === undefined || obj[header] === null ? "" : obj[header];
      });
      sheet.getRange(i + 1, 1, 1, spec.headers.length).setValues([row]);
      return;
    }
  }
  appendObject_(moduleKey, obj);
}

function findObject_(moduleKey, id) {
  if (!id) return null;
  return readObjects_(moduleKey, { includeDeleted: true })
    .find(function (row) { return String(row.id) === String(id); }) || null;
}

function normalizeRecord_(moduleKey, data, activeShift, now) {
  const spec = NIR_SHEETS[moduleKey];
  const record = {};
  spec.headers.forEach(function (header) {
    record[header] = data[header] === undefined || data[header] === null ? "" : data[header];
  });
  record.id = record.id || generateId_(moduleKey.substring(0, 3).toUpperCase());
  record.created_at = record.created_at || now;
  record.updated_at = now;
  record.plantao_id = record.plantao_id || (activeShift ? activeShift.id : "");
  if (record.turno) record.turno = normalizeTurno_(record.turno);
  if (!record.turno && activeShift) record.turno = activeShift.turno;
  return record;
}

function filterRows_(rows, filters) {
  const f = filters || {};
  const term = normalizeText_(f.search || "");
  return rows.filter(function (row) {
    if (f.plantao_id && row.plantao_id !== f.plantao_id) return false;
    if (f.turno && f.turno !== "TODOS" && row.turno && row.turno !== f.turno) return false;
    if (f.statuses && f.statuses.length && f.statuses.indexOf(row.status || row.status_leito || row.tipo_bloqueio) === -1) return false;
    if (f.specialties && f.specialties.length && !matchesAny_(row.especialidade || row.especialidade_destino || row.especialidade_origem, f.specialties)) return false;
    if (f.origens && f.origens.length && !matchesAny_(row.origem_regulacao || row.origem || row.unidade_origem, f.origens)) return false;
    if (f.unidade && !containsText_(row.unidade || row.leito_atual || "", f.unidade)) return false;
    if (f.onlyPending && isFinalStatus_(row.status || row.status_leito || row.desfecho)) return false;
    if (f.dateFrom || f.dateTo) {
      const rowDate = firstDate_(row, ["data_solicitacao", "data_programacao", "data_internacao", "data_inicio", "created_at"]);
      if (!rowDate) return false;
      if (f.dateFrom && rowDate < normalizeDate_(f.dateFrom)) return false;
      if (f.dateTo && rowDate > normalizeDate_(f.dateTo)) return false;
    }
    if (term) {
      const haystack = normalizeText_(Object.keys(row).map(function (key) { return row[key]; }).join(" "));
      if (haystack.indexOf(term) === -1) return false;
    }
    return true;
  });
}

function buildDashboard_(filters) {
  const regulations = filterRows_(readObjects_("regulations"), filters || {});
  const beds = filterRows_(readObjects_("beds"), filters || {});
  const procedures = filterRows_(readObjects_("procedures"), filters || {});
  const icu = filterRows_(readObjects_("icu"), filters || {});
  const blocks = filterRows_(readObjects_("blocks"), filters || {});

  const kpis = calculateKpis_(regulations, beds, procedures, icu, blocks);
  return {
    kpis: kpis,
    bySpecialty: groupCount_(regulations, "especialidade"),
    byStatus: groupCount_(regulations, "status"),
    byOrigin: groupCount_(regulations, "origem_regulacao"),
    latestRegulations: regulations.slice(-12).reverse(),
    pendingBeds: beds.filter(function (row) { return !isFinalStatus_(row.status_leito) && row.status_leito !== "LIVRE"; }).slice(-12).reverse(),
    pendingProcedures: procedures.filter(function (row) { return !isFinalStatus_(row.status); }).slice(-12).reverse(),
    icu: icu.slice(-12).reverse(),
    blocks: blocks.filter(function (row) { return row.status !== "ENCERRADO"; }).slice(-12).reverse()
  };
}

function generateReportInternal_(payload) {
  const filters = payload && payload.filters ? payload.filters : {};
  const regulations = filterRows_(readObjects_("regulations"), filters);
  const beds = filterRows_(readObjects_("beds"), filters);
  const procedures = filterRows_(readObjects_("procedures"), filters);
  const icu = filterRows_(readObjects_("icu"), filters);
  const blocks = filterRows_(readObjects_("blocks"), filters);
  const activeShift = getActiveShift_();
  const kpis = calculateKpis_(regulations, beds, procedures, icu, blocks);

  const report = {
    id: generateId_("REL"),
    plantao_id: filters.plantao_id || (activeShift ? activeShift.id : ""),
    generatedAt: nowIso_(),
    generatedBy: getUser_().email,
    filters: filters,
    kpis: kpis,
    sections: {
      resumo: buildSummaryLines_(kpis),
      regulacoes: regulations,
      leitos: beds,
      procedimentos: procedures,
      uti: icu,
      bloqueios: blocks,
      especialidades: groupCount_(regulations, "especialidade"),
      status: groupCount_(regulations, "status")
    }
  };
  report.text = buildReportText_(report);

  if (payload && payload.save) {
    appendObject_("reports", {
      id: report.id,
      plantao_id: report.plantao_id,
      gerado_em: report.generatedAt,
      gerado_por: report.generatedBy,
      filtros_json: jsonStringify_(filters),
      kpis_json: jsonStringify_(kpis),
      relatorio_texto: report.text,
      snapshot_json: jsonStringify_(report.sections)
    });
    writeLog_("relatorio", "RELATORIO_GERADO", report.id, null, report, "");
  }

  return report;
}

function buildCloseShiftSnapshotReport_(closedShift, payload) {
  const filters = Object.assign({}, payload && payload.filters ? payload.filters : {}, { plantao_id: closedShift.id });
  const generatedAt = nowIso_();
  const report = {
    id: generateId_("REL"),
    plantao_id: closedShift.id,
    generatedAt: generatedAt,
    generatedBy: getUser_().email,
    filters: filters,
    kpis: {},
    sections: {
      plantao: closedShift,
      resumo: [
        "Plantao encerrado em " + formatDateTime_(closedShift.encerrado_em),
        "Snapshot leve salvo para nao bloquear o encerramento."
      ]
    }
  };
  report.text = [
    "RELATORIO DE ENCERRAMENTO DE PLANTAO - NIR",
    "Gerado em: " + formatDateTime_(generatedAt),
    "Plantao: " + formatShiftIdForReport_(closedShift.id),
    "Data: " + formatDateOnly_(closedShift.data),
    "Turno: " + (closedShift.turno || "-"),
    "Encerrado por: " + (closedShift.encerrado_por || "-"),
    "Encerrado em: " + formatDateTime_(closedShift.encerrado_em),
    "",
    "O relatorio operacional detalhado pode ser gerado na aba Relatorio."
  ].join("\n");
  return report;
}

function calculateKpis_(regulations, beds, procedures, icu, blocks) {
  const reserved = regulations.filter(function (row) { return row.status === "RESERVADO"; }).length;
  const admitted = regulations.filter(function (row) { return row.status === "INTERNADO"; }).length;
  const denied = regulations.filter(function (row) { return row.status === "NEGADO"; }).length;
  const canceled = regulations.filter(function (row) { return row.status === "CANCELADO"; }).length;
  const pendingReg = regulations.filter(function (row) { return !isFinalStatus_(row.status); }).length;
  const occupiedBeds = beds.filter(function (row) { return row.status_leito && row.status_leito !== "LIVRE"; }).length;
  const blockedBeds = blocks.filter(function (row) { return row.status !== "ENCERRADO"; }).length;
  const waitingIcuDischarge = icu.filter(function (row) { return row.tipo === "ALTA_UTI_AGUARDANDO_LEITO"; }).length;

  return {
    regulacoes: regulations.length,
    reservados: reserved,
    internados: admitted,
    negados: denied,
    cancelados: canceled,
    pendentes: pendingReg,
    leitos_ocupados_ou_reservados: occupiedBeds,
    procedimentos_pendentes: procedures.filter(function (row) { return !isFinalStatus_(row.status); }).length,
    altas_uti_aguardando_leito: waitingIcuDischarge,
    bloqueios_abertos: blockedBeds,
    tempo_medio_reserva_horas: averageHours_(regulations, "vaga_dada_em", "reservado_em"),
    tempo_medio_chegada_horas: averageHours_(regulations, "reservado_em", "chegou_em"),
    permanencia_media_dias: averageNumber_(beds, "permanencia_dias")
  };
}

function buildSummaryLines_(kpis) {
  return [
    "Regulacoes no filtro: " + kpis.regulacoes,
    "Reservados: " + kpis.reservados + " | Internados: " + kpis.internados + " | Negados: " + kpis.negados + " | Cancelados: " + kpis.cancelados,
    "Pendencias abertas: " + kpis.pendentes + " | Procedimentos pendentes: " + kpis.procedimentos_pendentes,
    "Leitos ocupados/reservados em UIB: " + kpis.leitos_ocupados_ou_reservados + " | Bloqueios abertos: " + kpis.bloqueios_abertos,
    "Altas de UTI aguardando leito: " + kpis.altas_uti_aguardando_leito
  ];
}

function buildReportText_(report) {
  const only = report.filters && report.filters.reportModule ? report.filters.reportModule : "";
  const show = function (moduleKey) { return !only || only === moduleKey; };
  const lines = [];
  lines.push("RELATORIO DE PASSAGEM DE PLANTAO - NIR");
  lines.push("Gerado em: " + formatDateTime_(report.generatedAt));
  if (report.plantao_id) lines.push("Plantao: " + formatShiftIdForReport_(report.plantao_id));
  if (only) lines.push("Modulo: " + only);
  lines.push("");
  lines.push("RESUMO");
  report.sections.resumo.forEach(function (line) { lines.push("- " + line); });
  if (show("regulations")) {
    lines.push("");
    lines.push("REGULACOES POR STATUS");
    report.sections.status.forEach(function (item) { lines.push("- " + item.label + ": " + item.count); });
    lines.push("");
    lines.push("REGULACOES POR ESPECIALIDADE");
    report.sections.especialidades.forEach(function (item) { lines.push("- " + item.label + ": " + item.count); });
    lines.push("");
    lines.push("PENDENCIAS DE REGULACAO");
    report.sections.regulacoes.filter(function (row) { return !isFinalStatus_(row.status); }).slice(0, 40).forEach(function (row) {
      lines.push("- " + compactPatient_(row) + " | " + (row.especialidade || "-") + " | " + (row.status || "-") + " | " + (row.observacao || ""));
    });
  }
  if (show("beds")) {
    lines.push("");
    lines.push("UIB / LEITOS");
    report.sections.leitos.filter(function (row) { return row.status_leito !== "LIVRE"; }).slice(0, 40).forEach(function (row) {
      lines.push("- " + (row.unidade || "-") + " " + (row.leito || "-") + " | " + compactPatient_(row) + " | " + (row.status_leito || "-") + " | " + (row.pendencia || row.observacao || ""));
    });
  }
  if (show("procedures")) {
    lines.push("");
    lines.push("PROCEDIMENTOS E EXAMES");
    report.sections.procedimentos.filter(function (row) { return !isFinalStatus_(row.status); }).slice(0, 40).forEach(function (row) {
      lines.push("- " + (row.tipo || "-") + " | " + compactPatient_(row) + " | " + (row.procedimento || "-") + " | " + (row.data_programacao || "-") + " " + (row.hora_programacao || "") + " | " + (row.status || "-"));
    });
  }
  if (show("icu")) {
    lines.push("");
    lines.push("UTI / SRPA");
    report.sections.uti.slice(0, 40).forEach(function (row) {
      lines.push("- " + (row.tipo || "-") + " | " + compactPatient_(row) + " | " + (row.leito_atual || "-") + " -> " + (row.destino || "-") + " | " + (row.status || "-"));
    });
  }
  if (show("blocks")) {
    lines.push("");
    lines.push("BLOQUEIOS");
    report.sections.bloqueios.filter(function (row) { return row.status !== "ENCERRADO"; }).slice(0, 40).forEach(function (row) {
      lines.push("- " + (row.tipo_bloqueio || "-") + " | " + (row.unidade || "-") + " " + (row.leito || "-") + " | " + (row.motivo || row.observacao || ""));
    });
  }
  return lines.join("\n");
}

function formatShiftIdForReport_(shiftId) {
  return String(shiftId || "").replace(/(\d{4})(\d{2})(\d{2})/g, function (_, year, month, day) {
    return day + month + year;
  });
}

function compactPatient_(row) {
  return row.paciente_alias || row.paciente_nome || row.nome || row.prontuario || row.fastmedic || "Paciente sem identificador";
}

function getOptions_() {
  return {
    modules: [
      { key: "regulations", label: "Regulacao" },
      { key: "beds", label: "UIB / Leitos" },
      { key: "procedures", label: "Procedimentos" },
      { key: "icu", label: "UTI / SRPA" },
      { key: "blocks", label: "Bloqueios" }
    ],
    turns: ["MT", "SN", "TODOS"],
    origins: ["CRL", "SAMU", "LOCAL", "SIREG", "AMBULATORIO"],
    statuses: [
      "NOVO", "EM_AVALIACAO", "AGUARDANDO_REGULACAO", "RESERVADO", "FALTA_CHEGAR",
      "INTERNADO", "NEGADO", "CANCELADO", "AGENDADO", "REAGENDADO", "REALIZADO",
      "CONCLUIDO", "PENDENTE", "LIVRE", "BLOQUEADO", "ALTA_24H", "AGUARDA_LEITO"
    ],
    specialties: NIR_SPECIALTIES
  };
}

function writeLog_(moduleKey, action, recordId, before, after, observation) {
  try {
    appendObject_("log", {
      id: generateId_("LOG"),
      data_hora: nowIso_(),
      usuario: getUser_().email,
      plantao_id: after && after.plantao_id ? after.plantao_id : before && before.plantao_id ? before.plantao_id : getConfig_(NIR_ACTIVE_SHIFT_KEY),
      modulo: moduleKey,
      acao: action,
      registro_id: recordId || "",
      antes_json: before ? jsonStringify_(before) : "",
      depois_json: after ? jsonStringify_(after) : "",
      observacao: observation || ""
    });
  } catch (err) {
    Logger.log("Falha ao registrar log NIR: " + err);
  }
}

function normalizeCell_(value, header) {
  if (value === null || value === undefined) return "";
  if (Object.prototype.toString.call(value) === "[object Date]") {
    const pattern = isDateOnlyField_(header) ? "yyyy-MM-dd" : "yyyy-MM-dd'T'HH:mm:ss";
    return Utilities.formatDate(value, NIR_TIMEZONE, pattern);
  }
  return value;
}

function isDateOnlyField_(header) {
  return [
    "data",
    "data_solicitacao",
    "data_internacao",
    "data_programacao",
    "ultimo_contato",
    "data_inicio",
    "previsao"
  ].indexOf(header) !== -1;
}

function generateId_(prefix) {
  const raw = Utilities.getUuid().split("-")[0].toUpperCase();
  return (prefix || "NIR") + "-" + raw;
}

function nowIso_() {
  return Utilities.formatDate(new Date(), NIR_TIMEZONE, "yyyy-MM-dd'T'HH:mm:ss");
}

function normalizeDate_(value) {
  if (!value) return "";
  if (Object.prototype.toString.call(value) === "[object Date]") {
    return Utilities.formatDate(value, NIR_TIMEZONE, "yyyy-MM-dd");
  }
  const text = String(value).trim();
  if (/^\d{4}-\d{2}-\d{2}/.test(text)) return text.substring(0, 10);
  const match = text.match(/^(\d{1,2})[\/.-](\d{1,2})[\/.-](\d{2,4})/);
  if (match) {
    let year = Number(match[3]);
    if (year < 100) year += 2000;
    return [year, pad2_(match[2]), pad2_(match[1])].join("-");
  }
  const date = new Date(text);
  if (isNaN(date)) return "";
  return Utilities.formatDate(date, NIR_TIMEZONE, "yyyy-MM-dd");
}

function dayName_(value) {
  const date = new Date(normalizeDate_(value) + "T00:00:00");
  const names = ["Domingo", "Segunda-feira", "Terca-feira", "Quarta-feira", "Quinta-feira", "Sexta-feira", "Sabado"];
  return isNaN(date) ? "" : names[date.getDay()];
}

function normalizeTurno_(value) {
  const text = String(value || "").trim().toUpperCase();
  if (text === "M" || text === "MANHA") return "MT";
  if (text === "T" || text === "TARDE") return "MT";
  if (text === "N" || text === "NOITE") return "SN";
  if (text === "MT" || text === "SN") return text;
  return text || "MT";
}

function formatDateTime_(value) {
  if (!value) return "";
  const date = new Date(String(value).replace(" ", "T"));
  if (isNaN(date)) return String(value);
  return Utilities.formatDate(date, NIR_TIMEZONE, "dd/MM/yyyy HH:mm");
}

function formatDateOnly_(value) {
  const normalized = normalizeDate_(value);
  if (!normalized) return value ? String(value) : "";
  const parts = normalized.split("-");
  return parts.length === 3 ? [parts[2], parts[1], parts[0]].join("/") : normalized;
}

function pad2_(value) {
  return String(value).padStart(2, "0");
}

function jsonStringify_(value) {
  try {
    return JSON.stringify(value || {});
  } catch (err) {
    return "{}";
  }
}

function normalizeText_(value) {
  return String(value || "")
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .toLowerCase()
    .trim();
}

function containsText_(value, term) {
  return normalizeText_(value).indexOf(normalizeText_(term)) !== -1;
}

function matchesAny_(value, list) {
  const text = normalizeText_(value);
  return (list || []).some(function (item) { return text.indexOf(normalizeText_(item)) !== -1; });
}

function isFinalStatus_(status) {
  const text = normalizeText_(status).toUpperCase();
  return NIR_FINAL_STATUSES.some(function (item) { return normalizeText_(item).toUpperCase() === text; });
}

function firstDate_(row, keys) {
  for (let i = 0; i < keys.length; i++) {
    const normalized = normalizeDate_(row[keys[i]]);
    if (normalized) return normalized;
  }
  return "";
}

function averageHours_(rows, startKey, endKey) {
  const values = rows.map(function (row) {
    const start = toDate_(row[startKey]);
    const end = toDate_(row[endKey]);
    if (!start || !end) return null;
    return (end.getTime() - start.getTime()) / 36e5;
  }).filter(function (value) { return value !== null && !isNaN(value) && value >= 0; });
  if (!values.length) return 0;
  return Math.round((values.reduce(function (a, b) { return a + b; }, 0) / values.length) * 10) / 10;
}

function averageNumber_(rows, key) {
  const values = rows.map(function (row) { return Number(row[key]); })
    .filter(function (value) { return !isNaN(value); });
  if (!values.length) return 0;
  return Math.round((values.reduce(function (a, b) { return a + b; }, 0) / values.length) * 10) / 10;
}

function toDate_(value) {
  if (!value) return null;
  if (Object.prototype.toString.call(value) === "[object Date]") return isNaN(value) ? null : value;
  const date = new Date(String(value).replace(" ", "T"));
  return isNaN(date) ? null : date;
}

function groupCount_(rows, key) {
  const map = {};
  rows.forEach(function (row) {
    const label = row[key] || "Nao informado";
    map[label] = (map[label] || 0) + 1;
  });
  return Object.keys(map).map(function (label) {
    return { label: label, count: map[label] };
  }).sort(function (a, b) { return b.count - a.count || String(a.label).localeCompare(String(b.label)); });
}
