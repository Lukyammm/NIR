const PLANILHA_ID = "1vnnGQEkAjP9eTRLWWSb2lSngwGTRYa_rtk6DsG8HGqc";

function onOpen() {
  const ui = getSpreadsheet_().getUi();
  ui.createMenu("Ocorrências NIR")
    .addItem("Criar Estrutura Ocorrências", "criarEstruturaNIR")
    .addItem("Abrir Ocorrências (Sidebar)", "abrirWebAppSidebar")
    .addItem("Encerrar Plantão Atual", "encerrarPlantao")
    .addToUi();
}

function abrirWebAppSidebar() {
  const html = HtmlService.createHtmlOutputFromFile("index")
    .setTitle("Ocorrências do Plantão – NIR")
    .setWidth(1200);
  SpreadsheetApp.getUi().showSidebar(html);
}

function doGet() {
  return HtmlService.createHtmlOutputFromFile("index")
    .setTitle("Ocorrências do Plantão – NIR");
}

// ========================
// 1) CRIAÇÃO DE ABAS
// ========================
function criarEstruturaNIR() {
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

// ========================
// 2) PLANTÃO
// ========================

// Criar novo plantão COM equipe (usado pelo modal ao abrir plantão)
function criarPlantaoComEquipe(dados) {
  const ss = getSpreadsheet_();
  const aba = ss.getSheetByName("CONFIG_PLANTAO");
  if (!aba) return;

  const idAtual = aba.getRange("A2").getValue();
  if (idAtual) {
    // já tem plantão ativo, em teoria o front já bloqueou isso
    throw new Error("Já existe um plantão ativo. Encerre antes de abrir outro.");
  }

  const id = gerarID_();
  const agora = new Date();
  const timezone = Session.getScriptTimeZone() || "America/Sao_Paulo";
  const dataPlantao = normalizarDataSimples_(dados.data, timezone) || dados.data || "";

  const row = [
    id,
    dataPlantao,
    dados.diaSemana || "",
    dados.turno || "",
    dados.medico1 || "",
    dados.medico2 || "",
    dados.enf1 || "",
    dados.enf2 || "",
    dados.aux || "",
    agora,
    ""
  ];

  aba.getRange("A2:K2").setValues([row]);

  registrarEvento_("PLANTAO", id, "ABERTURA_PLANTAO", "Plantão aberto com equipe definida");
}

// Editar equipe/plantão já existente
function salvarPlantaoConfig(dados) {
  const ss = getSpreadsheet_();
  const aba = ss.getSheetByName("CONFIG_PLANTAO");
  if (!aba) return;

  const idAtual = aba.getRange("A2").getValue();
  if (!idAtual) {
    throw new Error("Nenhum plantão ativo para salvar.");
  }

  const timezone = Session.getScriptTimeZone() || "America/Sao_Paulo";
  const dataPlantao = normalizarDataSimples_(dados.data, timezone) || dados.data || "";

  const row = [
    idAtual,
    dataPlantao,
    dados.diaSemana || "",
    dados.turno || "",
    dados.medico1 || "",
    dados.medico2 || "",
    dados.enf1 || "",
    dados.enf2 || "",
    dados.aux || "",
    aba.getRange("J2").getValue(),
    aba.getRange("K2").getValue()
  ];

  aba.getRange("A2:K2").setValues([row]);
}

function encerrarPlantao() {
  const ss = getSpreadsheet_();
  const conf = ss.getSheetByName("CONFIG_PLANTAO");
  if (!conf) return;

  const plantaoId = conf.getRange("A2").getValue();
  if (!plantaoId) {
    SpreadsheetApp.getUi().alert("Nenhum plantão ativo em CONFIG_PLANTAO (A2).");
    return;
  }

  const encerramento = new Date();
  const hist = ss.getSheetByName("HIST_PLANTOES");

  const dados = conf.getRange("A2:K2").getValues()[0];
  dados[10] = encerramento; // Encerramento
  dados.push(Session.getActiveUser().getEmail());

  hist.appendRow(dados);
  registrarEvento_("PLANTAO", plantaoId, "ENCERRAMENTO_PLANTAO", "Plantão encerrado");

  conf.getRange("A2:K2").clearContent();
}

// ========================
// 3) FRONTEND: GET / INDICADORES
// ========================
function getAppData(viewType) {
  const ss = getSpreadsheet_();

  const viewMap = {
    "confirmadas": "RESERVA_CONFIRMADA",
    "vascular": "PROCEDIMENTO_VASCULAR",
    "negadas": "RESERVA_NEGADA",
    "anterior": "PLANTAO_ANTERIOR"
  };

  const principalSheet = viewMap[viewType] || "RESERVA_CONFIRMADA";

  return {
    plantao: getPlantaoConfig_(),
    indicadores: getIndicadores_(),
    principal: getTabela_(principalSheet),
    manutencao: getTabela_("BLOQUEADOS_MANUTENCAO"),
    isolamento: getTabela_("BLOQUEADOS_ISOLAMENTO"),
    principalNome: principalSheet
  };
}

function getPlantaoConfig_() {
  const ss = getSpreadsheet_();
  const aba = ss.getSheetByName("CONFIG_PLANTAO");
  if (!aba) return null;

  const valores = aba.getRange("A2:K2").getValues()[0];
  const possuiConteudo = valores.some((v) => v !== "" && v !== null);
  const tz = Session.getScriptTimeZone() || "America/Sao_Paulo";

  if (possuiConteudo && !valores[0]) {
    const idRecuperado = gerarID_();
    valores[0] = idRecuperado;
    aba.getRange("A2").setValue(idRecuperado);
    registrarEvento_(
      "PLANTAO",
      idRecuperado,
      "ID_RECUPERADO",
      "ID preenchido automaticamente ao detectar CONFIG_PLANTAO com dados sem ID"
    );
  }

  const aberturaDt = normalizarDataHora_(valores[9], tz);
  const encerramentoDt = normalizarDataHora_(valores[10], tz);
  const agora = new Date();
  const plantaoAtivo = Boolean(valores[0] || (possuiConteudo && (!encerramentoDt || agora < encerramentoDt)));

  return {
    id: plantaoAtivo ? valores[0] || "" : "",
    data: normalizarDataSimples_(valores[1], tz),
    diaSemana: valores[2] || "",
    turno: valores[3] || "",
    medico1: valores[4] || "",
    medico2: valores[5] || "",
    enf1: valores[6] || "",
    enf2: valores[7] || "",
    aux: valores[8] || "",
    abertura: aberturaDt || "",
    encerramento: encerramentoDt || ""
  };
}

function getTabela_(sheetName) {
  const ss = getSpreadsheet_();
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) return { headers: [], rows: [] };

  const lastCol = sheet.getLastColumn();
  const headers = lastCol > 0 ? sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(h => h || "") : [];

  if (headers.length === 0) {
    return { headers: [], rows: [] };
  }

  const dataRange = sheet.getDataRange();
  const values = dataRange.getValues();
  const linhas = values
    .slice(1)
    .filter(row => row.some(cell => cell !== "" && cell !== null))
    .map(row => {
      if (row.length < headers.length) {
        return row.concat(Array(headers.length - row.length).fill(""));
      }
      return row.slice(0, headers.length);
    });

  return { headers: headers, rows: linhas };
}

function getIndicadores_() {
  const ss = getSpreadsheet_();

  const confirmadas = ss.getSheetByName("RESERVA_CONFIRMADA");
  const vascular = ss.getSheetByName("PROCEDIMENTO_VASCULAR");
  const negadas = ss.getSheetByName("RESERVA_NEGADA");
  const anterior = ss.getSheetByName("PLANTAO_ANTERIOR");

  const countConfirmadas = confirmadas ? Math.max(confirmadas.getLastRow() - 1, 0) : 0;
  const countVascular = vascular ? Math.max(vascular.getLastRow() - 1, 0) : 0;
  const countNegadas = negadas ? Math.max(negadas.getLastRow() - 1, 0) : 0;
  const countAnterior = anterior ? Math.max(anterior.getLastRow() - 1, 0) : 0;

  return {
    alocados: countConfirmadas + countVascular,
    reservasConfirmadas: countConfirmadas,
    reservasCanceladas: countNegadas,
    admitidosUIB: countAnterior
  };
}

// ========================
// 4) INSERIR REGISTRO
// ========================
function adicionarRegistro(modulo, registro) {
  const ss = getSpreadsheet_();
  const plantao = getPlantaoConfig_();
  const plantaoId = plantao && plantao.id ? plantao.id : "";

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
      return;
  }

  const aba = ss.getSheetByName(abaNome);
  if (!aba) return;
  aba.appendRow(row);

  registrarEvento_(modulo, id, "INSERCAO", "Novo registro inserido via WebApp");

  if (histNome && plantaoId) {
    const hist = ss.getSheetByName(histNome);
    if (hist) {
      const linhaHist = [plantaoId, id].concat(row.slice(1));
      hist.appendRow(linhaHist);
    }
  }
}

// ========================
// 5) EXCLUIR REGISTRO
// ========================
function excluirRegistro(modulo, id) {
  const mapa = {
    "RESERVA_CONFIRMADA": { aba: "RESERVA_CONFIRMADA", hist: "HIST_RESERVA_CONFIRMADA" },
    "PROCEDIMENTO_VASCULAR": { aba: "PROCEDIMENTO_VASCULAR", hist: "HIST_PROCEDIMENTO_VASCULAR" },
    "RESERVA_NEGADA": { aba: "RESERVA_NEGADA", hist: "HIST_RESERVA_NEGADA" },
    "PLANTAO_ANTERIOR": { aba: "PLANTAO_ANTERIOR", hist: null },
    "BLOQUEADOS_MANUTENCAO": { aba: "BLOQUEADOS_MANUTENCAO", hist: "HIST_MANUTENCAO" },
    "BLOQUEADOS_ISOLAMENTO": { aba: "BLOQUEADOS_ISOLAMENTO", hist: "HIST_ISOLAMENTO" }
  };

  const conf = mapa[modulo];
  if (!conf) return;

  const ss = getSpreadsheet_();
  const sheet = ss.getSheetByName(conf.aba);
  if (!sheet) return;

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  const values = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();
  const plantao = getPlantaoConfig_();
  const plantaoId = plantao && plantao.id ? plantao.id : "";

  for (let i = 0; i < values.length; i++) {
    if (String(values[i][0]) === String(id)) {
      const rowIndex = i + 2;
      const rowData = values[i];

      if (conf.hist && plantaoId) {
        const histSheet = ss.getSheetByName(conf.hist);
        if (histSheet) {
          const registro = [plantaoId, id].concat(rowData.slice(1));
          if (conf.hist === "HIST_MANUTENCAO" || conf.hist === "HIST_ISOLAMENTO") {
            registro.push(new Date());
          }
          histSheet.appendRow(registro);
        }
      }

      sheet.deleteRow(rowIndex);
      registrarEvento_(modulo, id, "EXCLUSAO_MANUAL", "Registro excluído via WebApp");
      break;
    }
  }
}

// ========================
// 6) AUXILIARES
// ========================
function gerarID_() {
  return Utilities.getUuid().split("-")[0].toUpperCase();
}

function registrarEvento_(modulo, idRegistro, tipo, obs) {
  const ss = getSpreadsheet_();
  const log = ss.getSheetByName("LOG_NIR");
  if (!log) return;

  log.appendRow([
    gerarID_(),
    idRegistro,
    modulo,
    tipo,
    Session.getActiveUser().getEmail(),
    new Date(),
    obs || ""
  ]);
}

function normalizarDataHora_(valor, tz) {
  if (!valor) return null;
  const timezone = tz || Session.getScriptTimeZone() || "America/Sao_Paulo";

  if (Object.prototype.toString.call(valor) === "[object Date]") {
    if (isNaN(valor)) return null;
    return ajustarParaTimezone_(valor, timezone);
  }

  const str = String(valor).trim();
  if (!str) return null;

  const m = str.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})(?:[ T](\d{1,2}):(\d{1,2}))?/);
  if (m) {
    const dia = Number(m[1]);
    const mes = Number(m[2]) - 1;
    const ano = Number(m[3]);
    const hora = Number(m[4] || 0);
    const minuto = Number(m[5] || 0);
    return ajustarParaTimezone_(new Date(ano, mes, dia, hora, minuto), timezone);
  }

  const dt = new Date(str);
  if (isNaN(dt)) return null;
  return ajustarParaTimezone_(dt, timezone);
}

function normalizarDataSimples_(valor, tz) {
  const dt = normalizarDataHora_(valor, tz);
  if (!dt) return "";
  const timezone = tz || Session.getScriptTimeZone() || "America/Sao_Paulo";
  return Utilities.formatDate(dt, timezone, "dd/MM/yyyy");
}

function ajustarParaTimezone_(dateObj, tz) {
  if (!dateObj) return null;
  const timezone = tz || Session.getScriptTimeZone() || "America/Sao_Paulo";
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
