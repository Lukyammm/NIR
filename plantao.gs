function criarPlantaoComEquipe(dados) {
  assertAuthorized_();
  const ss = getSpreadsheet_();
  const aba = ss.getSheetByName("CONFIG_PLANTAO");
  if (!aba) return;

  const idAtual = aba.getRange("A2").getValue();
  if (idAtual) {
    throw new Error("Já existe um plantão ativo. Encerre antes de abrir outro.");
  }

  const id = gerarID_();
  const agora = new Date();
  const timezone = Session.getScriptTimeZone() || DEFAULT_TIMEZONE;
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

function salvarPlantaoConfig(dados) {
  assertAuthorized_();
  const ss = getSpreadsheet_();
  const aba = ss.getSheetByName("CONFIG_PLANTAO");
  if (!aba) return;

  const idAtual = aba.getRange("A2").getValue();
  if (!idAtual) {
    throw new Error("Nenhum plantão ativo para salvar.");
  }

  const timezone = Session.getScriptTimeZone() || DEFAULT_TIMEZONE;
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
  assertAuthorized_();
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
  dados[10] = encerramento;
  dados.push(Session.getActiveUser().getEmail());

  hist.appendRow(dados);
  registrarEvento_("PLANTAO", plantaoId, "ENCERRAMENTO_PLANTAO", "Plantão encerrado");

  conf.getRange("A2:K2").clearContent();
}

function getPlantaoConfig_() {
  const ss = getSpreadsheet_();
  const aba = ss.getSheetByName("CONFIG_PLANTAO");
  if (!aba) return null;

  const valores = aba.getRange("A2:K2").getValues()[0];
  const possuiConteudo = valores.some((v) => v !== "" && v !== null);
  const tz = Session.getScriptTimeZone() || DEFAULT_TIMEZONE;

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
  const possuiEncerramentoValido = Boolean(encerramentoDt && !isNaN(encerramentoDt));
  const plantaoAtivo = Boolean(
    valores[0] ||
    (possuiConteudo && (!possuiEncerramentoValido || agora < encerramentoDt))
  );

  return {
    id: plantaoAtivo ? String(valores[0] || "") : "",
    data: normalizarDataSimples_(valores[1], tz),
    diaSemana: valores[2] || "",
    turno: valores[3] || "",
    medico1: valores[4] || "",
    medico2: valores[5] || "",
    enf1: valores[6] || "",
    enf2: valores[7] || "",
    aux: valores[8] || "",
    abertura: aberturaDt || "",
    encerramento: encerramentoDt || "",
    ativo: plantaoAtivo,
    possuiConteudo
  };
}
