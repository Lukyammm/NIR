function adicionarRegistro(modulo, registro) {
  assertAuthorized_();
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

function excluirRegistro(modulo, id) {
  assertAuthorized_();
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
