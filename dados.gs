function getAppData(viewType) {
  assertAuthorized_();
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
