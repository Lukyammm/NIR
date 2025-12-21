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
  return Utilities.formatDate(dt, timezone, "dd/MM/yyyy");
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
