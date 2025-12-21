const DEFAULT_TIMEZONE = "America/Sao_Paulo";
const SPREADSHEET_ID_PROPERTY = "PLANILHA_ID";
const PLANILHA_ID =
  PropertiesService.getScriptProperties().getProperty(SPREADSHEET_ID_PROPERTY) ||
  "1vnnGQEkAjP9eTRLWWSb2lSngwGTRYa_rtk6DsG8HGqc";

const AUTHORIZED_USERS = [];
