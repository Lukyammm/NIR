function onOpen() {
  const ui = getSpreadsheet_().getUi();
  ui.createMenu("Ocorrências NIR")
    .addItem("Criar Estrutura Ocorrências", "criarEstruturaNIR")
    .addItem("Abrir Ocorrências (Sidebar)", "abrirWebAppSidebar")
    .addItem("Encerrar Plantão Atual", "encerrarPlantao")
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
