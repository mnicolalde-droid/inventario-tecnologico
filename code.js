function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("📄 Actas")
    .addItem("Generar Acta", "abrirFormulario")
    .addToUi();
}

function abrirFormulario() {
  const html = HtmlService.createHtmlOutputFromFile("index")
    .setWidth(700)
    .setHeight(800);
  SpreadsheetApp.getUi().showModalDialog(html, "Generador de Actas");
}
