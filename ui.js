function onOpen(_e) {
  SpreadsheetApp.getUi()
    .createMenu('Gerar')
    .addItem('Certificados', 'generateAndSendCertificatesFromSpreadsheet')
    .addToUi();
}
