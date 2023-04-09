function doGet(e) {
  const response = [{ status: e.parameter.wsName }]
  return ContentService.createTextOutput(JSON.stringify(response)).setMimeType(
    ContentService.MimeType.JSON
  )
}

function doPost(e) {}

function loadOperationsForm() {
  const htmlForSidebar = HtmlService.createTemplateFromFile('index')
  const htmlOutput = htmlForSidebar.evaluate()
  htmlOutput.setTitle('Operações na Carteira')

  const ui = SpreadsheetApp.getUi()
  ui.showSidebar(htmlOutput)
}

function loadSellForm() {
  const htmlForSidebar = HtmlService.createTemplateFromFile('sellForm')
  const htmlOutput = htmlForSidebar.evaluate()
  htmlOutput.setTitle('Entrada de Dados')

  const ui = SpreadsheetApp.getUi()
  ui.showSidebar(htmlOutput)
}

function createMenu() {
  const ui = SpreadsheetApp.getUi()
  const menu = ui.createMenu('Operações na Carteira')
  menu.addItem('Registrar operações', 'loadOperationsForm')
  // menu.addItem('Agrupar/Desdobrar', 'loadSellForm')
  menu.addToUi()
}

function onOpen() {
  createMenu()
}
