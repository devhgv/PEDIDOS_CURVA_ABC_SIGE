function showPedidosSidebar() {
  const htmlOutput = HtmlService.createHtmlOutputFromFile('AppScript/index') 
    .setWidth(400)  
    .setHeight(300); 
  SpreadsheetApp.getUi().showSidebar(htmlOutput);
}
