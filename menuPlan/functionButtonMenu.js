function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Menu') 
    .addItem('Atualizar', 'updatePedidos')  
    .addItem('Sobre', 'showAbout')  
    .addToUi();
}

function updatePedidos() {
  fetchAllPedidosFromMiddleware();
}

function showAbout() {
  const ui = SpreadsheetApp.getUi();
  const message = `Este módulo consome da API de pedidos da Portoreal, para criação de gráficos da Curva ABC/ABCD.`;
  ui.alert('Sobre o Sistema', message, ui.ButtonSet.OK);
}
