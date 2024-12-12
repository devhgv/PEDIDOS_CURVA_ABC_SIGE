function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Consumo de API') 
    .addItem('Iniciar/ Continuar Consumo API', 'updatePedidos')  
    .addItem('Sobre', 'showAbout')  
    .addToUi();
}

function updatePedidos() {
  fetchAllPedidosFromMiddleware();
}

function showAbout() {
  const ui = SpreadsheetApp.getUi();
  const message = `
    Este módulo consome dados da API de pedidos da Portoreal.

    Para utilizar corretamente:
    1. Clique em Iniciar o consumo da API;
    2. Aguarde até que o tempo limite de execução se encerre;
    3. Clique em Continuar o consumo da API, para os dados serem atualizados de onde parou.
  `;
  ui.alert('Sobre o Sistema:', message, ui.ButtonSet.OK);
}
