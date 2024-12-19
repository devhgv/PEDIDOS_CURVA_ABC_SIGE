function onOpen() {
  const ui = SpreadsheetApp.getUi();

  ui.createMenu('Consumo de API')
    .addItem('Iniciar/ Continuar Consumo API', 'updatePedidos')
    .addItem('Sobre', 'showAboutAPI')
    .addToUi();

  const curvaABCMenu = ui.createMenu('Curva ABC')
    .addItem('Funcionamento', 'showAboutCurva')
    .addSeparator() 
    .addItem('3 meses', 'curvaABC3Meses')
    .addItem('6 meses', 'curvaABC6Meses')
    .addItem('12 meses', 'curvaABC12Meses')
    .addItem('18 meses', 'curvaABC18Meses')
    .addItem('24 meses', 'curvaABC24Meses');
  
  curvaABCMenu.addToUi();
}


function updatePedidos() {
  fetchAllPedidosFromMiddleware(); 
}

function showAboutAPI() {
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

function showAboutCurva() {
  const ui = SpreadsheetApp.getUi();
  const message = `
    Nessa aba consta a Curva ABC/ ABCD de itens da Portoreal.

    Funcionamento:

    1. Ler os dados JSON da coluna "ItensJson" da aba Pedidos;
    2. Filtra os pedidos pelo campo "StatusSistema" (coluna J) que estejam como "Pedido Faturado";
    3. Soma a quantidade dos itens por "Descricao" (campo do JSON), agrupando os dados sem duplicidade;
    4. Permite filtrar os dados por um período (3, 6, 12, 18, 24 meses) com base na coluna "DataFaturada" (coluna C) da aba Pedidos;
    5. Insere os resultados na aba "Itens x Curva ABC/ABCD (qtd)", ordenando os itens pela quantidade total (do maior para o menor).
  `;
  ui.alert('Sobre a Curva:', message, ui.ButtonSet.OK);
}
