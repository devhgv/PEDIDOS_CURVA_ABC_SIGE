function onOpen() {
  const ui = SpreadsheetApp.getUi();

  ui.createMenu('Consumo de API')
  .addItem ('Funcionamento', 'showAboutAPI')
  .addSeparator()
  .addItem ('Iniciar ou Continuar Consumo API', 'runAPIProcess')
  .addToUi();

  const allCurveABC = ui.createMenu('Curvas ABC')
  .addItem('Funcionamento', 'showAboutCurve');
  allCurveABC.addToUi();
}

function showAboutAPI() {
  const ui = SpreadsheetApp.getUi();
  const message = `
  Este módulo consome dados da API de pedidos da Portoreal.

  Para utilizar corretamente:
  1. Clique em "Iniciar" o Consumo da API;
  2. Aguarde até que o tempo limite de execução se encerre;
  3. Clique em "Continuar Consumo API", para os dados serem atualizados de onde parou.
  
  Um pouco mais sobre:
  - Essa aplicação possui scripts que lidam com consumo de API do ERP SIGE, com paginação;
  - Ela inicia criando e formatando cabeçalhos na planilha ativa, após, começa consumir a API do SIGE com base os cabeçalhos definidos na Array em JSON;
  - Na última coluna, adiciona data/hora da atualização de consumo dos dados inseridos na planilha;
  - Os scripts foram criados para salvar o progresso e retomar de onde parou, para evitar duplicidade no consumo das páginas da API;
  - Para lidar com o tempo limite de execução do Google Apps Script (6 minutos), foi adicionado uma função para encerrar a execução automaticamente depois de 5 minutos, para maior controle da paginação;
  - Essa aplicação roda automaticamente através de um acionador (GAS) todos os sábados das 16h às 17h. 

  Obs: Na coluna ItensJSON, consta todos os Itens de um único pedido por fornecedor, isso alimentará as Curvas ABC/ ABCD deste módulo.  
  `;

ui.alert ('Sobre o Sistema:', message, ui.ButtonSet.OK);
}

function showAboutCurve() {
  const ui = SpreadsheetApp.getUi();
  const message = `
  Este módulo possui as Curvas ABC de diferentes períodos de meses da Portoreal.

  As Curvas são atualizadas todos os sábados após a API-SIGE ser atualizada, para que assim, também atualize os dados das Curvas ABC.

  Um pouco mais sobre:
  - Essa aplicação funciona em conjunto à aplicação de paginação da API;
  - Após atualização de dados de pedidos da API-SIGE, ela atualiza as Curvas ABC com novos dados;
  - Essa aplicação foi criada para apoio na tomada de decisão;
  - Essa aplicação roda automaticamente através de um acionador (GAS) todos os sábados das 18h às 19h. 
  `;
ui.alert ('Sobre o Módulo:', message, ui.ButtonSet.OK);
}



