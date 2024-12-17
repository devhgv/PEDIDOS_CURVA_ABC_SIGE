function generateABCItems() {
  const ui = SpreadsheetApp.getUi();

  const sheetPedidos = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Pedidos");
  const sheetABC = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Itens x Curva ABC/ ABCD (qtd)");

  // Limpa a aba "Itens x Curva ABC/ ABCD (qtd)"
  sheetABC.clearContents();

  // Configura cabeçalhos
  sheetABC.appendRow(["Descrição", "Quantidade Total"]);

  // Filtro de tempo: meses (valor fixo ou parametrizado)
  const monthsBack = 6; // Altere este valor conforme necessário
  const dateLimit = new Date();
  dateLimit.setMonth(dateLimit.getMonth() - monthsBack);

  // Mapeamento para somar Quantidades por Descrição
  const quantityMap = {};

  // Lê todos os dados da aba Pedidos
  const pedidosData = sheetPedidos.getDataRange().getValues();

  // Índices das colunas
  const idxDataFaturada = 2; // Coluna C
  const idxItensJson = 5; // Coluna F
  const idxStatus = 9; // Coluna J

  for (let i = 1; i < pedidosData.length; i++) {
    const status = pedidosData[i][idxStatus];
    const dataFaturada = pedidosData[i][idxDataFaturada];
    const itensJson = pedidosData[i][idxItensJson];

    // Filtra pedidos "Faturados" dentro do período
    if (status === "Pedido Faturado" && isDateWithinLimit(dataFaturada, dateLimit)) {
      try {
        const itens = JSON.parse(itensJson); // Parseia os itens JSON

        itens.forEach(item => {
          const descricao = item.Descricao || "Sem Descrição";
          const quantidade = parseFloat(item.Quantidade) || 0;

          // Soma a quantidade por Descrição
          if (!quantityMap[descricao]) {
            quantityMap[descricao] = 0;
          }
          quantityMap[descricao] += quantidade;
        });
      } catch (e) {
        Logger.log("Erro ao processar JSON: " + e.message);
      }
    }
  }

  // Ordena os itens pelo total de quantidade (decrescente)
  const sortedItems = Object.entries(quantityMap).sort((a, b) => b[1] - a[1]);

  // Insere os dados na aba ABC
  sortedItems.forEach(([descricao, quantidade]) => {
    sheetABC.appendRow([descricao, quantidade]);
  });

  // Alerta final (somente em contexto interativo)
  try {
    ui.alert(`Curva ABC gerada com sucesso para os últimos ${monthsBack} meses!`);
  } catch (e) {
    Logger.log(`Curva ABC gerada com sucesso para os últimos ${monthsBack} meses!`);
  }
}

// auxilia para validar datas dentro do período.

function isDateWithinLimit(dataFaturada, dateLimit) {
  const dateParts = dataFaturada.split("/");
  if (dateParts.length !== 3) return false;

  const date = new Date(`${dateParts[2]}-${dateParts[1]}-${dateParts[0]}`); // Formato "yyyy-MM-dd"
  return date >= dateLimit;
}
