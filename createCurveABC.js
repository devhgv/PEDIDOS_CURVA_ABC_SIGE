// FUNÇÃO PARA GERAR A CURVA ABC NO PERÍODO DE 3 MESES.
function createCurve3Month() {
  const tabOrigin = "consAPI";
  const tabDestination = "3 Meses";

  const tab = SpreadsheetApp.getActiveSpreadsheet();
  const sheetOrigem = tab.getSheetByName(tabOrigin);
  let sheetDestino = tab.getSheetByName(tabDestination);

  if (!sheetDestino) {
    sheetDestino = tab.insertSheet(tabDestination);
  }

  const dataOrigin = sheetOrigem.getDataRange().getValues();
  const headers = dataOrigin[0];

  const idxDataFaturada = headers.indexOf("DataFaturada");
  const idxStatusSistema = headers.indexOf("StatusSistema");
  const idxItensJson = headers.indexOf("ItensJSON");

  if (idxDataFaturada === -1 || idxStatusSistema === -1 || idxItensJson === -1) {
    throw new Error("Colunas necessárias não encontradas na aba consAPI.");
  }

  const dateCurrent = new Date();
  const dateLimit = new Date();
  dateLimit.setMonth(dateCurrent.getMonth() - 3);

  const dataFilter = dataOrigin.slice(1).filter(linha => {
    const dataFaturada = new Date(linha[idxDataFaturada]);
    return linha[idxStatusSistema] === "Pedido Faturado" && dataFaturada >= dateLimit;
  });

  const mapDescription = new Map();

  dataFilter.forEach(linha => {
    const dataFilter = linha[idxDataFaturada];
    const itensJSON = JSON.parse(linha[idxItensJson]);

    itensJSON.forEach(item => {
      const description = item.Descricao;
      const quantidade = item.Quantidade;

      if (!mapDescription.has(description)) {
        mapDescription.set(description, { description, dataFilter, quantidade: 0 });
      }
      mapDescription.get(description).quantidade += quantidade;
    });
  });

  const listItens = Array.from(mapDescription.values());
  listItens.sort((a, b) => b.quantidade - a.quantidade);

  const totalAmount = listItens.reduce((sum, item) => sum + item.quantidade, 0);
  let acumulado = 0;

  listItens.forEach(item => {
    const participacao = (item.quantidade / totalAmount) * 100;
    acumulado += participacao;
    item.participacao = participacao;
    item.acumulado = acumulado;
    item.classificacao =
      acumulado <= 80 ? "A" :
      acumulado <= 90 ? "B" :
      acumulado <= 95 ? "C" : "D";
  });

  sheetDestino.clearContents();
  sheetDestino.appendRow(["Data Faturada", "Descrição", "Quantidade Vendida", "Participação (%)", "Acumulado (%)", "Classificação Curva"]);

  listItens.forEach(item => {
    sheetDestino.appendRow([
      item.dataFilter,
      item.description,
      item.quantidade,
      item.participacao / 100, 
      item.acumulado / 100,   
      item.classificacao
    ]);
  });

  const ultimaRow = sheetDestino.getLastRow();
  sheetDestino.appendRow(["", "TOTAL", totalAmount, "", "", ""]);

  const range = sheetDestino.getRange(2, 4, ultimaRow, 2); 
  range.setNumberFormat("0.00%");

  Logger.log("Curva ABC criada com sucesso na aba '3 Meses'.");
}


// FUNÇÃO PARA GERAR A CURVA ABC NO PERÍODO DE 6 MESES.
function createCurve6Month() {
  const tabOrigin = "consAPI";
  const tabDestination = "6 Meses";

  const tab = SpreadsheetApp.getActiveSpreadsheet();
  const sheetOrigem = tab.getSheetByName(tabOrigin);
  let sheetDestino = tab.getSheetByName(tabDestination);

  if (!sheetDestino) {
    sheetDestino = tab.insertSheet(tabDestination);
  }

  const dataOrigin = sheetOrigem.getDataRange().getValues();
  const headers = dataOrigin[0];

  const idxDataFaturada = headers.indexOf("DataFaturada");
  const idxStatusSistema = headers.indexOf("StatusSistema");
  const idxItensJson = headers.indexOf("ItensJSON");

  if (idxDataFaturada === -1 || idxStatusSistema === -1 || idxItensJson === -1) {
    throw new Error("Colunas necessárias não encontradas na aba consAPI.");
  }

  const dateCurrent = new Date();
  const dateLimit = new Date();
  dateLimit.setMonth(dateCurrent.getMonth() - 6);

  const dataFilter = dataOrigin.slice(1).filter(linha => {
    const dataFaturada = new Date(linha[idxDataFaturada]);
    return linha[idxStatusSistema] === "Pedido Faturado" && dataFaturada >= dateLimit;
  });

  const mapDescription = new Map();

  dataFilter.forEach(linha => {
    const dataFilter = linha[idxDataFaturada];
    const itensJSON = JSON.parse(linha[idxItensJson]);

    itensJSON.forEach(item => {
      const description = item.Descricao;
      const quantidade = item.Quantidade;

      if (!mapDescription.has(description)) {
        mapDescription.set(description, { description, dataFilter, quantidade: 0 });
      }
      mapDescription.get(description).quantidade += quantidade;
    });
  });

  const listItens = Array.from(mapDescription.values());
  listItens.sort((a, b) => b.quantidade - a.quantidade);

  const totalAmount = listItens.reduce((sum, item) => sum + item.quantidade, 0);
  let acumulado = 0;

  listItens.forEach(item => {
    const participacao = (item.quantidade / totalAmount) * 100;
    acumulado += participacao;
    item.participacao = participacao;
    item.acumulado = acumulado;
    item.classificacao =
      acumulado <= 80 ? "A" :
      acumulado <= 90 ? "B" :
      acumulado <= 95 ? "C" : "D";
  });

  sheetDestino.clearContents();
  sheetDestino.appendRow(["Data Faturada", "Descrição", "Quantidade Vendida", "Participação (%)", "Acumulado (%)", "Classificação Curva"]);

  listItens.forEach(item => {
    sheetDestino.appendRow([
      item.dataFilter,
      item.description,
      item.quantidade,
      item.participacao / 100, 
      item.acumulado / 100,   
      item.classificacao
    ]);
  });

  const ultimaRow = sheetDestino.getLastRow();
  sheetDestino.appendRow(["", "TOTAL", totalAmount, "", "", ""]);

  const range = sheetDestino.getRange(2, 4, ultimaRow, 2); 
  range.setNumberFormat("0.00%");

  Logger.log("Curva ABC criada com sucesso na aba '6 Meses'.");
}


// FUNÇÃO PARA GERAR A CURVA ABC NO PERÍODO DE 12 MESES.
function createCurve12Month() {
  const tabOrigin = "consAPI";
  const tabDestination = "12 Meses";

  const tab = SpreadsheetApp.getActiveSpreadsheet();
  const sheetOrigem = tab.getSheetByName(tabOrigin);
  let sheetDestino = tab.getSheetByName(tabDestination);

  if (!sheetDestino) {
    sheetDestino = tab.insertSheet(tabDestination);
  }

  const dataOrigin = sheetOrigem.getDataRange().getValues();
  const headers = dataOrigin[0];

  const idxDataFaturada = headers.indexOf("DataFaturada");
  const idxStatusSistema = headers.indexOf("StatusSistema");
  const idxItensJson = headers.indexOf("ItensJSON");

  if (idxDataFaturada === -1 || idxStatusSistema === -1 || idxItensJson === -1) {
    throw new Error("Colunas necessárias não encontradas na aba consAPI.");
  }

  const dateCurrent = new Date();
  const dateLimit = new Date();
  dateLimit.setMonth(dateCurrent.getMonth() - 12);

  const dataFilter = dataOrigin.slice(1).filter(linha => {
    const dataFaturada = new Date(linha[idxDataFaturada]);
    return linha[idxStatusSistema] === "Pedido Faturado" && dataFaturada >= dateLimit;
  });

  const mapDescription = new Map();

  dataFilter.forEach(linha => {
    const dataFilter = linha[idxDataFaturada];
    const itensJSON = JSON.parse(linha[idxItensJson]);

    itensJSON.forEach(item => {
      const description = item.Descricao;
      const quantidade = item.Quantidade;

      if (!mapDescription.has(description)) {
        mapDescription.set(description, { description, dataFilter, quantidade: 0 });
      }
      mapDescription.get(description).quantidade += quantidade;
    });
  });

  const listItens = Array.from(mapDescription.values());
  listItens.sort((a, b) => b.quantidade - a.quantidade);

  const totalAmount = listItens.reduce((sum, item) => sum + item.quantidade, 0);
  let acumulado = 0;

  listItens.forEach(item => {
    const participacao = (item.quantidade / totalAmount) * 100;
    acumulado += participacao;
    item.participacao = participacao;
    item.acumulado = acumulado;
    item.classificacao =
      acumulado <= 80 ? "A" :
      acumulado <= 90 ? "B" :
      acumulado <= 95 ? "C" : "D";
  });

  sheetDestino.clearContents();
  sheetDestino.appendRow(["Data Faturada", "Descrição", "Quantidade Vendida", "Participação (%)", "Acumulado (%)", "Classificação Curva"]);

  listItens.forEach(item => {
    sheetDestino.appendRow([
      item.dataFilter,
      item.description,
      item.quantidade,
      item.participacao / 100, 
      item.acumulado / 100,   
      item.classificacao
    ]);
  });

  const ultimaRow = sheetDestino.getLastRow();
  sheetDestino.appendRow(["", "TOTAL", totalAmount, "", "", ""]);

  const range = sheetDestino.getRange(2, 4, ultimaRow, 2); 
  range.setNumberFormat("0.00%");

  Logger.log("Curva ABC criada com sucesso na aba '12 Meses'.");
}


// FUNÇÃO PARA GERAR A CURVA ABC NO PERÍODO DE 18 MESES.
function createCurve18Month() {
  const tabOrigin = "consAPI";
  const tabDestination = "18 Meses";

  const tab = SpreadsheetApp.getActiveSpreadsheet();
  const sheetOrigem = tab.getSheetByName(tabOrigin);
  let sheetDestino = tab.getSheetByName(tabDestination);

  if (!sheetDestino) {
    sheetDestino = tab.insertSheet(tabDestination);
  }

  const dataOrigin = sheetOrigem.getDataRange().getValues();
  const headers = dataOrigin[0];

  const idxDataFaturada = headers.indexOf("DataFaturada");
  const idxStatusSistema = headers.indexOf("StatusSistema");
  const idxItensJson = headers.indexOf("ItensJSON");

  if (idxDataFaturada === -1 || idxStatusSistema === -1 || idxItensJson === -1) {
    throw new Error("Colunas necessárias não encontradas na aba consAPI.");
  }

  const dateCurrent = new Date();
  const dateLimit = new Date();
  dateLimit.setMonth(dateCurrent.getMonth() - 18);

  const dataFilter = dataOrigin.slice(1).filter(linha => {
    const dataFaturada = new Date(linha[idxDataFaturada]);
    return linha[idxStatusSistema] === "Pedido Faturado" && dataFaturada >= dateLimit;
  });

  const mapDescription = new Map();

  dataFilter.forEach(linha => {
    const dataFilter = linha[idxDataFaturada];
    const itensJSON = JSON.parse(linha[idxItensJson]);

    itensJSON.forEach(item => {
      const description = item.Descricao;
      const quantidade = item.Quantidade;

      if (!mapDescription.has(description)) {
        mapDescription.set(description, { description, dataFilter, quantidade: 0 });
      }
      mapDescription.get(description).quantidade += quantidade;
    });
  });

  const listItens = Array.from(mapDescription.values());
  listItens.sort((a, b) => b.quantidade - a.quantidade);

  const totalAmount = listItens.reduce((sum, item) => sum + item.quantidade, 0);
  let acumulado = 0;

  listItens.forEach(item => {
    const participacao = (item.quantidade / totalAmount) * 100;
    acumulado += participacao;
    item.participacao = participacao;
    item.acumulado = acumulado;
    item.classificacao =
      acumulado <= 80 ? "A" :
      acumulado <= 90 ? "B" :
      acumulado <= 95 ? "C" : "D";
  });

  sheetDestino.clearContents();
  sheetDestino.appendRow(["Data Faturada", "Descrição", "Quantidade Vendida", "Participação (%)", "Acumulado (%)", "Classificação Curva"]);

  listItens.forEach(item => {
    sheetDestino.appendRow([
      item.dataFilter,
      item.description,
      item.quantidade,
      item.participacao / 100, 
      item.acumulado / 100,   
      item.classificacao
    ]);
  });

  const ultimaRow = sheetDestino.getLastRow();
  sheetDestino.appendRow(["", "TOTAL", totalAmount, "", "", ""]);

  const range = sheetDestino.getRange(2, 4, ultimaRow, 2); 
  range.setNumberFormat("0.00%");

  Logger.log("Curva ABC criada com sucesso na aba '18 Meses'.");
}


// FUNÇÃO PARA GERAR A CURVA ABC NO PERÍODO DE 24 MESES.
function createCurve24Month() {
  const tabOrigin = "consAPI";
  const tabDestination = "24 Meses";

  const tab = SpreadsheetApp.getActiveSpreadsheet();
  const sheetOrigem = tab.getSheetByName(tabOrigin);
  let sheetDestino = tab.getSheetByName(tabDestination);

  if (!sheetDestino) {
    sheetDestino = tab.insertSheet(tabDestination);
  }

  const dataOrigin = sheetOrigem.getDataRange().getValues();
  const headers = dataOrigin[0];

  const idxDataFaturada = headers.indexOf("DataFaturada");
  const idxStatusSistema = headers.indexOf("StatusSistema");
  const idxItensJson = headers.indexOf("ItensJSON");

  if (idxDataFaturada === -1 || idxStatusSistema === -1 || idxItensJson === -1) {
    throw new Error("Colunas necessárias não encontradas na aba consAPI.");
  }

  const dateCurrent = new Date();
  const dateLimit = new Date();
  dateLimit.setMonth(dateCurrent.getMonth() - 24);

  const dataFilter = dataOrigin.slice(1).filter(linha => {
    const dataFaturada = new Date(linha[idxDataFaturada]);
    return linha[idxStatusSistema] === "Pedido Faturado" && dataFaturada >= dateLimit;
  });

  const mapDescription = new Map();

  dataFilter.forEach(linha => {
    const dataFilter = linha[idxDataFaturada];
    const itensJSON = JSON.parse(linha[idxItensJson]);

    itensJSON.forEach(item => {
      const description = item.Descricao;
      const quantidade = item.Quantidade;

      if (!mapDescription.has(description)) {
        mapDescription.set(description, { description, dataFilter, quantidade: 0 });
      }
      mapDescription.get(description).quantidade += quantidade;
    });
  });

  const listItens = Array.from(mapDescription.values());
  listItens.sort((a, b) => b.quantidade - a.quantidade);

  const totalAmount = listItens.reduce((sum, item) => sum + item.quantidade, 0);
  let acumulado = 0;

  listItens.forEach(item => {
    const participacao = (item.quantidade / totalAmount) * 100;
    acumulado += participacao;
    item.participacao = participacao;
    item.acumulado = acumulado;
    item.classificacao =
      acumulado <= 80 ? "A" :
      acumulado <= 90 ? "B" :
      acumulado <= 95 ? "C" : "D";
  });

  sheetDestino.clearContents();
  sheetDestino.appendRow(["Data Faturada", "Descrição", "Quantidade Vendida", "Participação (%)", "Acumulado (%)", "Classificação Curva"]);

  listItens.forEach(item => {
    sheetDestino.appendRow([
      item.dataFilter,
      item.description,
      item.quantidade,
      item.participacao / 100, 
      item.acumulado / 100,   
      item.classificacao
    ]);
  });

  const ultimaRow = sheetDestino.getLastRow();
  sheetDestino.appendRow(["", "TOTAL", totalAmount, "", "", ""]);

  const range = sheetDestino.getRange(2, 4, ultimaRow, 2); 
  range.setNumberFormat("0.00%");

  Logger.log("Curva ABC criada com sucesso na aba '24 Meses'.");
}


