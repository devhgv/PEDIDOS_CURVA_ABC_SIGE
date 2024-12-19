function curvaABC3Meses() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const pedidosSheet = ss.getSheetByName("Pedidos");
  const itensSheet = ss.getSheetByName("3 Meses");

  const colDataFaturada = 2; // Coluna C
  const colItensJson = 5; // Coluna F
  const colStatusSistema = 9; // Coluna J

  const data = pedidosSheet.getDataRange().getValues();

  const mesesFiltro = "3"; // Adicione o número de meses.
  const hoje = new Date();
  const dataLimite = new Date(hoje.setMonth(hoje.getMonth() - mesesFiltro));

  const consolidado = {}; 

  // Itera sobre as linhas, ignorando o cabeçalho
  for (let i = 1; i < data.length; i++) {
    const statusSistema = data[i][colStatusSistema];
    const dataFaturadaStr = data[i][colDataFaturada];
    const itensJson = data[i][colItensJson];

    // Converte a data de string "DD/MM/AAAA" para objeto Date
    const partesData = dataFaturadaStr.split("/");
    const dataFaturada = new Date(partesData[2], partesData[1] - 1, partesData[0]);

    // Aplica filtro de data e status "Pedido Faturado"
    if (dataFaturada >= dataLimite && statusSistema === "Pedido Faturado" && isValidJson(itensJson)) {
      const itens = JSON.parse(itensJson);

      itens.forEach(item => {
        const descricao = item.Descricao; // Ajuste para o campo correto "Descricao"
        const quantidade = parseFloat(item.Quantidade) || 0;

        if (descricao) {
          // Agrupa as quantidades por descrição
          if (!consolidado[descricao]) {
            consolidado[descricao] = { quantidade: 0, dataFaturada: dataFaturadaStr };
          }
          consolidado[descricao].quantidade += quantidade;
        }
      });
    }
  }

  // Constrói os dados para inserção na planilha
  const output = Object.entries(consolidado).map(([descricao, dados]) => [
    dados.dataFaturada, descricao, dados.quantidade
  ]);

  output.sort((a, b) => b[2] - a[2]); //Ordena a quantidade de forma decrescente.

  if (output.length > 0) {
    itensSheet.getRange(2, 1, output.length, 3).setValues(output);
  } else {
    itensSheet.appendRow(["Não há dados para o período selecionado."]);
  }

  Logger.log("Consolidação concluída com sucesso!");
}


function curvaABC6Meses() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const pedidosSheet = ss.getSheetByName("Pedidos");
  const itensSheet = ss.getSheetByName("6 Meses");

  const colDataFaturada = 2; 
  const colItensJson = 5; 
  const colStatusSistema = 9;

  const data = pedidosSheet.getDataRange().getValues();

  const mesesFiltro = "6"; 
  const hoje = new Date();
  const dataLimite = new Date(hoje.setMonth(hoje.getMonth() - mesesFiltro));

  const consolidado = {}; 

  
  for (let i = 1; i < data.length; i++) {
    const statusSistema = data[i][colStatusSistema];
    const dataFaturadaStr = data[i][colDataFaturada];
    const itensJson = data[i][colItensJson];

   
    const partesData = dataFaturadaStr.split("/");
    const dataFaturada = new Date(partesData[2], partesData[1] - 1, partesData[0]);

    if (dataFaturada >= dataLimite && statusSistema === "Pedido Faturado" && isValidJson(itensJson)) {
      const itens = JSON.parse(itensJson);

      itens.forEach(item => {
        const descricao = item.Descricao; 
        const quantidade = parseFloat(item.Quantidade) || 0;

        if (descricao) {
          
          if (!consolidado[descricao]) {
            consolidado[descricao] = { quantidade: 0, dataFaturada: dataFaturadaStr };
          }
          consolidado[descricao].quantidade += quantidade;
        }
      });
    }
  }

  
  const output = Object.entries(consolidado).map(([descricao, dados]) => [
    dados.dataFaturada, descricao, dados.quantidade
  ]);

  output.sort((a, b) => b[2] - a[2]); 
  if (output.length > 0) {
    itensSheet.getRange(2, 1, output.length, 3).setValues(output);
  } else {
    itensSheet.appendRow(["Não há dados para o período selecionado."]);
  }

  Logger.log("Consolidação concluída com sucesso!");
}


function curvaABC12Meses() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const pedidosSheet = ss.getSheetByName("Pedidos");
  const itensSheet = ss.getSheetByName("12 Meses");

  const colDataFaturada = 2; 
  const colItensJson = 5; 
  const colStatusSistema = 9;

  const data = pedidosSheet.getDataRange().getValues();

  const mesesFiltro = "12"; 
  const hoje = new Date();
  const dataLimite = new Date(hoje.setMonth(hoje.getMonth() - mesesFiltro));

  const consolidado = {}; 

  
  for (let i = 1; i < data.length; i++) {
    const statusSistema = data[i][colStatusSistema];
    const dataFaturadaStr = data[i][colDataFaturada];
    const itensJson = data[i][colItensJson];

   
    const partesData = dataFaturadaStr.split("/");
    const dataFaturada = new Date(partesData[2], partesData[1] - 1, partesData[0]);

    if (dataFaturada >= dataLimite && statusSistema === "Pedido Faturado" && isValidJson(itensJson)) {
      const itens = JSON.parse(itensJson);

      itens.forEach(item => {
        const descricao = item.Descricao; 
        const quantidade = parseFloat(item.Quantidade) || 0;

        if (descricao) {
          
          if (!consolidado[descricao]) {
            consolidado[descricao] = { quantidade: 0, dataFaturada: dataFaturadaStr };
          }
          consolidado[descricao].quantidade += quantidade;
        }
      });
    }
  }

  
  const output = Object.entries(consolidado).map(([descricao, dados]) => [
    dados.dataFaturada, descricao, dados.quantidade
  ]);

  output.sort((a, b) => b[2] - a[2]); 
  if (output.length > 0) {
    itensSheet.getRange(2, 1, output.length, 3).setValues(output);
  } else {
    itensSheet.appendRow(["Não há dados para o período selecionado."]);
  }

  Logger.log("Consolidação concluída com sucesso!");
}


function curvaABC18Meses() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const pedidosSheet = ss.getSheetByName("Pedidos");
  const itensSheet = ss.getSheetByName("18 Meses");

  const colDataFaturada = 2; 
  const colItensJson = 5; 
  const colStatusSistema = 9;

  const data = pedidosSheet.getDataRange().getValues();

  const mesesFiltro = "18"; 
  const hoje = new Date();
  const dataLimite = new Date(hoje.setMonth(hoje.getMonth() - mesesFiltro));

  const consolidado = {}; 

  
  for (let i = 1; i < data.length; i++) {
    const statusSistema = data[i][colStatusSistema];
    const dataFaturadaStr = data[i][colDataFaturada];
    const itensJson = data[i][colItensJson];

   
    const partesData = dataFaturadaStr.split("/");
    const dataFaturada = new Date(partesData[2], partesData[1] - 1, partesData[0]);

    if (dataFaturada >= dataLimite && statusSistema === "Pedido Faturado" && isValidJson(itensJson)) {
      const itens = JSON.parse(itensJson);

      itens.forEach(item => {
        const descricao = item.Descricao; 
        const quantidade = parseFloat(item.Quantidade) || 0;

        if (descricao) {
          
          if (!consolidado[descricao]) {
            consolidado[descricao] = { quantidade: 0, dataFaturada: dataFaturadaStr };
          }
          consolidado[descricao].quantidade += quantidade;
        }
      });
    }
  }

  
  const output = Object.entries(consolidado).map(([descricao, dados]) => [
    dados.dataFaturada, descricao, dados.quantidade
  ]);

  output.sort((a, b) => b[2] - a[2]); 
  if (output.length > 0) {
    itensSheet.getRange(2, 1, output.length, 3).setValues(output);
  } else {
    itensSheet.appendRow(["Não há dados para o período selecionado."]);
  }

  Logger.log("Consolidação concluída com sucesso!");
}


function curvaABC24Meses() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const pedidosSheet = ss.getSheetByName("Pedidos");
  const itensSheet = ss.getSheetByName("24 Meses");

  const colDataFaturada = 2; 
  const colItensJson = 5; 
  const colStatusSistema = 9;

  const data = pedidosSheet.getDataRange().getValues();

  const mesesFiltro = "24"; 
  const hoje = new Date();
  const dataLimite = new Date(hoje.setMonth(hoje.getMonth() - mesesFiltro));

  const consolidado = {}; 

  
  for (let i = 1; i < data.length; i++) {
    const statusSistema = data[i][colStatusSistema];
    const dataFaturadaStr = data[i][colDataFaturada];
    const itensJson = data[i][colItensJson];

   
    const partesData = dataFaturadaStr.split("/");
    const dataFaturada = new Date(partesData[2], partesData[1] - 1, partesData[0]);

    if (dataFaturada >= dataLimite && statusSistema === "Pedido Faturado" && isValidJson(itensJson)) {
      const itens = JSON.parse(itensJson);

      itens.forEach(item => {
        const descricao = item.Descricao; 
        const quantidade = parseFloat(item.Quantidade) || 0;

        if (descricao) {
          
          if (!consolidado[descricao]) {
            consolidado[descricao] = { quantidade: 0, dataFaturada: dataFaturadaStr };
          }
          consolidado[descricao].quantidade += quantidade;
        }
      });
    }
  }

  
  const output = Object.entries(consolidado).map(([descricao, dados]) => [
    dados.dataFaturada, descricao, dados.quantidade
  ]);

  output.sort((a, b) => b[2] - a[2]); 
  if (output.length > 0) {
    itensSheet.getRange(2, 1, output.length, 3).setValues(output);
  } else {
    itensSheet.appendRow(["Não há dados para o período selecionado."]);
  }

  Logger.log("Consolidação concluída com sucesso!");
}



//Validação do JSON
function isValidJson(str) {
  try {
    JSON.parse(str);
    return true;
  } catch (e) {
    return false;
  }
}
