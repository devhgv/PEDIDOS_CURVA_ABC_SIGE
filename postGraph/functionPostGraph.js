function gerarCurvaABC() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const pedidosSheet = spreadsheet.getSheetByName("Pedidos");
  const curvaABCSheet = spreadsheet.getSheetByName("Curva ABC");

  if (!pedidosSheet || !curvaABCSheet) {
    SpreadsheetApp.getUi().alert("Certifique-se de que as planilhas 'Pedidos' e 'Curva ABC' existam.");
    return;
  }

  // Calcular a Curva ABC
  const resultados = calcularCurvaABC(pedidosSheet);

  // Limpar dados antigos e adicionar cabeçalho
  curvaABCSheet.clear();
  curvaABCSheet.appendRow(["ClienteCNPJ", "Total Vendido", "Percentual Acumulado (%)", "Categoria"]);

  // Atualizar planilha com novos dados
  curvaABCSheet.getRange(2, 1, resultados.length, 4).setValues(resultados);

  // Gerar gráfico da Curva ABC
  gerarGrafico(curvaABCSheet, resultados);

  SpreadsheetApp.getUi().alert("Curva ABC gerada com sucesso!");
}

function gerarCurvaABCD() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const pedidosSheet = spreadsheet.getSheetByName("Pedidos");
  const curvaABCSheet = spreadsheet.getSheetByName("Curva ABC");

  if (!pedidosSheet || !curvaABCSheet) {
    SpreadsheetApp.getUi().alert("Certifique-se de que as planilhas 'Pedidos' e 'Curva ABC' existam.");
    return;
  }

  let resultados = [];

  // Verificar se a Curva ABC já foi gerada
  const curvaABCData = curvaABCSheet.getDataRange().getValues();
  const hasCurveABC = curvaABCData.length > 1; // Verifica se há mais de uma linha (cabeçalho e dados)

  if (hasCurveABC) {
    // Caso Curva ABC já exista, expanda para ABCD
    const headers = curvaABCData[0]; // Cabeçalho
    const data = curvaABCData.slice(1); // Dados sem o cabeçalho
    const percentualIndex = headers.indexOf("Percentual Acumulado (%)");

    if (percentualIndex === -1) {
      SpreadsheetApp.getUi().alert("A coluna 'Percentual Acumulado (%)' não foi encontrada na Curva ABC.");
      return;
    }

    resultados = data.map((row) => {
      const percentual = parseFloat(row[percentualIndex]) || 0;
      const categoria = getCategoria(percentual);
      row[headers.length] = categoria; // Atualiza a última coluna (Categoria)
      return row;
    });
  } else {
    // Caso Curva ABC não exista, recalcular tudo a partir de "Pedidos"
    resultados = calcularCurvaABC(pedidosSheet);
    curvaABCSheet.clear(); // Limpa dados antigos
    curvaABCSheet.appendRow(["ClienteCNPJ", "Total Vendido", "Percentual Acumulado (%)", "Categoria", "Categoria ABCD"]);
  }

  // Atualizar planilha de Curva ABCD
  curvaABCSheet.getRange(2, 1, resultados.length, 5).setValues(resultados); // 5 colunas

  // Gerar gráfico da Curva ABCD
  gerarGrafico(curvaABCSheet, resultados);

  SpreadsheetApp.getUi().alert("Curva ABCD gerada com sucesso!");
}
