function consolidarItens() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const pedidosSheet = ss.getSheetByName("Pedidos");
  const itensSheet = ss.getSheetByName("Itens");

  // Limpa a aba "Itens" para atualizar os dados
  itensSheet.clearContents();
  itensSheet.appendRow(["Descrição", "Valor Total"]);

  // Obtem os dados das colunas relevantes da aba "Pedidos"
  const data = pedidosSheet.getDataRange().getValues();
  const colItensJson = 4; // Coluna E (Índice base 0)
  const colStatusSistema = 8; // Coluna I (Índice base 0)

  const consolidado = {}; // Objeto para armazenar soma de valores por descrição

  // Itera sobre as linhas, ignorando o cabeçalho
  for (let i = 1; i < data.length; i++) {
    const statusSistema = data[i][colStatusSistema];
    const itensJson = data[i][colItensJson];

    // Processa apenas linhas com "Pedido Faturado" e JSON válido
    if (statusSistema === "Pedido Faturado" && isValidJson(itensJson)) {
      const itens = JSON.parse(itensJson);

      // Itera pelos itens do JSON
      itens.forEach(item => {
        const descricao = item.Descricao;
        const valorUnitario = parseFloat(item.ValorUnitario) || 0;

        if (descricao) {
          // Soma os valores agrupados por descrição
          consolidado[descricao] = (consolidado[descricao] || 0) + valorUnitario;
        }
      });
    }
  }

  // Insere os dados consolidados na aba "Itens"
  const output = Object.entries(consolidado).map(([descricao, valorTotal]) => [descricao, valorTotal]);
  itensSheet.getRange(2, 1, output.length, 2).setValues(output);

  Logger.log("Consolidação concluída com sucesso!");
}

//Valida se uma string é um JSON válido
function isValidJson(str) { 
  try {
    JSON.parse(str);
    return true;
  } catch (e) {
    return false;
  }
}
