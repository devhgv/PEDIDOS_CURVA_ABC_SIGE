function consolidarItensQtd() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const pedidosSheet = ss.getSheetByName("Pedidos");
  const itensSheet = ss.getSheetByName("Curva ABC/ ABCD (qtd)");

  // Limpa a aba de itens antes de atualizar os dados
  itensSheet.clearContents();
  itensSheet.appendRow(["Data Faturada", "Descrição", "Quantidade Vendida"]);

  // Define os índices das colunas
  const colDataFaturada = 2; // Coluna C
  const colItensJson = 5; // Coluna F
  const colStatusSistema = 9; // Coluna J

  const data = pedidosSheet.getDataRange().getValues();

  const mesesFiltro = ""; // Adicione o número de meses.
  const hoje = new Date();
  const dataLimite = new Date(hoje.setMonth(hoje.getMonth() - mesesFiltro));

  const consolidado = {}; // Objeto para consolidar a soma de quantidades por descrição

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

      // Itera sobre os itens do JSON
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
