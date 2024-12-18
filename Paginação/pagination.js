function fetchPedidosWithPagination(API_BASE_URL, params, sheet, headers) {
  let currentPage = getLastPage(); // Obtém a última página salva
  let currentRow = getLastRow(); // Obtém a última linha salva

  while (true) {
    try {
      const url = `${API_BASE_URL}?page=${currentPage}`;
      Logger.log(`Processando página ${url}...`);
      const response = UrlFetchApp.fetch(url, params);
      const statusCode = response.getResponseCode();

      if (statusCode !== 200) {
        Logger.log(`Erro ao acessar a API: ${statusCode}`);
        Logger.log(`Resposta do servidor: ${response.getContentText()}`);
        break;
      }

      const responseText = response.getContentText();
      if (!isValidJSONString(responseText)) {
        Logger.log("Resposta não é JSON: " + responseText);
        break;
      }

      const pedidos = JSON.parse(responseText);
      if (!Array.isArray(pedidos) || pedidos.length === 0) {
        Logger.log("Nenhum pedido encontrado na página.");
        break;
      }

      const data = pedidos.map(pedido => [
        pedido.ID || "N/A",
        pedido.DataEnvio || "N/A",
        pedido.Cliente || "N/A",
        pedido.ClienteCNPJ || "N/A",
        JSON.stringify(pedido.Items || []), 
        pedido.Categoria || "N/A",
        pedido.Empresa || "N/A",
        pedido.ValorFinal || 0,
        pedido.StatusSistema || "N/A"
      ]);

      sheet.getRange(currentRow, 1, data.length, headers.length).setValues(data);
      currentRow += data.length;

      saveProgress(currentPage, currentRow);

      if (pedidos.length < 100) { // Verifica se a página retornou menos de 100 pedidos, pode ser alterado.
        Logger.log("Paginação concluída. Todos os dados foram processados.");
        break;
      }

      currentPage++;
    } catch (error) {
      Logger.log("Erro ao buscar os pedidos: " + error.message + "\nStack: " + error.stack);
      break;
    }
  }

  sheet.autoResizeColumns(1, headers.length);
}

