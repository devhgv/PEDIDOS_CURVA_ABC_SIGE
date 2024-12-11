function fetchPedidosWithPagination(API_BASE_URL, params, sheet, headers) {
  let currentPage = getLastPage();  // Obtém a última página salva
  const pageSize = 100; // Ajuste o tamanho da página conforme necessário

  // Verifica se os cabeçalhos já existem, evitando duplicação
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(headers);
    formatHeaders(sheet, headers.length);
  }

  while (true) {
    try {
      const response = UrlFetchApp.fetch(`${API_BASE_URL}&page=${currentPage}`, params);
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

      if (!pedidos || pedidos.length === 0) {
        Logger.log("Nenhum pedido encontrado.");
        break;
      }
      
      pedidos.forEach((pedido) => {
        // Verifica se o pedido tem a propriedade "Items" e é uma array
        if (pedido.Items && Array.isArray(pedido.Items)) {
          // Para cada item dentro do pedido, insere os dados do item como JSON na planilha
          pedido.Items.forEach((item) => {
            // Converte o objeto "item" para uma string JSON
            const itemJson = JSON.stringify(item);  // Transforma o objeto em uma string JSON

            // Agora adiciona esses dados na planilha
            sheet.appendRow([
              pedido.ID || "N/A", 
              pedido.DataEnvio || "N/A", 
              pedido.Cliente || "N/A", 
              pedido.ClienteCNPJ || "N/A", 
              itemJson,  // A célula com o objeto item em formato JSON
              pedido.ValorTotal || 0, 
              pedido.Categoria || "N/A", 
              pedido.Empresa || "N/A", 
              pedido.ValorFinal || 0, 
              pedido.StatusSistema || "N/A"
            ]);
          });
        } else {
          // Caso não existam itens, adiciona uma linha com dados básicos
          sheet.appendRow([
            pedido.ID || "N/A",
            pedido.DataEnvio || "N/A",
            pedido.Cliente || "N/A",
            pedido.ClienteCNPJ || "N/A",
            "Sem Código",
            "Sem Descrição",
            pedido.Categoria || "N/A",
            pedido.Empresa || "N/A",
            pedido.ValorFinal || 0,
            pedido.StatusSistema || "N/A"
          ]);
        }
      });

      // Salva o progresso (número da página atual)
      saveProgress(currentPage);

      currentPage++;

      // Se a página retornada não tiver pedidos, significa que não há mais dados para processar
      if (pedidos.length < pageSize) {
        break;
      }
    } catch (error) {
      Logger.log("Erro ao buscar os pedidos: " + error.message);
      break;
    }
  }

  // Auto ajuste das colunas
  sheet.autoResizeColumns(1, headers.length);
}
