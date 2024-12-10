function fetchPedidosWithPagination(API_BASE_URL, params, sheet, headers) {
  let currentPage = 1;
  const pageSize = 100; // Ajuste o tamanho da página conforme necessário

  while (true) {
    try {
      // Constrói a URL com base na página atual
      const url = `${API_BASE_URL}?page=${currentPage}`;  
      Logger.log("Requisitando dados da URL: " + url);

      // Faz a requisição
      const response = UrlFetchApp.fetch(url, params);
      const statusCode = response.getResponseCode();

      if (statusCode !== 200) {
        Logger.log(`Erro ao acessar a API: ${statusCode}`);
        Logger.log(`Resposta do servidor: ${response.getContentText()}`);
        saveProgress(currentPage);
        break;
      }

      const responseText = response.getContentText();
      if (!isValidJSONString(responseText)) {
        Logger.log("Resposta não é JSON: " + responseText);
        saveProgress(currentPage);
        break;
      }

      const pedidos = JSON.parse(responseText);

      if (!pedidos || pedidos.length === 0) {
        Logger.log("Nenhum pedido encontrado.");
        saveProgress(currentPage);
        break;
      }

      // Processa os dados dos pedidos
      pedidos.forEach((pedido) => {
        Logger.log("Buscando Dados dos Pedidos (...) ");
        if (pedido.Items && Array.isArray(pedido.Items)) {
          pedido.Items.forEach((item) => {

            const itemJson = JSON.stringify(item); //Transforma Itens em JSON
            
            sheet.appendRow([ 
              pedido.ID || "N/A", 
              pedido.DataEnvio || "N/A", 
              pedido.Cliente || "N/A", 
              pedido.ClienteCNPJ || "N/A", 
              itemJson,
              pedido.Categoria || "N/A", 
              pedido.Empresa || "N/A", 
              pedido.ValorFinal || 0, 
              pedido.StatusSistema || "N/A"
            ]);
          });
        } else {
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

      // Salva o progresso
      Logger.log("Página atual: " + currentPage);
      saveProgress(currentPage);
      currentPage++;

      // Condição de parada se não houver mais páginas
      if (pedidos.length < pageSize) {
        Logger.log("Todos os pedidos processados.");
        break;
      }

      // Verifica tempo restante
      if (Session.getScriptTimeRemaining() < 30) {
        Logger.log("Resta 30s para o tempo se esgotar. Salvando progresso na página " + currentPage);
        saveProgress(currentPage);
        break;
      }
    } catch (error) {
      Logger.log("Erro ao buscar os pedidos: " + error.message);
      saveProgress(currentPage);
      break;
    }
  }

  // Auto ajuste das colunas
  sheet.autoResizeColumns(1, headers.length);
}
