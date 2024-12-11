function fetchPedidosWithPagination(API_BASE_URL, params, sheet, headers) {
  let currentPage = 1; // Página inicial do consumo
  const pageSize = 100; // Tamanho da página, ajustável

  for (; ; currentPage++) { // Incrementa `currentPage` (outra página) automaticamente em cada iteração
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
        saveProgress(currentPage); // Salva progresso e sai do loop
        break;
      }

      const responseText = response.getContentText();
      if (!isValidJSONString(responseText)) {
        Logger.log("Resposta não é JSON: " + responseText);
        saveProgress(currentPage); // Salva progresso e sai do loop
        break;
      }

      const pedidos = JSON.parse(responseText);

      // Verifica se há dados retornados
      if (!pedidos || pedidos.length === 0) {
        Logger.log("Nenhum pedido encontrado na página " + currentPage);
        saveProgress(currentPage); // Salva progresso e encerra
        break;
      }

     // Processa os dados dos pedidos
pedidos.forEach((pedido) => {
  Logger.log("Processando dados dos pedidos (...)");

  // Verifica se o pedido tem itens e adiciona linha por linha
  if (pedido.Items && Array.isArray(pedido.Items) && pedido.Items.length > 0) {
    pedido.Items.forEach((item) => {
      const itemJson = JSON.stringify(item); // Converte o item em JSON
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
    // Adiciona um pedido sem itens
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
    Logger.log("Linha adicionada para o pedido sem itens.");
  }
});

      // Salva o progresso
      Logger.log("Página atual concluída: " + currentPage);
      saveProgress(currentPage);

      // Condição de parada: se o número de itens retornados for menor que `pageSize`
      if (pedidos.length < pageSize) {
        Logger.log("Todos os pedidos processados. Última página: " + currentPage);
        break;
      }

      // Verifica tempo restante
      if (Session.getScriptTimeRemaining() < 30) {
        Logger.log("Resta 30s para o tempo se esgotar. Salvando progresso na página " + currentPage);
        saveProgress(currentPage); // Salva progresso e sai do loop
        break;
      }
    } catch (error) {
      Logger.log("Erro ao buscar os pedidos: " + error.message);
      saveProgress(currentPage); // Salva progresso e sai do loop
      break;
    }
  }

  // Auto ajuste das colunas
  sheet.autoResizeColumns(1, headers.length);
}
