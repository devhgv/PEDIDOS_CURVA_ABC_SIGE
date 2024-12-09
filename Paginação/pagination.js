function fetchPedidosWithPagination(API_BASE_URL, params, sheet, headers) {
  let currentPage = getLastPage();  // Obtém a última página salva
  const pageSize = 100; // Ajuste o tamanho da página conforme necessário

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
        // Verifica se existem itens no pedido
        if (pedido.Items && Array.isArray(pedido.Items)) {
          pedido.Items.forEach((item) => {
            // Adiciona uma linha para cada item no pedido
            sheet.appendRow([
              pedido.ID || "N/A",
              pedido.DataEnvio || "N/A",
              pedido.Cliente || "N/A",
              pedido.ClienteCNPJ || "N/A",
              item.Codigo || "N/A",     
              item.Descricao || "N/A",
              item.ValorTotal || 0,   
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

// Função para salvar o progresso (última página visitada)
function saveProgress(page) {
  const scriptProperties = PropertiesService.getScriptProperties();
  scriptProperties.setProperty('lastPage', page.toString());  // Armazena a última página
}

// Função para obter a última página salva
function getLastPage() {
  const scriptProperties = PropertiesService.getScriptProperties();
  const lastPage = scriptProperties.getProperty('lastPage');
  return lastPage ? parseInt(lastPage) : 1;  // Se não houver progresso salvo, começa na página 1
}
