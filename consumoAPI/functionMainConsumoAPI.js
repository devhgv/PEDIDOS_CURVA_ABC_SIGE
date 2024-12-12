function fetchAllPedidosFromMiddleware() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  if (!sheet) {
    Logger.log("Erro: Não foi possível obter a planilha ativa.");
    return;
  }

  Logger.log("Criando cabeçalhos na planilha(...)");
  const API_BASE_URL = "https://api.sigecloud.com.br/request/Pedidos/GetTodosPedidos";

  const params = {
    method: "GET",
    headers: {
      "Accept": "application/json",
      "Content-Type": "application/json",
      "Authorization-Token": "700dae74083f9f420d73b6602d88c3687aca653de8da145b82eb61c4e6b7539227ee7f7efe2dbd21875714c55282ee5c401a4633ff76c2f8e530a3d8ba9501be7b84e090dc02ddc953075acb2c1b404cb5a526002697cd1bf63e95ea4e7e769a96556e387144d872c220c3917327e50b4afe6161e8994745cd3125d96c8542e2",
      "User": "encarregadoportoreal@gmail.com",
      "App": "API Felipe"
    },
    muteHttpExceptions: true
  };

  const headers = ["ID", "DataEnvio", "Cliente", "ClienteCNPJ", "ItensJson", "Categoria", "Empresa", "ValorFinal", "StatusSistema"];

  if (!sheet.getRange(1, 1, 1, headers.length).getValues()[0].some(val => val)) {
    sheet.appendRow(headers);
    formatHeaders(sheet, headers.length);
  }

  Logger.log("Buscando dados dos pedidos com paginação...");
  fetchPedidosWithPagination(API_BASE_URL, params, sheet, headers);
  Logger.log("Importação de pedidos concluída com sucesso!");
}

function formatHeaders(sheet, columnCount) {
  if (!sheet) {
    Logger.log("Erro: A planilha fornecida para formatar os cabeçalhos é indefinida.");
    return;
  }
  sheet.getRange(1, 1, 1, columnCount)
    .setFontWeight("bold")
    .setBackground("#D3D3D3");
  Logger.log("Cabeçalhos formatados com fundo cinza e texto em negrito.");
}

function saveProgress(currentPage, currentRow) {
  if (currentPage !== undefined && currentPage !== null) {
    PropertiesService.getScriptProperties().setProperty('LAST_PAGE', currentPage.toString());
    PropertiesService.getScriptProperties().setProperty('LAST_ROW', currentRow.toString());
    Logger.log(`Progresso salvo: Página ${currentPage}, Linha ${currentRow}`);
  } else {
    Logger.log("Erro: currentPage está indefinido ou nulo.");
  }
}

function getLastPage() {
  const lastPage = PropertiesService.getScriptProperties().getProperty('LAST_PAGE');
  return lastPage ? parseInt(lastPage, 10) : 1;
}

function getLastRow() {
  const lastRow = PropertiesService.getScriptProperties().getProperty('LAST_ROW');
  return lastRow ? parseInt(lastRow, 10) : 2; // Linha 2, assumindo cabeçalhos na 1
}

function isValidJSONString(str) {
  try {
    JSON.parse(str);
    return true;
  } catch (e) {
    return false;
  }
}
