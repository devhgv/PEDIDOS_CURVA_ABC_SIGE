function fetchAllPedidosFromMiddleware() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  if (!sheet) {
    Logger.log("Erro: Não foi possível obter a planilha ativa.");
    return;
  }
  
  const API_BASE_URL = "https://api.sigecloud.com.br/request/Pedidos/GetTodosPedidos?page=1";
  
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

  const headers = ["ID", "DataEnvio", "Cliente", "ClienteCNPJ", "Codigo", "Descricao","ValorTotal", "Categoria", "Empresa", "ValorFinal", "StatusSistema"];
  sheet.appendRow(headers);
  formatHeaders(sheet, headers.length);  
  fetchPedidosWithPagination(API_BASE_URL, params, sheet, headers);
  
  Logger.log("Importação de pedidos concluída com sucesso!");
}
