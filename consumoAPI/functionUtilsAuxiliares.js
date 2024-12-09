const API_URL = "https://api.sigecloud.com.br/request/Pedidos/GetTodosPedidos?page=1";

function formatHeaders(sheet, columnCount) {
  if (!sheet) {
    Logger.log("Erro: A planilha fornecida para formatar os cabeçalhos é indefinida.");
    return;
  }

  sheet.getRange(1, 1, 1, columnCount)
    .setFontWeight("bold")
    .setBackground("#D3D3D3");
}


function isValidJSONString(str) {
  try {
    JSON.parse(str);
    return true;
  } catch (e) {
    return false;
  }
}
