// FUNÇÃO PRINCIPAL PARA RODAR O CONSUMO E PAGINAÇÃO DA API.
function runAPIProcess() { //- Essa função roda automaticamente através de um acionador GAS todos os sábados das 16h às 17h. 
  try {
    Logger.log ("Iniciando o processo de consumo da API (...)");

    const tab = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    if (!tab) {
      Logger.log("Erro: Aba da planilha não localizada!");
      return;
    }

const headers_tab = ["ID", "DataEnvio", "DataFaturada", "Cliente", "ClienteCNPJ", "ItensJSON", "Categoria", "Empresa", "ValorFinal", "StatusSistema", "UltimaAtualizacaoDados"];
  if (!tab.getRange(1, 1, 1, headers_tab.length).getValues()[0].some(val => val)) {
    tab.appendRow(headers_tab);
    formatHeaders(tab,headers_tab.length);
  }

  const url_api = "https://api.sigecloud.com.br/request/Pedidos/GetTodosPedidos";
  const access_api = {
      method: "GET",
      headers: {"Accept": "application/json",
          "Content-Type": "application/json",
          "Authorization-token": "700dae74083f9f420d73b6602d88c3687aca653de8da145b82eb61c4e6b7539227ee7f7efe2dbd21875714c55282ee5c401a4633ff76c2f8e530a3d8ba9501be7b84e090dc02ddc953075acb2c1b404cb5a526002697cd1bf63e95ea4e7e769a96556e387144d872c220c3917327e50b4afe6161e8994745cd3125d96c8542e2",
          "User": "encarregadoportoreal@gmail.com",
          "App": "API Felipe"
          }
  };

paginationAPI(url_api, access_api, tab, headers_tab);
Logger.log ("Processo concluído com sucesso!");
  } catch (error) {
    Logger.log (`Erro no Processo: ${error.message}\nStack: ${error.stack}`);
  }
}

function updateAllCurveABC () { //- Essa função roda automaticamente através de um acionador GAS todos os sábados das 18h às 19h. 
  Logger.log ("Atualizando Curva ABC dos últimos 3 Meses (...)");
  createCurve3Month();
  Logger.log ("Atualizando Curva ABC dos últimos 6 Meses (...)");
  createCurve6Month();
  Logger.log ("Atualizando Curva ABC dos últimos 12 Meses (...)");
  createCurve12Month();
  Logger.log ("Atualizando Curva ABC dos últimos 18 Meses (...)");
  createCurve18Month();
  Logger.log ("Atualizando Curva ABC dos últimos 24 Meses (...)");
  createCurve24Month();
  Logger.log ("Todas as Curvas ABC foram atualizadas!")
}

