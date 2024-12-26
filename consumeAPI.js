// FUNÇÃO PARA CONSUMIR A API DO SIGE.
function consumeAPI() {
  const tab = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  if (!tab) {
    Logger.log ("Aba da Planilha não localizada!")
    return;
  }

const url_api = "https://api.sigecloud.com.br/request/Pedidos/GetTodosPedidos";

const access_api = {
     method: "GET",
     headers: {
        "Accept": "application/json",
        "Content-Type": "application/json",
        "Authorization-Token": "700dae74083f9f420d73b6602d88c3687aca653de8da145b82eb61c4e6b7539227ee7f7efe2dbd21875714c55282ee5c401a4633ff76c2f8e530a3d8ba9501be7b84e090dc02ddc953075acb2c1b404cb5a526002697cd1bf63e95ea4e7e769a96556e387144d872c220c3917327e50b4afe6161e8994745cd3125d96c8542e2",
        "User": "encarregadoportoreal@gmail.com",
        "App": "API Felipe"
     },
}

const headers_tab = ["ID", "DataEnvio", "DataFaturada", "Cliente", "ClienteCNPJ", "ItensJSON", "Categoria", "Empresa", "ValorFinal", "StatusSistema", "UltimaAtualizacaoDados"];

    if (!tab.getRange(1, 1, 1, headers_tab.length).getValues()[0].some(val => val)) {
      tab.appendRow(headers_tab);
      formatHeaders(tab, headers_tab.length);
    } 

    Logger.log("Buscando dados dos pedidos com paginação...");
    paginationAPI(tab, url_api, access_api, headers);
    Logger.log("Importação de pedidos concluída com sucesso!");
} 

// FUNÇÃO QUE FORMATA OS CABEÇALHOS CRIADOS NA FUNÇÃO ANTERIOR.
function formatHeaders (tab, columnCount) { 
    if (!tab) {
      Logger.log ("ERRO: A planilha fornecida para formatar os cabeçalhos não está disponível ou está indefinida.");
      return;
    } else {
      tab.getRange(1, 1, 1, columnCount)
      .setFontWeight("bold")
      .setBackground("#D3D3D3");
      Logger.log("Cabeçalhos criados e formatados na Planilha - Aba: consAPI")
    }
}

// FUNÇÕES PARA GERENCIAR O PROGRESSO DAS PÁGINAS DA API E DAS LINHAS DA PLANILHA (SALVA E RETOMA).
function getLasPage() {
    const lastPage = PropertiesService.getScriptProperties().getProperty("LAST_PAGE");
    return lastPage ? parseInt(lastPage, 10) : 1;
}

function getLastRow() {
    const lastRow = PropertiesService.getScriptProperties().getProperty("LAST_ROW");
    return lastRow ? parseInt(lastRow, 10) : 2;
 }

function saveProgress(currentPage, currentRow) {
    if (currentPage !== undefined && currentPage !== null) {
       PropertiesService.getScriptProperties().setProperty("LAST_PAGE", currentPage.toString());
       PropertiesService.getScriptProperties().setProperty("LAST_ROW", currentRow.toString());
       Logger.log (`Progresso salvo: Página ${currentPage} na Linha ${currentRow}`);
  } else {
       Logger.log ("ERRO: A página não está definida ou não existe!");
  }
}

// FUNÇÃO PARA VALIDAR SE O RETORNO DA STRING ESTÁ EM JSON.
function validJSON (str) {
  try {
    JSON.parse(str);
    return true;
  } catch (e) {
    return false;
  }
}

// FUNÇÃO PARA FORMATAR OS VALORES DA COLUNA "DATA ENVIO".
function formatDate(dateString) {
  const date = new Date(dateString);
  if (isNaN(date)) {
    return "Data inválida";
  }
  const options = { year: 'numeric', month: '2-digit', day: '2-digit' };
  return date.toLocaleDateString('pt-BR', options);
}

// FUNÇÃO PARA LIDAR COM A PAGINAÇÃO DA API.
function paginationAPI(url_api, access_api, tab, headers_tab) {
 const startTime = new Date().getTime(); 
 const maxExecutionTime = 5 * 60 * 1000;

    let currentPage = getLasPage();
    let currentRow = getLastRow();

    const dataAtualizacao = new Date();

    while(true) {
      try {
        const verificationTime = new Date().getTime() - startTime;
        if (verificationTime >= maxExecutionTime) {
          Logger.log('Tempo Limite de Execução atingido para este script. Parando o processo de paginação (...)');
          saveProgress (currentPage, currentRow);
          break;
        }

        const url = `${url_api}?page=${currentPage}`;
        Logger.log (`Buscando dados da página ${currentPage} (...)`);
        const response = UrlFetchApp.fetch(url, access_api);
        const statusCode = response.getResponseCode();

        if (statusCode !== 200) {
          Logger.log(`Erro ao acessar a API: ${statusCode}`);
          Logger.log(`Resposta do servidor: ${response.getContentText()}`);
          break;
        }
      
        const responseText = response.getContentText();
        if (!validJSON(responseText)) {
          Logger.log ("Resposta não é JSON válida: "+ responseText);
          break;
        } 

        const pedidos = JSON.parse(responseText);
        if (!Array.isArray(pedidos) || pedidos.length === 0) {
          Logger.log ("Nenhum pedido encontrado na página.");
          break;
        }

        const data = pedidos.map(pedido => [
        pedido.ID || "N/A",
        pedido.DataEnvio || "N/A",
        formatDate(pedido.DataEnvio || "N/A"),
        pedido.Cliente || "N/A",
        pedido.ClienteCNPJ || "N/A",
        JSON.stringify(pedido.Items || []),
        pedido.Categoria || "N/A",
        pedido.Empresa || "N/A",
        pedido.ValorFinal || 0,
        pedido.StatusSistema || "N/A",
        dataAtualizacao
        ]);

  tab.getRange(currentRow, 1, data.length, headers_tab.length).setValues(data);
  currentRow += data.length;

  saveProgress(currentPage, currentRow);
  Logger.log(`Página ${currentPage} processada com sucesso.`);
  currentPage++;
      } catch (error) {
        Logger.log (`Erro ao processar página ${currentPage}: ${error.message}\nStack: ${error.stack}`);
        break;
      }
    }
  }

 




























