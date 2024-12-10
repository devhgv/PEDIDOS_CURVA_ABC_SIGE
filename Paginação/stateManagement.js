
function saveProgress(currentPage) {
  if (currentPage !== undefined && currentPage !== null) {
    PropertiesService.getScriptProperties().setProperty('LAST_PAGE', currentPage.toString());
    Logger.log("Progresso salvo: Página " + currentPage);
  } else {
    Logger.log("Erro: currentPage está indefinido ou nulo.");
  }
}


function getLastPage() {
  const lastPage = PropertiesService.getScriptProperties().getProperty('LAST_PAGE');
  if (lastPage !== null) {
    Logger.log("Progresso recuperado: Página " + lastPage);
    return parseInt(lastPage, 10); 
  } else {
    Logger.log("Nenhum progresso encontrado. Começando pela página 1.");
    return 1;  // Se não houver um valor salvo, começa pela página 1
  }
}

