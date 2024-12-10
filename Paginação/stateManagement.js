function saveProgress(currentPage) {
  if (currentPage !== undefined && currentPage !== null) {
    PropertiesService.getScriptProperties().setProperty('LAST_PAGE', currentPage.toString());
  } else {
    Logger.log("Erro: currentPage está indefinido ou nulo.");
  }
}

function getLastPage() {
  const lastPage = PropertiesService.getScriptProperties().getProperty('LAST_PAGE');
  if (lastPage !== null) {
    return parseInt(lastPage, 10); 
  } else {
    return 1;  // Se não houver um valor salvo, começa pela página 1
  }
}
