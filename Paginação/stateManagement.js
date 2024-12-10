function saveProgress(currentPage) {
  if (currentPage !== undefined && currentPage !== null) {
    PropertiesService.getScriptProperties().setProperty('LAST_PAGE', currentPage.toString());
    Logger.log(`Progresso salvo: página ${currentPage}`);
  } else {
    Logger.log("Erro: currentPage está indefinido ou nulo.");
  }
}


function getLastPage() {
  const lastPage = PropertiesService.getScriptProperties().getProperty('LAST_PAGE');
  if (lastPage !== null) {
    const page = parseInt(lastPage, 10);
    Logger.log(`Última página recuperada: ${page}`);
    return page; 
  } else {
    Logger.log("Nenhuma página salva anteriormente. Iniciando na página 1.");
    return 1; 
  }
}
