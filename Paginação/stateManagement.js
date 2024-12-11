function saveProgress(currentPage) {
  if (Number.isInteger(currentPage) && currentPage > 0) {
    PropertiesService.getScriptProperties().setProperty('LAST_PAGE', currentPage.toString());
    Logger.log("Progresso salvo: Página " + currentPage);
  } else {
    Logger.log("Erro: currentPage inválido ao salvar progresso. Valor recebido: " + currentPage);
  }
}

function getLastPage() {
  const lastPage = PropertiesService.getScriptProperties().getProperty('LAST_PAGE');
  if (lastPage !== null && !isNaN(Number(lastPage)) && Number(lastPage) > 0) {
    Logger.log("Progresso recuperado: Página " + lastPage);
    return Number(lastPage); 
  } else {
    Logger.log("Nenhum progresso válido encontrado. Começando pela página 1.");
    return 1; // Valor padrão se não houver progresso salvo ou valor inválido
  }
}

