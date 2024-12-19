function main() {
  fetchAllPedidosFromMiddleware(); 
  updateVerification();
}
// Para execuções manuais

function gerarTodasCurvasABC(){ // Adicionado acionador GAS, para todo sábado das 18h às 19h.
  curvaABC3Meses();
  curvaABC6Meses();
  curvaABC12Meses();
  curvaABC18Meses();
  curvaABC24Meses();
  Logger.log("Todas as curvas ABC foram atualizadas com sucesso!");
}