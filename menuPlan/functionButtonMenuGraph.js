function onOpenGraph() {
  const ui = SpreadsheetApp.getUi();

  // Adiciona o menu principal para gráficos
  ui.createMenu("Menu Gráfico")
    .addItem("Gerar Curva ABC", "gerarCurvaABC")
    .addItem("Gerar Curva ABCD", "gerarCurvaABCD")
    .addItem("Limpar Gráficos", "clearGraphsAndData")
    .addToUi();

  addUpdateMenuForABC(); 
}

function addUpdateMenuForABC() {
  const ui = SpreadsheetApp.getUi();
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

  if (sheet.getName() === "Curva ABC") {
    ui.createMenu("Atualizar")
      .addItem("Atualizar Curva ABC", "atualizarCurvaABC")
      .addItem("Atualizar Curva ABCD", "atualizarCurvaABCD")
      .addToUi();
  }
}
