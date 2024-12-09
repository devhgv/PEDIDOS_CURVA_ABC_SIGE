function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu("Menu")
    .addItem("Atualizar", "fetchAllPedidosFromMiddleware")  // Atualizar dados com novos pedidos
    .addItem("Sobre", "showAbout")
    .addToUi();
}

function showAbout() {
  const ui = SpreadsheetApp.getUi();
  ui.alert(
    "Sobre",
    "Este metódo exibe os pedidos do sistema. Utilize o Botão 'Atualizar' para carregar os pedidos mais recentes dos últimos meses.",
    ui.ButtonSet.OK
  );
}
