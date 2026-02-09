function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('🚀 Painel de Controle')
    .addItem('📊 Atualizar Dashboard', 'construirDashboard')
    .addSeparator()
    .addItem('🔥 Enviar Prioridades para Operação', 'sincronizarPrioridades')
    .addToUi();
}