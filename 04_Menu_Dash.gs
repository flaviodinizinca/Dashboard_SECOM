/**
 * 04_Menu_Dash.gs
 * Menu centralizado do Dashboard.
 */

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  
  ui.createMenu('📊 Dashboard SECOM')
    .addItem('🔄 Atualizar Tudo (Geral)', 'atualizarDashboardGeral') 
    .addItem('🛠️ Atualizar Saneamento', 'atualizarDashboardSaneamento')
    .addSeparator()
    .addItem('🔥 Sincronizar Prioridades', 'sincronizarPrioridades')
    .addToUi();
}

/**
 * Função Wrapper para garantir a execução segura.
 */
function atualizarDashboardGeral() {
  const ui = SpreadsheetApp.getUi();
  
  try {
    // Tenta executar a função principal que está no 03_Builder.gs
    if (typeof construirDashboard === 'function') {
      construirDashboard();
      ui.alert("Atualização Geral concluída com sucesso!");
    } else {
      ui.alert("Erro: A função 'construirDashboard' não foi encontrada. Verifique se o arquivo 03_Builder.gs está salvo.");
    }
  } catch (e) {
    ui.alert("Erro crítico ao tentar atualizar: " + e.message);
  }
}