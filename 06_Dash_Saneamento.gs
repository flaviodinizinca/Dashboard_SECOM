/**
 * 06_Dash_Saneamento.gs
 * Agora aponta para a PLANILHA EXCLUSIVA DE SANEAMENTO.
 */
function atualizarDashboardSaneamento() {
  const ssDash = SpreadsheetApp.getActiveSpreadsheet();
  
  // ID DA NOVA PLANILHA DE SANEAMENTO
  const ID_PLANILHA_SANEAMENTO = "1TxyCWwg9IBZpEh9g6E_PgNUx5ucR_CwlTCaS_eXihTs"; 

  let ssSan;
  try {
    ssSan = SpreadsheetApp.openById(ID_PLANILHA_SANEAMENTO);
  } catch (e) {
    SpreadsheetApp.getUi().alert("Erro ao abrir Planilha Saneamento: " + e.message);
    return;
  }

  const todasAbas = ssSan.getSheets();
  let dados = [];

  // Pula abas de configuração
  const ignorar = ["Config_Saneamento", "Página1"]; 

  todasAbas.forEach(aba => {
    if (!ignorar.includes(aba.getName())) {
      const lastRow = aba.getLastRow();
      if (lastRow > 1) {
        const valores = aba.getRange(2, 1, lastRow - 1, 11).getValues();
        valores.forEach(lin => {
          dados.push([
            aba.getName(), // Responsável
            lin[0], lin[1], lin[4], lin[8], lin[10] // Proc, Data, Obj, Encerrado, Status
          ]);
        });
      }
    }
  });

  // Atualiza Resumo no Dashboard
  let abaResumo = ssDash.getSheetByName("Resumo Saneamento");
  if (!abaResumo) abaResumo = ssDash.insertSheet("Resumo Saneamento");
  else abaResumo.clear();

  const header = ["Responsável", "Processo", "Data", "Objeto", "Encerrado?", "Status"];
  abaResumo.getRange(1,1,1,6).setValues([header]).setFontWeight("bold").setBackground("#E69138").setFontColor("white");
  
  if (dados.length > 0) {
    abaResumo.getRange(2, 1, dados.length, 6).setValues(dados);
    abaResumo.getRange(2, 3, dados.length, 1).setNumberFormat("dd/mm/yyyy");
  }
}