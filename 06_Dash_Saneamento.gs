/**
 * 06_Dash_Saneamento.gs
 * Extrai dados exclusivamente das abas de SANEAMENTO.
 */

function atualizarDashboardSaneamento() {
  const ssDash = SpreadsheetApp.getActiveSpreadsheet();
  
  // ID DA PLANILHA DE CONTROLE (Fornecido por você)
  const ID_PLANILHA_CONTROLE = "1n6l2ofxEvQTrZ49IY7b30U_dcUqb-MuAbVaW890S6ng"; 

  let ssControle;
  try {
    ssControle = SpreadsheetApp.openById(ID_PLANILHA_CONTROLE);
  } catch (e) {
    SpreadsheetApp.getUi().alert(
      "Erro ao abrir planilha de Controle.\n" +
      "Verifique se você tem permissão de acesso a ela.\n" +
      "ID: " + ID_PLANILHA_CONTROLE + "\n" +
      "Detalhe: " + e.message
    );
    return;
  }

  const todasAbas = ssControle.getSheets();
  let dadosSaneamento = [];

  todasAbas.forEach(aba => {
    // Procura abas que tenham "(Saneamento)" no nome
    if (aba.getName().includes("(Saneamento)")) {
      const lastRow = aba.getLastRow();
      
      // Verifica se tem dados além do cabeçalho
      if (lastRow > 1) {
        try {
          // Pega colunas A até K (11 colunas) conforme estrutura definida
          const valores = aba.getRange(2, 1, lastRow - 1, 11).getValues();
          
          valores.forEach(linha => {
            const responsavel = aba.getName().replace(" (Saneamento)", "").trim();
            
            dadosSaneamento.push([
              responsavel, 
              linha[0], // Processo
              linha[1], // Data Chegada
              linha[4], // Objeto
              linha[8], // Encerrado?
              linha[10] // Status
            ]);
          });
        } catch (erroLeitura) {
          console.error("Erro ao ler aba " + aba.getName() + ": " + erroLeitura.message);
        }
      }
    }
  });

  // Atualiza a aba de Resumo no Dashboard
  let abaResumo = ssDash.getSheetByName("Resumo Saneamento");
  if (!abaResumo) {
    abaResumo = ssDash.insertSheet("Resumo Saneamento");
  } else {
    abaResumo.clear();
  }

  const cabecalho = ["Responsável", "Processo", "Data Chegada", "Objeto", "Encerrado?", "Status"];
  abaResumo.getRange(1, 1, 1, cabecalho.length)
    .setValues([cabecalho])
    .setFontWeight("bold")
    .setBackground("#E69138") // Laranja Saneamento
    .setFontColor("white");

  if (dadosSaneamento.length > 0) {
    abaResumo.getRange(2, 1, dadosSaneamento.length, cabecalho.length).setValues(dadosSaneamento);
    abaResumo.autoResizeColumns(1, cabecalho.length);
    
    // Formatação de Dados
    abaResumo.getRange(2, 2, dadosSaneamento.length, 1).setNumberFormat("@"); // Processo como texto
    abaResumo.getRange(2, 3, dadosSaneamento.length, 1).setNumberFormat("dd/mm/yyyy"); // Data
  }

  SpreadsheetApp.getUi().alert(`Dashboard de Saneamento atualizado com ${dadosSaneamento.length} registros.`);
}