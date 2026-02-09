/**
 * 06_Dash_Saneamento.gs
 * Extrai dados exclusivamente das abas de SANEAMENTO.
 */

function atualizarDashboardSaneamento() {
  const ssDash = SpreadsheetApp.getActiveSpreadsheet();
  
  // ID da Planilha Operacional (Controle)
  const ID_PLANILHA_CONTROLE = "1n6l2ofxEvQTrZ49IY7b30U_dcUqb-MuAbVaW890S6ng"; // <--- COLOQUE SEU ID AQUI

  let ssControle;
  try {
    ssControle = SpreadsheetApp.openById(ID_PLANILHA_CONTROLE);
  } catch (e) {
    SpreadsheetApp.getUi().alert("Erro ao abrir planilha de Controle.");
    return;
  }

  const todasAbas = ssControle.getSheets();
  let dadosSaneamento = [];

  // 1. VARREDURA DE ABAS
  todasAbas.forEach(aba => {
    // Identifica abas de saneamento pelo nome
    if (aba.getName().includes("(Saneamento)")) {
      const lastRow = aba.getLastRow();
      if (lastRow > 1) {
        // Pega colunas A até K (11 colunas)
        const valores = aba.getRange(2, 1, lastRow - 1, 11).getValues();
        
        valores.forEach(linha => {
          // Adiciona o nome do responsável (Nome da aba sem o sufixo)
          const responsavel = aba.getName().replace(" (Saneamento)", "").trim();
          
          // Estrutura do objeto para o Dashboard
          dadosSaneamento.push([
            responsavel,       // Coluna 1: Responsável
            linha[0],          // Coluna 2: Processo
            linha[1],          // Coluna 3: Data Chegada
            linha[4],          // Coluna 4: Objeto
            linha[8],          // Coluna 5: Encerrado?
            linha[10]          // Coluna 6: Status
          ]);
        });
      }
    }
  });

  // 2. ATUALIZAR ABA DE RESUMO NO DASHBOARD
  let abaResumo = ssDash.getSheetByName("Resumo Saneamento");
  if (!abaResumo) {
    abaResumo = ssDash.insertSheet("Resumo Saneamento");
  } else {
    abaResumo.clear(); // Limpa dados antigos
  }

  // Cabeçalho do Resumo
  const cabecalho = ["Responsável", "Processo", "Data Chegada", "Objeto", "Encerrado?", "Status"];
  abaResumo.getRange(1, 1, 1, cabecalho.length).setValues([cabecalho]).setFontWeight("bold").setBackground("#E69138").setFontColor("white");

  // Insere os dados
  if (dadosSaneamento.length > 0) {
    abaResumo.getRange(2, 1, dadosSaneamento.length, cabecalho.length).setValues(dadosSaneamento);
    
    // Formatação Básica
    abaResumo.autoResizeColumns(1, cabecalho.length);
    abaResumo.getRange(2, 2, dadosSaneamento.length, 1).setNumberFormat("@"); // Processo como texto
    abaResumo.getRange(2, 3, dadosSaneamento.length, 1).setNumberFormat("dd/mm/yyyy"); // Data
  }

  SpreadsheetApp.getUi().alert(`Dashboard de Saneamento atualizado com ${dadosSaneamento.length} registros.`);
}

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('📊 Dashboard SECOM')
    .addItem('Atualizar Tudo', 'atualizarDashboardGeral') // Sua função existente
    .addItem('Atualizar Saneamento', 'atualizarDashboardSaneamento')
    .addToUi();
}