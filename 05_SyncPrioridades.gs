/**
 * 05_SyncPrioridades.gs
 * Sincroniza e gerencia as prioridades com disparos específicos (SECOM/DISUP).
 * Deve estar na Planilha DASHBOARD.
 */

// Cria a lista suspensa (Janaina/Julio) automaticamente ao editar a Coluna A
function onEdit(e) {
  const range = e.range;
  const sheet = range.getSheet();
  
  // Se estiver editando a guia "Prioridades", Coluna A (Processo), Linha > 1
  if (sheet.getName() === "Prioridades" && range.getColumn() === 1 && range.getRow() > 1) {
    const celulaDisparo = sheet.getRange(range.getRow(), 2); // Coluna B
    
    // Cria a validação na Coluna B
    const regra = SpreadsheetApp.newDataValidation()
      .requireValueInList(['Janaina', 'Julio'], true)
      .setAllowInvalid(false)
      .build();
      
    celulaDisparo.setDataValidation(regra);
  }
}

/**
 * Sincroniza a lista de "Prioridades" do Dashboard com a Planilha Operacional.
 */
function sincronizarPrioridades() {
  const ssDash = SpreadsheetApp.getActiveSpreadsheet();
  const abaPrioridades = ssDash.getSheetByName("Prioridades");
  
  if (!abaPrioridades) {
    SpreadsheetApp.getUi().alert('A guia "Prioridades" não foi encontrada neste Dashboard.');
    return;
  }

  // ID da Planilha de Controle de Processos (Operacional)
  // Substitua pelo ID correto da sua planilha 15W847YN...
  const ID_PLANILHA_OPERACIONAL = "15W847YN-SEU-ID-AQUI"; 

  // 1. Pega a lista de Processos e Disparos
  const ultimaLinha = abaPrioridades.getLastRow();
  let mapaPrioridades = {}; // Objeto para armazenar Processo -> Disparo

  if (ultimaLinha >= 2) {
    // Pega Coluna A (Processo) e Coluna B (Disparo)
    const dados = abaPrioridades.getRange(2, 1, ultimaLinha - 1, 2).getValues();
    
    dados.forEach(linha => {
      const proc = String(linha[0]).trim();
      const disparo = String(linha[1]).trim(); // Janaina ou Julio
      if (proc !== "") {
        mapaPrioridades[proc] = disparo;
      }
    });
  }

  // 2. Abre a Planilha Operacional
  let ssOperacional;
  try {
    ssOperacional = SpreadsheetApp.openById(ID_PLANILHA_OPERACIONAL);
  } catch (e) {
    SpreadsheetApp.getUi().alert("Erro ao abrir planilha Operacional. Verifique o ID no script.");
    return;
  }

  const todasAbas = ssOperacional.getSheets();
  
  // Abas que não devem ser verificadas
  const ABAS_IGNORAR = ["ToFor", "Modal_Config", "Resumo", "Dashboard", "Config"]; 

  let marcados = 0;
  let desmarcados = 0;

  // 3. Varre todas as abas de compradores
  todasAbas.forEach(aba => {
    const nomeAba = aba.getName();
    
    if (!ABAS_IGNORAR.includes(nomeAba)) {
      const lastRow = aba.getLastRow();
      
      if (lastRow > 1) {
        // Lê os processos existentes na aba (Coluna A)
        const rangeProcessos = aba.getRange(2, 1, lastRow - 1, 1);
        const valoresProcessos = rangeProcessos.getValues().flat().map(String);
        
        valoresProcessos.forEach((processo, index) => {
          const linhaReal = index + 2;
          const colIndexPrioridade = 10; // Coluna J (10) - Conforme 02_Guias.gs
          const procLimpo = processo.trim();
          
          const celulaStatus = aba.getRange(linhaReal, colIndexPrioridade);
          
          if (mapaPrioridades.hasOwnProperty(procLimpo)) {
            // === É UMA PRIORIDADE ===
            const quemDisparou = mapaPrioridades[procLimpo];
            let textoStatus = "PRIORIDADE";
            
            if (quemDisparou === "Julio") textoStatus = "prioridade SECOM";
            else if (quemDisparou === "Janaina") textoStatus = "prioridade DISUP";

            // Se o status for diferente, atualiza
            if (celulaStatus.getValue() !== textoStatus) {
              celulaStatus.setValue(textoStatus).setFontWeight("bold").setFontColor("red");
              
              // Opcional: Pintar a linha de amarelo para destacar
              aba.getRange(linhaReal, 1, 1, 20).setBackground("#FFF2CC");
              marcados++;
            }
            
          } else {
            // === NÃO É PRIORIDADE ===
            const valorAtual = celulaStatus.getValue();
            
            // Se tiver marcado como prioridade antiga, limpa
            if (valorAtual === "prioridade SECOM" || valorAtual === "prioridade DISUP") {
              celulaStatus.clearContent();
              // Opcional: Resetar cor da linha (precisaria da lógica de cores padrão)
              // aba.getRange(linhaReal, 1, 1, 20).setBackground("white"); 
              desmarcados++;
            }
          }
        });
      }
    }
  });
  
  SpreadsheetApp.getUi().alert(`Sincronização Concluída!\n\n🔥 Atualizados: ${marcados}\n🧹 Limpos: ${desmarcados}`);
}