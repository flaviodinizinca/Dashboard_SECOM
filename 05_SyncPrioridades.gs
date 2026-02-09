/**
 * Sincroniza a lista de "Prioridades" do Dashboard com a Planilha Operacional.
 * AGORA COM LIMPEZA: Remove a prioridade se o processo sair da lista.
 */
function sincronizarPrioridades() {
  const ssDash = SpreadsheetApp.getActiveSpreadsheet();
  const abaPrioridades = ssDash.getSheetByName("Prioridades");
  
  if (!abaPrioridades) {
    SpreadsheetApp.getUi().alert('A guia "Prioridades" não foi encontrada neste Dashboard.');
    return;
  }

  // 1. Pega a lista de Processos Prioritários
  const ultimaLinha = abaPrioridades.getLastRow();
  let listaPrioridades = [];
  
  if (ultimaLinha >= 2) {
    // Pega os valores, converte para String e remove espaços em branco
    const valoresBrutos = abaPrioridades.getRange(2, 1, ultimaLinha - 1, 1).getValues();
    listaPrioridades = valoresBrutos.flat().map(p => String(p).trim()).filter(p => p !== "");
  }

  // 2. Abre a Planilha Operacional
  const ssOperacional = SpreadsheetApp.openById(ID_PLANILHA_FONTE);
  const todasAbas = ssOperacional.getSheets();
  
  let marcados = 0;
  let desmarcados = 0;

  // 3. Varre todas as abas
  todasAbas.forEach(aba => {
    const nomeAba = aba.getName();
    
    if (!ABAS_IGNORAR.includes(nomeAba)) {
      const lastRow = aba.getLastRow();
      
      if (lastRow > 1) {
        // Lê os processos existentes na aba
        const rangeProcessos = aba.getRange(2, 1, lastRow - 1, 1);
        const valoresProcessos = rangeProcessos.getValues().flat().map(String);
        
        // Loop linha a linha
        valoresProcessos.forEach((processo, index) => {
          const linhaReal = index + 2;
          const colunaPrioridade = 20; // Coluna T
          
          if (listaPrioridades.includes(processo.trim())) {
            // === CASO 1: É PRIORIDADE ===
            // Verifica se já está marcado para não reescrever sem necessidade (otimização)
            const valorAtual = aba.getRange(linhaReal, colunaPrioridade).getValue();
            
            if (valorAtual !== "🔥 PRIORIDADE MÁXIMA") {
              // Marca
              aba.getRange(linhaReal, colunaPrioridade).setValue("🔥 PRIORIDADE MÁXIMA").setFontWeight("bold").setFontColor("red");
              aba.getRange(linhaReal, 1, 1, 20).setBackground("#FFF2CC"); // Amarelo suave
              marcados++;
            }
            
          } else {
            // === CASO 2: NÃO É (OU DEIXOU DE SER) PRIORIDADE ===
            // Precisamos limpar, mas somente se estiver marcado como prioridade
            // Isso evita apagar outras formatações da planilha
            
            const rangeStatus = aba.getRange(linhaReal, colunaPrioridade);
            const valorStatus = rangeStatus.getValue();
            
            if (valorStatus === "🔥 PRIORIDADE MÁXIMA") {
              // Limpa o texto da prioridade
              rangeStatus.clearContent();
              
              // Restaura a cor de fundo original (Zebrado ou Branco)
              // Logica simples: Se for par #FFFFFF, impar #F3F3F3 (conforme padrão do relatório)
              // Ou simplesmente limpamos o background para o padrão da guia
              
              // Vamos reaplicar a cor padrão do "Bloco Geral" e "Bloco Datas" para não ficar feio
              // Como a linha inteira foi pintada de amarelo, precisamos restaurar por partes
              
              const rangeLinha = aba.getRange(linhaReal, 1, 1, 20);
              
              // Restaura cores padrão dos blocos (Baseado no 02_Guias da Operacional)
              aba.getRange(linhaReal, 1, 1, 4).setBackground("#FFFFFF");   // A-D (Geral) fica branco
              aba.getRange(linhaReal, 5, 1, 4).setBackground("#FFFFFF");   // E-H (SECOM) fica branco
              aba.getRange(linhaReal, 9, 1, 8).setBackground("#FFFFFF");   // I-P (Comprador) fica branco
              aba.getRange(linhaReal, 17, 1, 2).setBackground("#FFFFFF");  // Q-R (Finalização) fica branco
              aba.getRange(linhaReal, 19, 1, 1).setBackground("#FFFFFF");  // S (Prazo) fica branco
              aba.getRange(linhaReal, 20, 1, 1).setBackground("#EA9999");  // T (Prioridade) volta a ser vermelho fundo padrão (vazio)
              
              desmarcados++;
            }
          }
        });
      }
    }
  });
  
  SpreadsheetApp.getUi().alert(`Sincronização Concluída!\n\n🔥 Marcados: ${marcados}\n🧹 Limpos: ${desmarcados}`);
}