/**
 * Busca dados detalhados para o relatório analítico.
 */
function extrairDadosOperacionais() {
  const ssFonte = SpreadsheetApp.openById(ID_PLANILHA_FONTE);
  const todasAbas = ssFonte.getSheets();
  let dadosConsolidados = [];

  todasAbas.forEach(aba => {
    const nomeComprador = aba.getName();
    
    if (!ABAS_IGNORAR.includes(nomeComprador)) {
      const ultimaLinha = aba.getLastRow();
      const ultimaColuna = aba.getMaxColumns();
      
      if (ultimaLinha > 1 && ultimaColuna >= 19) {
        // Pega colunas A(1) até S(19)
        const valores = aba.getRange(2, 1, ultimaLinha - 1, 19).getValues();
        
        valores.forEach(linha => {
          if (linha[0]) {
            // Mapeamento de Colunas (Índices começam em 0)
            const processo = linha[0];      // Col A
            const objeto = linha[1];        // Col B
            const modalidade = linha[2];    // Col C
            const qtdItens = Number(linha[3]) || 0;
            
            // Datas Importantes
            const dataRecComprador = linha[8];  // Col I (Recebimento)
            const dataInicioPesq = linha[9];    // Col J (Início Pesquisa)
            const dataEnvioChefia = linha[16];  // Col Q (Envio Chefia)
            const dataEnvioCoage = linha[17];   // Col R (Fim/COAGE)
            
            // --- CÁLCULOS DE TEMPOS (ANALÍTICO) ---
            const agora = new Date();
            
            // 1. Tempo de Reação (Recebimento -> Início)
            let tempoReacao = 0;
            if (dataRecComprador instanceof Date && dataInicioPesq instanceof Date) {
               tempoReacao = diffDias(dataRecComprador, dataInicioPesq);
            } else if (dataRecComprador instanceof Date && !dataInicioPesq) {
               tempoReacao = diffDias(dataRecComprador, agora); // Está parado esperando início
            }

            // 2. Tempo de Execução (Início -> Chefia)
            let tempoExecucao = 0;
            if (dataInicioPesq instanceof Date && dataEnvioChefia instanceof Date) {
               tempoExecucao = diffDias(dataInicioPesq, dataEnvioChefia);
            } else if (dataInicioPesq instanceof Date && !dataEnvioChefia) {
               tempoExecucao = diffDias(dataInicioPesq, agora); // Está em execução
            }

            // 3. Tempo de Chefia (Chefia -> COAGE)
            let tempoChefia = 0;
            if (dataEnvioChefia instanceof Date && dataEnvioCoage instanceof Date) {
               tempoChefia = diffDias(dataEnvioChefia, dataEnvioCoage);
            } else if (dataEnvioChefia instanceof Date && !dataEnvioCoage) {
               tempoChefia = diffDias(dataEnvioChefia, agora); // Parado na chefia
            }

            // Total Dias Decorridos
            let diasTotais = 0;
            const dataFim = (dataEnvioCoage instanceof Date) ? dataEnvioCoage : agora;
            if (dataRecComprador instanceof Date) {
              diasTotais = diffDias(dataRecComprador, dataFim);
            }

            // Status do Prazo
            let statusPrazo = "No Prazo";
            if (diasTotais > 60) statusPrazo = "Atrasado";
            else if (diasTotais > 50 && !(dataEnvioCoage instanceof Date)) statusPrazo = "Alerta";
            if (dataEnvioCoage instanceof Date) statusPrazo = "Concluído";

            dadosConsolidados.push({
              comprador: nomeComprador,
              processo: processo,
              objeto: objeto,
              modalidade: modalidade,
              qtdItens: qtdItens,
              diasTotais: diasTotais,
              tempoReacao: tempoReacao,
              tempoExecucao: tempoExecucao,
              tempoChefia: tempoChefia,
              statusPrazo: statusPrazo
            });
          }
        });
      }
    }
  });
  
  // Ordena pelos mais demorados primeiro
  dadosConsolidados.sort((a, b) => b.diasTotais - a.diasTotais);
  
  return dadosConsolidados;
}

// Função auxiliar simples para diferença de dias
function diffDias(d1, d2) {
  const diff = Math.abs(d2 - d1);
  return Math.ceil(diff / (1000 * 60 * 60 * 24));
}