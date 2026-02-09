/**
 * Orquestrador principal.
 */
function construirDashboard() {
  const dados = extrairDadosOperacionais();
  const metricas = calcularMetricas(dados);
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  construirRelatorioAnalitico(ss, dados);
  construirPainelGrafico(ss, metricas);
}

/**
 * ABA 1: Relatório Analítico Detalhado
 */
function construirRelatorioAnalitico(ss, dados) {
  let aba = ss.getSheetByName(ABA_RELATORIO);
  if (!aba) aba = ss.insertSheet(ABA_RELATORIO);
  aba.clear();
  
  // Cabeçalhos
  const cabecalhos = [
    "Comprador", "Processo", "Modalidade", "Objeto", "Qtd Itens",
    "Dias Reação", "Dias Execução", "Dias Chefia", 
    "TOTAL DIAS", "Status", "Velocidade (Itens/Dia)"
  ];
  
  const rangeCab = aba.getRange(1, 1, 1, cabecalhos.length);
  rangeCab.setValues([cabecalhos])
    .setFontWeight("bold")
    .setBackground(CORES.CABECALHO_TABELA)
    .setFontColor(CORES.TEXTO_BRANCO)
    .setHorizontalAlignment("center");

  if (dados.length === 0) return;

  const matrizDados = dados.map(d => {
    let velocidade = 0;
    if (d.diasTotais > 0) velocidade = (d.qtdItens / d.diasTotais).toFixed(2);
    
    return [
      d.comprador, d.processo, d.modalidade, d.objeto, d.qtdItens,
      d.tempoReacao, d.tempoExecucao, d.tempoChefia, 
      d.diasTotais, d.statusPrazo, velocidade
    ];
  });

  const rangeDados = aba.getRange(2, 1, matrizDados.length, cabecalhos.length);
  rangeDados.setValues(matrizDados);
  rangeDados.setVerticalAlignment("middle");
  
  // Formatação Condicional
  const coresFundo = [];
  const coresFonte = [];
  
  for (let i = 0; i < dados.length; i++) {
    let rowColor = (i % 2 == 0) ? CORES.LINHA_PAR : CORES.LINHA_IMPAR;
    let statusColor = rowColor;
    let statusFont = "#000000";

    if (dados[i].statusPrazo === "Atrasado") {
      statusColor = CORES.NEGATIVO; statusFont = "#FFFFFF";
    } else if (dados[i].statusPrazo === "Alerta") {
      statusColor = CORES.ALERTA;
    } else if (dados[i].statusPrazo === "Concluído") {
      statusColor = "#CEEAD6";
    }

    let linhaCores = Array(cabecalhos.length).fill(rowColor);
    let linhaFontes = Array(cabecalhos.length).fill("#000000");

    linhaCores[9] = statusColor;
    linhaFontes[9] = statusFont;

    // Destaques Analíticos
    if (dados[i].tempoReacao > 5) linhaCores[5] = "#F4CCCC"; // Vermelho Suave
    if (dados[i].tempoChefia > 5) linhaCores[7] = "#FFF2CC"; // Amarelo Suave

    coresFundo.push(linhaCores);
    coresFonte.push(linhaFontes);
  }
  
  rangeDados.setBackgrounds(coresFundo);
  rangeDados.setFontColors(coresFonte);
  
  // Ajustes Finais
  aba.setFrozenRows(1);
  aba.setColumnWidth(2, 140); 
  aba.setColumnWidth(4, 250);
}

/**
 * ABA 2: Painel Gráfico com PODIUM
 */
function construirPainelGrafico(ss, metricas) {
  let aba = ss.getSheetByName(ABA_GRAFICOS);
  if (!aba) aba = ss.insertSheet(ABA_GRAFICOS);
  
  // Remove gráficos antigos
  const graficosAntigos = aba.getCharts();
  graficosAntigos.forEach(g => aba.removeChart(g));
  
  aba.clear(); // Limpa células e formatação
  
  // Garante colunas suficientes
  if (aba.getMaxColumns() < 40) aba.insertColumnsAfter(aba.getMaxColumns(), 40 - aba.getMaxColumns());
  
  aba.setHiddenGridlines(true);
  aba.setTabColor(CORES.AZUL_SECOM);

  // Título do Painel
  aba.getRange("B2").setValue("PAINEL GERENCIAL SECOM").setFontSize(24).setFontWeight("bold").setFontColor(CORES.CABECALHO_TABELA);
  aba.getRange("B3").setValue("Ranking baseado em EFICIÊNCIA (Itens processados por dia trabalhado)").setFontStyle("italic").setFontColor("#666666");

  // ==========================================
  // 1. CONSTRUÇÃO DO PODIUM (TABELA VISUAL)
  // ==========================================
  
  // Prepara dados do Podium (Ordenado por Eficiência)
  let podiumArr = [];
  for (let c in metricas.porComprador) {
    podiumArr.push({
      nome: c,
      eficiencia: metricas.porComprador[c].eficiencia, // Itens/Dia
      total: metricas.porComprador[c].somaItens
    });
  }
  // Ordena do maior para o menor
  podiumArr.sort((a, b) => b.eficiencia - a.eficiencia);
  
  // Desenha a Tabela na Coluna L (L5)
  const startRow = 5;
  const startCol = 12; // Col L
  
  // Título da Tabela
  aba.getRange(startRow - 1, startCol).setValue("🏆 RANKING DE PRODUTIVIDADE").setFontWeight("bold").setFontSize(14);
  
  // Cabeçalhos
  const headerPodium = [["#", "Comprador", "Velocidade", "Vol. Total"]];
  const headerRange = aba.getRange(startRow, startCol, 1, 4);
  headerRange.setValues(headerPodium)
    .setBackground(CORES.CABECALHO_TABELA)
    .setFontColor("white")
    .setFontWeight("bold")
    .setHorizontalAlignment("center");
    
  // Preenche dados
  if (podiumArr.length > 0) {
    const dadosPodium = podiumArr.map((p, index) => [
      (index + 1) + "º", 
      p.nome, 
      p.eficiencia + " itens/dia", 
      p.total + " itens"
    ]);
    
    const bodyRange = aba.getRange(startRow + 1, startCol, dadosPodium.length, 4);
    bodyRange.setValues(dadosPodium);
    bodyRange.setHorizontalAlignment("center");
    bodyRange.setBorder(true, true, true, true, true, true, "#E0E0E0", SpreadsheetApp.BorderStyle.SOLID);
    
    // Estilização do TOP 3 (Ouro, Prata, Bronze)
    for (let i = 0; i < dadosPodium.length; i++) {
      const rowRange = aba.getRange(startRow + 1 + i, startCol, 1, 4);
      
      if (i === 0) { // 1º Lugar - Ouro
        rowRange.setBackground("#FFF2CC").setFontWeight("bold"); // Fundo Amarelo Claro
        aba.getRange(startRow + 1 + i, startCol).setFontColor("#BF9000"); // Texto Dourado no nº
      } else if (i === 1) { // 2º Lugar - Prata
        rowRange.setBackground("#F3F3F3"); // Cinza Claro
        aba.getRange(startRow + 1 + i, startCol).setFontColor("#7F7F7F");
      } else if (i === 2) { // 3º Lugar - Bronze
        rowRange.setBackground("#FCE5CD"); // Laranja Claro
        aba.getRange(startRow + 1 + i, startCol).setFontColor("#B45F06");
      }
    }
  }
  
  // Ajuste de colunas do Podium
  aba.setColumnWidth(startCol, 40);   // #
  aba.setColumnWidth(startCol+1, 150); // Nome
  aba.setColumnWidth(startCol+2, 120); // Velocidade
  aba.setColumnWidth(startCol+3, 100); // Total

  // ==========================================
  // 2. DADOS AUXILIARES PARA GRÁFICOS (OCULTOS)
  // ==========================================
  
  // Dados Volume (Gráfico 1)
  let rankArr = [["Comprador", "Total Itens"]];
  for (let c in metricas.porComprador) {
    rankArr.push([c, metricas.porComprador[c].somaItens]); 
  }
  rankArr.sort((a, b) => b[1] - a[1]); 
  aba.getRange(1, 26, rankArr.length, 2).setValues(rankArr); 

  // Dados Prazos (Gráfico 2)
  let prazoArr = [["Status", "Qtd"]];
  prazoArr.push(["No Prazo", metricas.prazoGeral["No Prazo"]]);
  prazoArr.push(["Alerta", metricas.prazoGeral["Alerta"]]);
  prazoArr.push(["Atrasado", metricas.prazoGeral["Atrasado"]]);
  prazoArr.push(["Concluído", metricas.prazoGeral["Concluído"]]);
  aba.getRange(1, 29, 5, 2).setValues(prazoArr); 

  // Dados Dispersão (Gráfico 3)
  let dispArr = [["Itens", "Dias"]];
  if (metricas.dispersao.length > 0) {
    metricas.dispersao.forEach(d => dispArr.push(d));
    aba.getRange(1, 32, dispArr.length, 2).setValues(dispArr); 
  }

  // ==========================================
  // 3. CRIAÇÃO DOS GRÁFICOS
  // ==========================================
  
  // Gráfico 1: Volume Total
  const chart1 = aba.newChart().setChartType(Charts.ChartType.BAR)
    .addRange(aba.getRange(1, 26, rankArr.length, 2))
    .setPosition(5, 2, 0, 0) // B5
    .setOption('title', 'Volume Total de Trabalho (Itens)')
    .setOption('legend', {position: 'none'})
    .setOption('colors', [CORES.AZUL_SECOM])
    .setOption('width', 400).setOption('height', 300)
    .build();
  aba.insertChart(chart1);

  // Gráfico 2: Status Prazos
  const chart2 = aba.newChart().setChartType(Charts.ChartType.PIE)
    .addRange(aba.getRange(1, 29, 5, 2))
    .setPosition(5, 7, 0, 0) // G5 (Aprox)
    .setOption('title', 'Saúde dos Prazos')
    .setOption('colors', [CORES.POSITIVO, CORES.ALERTA, CORES.NEGATIVO, "#46BDC6"]) 
    .setOption('width', 400).setOption('height', 300)
    .build();
  aba.insertChart(chart2);

  // Gráfico 3: Dispersão (Abaixo)
  const chart3 = aba.newChart().setChartType(Charts.ChartType.SCATTER)
    .addRange(aba.getRange(1, 32, dispArr.length, 2))
    .setPosition(23, 2, 0, 0) // B23
    .setOption('title', 'Análise de Complexidade: Itens vs Dias')
    .setOption('hAxis', {title: 'Qtd Itens'})
    .setOption('vAxis', {title: 'Dias Gastos'})
    .setOption('legend', {position: 'none'})
    .setOption('width', 900).setOption('height', 350)
    .build();
  aba.insertChart(chart3);
  
  // Limpeza Visual
  aba.hideColumns(26, 15);
}