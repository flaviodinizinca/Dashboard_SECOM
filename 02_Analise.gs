/**
 * Agrega métricas e CALCULAS A EFICIÊNCIA (Itens/Dia).
 */
function calcularMetricas(dados) {
  const metricas = {
    porComprador: {},
    prazoGeral: {"No Prazo": 0, "Alerta": 0, "Atrasado": 0, "Concluído": 0},
    dispersao: [] 
  };

  dados.forEach(item => {
    // 1. Agrupamento Por Comprador
    if (!metricas.porComprador[item.comprador]) {
      metricas.porComprador[item.comprador] = { 
        qtdProcessos: 0, 
        somaItens: 0, 
        somaDias: 0  // Importante para a média
      };
    }
    
    // Acumula os valores
    metricas.porComprador[item.comprador].qtdProcessos++;
    metricas.porComprador[item.comprador].somaItens += item.qtdItens;
    metricas.porComprador[item.comprador].somaDias += (item.diasTotais || 0);
    
    // 2. Contagem do Status Geral
    if (metricas.prazoGeral.hasOwnProperty(item.statusPrazo)) {
      metricas.prazoGeral[item.statusPrazo]++;
    } else {
      metricas.prazoGeral["No Prazo"]++;
    }
    
    // 3. Dados para Dispersão
    if (item.qtdItens > 0 && item.diasTotais > 0) {
      metricas.dispersao.push([item.qtdItens, item.diasTotais]);
    }
  });

  // --- BLOCO QUE ESTAVA FALTANDO ---
  // Calcula as médias finais e a eficiência de cada um
  for (let c in metricas.porComprador) {
    const obj = metricas.porComprador[c];
    
    // Eficiência = Total de Itens / Total de Dias Gastos
    let eficiencia = 0;
    if (obj.somaDias > 0) {
      eficiencia = parseFloat((obj.somaItens / obj.somaDias).toFixed(2));
    }
    
    // Salva o cálculo dentro do objeto para o Builder usar
    obj.eficiencia = eficiencia;
  }

  return metricas;
}