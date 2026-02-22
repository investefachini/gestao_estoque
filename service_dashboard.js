/**
 * ============================================================================
 * ARQUIVO: service_dashboard.gs
 * RESPONSABILIDADE: Consultar os repositórios e formatar os pacotes de dados
 * massivos que alimentam a tabela inicial e o Cockpit BI.
 * ============================================================================
 */

function getDadosEstoque() {
  const ss = SpreadsheetApp.openById(CONFIG.ID_PLANILHA);
  const sheet = ss.getSheetByName(CONFIG.ABA_ESTOQUE);
  const dados = sheet.getDataRange().getDisplayValues();
  
  if (dados.length < 2) return [];
  const headers = dados[0];
  const I = { 
      RESERVA: headers.indexOf("RESERVA"), PATIO: headers.indexOf("PATIO"), FAMILIA: headers.indexOf("FAMILIA"), 
      MODELO: headers.indexOf("MODELO"), UP: headers.indexOf("UP"), OBS: headers.indexOf("OBS.:"), 
      VARIANTE: headers.indexOf("VARIANTE"), ANO: headers.indexOf("ANO/MOD"), FZ: headers.indexOf("FZ"), 
      COR: headers.indexOf("COR"), TABELA: headers.indexOf("TABELA 26/26"), OPORTUNIDADE: headers.indexOf("OPORTUNIDADE"),
      CUSTO: headers.indexOf("CUSTO") > -1 ? headers.indexOf("CUSTO") : 11,
      DATA_COMPRA: headers.indexOf("DATA COMPRA") > -1 ? headers.indexOf("DATA COMPRA") : 2
  };
  
  const mapaReserva = getMapaStatusVendedor_(ss);
  const mapaTransf = getMapaTransferenciasAtivas_(ss);
  const mapaObs = getMapaObservacoes_(ss); 

  // O loop (map) começa aqui, e agora tudo vai funcionar:
  return dados.slice(1).filter(r => r[I.FZ] && String(r[I.FZ]).trim() !== "").map(linha => {
    const fz = String(linha[I.FZ]).trim().toUpperCase();
    let statusShow = "LIVRE";
    let isRenovada = false;
    let dataExp = "-", revendaRes = "-", vendRes = "-", tempRes = "-";
    
    const infoRes = mapaReserva[fz];
    if (infoRes && (infoRes.status === "RESERVADO" || infoRes.status === "PENDENTE")) { 
        statusShow = infoRes.status; 
        if (infoRes.avisos === "Reserva Renovada") isRenovada = true;
        
        dataExp = infoRes.expiracao; revendaRes = infoRes.revenda;
        vendRes = infoRes.vendedor; tempRes = infoRes.temperatura;
    } 
    else if (String(linha[I.RESERVA]).trim().toUpperCase() === "VENDIDO") { statusShow = "VENDIDO"; }
    
    // --- LÓGICA DO MODAL INFOS (Aba OBS) COLOCADA NO LUGAR CORRETO ---
    let obsFinal = linha[I.OBS] || ""; 
    let balao = ""; 

    if (mapaObs[fz]) {
        if (mapaObs[fz].observacao) obsFinal = mapaObs[fz].observacao; // Sobrescreve a OBS antiga
        
        balao = `<hr style="margin: 8px 0; border-color: inherit; opacity: 0.2;">
                 <b>NEG REVERSÃO:</b> ${mapaObs[fz].negReversao}<br>
                 <b>FUNDO ESTRELA:</b> ${mapaObs[fz].fundoEstrela}<br>
                 <b>COMISSÃO:</b> ${mapaObs[fz].comissao}<br>
                 <b>RETIRADA:</b> ${mapaObs[fz].retirada}<br>
                 <b>PROG:</b> ${mapaObs[fz].programacao}<br>
                 <b>BÔNUS:</b> ${mapaObs[fz].bonificacao}`;
    }

    // Retorna a linha completa para o front-end, incluindo a nova coluna 20 (balão)
    return [ 
        statusShow, linha[I.PATIO]||"", linha[I.FAMILIA]||"", linha[I.MODELO]||"", 
        linha[I.UP]||"", linha[I.VARIANTE]||"", linha[I.ANO]||"", linha[I.FZ]||"", 
        linha[I.COR]||"", linha[I.TABELA]||"", linha[I.OPORTUNIDADE]||"", obsFinal, 
        mapaTransf.has(fz) ? "SOLICITADA" : "",
        linha[I.CUSTO]||"", linha[I.DATA_COMPRA]||"", isRenovada,
        dataExp, revendaRes, vendRes, tempRes, balao 
    ];
  });
}

function getDadosDashboardCompleto() {
  try {
    const ss = SpreadsheetApp.openById(CONFIG.ID_PLANILHA);
    const shE = ss.getSheetByName(CONFIG.ABA_ESTOQUE); 
    const shH = ss.getSheetByName(CONFIG.ABA_HISTORICO); 
    const shT = ss.getSheetByName(CONFIG.ABA_TRANSF);
    const shVendas = ss.getSheetByName(CONFIG.ABA_VENDAS_NBS);
    
    const dadosEstoque = []; const dadosReservas = []; const dadosTransf = []; const dadosVendas = [];
    let headersReservas = []; 
    
    if (shE) {
      const dados = shE.getDataRange().getDisplayValues(); const h = dados[0]; const mapaReservas = getMapaStatusVendedor_(ss);
      for (let i = 1; i < dados.length; i++) {
        const fz = String(dados[i][h.indexOf("FZ")]).trim().toUpperCase(); if (!fz) continue; let status = "LIVRE"; if (mapaReservas[fz]) { status = mapaReservas[fz].status; }
        dadosEstoque.push({ fz: fz, status: status, patio: dados[i][h.indexOf("PATIO")] || "", familia: dados[i][h.indexOf("FAMILIA")] || "", modelo: dados[i][h.indexOf("MODELO")] || "", cor: dados[i][h.indexOf("COR")] || "", custo: dados[i][h.indexOf("TABELA 26/26")] || "", dataCompra: dados[i][h.indexOf("DATA COMPRA")] || "" });
      }
    }
    
    if (shH) { 
        const dados = shH.getDataRange().getDisplayValues(); 
        headersReservas = dados[0]; 
        for (let i = 1; i < dados.length; i++) { dadosReservas.push(dados[i]); } 
    }
    
    if (shT) { const dados = shT.getDataRange().getDisplayValues(); for (let i = 1; i < dados.length; i++) { dadosTransf.push({ data: dados[i][0] || "", fz: dados[i][1] || "", modelo: dados[i][2] || "", origem: dados[i][3] || "", destino: dados[i][4] || "", solicitante: dados[i][5] || "", direta: dados[i][7] || "" }); } }
    
    if (shVendas) {
        const dadosV = shVendas.getDataRange().getDisplayValues();
        const hV = dadosV[0];
        const idx = {
            revenda: hV.indexOf("EMPRESA_VEICULO"), dataFat: hV.indexOf("DATA VENDA CORRIJIDA"),
            mes: hV.indexOf("Mês Venda"), ano: hV.indexOf("Ano Venda"), fz: hV.indexOf("FZ"),
            valFat: hV.indexOf("VALOR VENDA2"), diasCV: hV.indexOf("DIAS ENTRE COMPRA E VENDA")
        };

        for (let i = 1; i < dadosV.length; i++) {
            const fzVenda = idx.fz > -1 ? String(dadosV[i][idx.fz]).trim().toUpperCase() : "";
            if (fzVenda) {
                dadosVendas.push({ 
                    revenda: idx.revenda > -1 ? dadosV[i][idx.revenda] : "", 
                    dataFaturamento: idx.dataFat > -1 ? dadosV[i][idx.dataFat] : "", 
                    mesVenda: idx.mes > -1 ? dadosV[i][idx.mes] : "", anoVenda: idx.ano > -1 ? dadosV[i][idx.ano] : "",
                    fz: fzVenda, valorFaturado: idx.valFat > -1 ? dadosV[i][idx.valFat] : "",
                    diasCompraVenda: idx.diasCV > -1 ? dadosV[i][idx.diasCV] : ""
                });
            }
        }
    }

    return JSON.stringify({ estoque: dadosEstoque, reservas: dadosReservas, headersReservas: headersReservas, transferencias: dadosTransf, vendasNBS: dadosVendas });
  } catch(e) { return JSON.stringify({ error: e.message }); }
}