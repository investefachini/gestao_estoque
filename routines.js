/**
 * ============================================================================
 * ARQUIVO: routines.gs
 * RESPONSABILIDADE: Robôs de automação (CRON Jobs) que rodam em segundo plano
 * para limpar reservas expiradas e promover a fila de espera.
 * ============================================================================
 */

/**
 * FUNÇÃO MESTRA: Esta é a função que o gatilho (Trigger) de tempo vai chamar.
 */
function executarRotinasDeFundo() {
  console.log("Iniciando rotinas de segundo plano...");
  try { verificarRotinaExpiracao(); } catch(e) { console.error("Erro Expiração: " + e.message); }
  try { processarFilaDeEspera(); } catch(e) { console.error("Erro Fila: " + e.message); }
  console.log("Rotinas finalizadas com sucesso.");
}

function verificarRotinaExpiracao() {
  const ss = SpreadsheetApp.openById(CONFIG.ID_PLANILHA);
  const shH = ss.getSheetByName(CONFIG.ABA_HISTORICO);
  const shE = ss.getSheetByName(CONFIG.ABA_ESTOQUE);
  
  if (!shH || !shE) return;

  const dados = shH.getDataRange().getValues();
  const agora = new Date();
  
  for (let i = 1; i < dados.length; i++) {
    const status = String(dados[i][1]).trim().toUpperCase();
    const dataExp = parseDataBR(dados[i][2]);
    
    if ((status === "RESERVADO" || status === "PENDENTE") && dataExp && dataExp < agora) {
      const linha = i + 1;
      const fz = dados[i][4];
      
      shH.getRange(linha, 2).setValue("EXPIRADO");
      shH.getRange(linha, 3).setValue("");
      const timestamp = Utilities.formatDate(new Date(), CONFIG.FUSO_HORARIO, "dd/MM/yyyy HH:mm");
      shH.getRange(linha, 4).setValue(`Expiração Automática em ${timestamp}`);

      let infoEstoque = { modelo: dados[i][22] || "N/D", familia: "-", cor: "-", up: "-", variante: "-", ano: "-" };
      const tf = shE.createTextFinder(fz).matchEntireCell(true).findNext();
      
      if (tf) {
          const rowE = tf.getRow();
          const headersE = shE.getRange(1, 1, 1, shE.getLastColumn()).getValues()[0];
          const valuesE = shE.getRange(rowE, 1, 1, shE.getLastColumn()).getValues()[0];
          infoEstoque = {
              familia: valuesE[headersE.indexOf("FAMILIA")] || "-",
              modelo: valuesE[headersE.indexOf("MODELO")] || infoEstoque.modelo,
              ano: valuesE[headersE.indexOf("ANO/MOD")] || "-",
              cor: valuesE[headersE.indexOf("COR")] || "-",
              up: valuesE[headersE.indexOf("UP")] || "-",
              variante: valuesE[headersE.indexOf("VARIANTE")] || "-",
              obs: valuesE[headersE.indexOf("OBS.:")] || "-"
          };
      }

      const p = {
          fz: fz, revenda: dados[i][5], pedido: dados[i][6], responsavel: dados[i][7],
          email: dados[i][8], emailCC: dados[i][9], emailCoord: dados[i][10],
          valorNegociado: dados[i][14], cliente: dados[i][23], motivo: "PRAZO DE VALIDADE ESGOTADO",
          ...infoEstoque
      };
      
      try { enviarEmailExpiracao_Premium(p); } catch(e) {}
    }
  }
}

function processarFilaDeEspera() {
  try {
    const ss = SpreadsheetApp.openById(CONFIG.ID_PLANILHA);
    const shH = ss.getSheetByName(CONFIG.ABA_HISTORICO);
    if (!shH) return;
    
    const dados = shH.getDataRange().getValues();
    const headers = dados[0];
    const idx = { fz: headers.indexOf("FZ"), status: headers.indexOf("STATUS_ATUAL") };
    const fzMap = {};
    
    for (let i = 1; i < dados.length; i++) {
      const fz = String(dados[i][idx.fz]).trim().toUpperCase();
      const status = String(dados[i][idx.status]).trim().toUpperCase();
      
      if (!fz) continue;
      
      if (!fzMap[fz]) fzMap[fz] = { temReservaAtiva: false, vendido: false, filaEspera: [] };
      
      if (status === "RESERVADO" || status === "PENDENTE") fzMap[fz].temReservaAtiva = true;
      else if (status === "VENDIDO" || status === "FATURADO") fzMap[fz].vendido = true;
      else if (status === "FILA DE ESPERA") fzMap[fz].filaEspera.push(i + 1);
    }
    
    for (const fz in fzMap) {
      const obj = fzMap[fz];
      if (!obj.temReservaAtiva && !obj.vendido && obj.filaEspera.length > 0) {
        processarFilaDeEsperaParaFZ(ss, fz);
      }
    }
  } catch(e) {}
}

function processarFilaDeEsperaParaFZ(ss, fzAlvo) {
  try {
    const shH = ss.getSheetByName(CONFIG.ABA_HISTORICO);
    if (!shH) return;
    
    const dados = shH.getDataRange().getValues();
    const headers = dados[0];
    const idx = { fz: headers.indexOf("FZ"), status: headers.indexOf("STATUS_ATUAL"), dataExp: headers.indexOf("DATA_EXPIRACAO"), avisos: headers.indexOf("AVISOS") };
    
    for (let i = 1; i < dados.length; i++) {
      const fz = String(dados[i][idx.fz]).trim().toUpperCase();
      const status = String(dados[i][idx.status]).trim().toUpperCase();
      
      if (fz === fzAlvo && status === "FILA DE ESPERA") {
        const novaExpiracao = new Date();
        novaExpiracao.setDate(novaExpiracao.getDate() + 7);
        const dataExpFormatada = Utilities.formatDate(novaExpiracao, CONFIG.FUSO_HORARIO, "dd/MM/yyyy HH:mm");
        
        shH.getRange(i + 1, idx.status + 1).setValue("RESERVADO");
        shH.getRange(i + 1, idx.dataExp + 1).setValue(dataExpFormatada);
        shH.getRange(i + 1, idx.avisos + 1).setValue("Promovido da Fila");
        
        try {
          const dadosReserva = {
            fz: fz, revenda: dados[i][headers.indexOf("REVENDA")] || "", responsavel: dados[i][headers.indexOf("VENDEDOR_NOME")] || "",
            email: dados[i][headers.indexOf("VENDEDOR_EMAIL")] || "", emailCC: dados[i][headers.indexOf("GERENTE_EMAIL")] || "",
            emailCoord: dados[i][headers.indexOf("COORDENADOR_EMAIL")] || "", valorNegociado: dados[i][headers.indexOf("VALOR_NEGOCIADO")] || "",
            pagto: dados[i][headers.indexOf("FORMA_PAGTO")] || "", instituicao: dados[i][headers.indexOf("INSTITUICAO")] || ""
          };
          
          const dadosEstoque = obterDadosVeiculo_(ss, fz);
          enviarEmailFilaOuReserva(Object.assign(dadosReserva, { modelo: dadosEstoque.modelo, familia: dadosEstoque.familia }), "PENDENTE");
        } catch(e) {}
        
        SpreadsheetApp.flush();
        return; // Só promove a primeira entrada da fila
      }
    }
  } catch(e) {}
}