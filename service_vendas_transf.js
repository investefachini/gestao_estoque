/**
 * ============================================================================
 * ARQUIVO: service_vendas_transf.gs
 * RESPONSABILIDADE: Regras de negócio do fluxo de Fechamento de Vendas, 
 * Geração de Pastas no Drive e Transferência de Pátios.
 * ============================================================================
 */

function buscarDadosVeiculoVenda(fz) {
  try {
    const ss = SpreadsheetApp.openById(CONFIG.ID_PLANILHA);
    const mapa = getMapaReservasAtivasDetalhado_(ss);
    const fzUpper = String(fz).trim().toUpperCase();
    
    if (!mapa[fzUpper]) return { success: false, message: "Veículo não possui reserva ativa." };
    const reserva = mapa[fzUpper];
    
    return { success: true, valor: reserva.valorVenda, pagto: reserva.pagto, banco: reserva.banco, vendedor: reserva.vendedor, revenda: reserva.revenda };
  } catch (e) { return { success: false, message: "Erro: " + e.message }; }
}

function processarVenda(formData) {
  try {
    const ss = SpreadsheetApp.openById(CONFIG.ID_PLANILHA);
    const fz = String(formData.vFz).trim().toUpperCase();
    const mapa = getMapaReservasAtivasDetalhado_(ss);
    if (!mapa[fz]) return { success: false, message: "Veículo não possui reserva ativa." };
    
    atualizarStatusReserva_(ss, fz, "VENDIDO");
    removerTransferenciaPendente_(ss, fz);
    
    const pastaRaiz = garantirPastaRaiz_();
    const pastaMes = garantirPastaMes_(pastaRaiz);
    const pastaFZ = pastaMes.createFolder("FZ_" + fz);
    
    const arquivos = [];
    if (formData.vFilePedido) arquivos.push(salvarArquivo_(formData.vFilePedido, pastaFZ, "Pedido"));
    if (formData.vFileNF) arquivos.push(salvarArquivo_(formData.vFileNF, pastaFZ, "NF"));
    if (formData.vFileAuth) arquivos.push(salvarArquivo_(formData.vFileAuth, pastaFZ, "Autorizacao"));
    if (formData.vFileDoc) arquivos.push(salvarArquivo_(formData.vFileDoc, pastaFZ, "Documentos"));
    
    const shV = ss.getSheetByName(CONFIG.ABA_VENDAS) || ss.insertSheet(CONFIG.ABA_VENDAS);
    if (shV.getLastRow() === 0) shV.appendRow(["DATA_VENDA","FZ","CLIENTE","CNPJ_CPF","CELULAR","EMAIL","EMAIL_FISCAL","AGENTE_NOME","AGENTE_EMAIL","LINK_PASTA"]);
    
    shV.appendRow([
      new Date(), fz, formData.vClienteNome || "", formData.vCnpj || "", formData.vCelular || "",
      formData.vEmail || "", formData.vEmailFiscal || "", formData.vAgenteNome || "", formData.vAgenteEmail || "", pastaFZ.getUrl()
    ]);
    return { success: true, message: "Venda registrada com sucesso!" };
  } catch(e) { return { success: false, message: "Erro: " + e.message }; }
}

function buscarDadosVeiculoTransf(fz) {
  try {
    const ss = SpreadsheetApp.openById(CONFIG.ID_PLANILHA);
    const shE = ss.getSheetByName(CONFIG.ABA_ESTOQUE);
    const shH = ss.getSheetByName(CONFIG.ABA_HISTORICO);
    const fzNorm = String(fz).trim().toUpperCase();
    
    const dadosEstoque = shE.getDataRange().getDisplayValues();
    const headersEstoque = dadosEstoque[0];
    const idxEstoque = { FZ: headersEstoque.indexOf("FZ"), MODELO: headersEstoque.indexOf("MODELO"), PATIO: headersEstoque.indexOf("PATIO") };
    const linhaEstoque = dadosEstoque.find(r => String(r[idxEstoque.FZ]).trim().toUpperCase() === fzNorm);
    
    if (!linhaEstoque) return { success: false, message: "❌ Veículo não encontrado no estoque!" };
    
    let dadosReserva = null;
    if (shH) {
      const dadosHistorico = shH.getDataRange().getDisplayValues();
      const headersHistorico = dadosHistorico[0];
      const idxHistorico = { fz: headersHistorico.indexOf("FZ"), status: headersHistorico.indexOf("STATUS_ATUAL"), revenda: headersHistorico.indexOf("REVENDA"), vendedor: headersHistorico.indexOf("VENDEDOR_NOME") };
      
      for (let i = dadosHistorico.length - 1; i >= 1; i--) {
        const fzLinha = String(dadosHistorico[i][idxHistorico.fz]).trim().toUpperCase();
        const statusLinha = String(dadosHistorico[i][idxHistorico.status]).trim().toUpperCase();
        if (fzLinha === fzNorm) {
          if (statusLinha === "RESERVADO" || statusLinha === "PENDENTE") { dadosReserva = { revenda: dadosHistorico[i][idxHistorico.revenda] || "", vendedor: dadosHistorico[i][idxHistorico.vendedor] || "" }; break; }
        }
      }
    }
    
    return { success: true, modelo: linhaEstoque[idxEstoque.MODELO] || "", origem: linhaEstoque[idxEstoque.PATIO] || "", destino: dadosReserva ? dadosReserva.revenda : "", vendedor: dadosReserva ? dadosReserva.vendedor : "" };
  } catch(e) { return { success: false, message: `Erro ao buscar: ${e.message}` }; }
}

function solicitarTransferencia(params) {
  try {
    const ss = SpreadsheetApp.openById(CONFIG.ID_PLANILHA);
    const shT = ss.getSheetByName(CONFIG.ABA_TRANSF) || ss.insertSheet(CONFIG.ABA_TRANSF);
    if (shT.getLastRow() === 0) shT.appendRow(["DATA","FZ","MODELO","ORIGEM","DESTINO","SOLICITANTE","EMAIL","ENTREGA_DIRETA"]);
    
    shT.appendRow([ Utilities.formatDate(new Date(), CONFIG.FUSO_HORARIO, "dd/MM/yyyy HH:mm"), params.fz || "", params.modelo || "", params.origem || "", params.destino || "", params.nome || "", params.email || "", params.entregaDireta ? "SIM" : "NÃO" ]);
    
    const emailsOrigem = getEmailsPorPatio_(ss, params.origem);
    const emailsDestino = getEmailsPorPatio_(ss, params.destino);
    const todosEmails = [...new Set([...emailsOrigem, ...emailsDestino])];
    
    enviarEmailTransferencia(params, todosEmails.join(","));
    return { success: true, message: "Transferência solicitada com sucesso!" };
  } catch(e) { return { success: false, message: "Erro: " + e.message }; }
}