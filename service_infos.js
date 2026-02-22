function buscarDadosInfos(fz) {
  try {
    const ss = SpreadsheetApp.openById(CONFIG.ID_PLANILHA);
    const mapa = getMapaObservacoes_(ss);
    const fzNorm = String(fz).trim().toUpperCase();
    if (mapa[fzNorm]) return { success: true, dados: mapa[fzNorm] };
    return { success: true, dados: null, message: "Veículo sem histórico. Pode preencher." };
  } catch (e) { return { success: false, message: e.message }; }
}

function salvarDadosInfos(p) {
  const lock = LockService.getScriptLock();
  try { lock.waitLock(10000); } catch (e) { return { success: false, message: "Sistema ocupado." }; }
  try {
    const ss = SpreadsheetApp.openById(CONFIG.ID_PLANILHA);
    const sh = garantirAbaObs_(ss);
    sh.appendRow([ String(p.fz).trim().toUpperCase(), new Date(), p.negReversao?"SIM":"NÃO", p.fundoEstrela?"SIM":"NÃO", p.comissao||"", p.retirada||"", p.programacao||"", p.bonificacao||"", p.observacao||"", Session.getActiveUser().getEmail() ]);
    return { success: true, message: "✅ Informações gravadas (Histórico atualizado)!" };
  } catch(e) { return { success: false, message: e.message }; } finally { lock.releaseLock(); }
}