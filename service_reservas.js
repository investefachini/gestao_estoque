/**
 * ============================================================================
 * ARQUIVO: service_reservas.gs
 * RESPONSABILIDADE: Regras de negócio exclusivas do fluxo de Reservas 
 * (Solicitar, Cancelar, Renovar e Fila de Espera).
 * ============================================================================
 */

function solicitarReserva(p) {
  const lock = LockService.getScriptLock();
  try { lock.waitLock(10000); } catch (e) { return { success: false, message: "Sistema ocupado. Tente novamente." }; }

  try {
    const ss = SpreadsheetApp.openById(CONFIG.ID_PLANILHA);
    const shE = ss.getSheetByName(CONFIG.ABA_ESTOQUE);
    const tf = shE.createTextFinder(p.fz).matchEntireCell(true).findNext();
    if (!tf) throw new Error("Veículo não encontrado no estoque.");
    
    const linha = tf.getRow();
    const headers = shE.getRange(1, 1, 1, shE.getLastColumn()).getValues()[0];
    const dadosLinha = shE.getRange(linha, 1, 1, shE.getLastColumn()).getValues()[0];
    
    const infoEstoque = {
        familia: dadosLinha[headers.indexOf("FAMILIA")]||"-", modelo: dadosLinha[headers.indexOf("MODELO")]||"-",
        ano: dadosLinha[headers.indexOf("ANO/MOD")]||"-", cor: dadosLinha[headers.indexOf("COR")]||"-",
        up: dadosLinha[headers.indexOf("UP")]||"-", variante: dadosLinha[headers.indexOf("VARIANTE")]||"-",
        obs: dadosLinha[headers.indexOf("OBS.:")]||"-"
    };

    const map = getMapaStatusVendedor_(ss);
    const fzTratado = String(p.fz).trim().toUpperCase();
    const st = map[fzTratado] ? String(map[fzTratado].status).trim().toUpperCase() : "LIVRE";
    
    if (st === "VENDIDO") return { success: false, message: "❌ Este veículo já foi VENDIDO e não aceita reservas." };

    const acao = (st === "LIVRE" || st === "CANCELADO" || st === "EXPIRADO") ? "PENDENTE" : "FILA DE ESPERA";
    const shH = garantirHistorico_();
    
    const exp = new Date(); exp.setDate(exp.getDate() + 7);
    const dataExp = (acao === "PENDENTE") ? exp : ""; 
    let df = p.dataFat; if(df && df.includes('-')) { const x=df.split('-'); df = `${x[2]}/${x[1]}/${x[0]}`; }

    shH.appendRow([
        new Date(), acao, dataExp, 0, p.fz, p.revenda, p.pedido, p.responsavel, p.email, p.emailCC, p.emailCoord, 
        df, p.pagto, p.instituicao, p.valorNegociado, p.temEntrada?"SIM":"NÃO", p.valorEntrada, p.diasPrazo, 
        p.temSeminovo?"SIM":"NÃO", p.valorSeminovo, "", "", infoEstoque.modelo, p.cliente, p.credito ? "SIM" : "NÃO", p.temperatura
    ]);

    try { enviarEmailFilaOuReserva({...p, ...infoEstoque}, acao); } catch(e) { console.error("Erro email: " + e.message); }

    let msgRetorno = acao === "PENDENTE" ? "✅ Reserva realizada com SUCESSO!" : "⚠️ Veículo já reservado. Você entrou na FILA DE ESPERA.";
    return { success: true, message: msgRetorno };
  } catch (erro) { return { success: false, message: "Erro ao salvar: " + erro.message };
  } finally { lock.releaseLock(); }
}

function buscarDadosReservaAtiva(fz) {
  const ss = SpreadsheetApp.openById(CONFIG.ID_PLANILHA);
  const fzTratado = String(fz).trim().toUpperCase();
  const map = getMapaReservasAtivasDetalhado_(ss);
  const dados = map[fzTratado];

  if (!dados) return { success: false, message: "Não encontrei reserva ativa para este FZ." };
  if (dados.status !== "RESERVADO" && dados.status !== "PENDENTE") return { success: false, message: `Status atual (${dados.status}) não permite cancelamento.` };

  const shE = ss.getSheetByName(CONFIG.ABA_ESTOQUE);
  const tf = shE.createTextFinder(fzTratado).matchEntireCell(true).findNext();
  let infoExtra = { modelo: "N/D", familia: "N/D" };
  
  if(tf) {
     const linha = tf.getRow();
     const headers = shE.getRange(1, 1, 1, shE.getLastColumn()).getValues()[0];
     const dadosLinha = shE.getRange(linha, 1, 1, shE.getLastColumn()).getValues()[0];
     infoExtra = {
        familia: dadosLinha[headers.indexOf("FAMILIA")]||"-", modelo: dadosLinha[headers.indexOf("MODELO")]||"-",
        ano: dadosLinha[headers.indexOf("ANO/MOD")]||"-", cor: dadosLinha[headers.indexOf("COR")]||"-",
        up: dadosLinha[headers.indexOf("UP")]||"-", variante: dadosLinha[headers.indexOf("VARIANTE")]||"-",
        obs: dadosLinha[headers.indexOf("OBS.:")]||"-"
     };
  }

  return { success: true, dados: { fz: fzTratado, revenda: dados.revenda, vendedor: dados.vendedor, cliente: dados.cliente || "", valorNegociado: dados.valorVenda, ...infoExtra } };
}

function processarCancelamentoReserva(p) {
  const lock = LockService.getScriptLock();
  try { lock.waitLock(10000); } catch (e) { return { success: false, message: "Sistema ocupado." }; }
  try {
    const ss = SpreadsheetApp.openById(CONFIG.ID_PLANILHA); 
    const shPatios = ss.getSheetByName(CONFIG.ABA_PATIOS); 
    const dadosPatios = shPatios.getDataRange().getDisplayValues(); let senhaCorreta = null;
    
    for(let i=1; i<dadosPatios.length; i++) { 
        if(String(dadosPatios[i][0]).trim().toUpperCase() === String(p.revenda).trim().toUpperCase()) { 
            senhaCorreta = String(dadosPatios[i][2]).trim(); break; 
        } 
    }
    
    if (!senhaCorreta) throw new Error(`Senha não configurada para a revenda: ${p.revenda}.`);
    if (String(p.senha).trim() !== senhaCorreta) throw new Error("⛔ SENHA INCORRETA! Cancelamento não autorizado.");
    
    const shH = ss.getSheetByName(CONFIG.ABA_HISTORICO); 
    const dadosH = shH.getDataRange().getValues(); 
    const fzAlvo = String(p.fz).trim().toUpperCase(); 
    let linhaEncontrada = -1; let emailsOriginais = {};
    
    // Ignora "FILA DE ESPERA" e cancela apenas a reserva ativa
    for (let i = dadosH.length - 1; i >= 1; i--) {
        const fzLinha = String(dadosH[i][4]).trim().toUpperCase(); 
        const statusLinha = String(dadosH[i][1]).trim().toUpperCase();
        
        if (fzLinha === fzAlvo && (statusLinha === "RESERVADO" || statusLinha === "PENDENTE")) { 
            linhaEncontrada = i + 1; 
            emailsOriginais = { vendedor: dadosH[i][8], gerente: dadosH[i][9], coord: dadosH[i][10] }; break; 
        }
    }
    
    if (linhaEncontrada === -1) throw new Error("Não foi encontrada uma reserva ATIVA para este FZ para ser cancelada."); 
    
    shH.getRange(linhaEncontrada, 2).setValue("CANCELADO"); 
    shH.getRange(linhaEncontrada, 3).setValue(""); 
    shH.getRange(linhaEncontrada, 4).setValue(p.motivo);
    
    try { enviarEmailCancelamento({ ...p, emailsDestino: emailsOriginais }); } catch(e) {}
    return { success: true, message: "Reserva CANCELADA com sucesso (Histórico atualizado)." };
  } catch (erro) { return { success: false, message: erro.message }; } finally { lock.releaseLock(); }
}

function buscarDadosRenovacao(fz) {
  try {
    const ss = SpreadsheetApp.openById(CONFIG.ID_PLANILHA);
    const shH = ss.getSheetByName(CONFIG.ABA_HISTORICO);
    const dados = shH.getDataRange().getDisplayValues();
    const headers = dados[0];
    const fzAlvo = String(fz).trim().toUpperCase();
    
    const idx = {
        FZ: headers.indexOf("FZ"), STATUS: headers.indexOf("STATUS_ATUAL"), AVISOS: headers.indexOf("AVISOS"),
        REVENDA: headers.indexOf("REVENDA"), VEND: headers.indexOf("VENDEDOR_NOME"), EMAIL: headers.indexOf("VENDEDOR_EMAIL"),
        EMAIL_CC: headers.indexOf("GERENTE_EMAIL"), EMAIL_COORD: headers.indexOf("COORDENADOR_EMAIL"), PAGTO: headers.indexOf("FORMA_PAGTO"),
        INST: headers.indexOf("INSTITUICAO"), VALOR: headers.indexOf("VALOR_NEGOCIADO"), CLIENTE: headers.indexOf("CLIENTE"),
        CREDITO: headers.indexOf("CREDITO_APROVADO"), TEMP: headers.indexOf("TERMOMETRO") 
    };

    let ultimaLinhaFz = null; let lastStatusFound = null; 

    for (let i = dados.length - 1; i >= 1; i--) {
      const linhaFz = String(dados[i][idx.FZ]).trim().toUpperCase();
      if (linhaFz === fzAlvo) {
        const st = String(dados[i][idx.STATUS]).trim().toUpperCase();
        if (st !== "FILA DE ESPERA" && !lastStatusFound) lastStatusFound = st;
        if (st === "RESERVADO" || st === "PENDENTE") { ultimaLinhaFz = dados[i]; break; }
      }
    }

    if (!ultimaLinhaFz) {
        if (lastStatusFound === "CANCELADO") return { success: false, message: "FZ com reserva cancelada" };
        if (lastStatusFound === "EXPIRADO") return { success: false, message: "FZ com reserva expirada" };
        return { success: false, message: "FZ não tem reserva ativa" };
    }

    const avisos = String(ultimaLinhaFz[idx.AVISOS]).trim();
    if (avisos === "Reserva Renovada") return { success: false, message: "Período de renovação máximo atingido. Procure o administrador do portal." };

    const shE = ss.getSheetByName(CONFIG.ABA_ESTOQUE);
    const tf = shE.createTextFinder(fzAlvo).matchEntireCell(true).findNext();
    let modelo = "N/D", familia = "N/D";
    if(tf) {
        const rowE = tf.getRow();
        const headersE = shE.getRange(1, 1, 1, shE.getLastColumn()).getValues()[0];
        const valsE = shE.getRange(rowE, 1, 1, shE.getLastColumn()).getValues()[0];
        modelo = valsE[headersE.indexOf("MODELO")] || "N/D";
        familia = valsE[headersE.indexOf("FAMILIA")] || "N/D";
    }

    return { 
        success: true, 
        dados: {
            modelo: modelo, familia: familia, revenda: ultimaLinhaFz[idx.REVENDA], vendedor: ultimaLinhaFz[idx.VEND],
            email: ultimaLinhaFz[idx.EMAIL], emailCC: ultimaLinhaFz[idx.EMAIL_CC], emailCoord: ultimaLinhaFz[idx.EMAIL_COORD],
            pagto: ultimaLinhaFz[idx.PAGTO], instituicao: ultimaLinhaFz[idx.INST], valor: ultimaLinhaFz[idx.VALOR],
            cliente: idx.CLIENTE > -1 ? ultimaLinhaFz[idx.CLIENTE] : "", credito: ultimaLinhaFz[idx.CREDITO], temperatura: ultimaLinhaFz[idx.TEMP]
        }
    };
  } catch(e) { return { success: false, message: e.message }; }
}

function processarRenovacaoReserva(p) {
  const lock = LockService.getScriptLock();
  try { lock.waitLock(10000); } catch (e) { return { success: false, message: "Sistema ocupado." }; }

  try {
    const ss = SpreadsheetApp.openById(CONFIG.ID_PLANILHA);
    const shPatios = ss.getSheetByName(CONFIG.ABA_PATIOS);
    const dadosPatios = shPatios.getDataRange().getDisplayValues();
    let senhaCorreta = null;
    for(let i=1; i<dadosPatios.length; i++) {
        if(String(dadosPatios[i][0]).trim().toUpperCase() === String(p.revenda).trim().toUpperCase()) {
            senhaCorreta = String(dadosPatios[i][2]).trim(); break;
        }
    }
    if (!senhaCorreta) throw new Error("Senha não configurada para a revenda.");
    if (String(p.senha).trim() !== senhaCorreta) throw new Error("⛔ SENHA INCORRETA! Renovação não autorizada.");

    const shH = ss.getSheetByName(CONFIG.ABA_HISTORICO);
    const dadosH = shH.getDataRange().getValues();
    const headers = dadosH[0];
    const fzAlvo = String(p.fz).trim().toUpperCase();
    
    const idxFz = headers.indexOf("FZ");
    const idxStatus = headers.indexOf("STATUS_ATUAL");
    
    let linhaEncontrada = -1; let dataCriacaoReserva = null;

    for (let i = dadosH.length - 1; i >= 1; i--) {
        if (String(dadosH[i][idxFz]).trim().toUpperCase() === fzAlvo) {
            const st = String(dadosH[i][idxStatus]).trim().toUpperCase();
            if (st === "RESERVADO" || st === "PENDENTE") {
                linhaEncontrada = i + 1; dataCriacaoReserva = new Date(dadosH[i][0]); break;
            }
        }
    }

    if (linhaEncontrada === -1) throw new Error("Reserva não encontrada para edição.");
    
    const novaExpiracao = new Date(dataCriacaoReserva.getTime());
    novaExpiracao.setDate(novaExpiracao.getDate() + 14);
    const dataExpFormatada = Utilities.formatDate(novaExpiracao, CONFIG.FUSO_HORARIO, "dd/MM/yyyy HH:mm");
    
    const getCol = (name, fallback) => headers.indexOf(name) > -1 ? headers.indexOf(name) + 1 : fallback;

    shH.getRange(linhaEncontrada, getCol("STATUS_ATUAL", 2)).setValue("RESERVADO");
    shH.getRange(linhaEncontrada, getCol("DATA_EXPIRACAO", 3)).setValue(dataExpFormatada); 
    shH.getRange(linhaEncontrada, getCol("AVISOS", 4)).setValue("Reserva Renovada"); 
    shH.getRange(linhaEncontrada, getCol("FORMA_PAGTO", 13)).setValue(p.novoPagto); 
    shH.getRange(linhaEncontrada, getCol("INSTITUICAO", 14)).setValue(p.novaInst); 
    shH.getRange(linhaEncontrada, getCol("VALOR_NEGOCIADO", 15)).setValue(p.novoValor);
    
    const colCliente = getCol("CLIENTE", 24);
    if (colCliente > -1) shH.getRange(linhaEncontrada, colCliente).setValue(p.novoCliente);
    
    shH.getRange(linhaEncontrada, getCol("CREDITO_APROVADO", 25)).setValue(p.novoCredito);
    shH.getRange(linhaEncontrada, getCol("TERMOMETRO", 26)).setValue(p.novaTemp); 
    shH.getRange(linhaEncontrada, 28).setValue(p.motivo); // Motivo Col AB (28)

    SpreadsheetApp.flush();
    try { enviarEmailRenovacao(p, dataExpFormatada); } catch(e) {}

    return { success: true, message: "Reserva RENOVADA com sucesso até " + dataExpFormatada.split(" ")[0] };
  } catch (erro) { return { success: false, message: erro.message }; } finally { lock.releaseLock(); }
}