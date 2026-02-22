/**
 * ============================================================================
 * ARQUIVO: repository.gs
 * RESPONSABILIDADE: Centralizar todas as leituras e gravações no Google Sheets
 * e no Google Drive. Nenhuma regra de negócio (IFs complexos de venda) fica aqui.
 * ============================================================================
 */

function registrarAcesso_(email) { 
  try { 
    const ss = SpreadsheetApp.openById(CONFIG.ID_PLANILHA);
    let sh = ss.getSheetByName(CONFIG.ABA_LOGS); 
    if (!sh) { 
      sh = ss.insertSheet(CONFIG.ABA_LOGS); 
      sh.appendRow(["DATA","EMAIL","TIPO"]);
    } 
    sh.appendRow([new Date(), email, "ACESSO"]); 
  } catch(e){} 
}

function garantirHistorico_() { 
  const ss = SpreadsheetApp.openById(CONFIG.ID_PLANILHA); 
  let sh = ss.getSheetByName(CONFIG.ABA_HISTORICO); 
  if (!sh) { 
    sh = ss.insertSheet(CONFIG.ABA_HISTORICO); 
    sh.appendRow([
      "DATA","STATUS_ATUAL","DATA_EXPIRACAO","AVISOS","FZ","REVENDA","NUM_PEDIDO",
      "VENDEDOR_NOME","VENDEDOR_EMAIL","GERENTE_EMAIL","COORDENADOR_EMAIL","DATA_PREV_FAT",
      "FORMA_PAGTO","INSTITUICAO","VALOR_NEGOCIADO","TEM_ENTRADA","VALOR_ENTRADA",
      "DIAS_PRAZO_DENIGRIS","TEM_SEMINOVO","VALOR_CAP_SEMINOVO", "FATURADO?", 
      "DIAS NA RESERVA", "MODELO", "CLIENTE", "CREDITO_APROVADO", "TEMPERATURA"
    ]);
  } 
  return sh; 
}

function getMapaStatusVendedor_(ss) { 
  const shH = ss.getSheetByName(CONFIG.ABA_HISTORICO); 
  const map = {};
  if (!shH) return map; 
  const d = shH.getDataRange().getDisplayValues(); 
  const h = d[0];
  const idx = { 
      FZ: h.indexOf("FZ"), ST: h.indexOf("STATUS_ATUAL"), VEND: h.indexOf("VENDEDOR_NOME"), 
      AVISOS: h.indexOf("AVISOS"), EXPIRA: h.indexOf("DATA_EXPIRACAO"), REV: h.indexOf("REVENDA"), 
      TEMP: h.indexOf("TERMOMETRO") 
  };
  for (let i = d.length - 1; i >= 1; i--) { 
    const fz = String(d[i][idx.FZ]).trim().toUpperCase();
    const st = String(d[i][idx.ST]).trim().toUpperCase(); 
    if (st === "FILA DE ESPERA") continue;
    if(fz && !map[fz]) {
        map[fz] = { 
            status: st, 
            vendedor: idx.VEND > -1 ? d[i][idx.VEND] : "-", 
            avisos: idx.AVISOS > -1 ? String(d[i][idx.AVISOS]).trim() : "",
            expiracao: idx.EXPIRA > -1 ? d[i][idx.EXPIRA] : "-",
            revenda: idx.REV > -1 ? d[i][idx.REV] : "-",
            temperatura: idx.TEMP > -1 ? d[i][idx.TEMP] : "-"
        }; 
    }
  } 
  return map;
}

function getMapaReservasAtivasDetalhado_(ss) { 
  const shH = ss.getSheetByName(CONFIG.ABA_HISTORICO); 
  const map = {}; 
  if (!shH) return map;
  const d = shH.getDataRange().getDisplayValues(); 
  const h = d[0]; 
  const idx = { 
    fz: h.indexOf("FZ"), status: h.indexOf("STATUS_ATUAL"), valorNeg: h.indexOf("VALOR_NEGOCIADO"), 
    pagto: h.indexOf("FORMA_PAGTO"), banco: h.indexOf("INSTITUICAO"), revenda: h.indexOf("REVENDA"), 
    vendedor: h.indexOf("VENDEDOR_NOME"), cliente: h.indexOf("CLIENTE"), entrada: h.indexOf("VALOR_ENTRADA"), 
    temEntrada: h.indexOf("TEM_ENTRADA"), prazo: h.indexOf("DIAS_PRAZO_DENIGRIS"), seminovo: h.indexOf("VALOR_CAP_SEMINOVO"), 
    temTroca: h.indexOf("TEM_SEMINOVO"), dataPrev: h.indexOf("DATA_PREV_FAT"), pedido: h.indexOf("NUM_PEDIDO") 
  }; 
  
  for (let i = d.length - 1; i >= 1; i--) { 
    const fz = String(d[i][idx.fz]).trim().toUpperCase();
    const st = String(d[i][idx.status]).trim().toUpperCase(); 
    if (st === "FILA DE ESPERA" || !fz || map[fz]) continue; 
    
    if (st === "RESERVADO" || st === "PENDENTE") {
      map[fz] = { 
        fz: fz, status: st, valorVenda: d[i][idx.valorNeg], pagto: d[i][idx.pagto], banco: d[i][idx.banco], 
        revenda: d[i][idx.revenda], vendedor: d[i][idx.vendedor], cliente: idx.cliente > -1 ? d[i][idx.cliente] : "",
        entrada: d[i][idx.entrada], temEntrada: d[i][idx.temEntrada], prazo: d[i][idx.prazo], seminovo: d[i][idx.seminovo], 
        temTroca: d[i][idx.temTroca], dataPrev: d[i][idx.dataPrev], pedido: d[i][idx.pedido] 
      };
    }
  } 
  return map; 
}

function getMapaTransferenciasAtivas_(ss) { 
  const sh = ss.getSheetByName(CONFIG.ABA_TRANSF);
  const set = new Set(); 
  if (!sh) return set; 
  const d = sh.getDataRange().getDisplayValues();
  for (let i = 1; i < d.length; i++) { 
    const fz = String(d[i][1]).trim().toUpperCase(); 
    if(fz) set.add(fz);
  } 
  return set; 
}

function getEmailsPorPatio_(ss, nomePatio) { 
  const sh = ss.getSheetByName(CONFIG.ABA_PATIOS); 
  if (!sh) return [];
  const dados = sh.getDataRange().getDisplayValues(); 
  const patioNorm = String(nomePatio).trim().toUpperCase(); 
  for (let i = 1; i < dados.length; i++) { 
    if (String(dados[i][0]).trim().toUpperCase() === patioNorm) { 
      return dados[i][1].split(",").map(e => e.trim());
    } 
  } 
  return []; 
}

function getListaEmails_() {
  try {
    const ss = SpreadsheetApp.openById(CONFIG.ID_PLANILHA);
    const sh = ss.getSheetByName(CONFIG.ABA_EMAILS);
    if (!sh) return [];
    
    const data = sh.getDataRange().getValues();
    const lista = [];
    for (let i = 1; i < data.length; i++) {
      if (data[i][0]) {
        lista.push({
          nome: String(data[i][0]).trim(),
          email: String(data[i][1] || "").trim(),
          gerente: String(data[i][2] || "").trim(),
          coord: String(data[i][3] || "").trim()
        });
      }
    }
    return lista;
  } catch (e) { return []; }
}

function obterDadosVeiculo_(ss, fz) {
  try {
    const sheet = ss.getSheetByName(CONFIG.ABA_ESTOQUE);
    const dados = sheet.getDataRange().getDisplayValues();
    const headers = dados[0];
    const idx = { fz: headers.indexOf("FZ"), modelo: headers.indexOf("MODELO"), familia: headers.indexOf("FAMILIA") };
    
    for (let i = 1; i < dados.length; i++) {
      if (String(dados[i][idx.fz]).trim().toUpperCase() === fz) {
        return { modelo: dados[i][idx.modelo] || "N/I", familia: dados[i][idx.familia] || "N/I" };
      }
    }
    return { modelo: "Não encontrado", familia: "Não encontrada" };
  } catch(e) { return { modelo: "Erro", familia: "Erro" }; }
}

function atualizarStatusReserva_(ss, fz, novoStatus) {
  try {
    const shH = ss.getSheetByName(CONFIG.ABA_HISTORICO);
    if (!shH) return;
    const dados = shH.getDataRange().getValues();
    const headers = dados[0];
    const idx = { fz: headers.indexOf("FZ"), status: headers.indexOf("STATUS_ATUAL"), avisos: headers.indexOf("AVISOS") };
    
    let linhaPrincipal = null;
    for (let i = dados.length - 1; i >= 1; i--) {
      const fzLinha = String(dados[i][idx.fz]).trim().toUpperCase();
      const statusLinha = String(dados[i][idx.status]).trim().toUpperCase();
      if (fzLinha === fz && (statusLinha === "RESERVADO" || statusLinha === "PENDENTE")) {
        linhaPrincipal = i + 1; break;
      }
    }
    if (linhaPrincipal) {
      shH.getRange(linhaPrincipal, idx.status + 1).setValue(novoStatus);
      shH.getRange(linhaPrincipal, idx.avisos + 1).setValue("Status Atualizado");
      SpreadsheetApp.flush();
    }
  } catch(e) { console.error("Erro ao atualizar status:", e); }
}

function removerTransferenciaPendente_(ss, fz) { 
  const sh = ss.getSheetByName(CONFIG.ABA_TRANSF);
  if(!sh) return; 
  const dados = sh.getDataRange().getValues(); 
  for (let i = dados.length - 1; i >= 1; i--) { 
    if (String(dados[i][1]).trim().toUpperCase() === fz) { 
      sh.deleteRow(i + 1);
    } 
  } 
}

// =============================================================
// REPOSITÓRIO DA ABA "OBS" E FUNÇÕES DO MODAL INFOS
// =============================================================
function garantirAbaObs_(ss) {
  let sh = ss.getSheetByName("OBS");
  if (!sh) {
    sh = ss.insertSheet("OBS");
    sh.appendRow(["FZ", "DATA_HORA", "NEG_REVERSAO", "FUNDO_ESTRELA", "COMISSAO", "RETIRADA", "PROGRAMACAO", "BONIFICACAO", "OBSERVACAO", "RESPONSAVEL"]);
  }
  return sh;
}

function getMapaObservacoes_(ss) {
  const sh = ss.getSheetByName("OBS");
  const map = {};
  if (!sh) return map;
  const dados = sh.getDataRange().getDisplayValues();
  for (let i = 1; i < dados.length; i++) {
    const fz = String(dados[i][0]).trim().toUpperCase();
    if (fz) {
      map[fz] = {
        negReversao: dados[i][2], fundoEstrela: dados[i][3], comissao: dados[i][4],
        retirada: dados[i][5], programacao: dados[i][6], bonificacao: dados[i][7], observacao: dados[i][8]
      };
    }
  }
  return map;
}

// -------------------------------------------------------------
// OPERAÇÕES DO GOOGLE DRIVE (Gerenciamento de Pastas)
// -------------------------------------------------------------
function garantirPastaRaiz_() { 
  const i = DriveApp.getFoldersByName(CONFIG.NOME_PASTA_RAIZ); 
  if(i.hasNext()) return i.next(); 
  return DriveApp.createFolder(CONFIG.NOME_PASTA_RAIZ);
}

function garantirPastaMes_(pastaRaiz) { 
  const mesAno = Utilities.formatDate(new Date(), CONFIG.FUSO_HORARIO, "yyyy-MM"); 
  const it = pastaRaiz.getFoldersByName(mesAno); 
  if (it.hasNext()) return it.next();
  return pastaRaiz.createFolder(mesAno); 
}

function salvarArquivo_(blob, pasta, prefixo) { 
  if (!blob || blob.length === 0) return "-";
  const nomeArq = prefixo + "_" + blob.getName(); 
  return pasta.createFile(blob).setName(nomeArq).getUrl(); 
}

// -------------------------------------------------------------
// UTILITÁRIOS GERAIS
// -------------------------------------------------------------
function parseDataBR(dataStr) { 
  if (!dataStr) return null;
  if (dataStr instanceof Date) return dataStr; 
  try { 
    const parts = String(dataStr).split(' ');
    const dateParts = parts[0].split('/'); 
    if (dateParts.length < 3) return null; 
    let hora=0, min=0;
    if(parts.length>1){ const t=parts[1].split(':'); hora=parseInt(t[0]); min=parseInt(t[1]); } 
    let ano = parseInt(dateParts[2]);
    if(ano<100) ano+=2000; 
    return new Date(ano, parseInt(dateParts[1])-1, parseInt(dateParts[0]), hora, min); 
  } catch(e) { return null; } 
}

// =============================================================
// REPOSITÓRIO DA ABA "OBS" (INFORMAÇÕES ADICIONAIS)
// =============================================================

function garantirAbaObs_(ss) {
  let sh = ss.getSheetByName("OBS");
  if (!sh) {
    sh = ss.insertSheet("OBS");
    sh.appendRow(["FZ", "DATA_HORA", "NEG_REVERSAO", "FUNDO_ESTRELA", "COMISSAO", "RETIRADA", "PROGRAMACAO", "BONIFICACAO", "OBSERVACAO", "RESPONSAVEL"]);
  }
  return sh;
}

function getMapaObservacoes_(ss) {
  const sh = ss.getSheetByName("OBS");
  const map = {};
  if (!sh) return map;
  
  const dados = sh.getDataRange().getDisplayValues();
  // Lê de cima para baixo. Se houver FZ duplicado, a última linha lida (a mais recente) é a que fica no mapa.
  for (let i = 1; i < dados.length; i++) {
    const fz = String(dados[i][0]).trim().toUpperCase();
    if (fz) {
      map[fz] = {
        negReversao: dados[i][2],
        fundoEstrela: dados[i][3],
        comissao: dados[i][4],
        retirada: dados[i][5],
        programacao: dados[i][6],
        bonificacao: dados[i][7],
        observacao: dados[i][8]
      };
    }
  }
  return map;
}