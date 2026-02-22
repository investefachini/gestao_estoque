/**
 * ============================================================================
 * SISTEMA CENTRAL DE ESTOQUE - BACKEND PRINCIPAL
 * Este arquivo funciona apenas como Roteador (Porta de Entrada).
 * A lógica de negócio está separada nos arquivos 'service_*.gs'
 * ============================================================================
 */

function doGet() {
  const emailLogado = Session.getActiveUser().getEmail();
  
  // Opcional: Registrar quem acessou
  registrarAcesso_(emailLogado);

  if (!emailLogado) return HtmlService.createHtmlOutput("<h3>Acesso Negado. Faça login no Google.</h3>");

  const tpl = HtmlService.createTemplateFromFile("index");
  
  // Ao invés do doGet() travar para buscar dados, ele apenas solicita ao service_dashboard
  tpl.dadosIniciais = JSON.stringify(getDadosEstoque());
  tpl.listaEmails = JSON.stringify(getListaEmails_());

  return tpl.evaluate()
    .setTitle("PORTAL GESTÃO DE ESTOQUE")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no');
}

function include(filename) { 
  return HtmlService.createHtmlOutputFromFile(filename).getContent(); 
}