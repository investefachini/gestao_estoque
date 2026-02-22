/**
 * ============================================================================
 * ARQUIVO: config.gs
 * RESPONSABILIDADE: Centralizar todas as variáveis globais, IDs de banco de 
 * dados, senhas (no futuro) e constantes do sistema.
 * ============================================================================
 */

const CONFIG = {
  // Configurações de Banco de Dados (IDs e Abas)
  ID_PLANILHA: "1m55U9vs4pxRZhQtOO4XpZuuAyEFqNVFMG1nuHEf7M3k", // ID Cópia Teste (Homologação)
  ABA_ESTOQUE: "ESTOQUE",
  ABA_HISTORICO: "HISTORICO_RESERVAS",
  ABA_LOGS: "LOG_ACESSOS",
  ABA_PATIOS: "RESPONSAVEIS_PATIO",
  ABA_TRANSF: "HISTORICO_TRANSFERENCIAS",
  ABA_VENDAS: "VENDAS",
  ABA_EMAILS: "LISTA_EMAILS",
  ABA_VENDAS_NBS: "VENDAS NBS", // <-- Adicionado aqui para centralizar
  
  // Configurações de Sistema e Integração
  NOME_PASTA_RAIZ: "ARQUIVOS_VENDAS_FZ",
  EMAIL_GESTOR: "fabio.fachini@denigris.com.br", 
  FUSO_HORARIO: "America/Sao_Paulo"
};