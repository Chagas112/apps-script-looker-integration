master.pipeline.gs

/**
 * Ponto de entrada: roda tudo para alimentar o Looker.
 * (Menu: "Looker / Atualizar tudo")
 */
function attgeral() {
  importarEventosAgenda();
  attrefe();
  attlooker();
}

/** Atualiza tudo da aba Referências */
function attrefe() {
  atualizarDataInicioClientes();
  atualizarEngajamentoPorFerramenta();
  indicarReuniaoMaisAtual();
  padronizarTipoReuniao();
  adicionarStatusReuniao();
  atualizarEtapa();
  atualizarConsultoras();
}

/** Atualiza tudo das abas por ferramenta (Looker tabs) */
function attlooker() {
  criarAbasPorFerramenta();
  criarAbaAnalytics();
  criarAbaForcaDeVendas();
  criarAbaCRM();
  criarAbaPortalB2B();
}

/**
 * Cria um "botão" no topo (menu).
 * Não precisa inserir desenho/shape na planilha.
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("Looker")
    .addItem("Atualizar tudo", "attgeral")
    .addSeparator()
    .addItem("Atualizar Background (Agenda)", "importarEventosAgenda")
    .addItem("Atualizar Referências", "attrefe")
    .addItem("Atualizar Abas Looker", "attlooker")
    .addToUi();
}

