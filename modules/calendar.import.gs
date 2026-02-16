// Funções Relacionadas ao Background (Calendar → Background)

function importarEventosAgenda() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var configuracoesSheet = spreadsheet.getSheetByName(SHEETS.CONFIG);
  var backgroundSheet = spreadsheet.getSheetByName(SHEETS.BACKGROUND);
  var referenciasSheet = spreadsheet.getSheetByName(SHEETS.REFERENCES);

  // Cria a aba "Background" se ela não existir
  if (!backgroundSheet) {
    backgroundSheet = spreadsheet.insertSheet(SHEETS.BACKGROUND);
    backgroundSheet.getRange(1, 1).setValue("A aba 'Background' foi criada automaticamente. Preencha os dados necessários.");
    SpreadsheetApp.getUi().alert("A aba 'Background' não foi encontrada e foi criada automaticamente.");
    return;
  }

  // Carrega a lista de clientes padronizados da aba "Referências"
  var referenciasData = referenciasSheet.getRange("A:A").getValues().flat().filter(String);
  var clientesOficiais = referenciasData.map(nome => normalizeText(nome));

  // ✅ EXEMPLO GENÉRICO (portfólio-safe)
  // Troque por sua “biblioteca” real quando quiser.
  var mapeamentoNomes = {
    "EMPRESA A": "EMPRESA A OFICIAL LTDA",
    "EMPRESA B": "EMPRESA B OFICIAL SA",
    "NOME COM ERRO": "NOME CORRIGIDO",
    "SIGLA X": "SIGLA X - NOME COMPLETO"
  };

  // Obtém dados da configuração
  var idAgenda = configuracoesSheet.getRange("B1").getValue();
  var dataInicial = configuracoesSheet.getRange("B2").getValue();
  var dataFinal = configuracoesSheet.getRange("B3").getValue();

  if (!idAgenda || !dataInicial || !dataFinal) {
    SpreadsheetApp.getUi().alert("Por favor, preencha todos os campos: ID da Agenda, Data Inicial e Data Final.");
    return;
  }

  try {
    var eventos = CalendarApp.getCalendarById(idAgenda).getEvents(new Date(dataInicial), new Date(dataFinal));

    backgroundSheet.clearContents();
    backgroundSheet.getRange(1, 1, 1, 10).setValues([[
      "Setor", "Nome da Empresa", "Organizador", "Co-participantes", "Tipo da Reunião",
      "Data", "Hora Início", "Hora Término", "Duração (minutos)", "Presença"
    ]]);

    var dados = [];
    for (var i = 0; i < eventos.length; i++) {
      var titulo = eventos[i].getTitle();
      var dataInicio = eventos[i].getStartTime();
      var dataFim = eventos[i].getEndTime();

      if (!dataInicio || !dataFim) continue;

      var duracaoMinutos = Math.round((dataFim - dataInicio) / (1000 * 60));
      var infoSeparada = separarTitulo(titulo);
      var setor = infoSeparada.setor;

      // Padroniza o nome da empresa usando o mapeamento e as referências
      var nomeEmpresa = padronizarNomeCliente(infoSeparada.nomeEmpresa, clientesOficiais, mapeamentoNomes);
      var tipoReuniao = infoSeparada.tipoReuniao;

      var organizadorEmail = (eventos[i].getCreators() || []).join(", ");
      var organizador = extrairNomeDoEmail(organizadorEmail);
      var coParticipantes = extrairCoParticipantes(eventos[i]);

      // Determina o status com base no título
      var status = determinarPresenca(
        titulo,
        Utilities.formatDate(dataInicio, Session.getScriptTimeZone(), "yyyy-MM-dd"),
        Utilities.formatDate(dataInicio, Session.getScriptTimeZone(), "HH:mm")
      );

      dados.push([
        setor,
        nomeEmpresa,
        organizador,
        coParticipantes,
        tipoReuniao,
        Utilities.formatDate(dataInicio, Session.getScriptTimeZone(), "dd/MM/yyyy"),
        Utilities.formatDate(dataInicio, Session.getScriptTimeZone(), "HH:mm"),
        Utilities.formatDate(dataFim, Session.getScriptTimeZone(), "HH:mm"),
        duracaoMinutos,
        status
      ]);
    }

    if (dados.length > 0) {
      backgroundSheet.getRange(2, 1, dados.length, dados[0].length).setValues(dados);
    } else {
      SpreadsheetApp.getUi().alert("Nenhum evento encontrado no período informado.");
    }
  } catch (e) {
    SpreadsheetApp.getUi().alert("Erro ao buscar eventos: " + e.message);
  }
}

function separarTitulo(titulo) {
  var padrao = /(.+?)\s*-\s*(.+?)\s*-\s*(.+?)(?:\s*-\s*.*)?$/;
  var resultado = titulo.match(padrao);

  if (resultado) {
    return {
      setor: resultado[1].trim(),
      nomeEmpresa: resultado[2].trim(),
      tipoReuniao: resultado[3].trim()
    };
  }
  return { setor: "Erro", nomeEmpresa: "Erro", tipoReuniao: "Erro" };
}

function padronizarNomeCliente(nomeEmpresa, clientesOficiais, mapeamentoNomes) {
  var nomeNormalizado = normalizeText(nomeEmpresa);

  if (mapeamentoNomes[nomeNormalizado]) {
    return mapeamentoNomes[nomeNormalizado];
  }

  var clientePadronizado = clientesOficiais.find(cliente =>
    nomeNormalizado.includes(cliente) || cliente.includes(nomeNormalizado)
  );

  return clientePadronizado || (nomeEmpresa || "").toString().trim();
}

function extrairNomeDoEmail(email) {
  if (!email) return "";
  var nomeParte = email.split("@")[0];
  var nomeSemPonto = nomeParte.split(".")[0];
  return nomeSemPonto.charAt(0).toUpperCase() + nomeSemPonto.slice(1).toLowerCase();
}

function extrairCoParticipantes(evento) {
  var coParticipantes = [];
  var participantes = evento.getGuestList();

  for (var i = 0; i < participantes.length; i++) {
    var email = participantes[i].getEmail();
    var status = participantes[i].getGuestStatus();

    // ✅ domínio ajustado
    if (email.endsWith("@jovendas.com") && status == CalendarApp.GuestStatus.YES) {
      coParticipantes.push(extrairNomeDoEmail(email));
    }
  }
  return coParticipantes.join(", ");
}

function determinarPresenca(titulo, data, horaInicio) {
  var hoje = new Date();
  var dataHoraReuniao = new Date(data + " " + horaInicio);

  if (titulo.includes("NoShow")) return "Não Compareceu";
  if (dataHoraReuniao > hoje) return "Marcada";
  return "Realizada";
}
