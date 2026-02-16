// Funções Relacionadas a Referências

function atualizarDataInicioClientes() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const referenciasSheet = ss.getSheetByName(SHEETS.REFERENCES);
  const analiseSheet = ss.getSheetByName(SHEETS.ANALYSIS);

  if (!referenciasSheet || !analiseSheet) {
    SpreadsheetApp.getUi().alert("Aba 'Referências' ou 'Análise' não foi encontrada.");
    return;
  }

  const referenciasData = referenciasSheet.getDataRange().getValues();
  const analiseData = analiseSheet.getDataRange().getValues();

  const referenciasHeader = referenciasData[0];
  const colClienteRef = referenciasHeader.indexOf("Clientes");

  let colDataInicioRef = referenciasHeader.indexOf("Data Início");
  if (colDataInicioRef === -1) {
    colDataInicioRef = referenciasSheet.getLastColumn();
    referenciasSheet.getRange(1, colDataInicioRef + 1).setValue("Data Início");
  }

  const analiseHeader = analiseData[0].map(h => (h || "").toString().trim().toUpperCase());
  const colClienteAnalise = analiseHeader.indexOf("CLIENTE");
  const colDataInicioAnalise = analiseHeader.indexOf("DATA INICIO");

  if (colClienteAnalise === -1 || colDataInicioAnalise === -1) {
    SpreadsheetApp.getUi().alert("Colunas 'Cliente' ou 'Data Início' não encontradas na aba 'Análise'.");
    return;
  }

  const mapaDatasInicio = {};
  analiseData.slice(1).forEach(row => {
    const cliente = normalizeText(row[colClienteAnalise]);
    const dataInicio = row[colDataInicioAnalise];
    if (cliente && dataInicio) mapaDatasInicio[cliente] = dataInicio;
  });

  referenciasData.slice(1).forEach((row, i) => {
    const cliente = normalizeText(row[colClienteRef]);
    const dataInicio = mapaDatasInicio[cliente] || "";
    referenciasSheet.getRange(i + 2, colDataInicioRef + 1).setValue(dataInicio);
  });
}

function atualizarEngajamentoPorFerramenta() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var referenciasSheet = ss.getSheetByName(SHEETS.REFERENCES);
  var analiseSheet = ss.getSheetByName(SHEETS.ANALYSIS);

  if (!referenciasSheet || !analiseSheet) {
    SpreadsheetApp.getUi().alert("Aba 'Referências' ou 'Análise' não foi encontrada.");
    return;
  }

  var referenciasData = referenciasSheet.getDataRange().getValues();
  var analiseData = analiseSheet.getDataRange().getValues();

  var headers = referenciasData[0];
  var colClientes = headers.indexOf("Clientes");

  function verificarOuCriarColuna(nomeColuna) {
    let colIndex = headers.indexOf(nomeColuna);
    if (colIndex === -1) {
      colIndex = referenciasSheet.getLastColumn();
      referenciasSheet.getRange(1, colIndex + 1).setValue(nomeColuna);
      headers.push(nomeColuna);
    }
    return colIndex;
  }

  var colAnalytics = verificarOuCriarColuna("Analytics");
  var colForcaDeVendas = verificarOuCriarColuna("Força de Vendas");
  var colCRM = verificarOuCriarColuna("CRM");
  var colPortalB2B = verificarOuCriarColuna("Portal B2B");
  var colEmAdocao = verificarOuCriarColuna("Em Adoção");

  var engajamentoPorCliente = {};

  var ferramentas = {
    "analytics": "Analytics",
    "forçadevendas": "Força de Vendas",
    "crm": "CRM",
    "portalb2b": "Portal B2B"
  };

  Object.keys(ferramentas).forEach(ferramenta => {
    var aba = ss.getSheetByName(ferramentas[ferramenta]);
    if (!aba) return;

    var dataFerramenta = aba.getDataRange().getValues();
    var headersFerramenta = dataFerramenta[0].map(h => (h || "").toString().trim().toLowerCase());
    var colCliente = headersFerramenta.indexOf("cliente");
    var colStatus = headersFerramenta.indexOf("status");
    if (colCliente === -1 || colStatus === -1) return;

    dataFerramenta.slice(1).forEach(row => {
      var cliente = normalizeText(row[colCliente]);
      var status = (row[colStatus] || "").toString().trim();

      if (!engajamentoPorCliente[cliente]) engajamentoPorCliente[cliente] = {};
      engajamentoPorCliente[cliente][ferramenta] = status || STATUS.NO_TOOL;
    });
  });

  var clientesEmAdocao = new Set(analiseData.slice(1).map(row => normalizeText(row[0])));

  referenciasData.slice(1).forEach((row, i) => {
    var cliente = normalizeText(row[colClientes]);
    var engajamento = engajamentoPorCliente[cliente] || {};

    referenciasSheet.getRange(i + 2, colAnalytics + 1).setValue(engajamento["analytics"] || STATUS.NO_TOOL);
    referenciasSheet.getRange(i + 2, colForcaDeVendas + 1).setValue(engajamento["forçadevendas"] || STATUS.NO_TOOL);
    referenciasSheet.getRange(i + 2, colCRM + 1).setValue(engajamento["crm"] || STATUS.NO_TOOL);
    referenciasSheet.getRange(i + 2, colPortalB2B + 1).setValue(engajamento["portalb2b"] || STATUS.NO_TOOL);

    referenciasSheet.getRange(i + 2, colEmAdocao + 1).setValue(clientesEmAdocao.has(cliente) ? "Em Adoção" : "Não Está em Adoção");
  });
}

function indicarReuniaoMaisAtual() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var backgroundSheet = ss.getSheetByName(SHEETS.BACKGROUND);
  var referenciasSheet = ss.getSheetByName(SHEETS.REFERENCES);

  if (!backgroundSheet || !referenciasSheet) {
    SpreadsheetApp.getUi().alert("As abas 'Background' ou 'Referências' não foram encontradas.");
    return;
  }

  var backgroundData = backgroundSheet.getDataRange().getValues();
  var referenciasData = referenciasSheet.getDataRange().getValues();

  if (backgroundData.length <= 1 || referenciasData.length <= 1) {
    SpreadsheetApp.getUi().alert("As abas 'Background' ou 'Referências' não possuem dados suficientes.");
    return;
  }

  var bgHeaders = backgroundData[0].map(h => (h || "").toString().trim().toLowerCase());
  var colEmpresaBG = bgHeaders.indexOf("nome da empresa");
  var colTipoReuniaoBG = bgHeaders.indexOf("tipo da reunião");

  var refHeaders = referenciasData[0].map(h => (h || "").toString().trim().toLowerCase());
  var colClientesRef = refHeaders.indexOf("clientes");
  var colReuniaoRef = 7; // Coluna H

  if (colEmpresaBG === -1 || colTipoReuniaoBG === -1 || colClientesRef === -1) {
    SpreadsheetApp.getUi().alert("Colunas necessárias não foram encontradas nas abas 'Background' ou 'Referências'.");
    return;
  }

  var reunioesMaisAtuais = {};
  backgroundData.slice(1).forEach(row => {
    var empresa = (row[colEmpresaBG] || "").toString().trim();
    var tipoReuniao = (row[colTipoReuniaoBG] || "").toString().trim();
    if (empresa && tipoReuniao) reunioesMaisAtuais[empresa] = tipoReuniao;
  });

  referenciasSheet.getRange(1, colReuniaoRef + 1).setValue("Tipo de Reunião").setFontWeight("bold");

  var clientesRefColumn = referenciasData.slice(1).map(row => (row[colClientesRef] || "").toString().trim());
  var reunioesRefColumn = clientesRefColumn.map(cliente => [reunioesMaisAtuais[cliente] || "Sem Reunião"]);

  referenciasSheet.getRange(2, colReuniaoRef + 1, reunioesRefColumn.length, 1).setValues(reunioesRefColumn);
}

function padronizarTipoReuniao() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var referenciasSheet = ss.getSheetByName(SHEETS.REFERENCES);

  if (!referenciasSheet) {
    SpreadsheetApp.getUi().alert("Aba 'Referências' não encontrada.");
    return;
  }

  var data = referenciasSheet.getDataRange().getValues();
  var headers = data[0].map(h => (h || "").toString().trim().toLowerCase());
  var colTipoReuniao = headers.indexOf("tipo de reunião");

  if (colTipoReuniao === -1) {
    SpreadsheetApp.getUi().alert("Coluna 'Tipo de Reunião' não encontrada.");
    return;
  }

  var conversaoReuniao = {
    "1º reunião": "1° Reunião", "1° reunião": "1° Reunião",
    "2º reunião": "2° Reunião", "2° reunião": "2° Reunião",
    "3º reunião": "3° Reunião", "3° reunião": "3° Reunião",
    "4º reunião": "4° Reunião", "4° reunião": "4° Reunião",
    "5º reunião": "5° Reunião", "5° reunião": "5° Reunião",
    "6º reunião": "6° Reunião", "6° reunião": "6° Reunião",
    "7º reunião": "7° Reunião", "7° reunião": "7° Reunião",
    "8º reunião": "8° Reunião", "8° reunião": "8° Reunião",
    "9º reunião": "9° Reunião", "9° reunião": "9° Reunião",
    "10º reunião": "10° Reunião", "10° reunião": "10° Reunião",
    "11º reunião": "11° Reunião", "11° reunião": "11° Reunião",
    "12º reunião": "12° Reunião", "12° reunião": "12° Reunião"
  };

  var updates = [];
  for (var i = 1; i < data.length; i++) {
    var tipoReuniao = data[i][colTipoReuniao];

    if (tipoReuniao) {
      var tipoCorrigido = tipoReuniao.toString().trim();
      tipoCorrigido = tipoCorrigido.replace(/º/g, "°");
      tipoCorrigido = conversaoReuniao[tipoCorrigido.toLowerCase()] || tipoCorrigido;
      updates.push([tipoCorrigido]);
    } else {
      updates.push([""]);
    }
  }

  referenciasSheet.getRange(2, colTipoReuniao + 1, updates.length, 1).setValues(updates);
}

function adicionarStatusReuniao() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var referenciasSheet = ss.getSheetByName(SHEETS.REFERENCES);

  if (!referenciasSheet) {
    SpreadsheetApp.getUi().alert("Aba 'Referências' não encontrada.");
    return;
  }

  var data = referenciasSheet.getDataRange().getValues();
  var headers = data[0].map(h => (h || "").toString().trim().toLowerCase());
  var colTipoReuniao = headers.indexOf("tipo de reunião");
  var colStatus = headers.indexOf("status");

  if (colStatus === -1) {
    colStatus = referenciasSheet.getLastColumn();
    referenciasSheet.getRange(1, colStatus + 1).setValue("Status");
  }

  var updates = [];
  for (var i = 1; i < data.length; i++) {
    var tipoReuniao = (data[i][colTipoReuniao] || "").toString().trim();
    var numeroReuniao = parseInt(tipoReuniao.replace(/[^0-9]/g, ""), 10);

    var status = "Sem Classificação";
    if (numeroReuniao === 1 || numeroReuniao === 2) status = "Iniciando";
    else if (numeroReuniao === 3) status = "Em Progresso";
    else if (numeroReuniao >= 4) status = "Finalizando";

    updates.push([status]);
  }

  referenciasSheet.getRange(2, colStatus + 1, updates.length, 1).setValues(updates);
}

function atualizarEtapa() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var cuboSheet = ss.getSheetByName(SHEETS.CUBO);
  var referenciasSheet = ss.getSheetByName(SHEETS.REFERENCES);

  if (!cuboSheet || !referenciasSheet) {
    Logger.log("❌ Uma ou ambas as abas não foram encontradas.");
    return;
  }

  var cuboData = cuboSheet.getDataRange().getValues();
  var cuboClientes = {}; // CNPJ -> Etapa

  for (var i = 1; i < cuboData.length; i++) {
    var etapa = cuboData[i][1];
    var cnpj = formatarCNPJ(cuboData[i][5]);

    if (cnpj && cnpj !== "00000000000000") {
      if (!(cnpj in cuboClientes) || cuboClientes[cnpj].etapa === "Encerrado") {
        cuboClientes[cnpj] = { etapa: etapa };
      }
    }
  }

  var referenciasData = referenciasSheet.getDataRange().getValues();
  var header = referenciasData[0];

  var etapaIndex = header.indexOf("Etapa");
  if (etapaIndex === -1) {
    referenciasSheet.getRange(1, header.length + 1).setValue("Etapa");
    etapaIndex = header.length;
  }

  var updates = [];
  for (var j = 1; j < referenciasData.length; j++) {
    var cnpjRef = formatarCNPJ(referenciasData[j][1]);
    var etapaVal = cuboClientes[cnpjRef] ? cuboClientes[cnpjRef].etapa : "Não Encontrado";
    updates.push([etapaVal]);
  }

  referenciasSheet.getRange(2, etapaIndex + 1, updates.length, 1).setValues(updates);
}

function atualizarConsultoras() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var cuboSheet = ss.getSheetByName(SHEETS.CUBO);
  var referenciasSheet = ss.getSheetByName(SHEETS.REFERENCES);

  if (!cuboSheet || !referenciasSheet) {
    Logger.log("❌ Uma ou ambas as abas não foram encontradas.");
    return;
  }

  var cuboData = cuboSheet.getDataRange().getValues();
  var cuboClientes = {}; // CNPJ -> consultora

  for (var i = 1; i < cuboData.length; i++) {
    var cnpj = formatarCNPJ(cuboData[i][5]);
    var consultora = formatarNomeConsultora(cuboData[i][2]);
    var etapa = cuboData[i][1];

    if (cnpj && cnpj !== "00000000000000") {
      if (!(cnpj in cuboClientes) || cuboClientes[cnpj].etapa === "Encerrado") {
        cuboClientes[cnpj] = { consultora: consultora, etapa: etapa };
      }
    }
  }

  var referenciasData = referenciasSheet.getDataRange().getValues();
  var header = referenciasData[0];

  var consultorasIndex = header.indexOf("Consultoras");
  if (consultorasIndex === -1) {
    referenciasSheet.getRange(1, header.length + 1).setValue("Consultoras");
    consultorasIndex = header.length;
  }

  var updates = [];
  for (var j = 1; j < referenciasData.length; j++) {
    var cnpjRef = formatarCNPJ(referenciasData[j][1]);
    var consultoraVal = cuboClientes[cnpjRef] ? cuboClientes[cnpjRef].consultora : "Não Encontrado";
    updates.push([consultoraVal]);
  }

  referenciasSheet.getRange(2, consultorasIndex + 1, updates.length, 1).setValues(updates);
}

function formatarNomeConsultora(nome) {
  if (!nome) return "Não Informado";
  nome = nome.toString().trim();
  if (nome.startsWith("CX - ")) return nome.substring(5).trim();
  return nome;
}

// Mantive seus “atalhos” também (compatibilidade)
function attrefe() {
  atualizarEngajamentoPorFerramenta();
  indicarReuniaoMaisAtual();
  adicionarStatusReuniao();
  atualizarEtapa();
  atualizarConsultoras();
  padronizarTipoReuniao();
}
