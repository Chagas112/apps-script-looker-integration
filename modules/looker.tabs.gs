// Funções Relacionadas a Abas por Ferramenta (Looker / BI)

function criarAbasPorFerramenta() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var analiseSheet = ss.getSheetByName(SHEETS.ANALYSIS);
  var pesoSheet = ss.getSheetByName(SHEETS.WEIGHT);

  if (!analiseSheet || !pesoSheet) {
    SpreadsheetApp.getUi().alert("A aba 'Analise' ou 'Peso da Nota' não foi encontrada.");
    return;
  }

  var data = analiseSheet.getDataRange().getValues();
  var pesosData = pesoSheet.getDataRange().getValues();

  if (data.length <= 1 || pesosData.length <= 1) {
    SpreadsheetApp.getUi().alert("As abas 'Analise' ou 'Peso da Nota' não possuem dados suficientes.");
    return;
  }

  var headers = data[0].map(h => (h || "").toString().trim().toLowerCase());
  var colCliente = headers.indexOf("cliente");
  var colModulo = headers.indexOf("módulo");
  var colCriterio = headers.indexOf("critérios");
  var colAnalise = headers.indexOf("análise - utilizado");
  var colPosOng = headers.indexOf("pós-ong utilizado");
  var col30Dias = headers.indexOf("30 dias");
  var col90Dias = headers.indexOf("90 dias");

  if ([colCliente, colModulo, colCriterio, colAnalise, colPosOng, col30Dias, col90Dias].some(idx => idx === -1)) {
    SpreadsheetApp.getUi().alert("Uma ou mais colunas essenciais não foram encontradas na aba 'Analise'.");
    return;
  }

  var pesos = {};
  pesosData.slice(1).forEach(row => {
    var modulo = (row[0] || "").toString().toLowerCase().trim();
    var criterio = (row[1] || "").toString().toLowerCase().trim();
    var peso = parseFloat(row[2]) || 0;

    if (!pesos[modulo]) pesos[modulo] = {};
    pesos[modulo][criterio] = peso;
  });

  var modulos = Object.keys(pesos);

  modulos.forEach(modulo => {
    var abaNome = modulo.charAt(0).toUpperCase() + modulo.slice(1);
    var moduloSheet = ss.getSheetByName(abaNome);

    if (!moduloSheet) moduloSheet = ss.insertSheet(abaNome);
    else moduloSheet.clear();

    var criterios = Object.keys(pesos[modulo]);
    var headersModulo = ["Cliente"].concat(criterios).concat(["Status"]);
    moduloSheet.getRange(1, 1, 1, headersModulo.length).setValues([headersModulo])
      .setFontWeight("bold")
      .setHorizontalAlignment("center");

    var detalhes = [];
    var rows = data.slice(1);

    rows.forEach(row => {
      var cliente = row[colCliente];
      var moduloAtual = row[colModulo] ? row[colModulo].toString().toLowerCase().trim() : "";
      var criterio = row[colCriterio] ? row[colCriterio].toString().toLowerCase().trim() : "";

      if (moduloAtual === modulo) {
        var valorMaisRecente = row[col90Dias] || row[col30Dias] || row[colPosOng] || row[colAnalise] || "Vazio";
        valorMaisRecente = valorMaisRecente === 2 ? STATUS.ENGAGED : valorMaisRecente === 1 ? STATUS.NOT_ENGAGED : STATUS.NO_TOOL;

        var linhaCliente = detalhes.find(d => d[0] === cliente) || [cliente].concat(Array(criterios.length).fill("Vazio"));
        var indiceCriterio = criterios.indexOf(criterio);

        if (indiceCriterio !== -1) {
          linhaCliente[indiceCriterio + 1] = valorMaisRecente;
        }

        if (!detalhes.some(d => d[0] === cliente)) detalhes.push(linhaCliente);
      }
    });

    detalhes.forEach(linha => {
      var pesoTotal = 100;
      var pesoEngajado = 0;
      var temVazio = false;
      var todosSemFerramenta = true;

      criterios.forEach((crit, index) => {
        if (linha[index + 1] === STATUS.ENGAGED) pesoEngajado += pesos[modulo][crit] || 0;
        if (linha[index + 1] === "Vazio") temVazio = true;
        if (linha[index + 1] !== STATUS.NO_TOOL) todosSemFerramenta = false;
      });

      var porcentagem = (pesoEngajado / pesoTotal) * 100;

      if (todosSemFerramenta) linha.push(STATUS.NO_TOOL);
      else if (temVazio) linha.push(STATUS.NOT_ENGAGED);
      else if (porcentagem >= 60) linha.push(STATUS.ENGAGED);
      else if (porcentagem >= 41) linha.push(STATUS.PARTIAL);
      else linha.push(STATUS.NOT_ENGAGED);
    });

    if (detalhes.length > 0) {
      moduloSheet.getRange(2, 1, detalhes.length, detalhes[0].length).setValues(detalhes);
    } else {
      moduloSheet.getRange(2, 1).setValue("Nenhum dado encontrado.");
    }

    moduloSheet.autoResizeColumns(1, headersModulo.length);
    moduloSheet.getRange(1, 1, detalhes.length + 1, headersModulo.length)
      .setBorder(true, true, true, true, true, true)
      .setHorizontalAlignment("center");
  });
}

function columnToNumber(column) {
  var base = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";
  return column.split("").reduce((acc, char) => acc * 26 + base.indexOf(char) + 1, 0);
}

function criarAbaAnalytics() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var analiseSheet = ss.getSheetByName(SHEETS.ANALYSIS);

  if (!analiseSheet) {
    SpreadsheetApp.getUi().alert("A aba 'Analise' não foi encontrada.");
    return;
  }

  var data = analiseSheet.getDataRange().getValues();
  if (data.length <= 1) {
    SpreadsheetApp.getUi().alert("A aba 'Analise' não possui dados suficientes.");
    return;
  }

  var headers = data[0].map(h => (h || "").toString().trim().toLowerCase());
  var colAnaliseUtilizado = headers.indexOf("análise - utilizado");
  var colPosOngUtilizado = headers.indexOf("pós-ong utilizado");
  var col30Dias = headers.indexOf("30 dias");
  var col90Dias = headers.indexOf("90 dias");

  if ([colAnaliseUtilizado, colPosOngUtilizado, col30Dias, col90Dias].some(idx => idx === -1)) {
    SpreadsheetApp.getUi().alert("Uma ou mais colunas essenciais não foram encontradas na aba 'Analise'.");
    return;
  }

  var criterios = ["consulta de vendas região/estado", "consulta comparativa", "consulta de vendas", "contrato (exec.)", " painel da", "consulta contratos", "gestão", "manutenção", "justificativas"];
  var abaNome = "Analytics";
  var moduloSheet = ss.getSheetByName(abaNome);

  if (!moduloSheet) {
    SpreadsheetApp.getUi().alert("A aba '" + abaNome + "' não foi encontrada.");
    return;
  }

  var startColumn = 12;

  criterios.forEach(criterio => {
    var criterioNome = `analytics_${criterio}`;
    var headersResumo = [
      criterioNome,
      `ANÁLISE - UTILIZADO_${criterioNome}`,
      `PÓS-ONG UTILIZADO_${criterioNome}`,
      `30 DIAS_${criterioNome}`,
      `90 DIAS_${criterioNome}`
    ];
    moduloSheet.getRange(1, startColumn, 1, headersResumo.length).setValues([headersResumo]).setFontWeight("bold");

    var analiseResultados = [
      ["Engajado",
        contarPorCriterio("analytics", criterio, colAnaliseUtilizado, 2),
        contarPorCriterio("analytics", criterio, colPosOngUtilizado, 2),
        contarPorCriterio("analytics", criterio, col30Dias, 2),
        contarPorCriterio("analytics", criterio, col90Dias, 2)
      ],
      ["Não engajado",
        contarPorCriterio("analytics", criterio, colAnaliseUtilizado, 1),
        contarPorCriterio("analytics", criterio, colPosOngUtilizado, 1),
        contarPorCriterio("analytics", criterio, col30Dias, 1),
        contarPorCriterio("analytics", criterio, col90Dias, 1)
      ]
    ];

    moduloSheet.getRange(2, startColumn, analiseResultados.length, headersResumo.length).setValues(analiseResultados);
    startColumn += headersResumo.length;
  });

  moduloSheet.autoResizeColumns(12, startColumn - 12);
}

function criarAbaForcaDeVendas() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var analiseSheet = ss.getSheetByName(SHEETS.ANALYSIS);

  if (!analiseSheet) {
    SpreadsheetApp.getUi().alert("A aba 'Analise' não foi encontrada.");
    return;
  }

  var data = analiseSheet.getDataRange().getValues();
  if (data.length <= 1) {
    SpreadsheetApp.getUi().alert("A aba 'Analise' não possui dados suficientes.");
    return;
  }

  var headers = data[0].map(h => (h || "").toString().trim().toLowerCase());
  var colAnaliseUtilizado = headers.indexOf("análise - utilizado");
  var colPosOngUtilizado = headers.indexOf("pós-ong utilizado");
  var col30Dias = headers.indexOf("30 dias");
  var col90Dias = headers.indexOf("90 dias");

  if ([colAnaliseUtilizado, colPosOngUtilizado, col30Dias, col90Dias].some(idx => idx === -1)) {
    SpreadsheetApp.getUi().alert("Uma ou mais colunas essenciais não foram encontradas na aba 'Analise'.");
    return;
  }

  var criterios = ["emitir pedido", "mapas", "gestão de clientes", "analise de resultados", "comissões", "imagens", "hist. compra cli"];
  var abaNome = "Força de Vendas";
  var moduloSheet = ss.getSheetByName(abaNome);

  if (!moduloSheet) {
    SpreadsheetApp.getUi().alert("A aba '" + abaNome + "' não foi encontrada.");
    return;
  }

  var startColumn = 10;

  criterios.forEach(criterio => {
    var criterioNome = `força_de_vendas_${criterio}`;
    var headersResumo = [
      criterioNome,
      `ANÁLISE - UTILIZADO_${criterioNome}`,
      `PÓS-ONG UTILIZADO_${criterioNome}`,
      `30 DIAS_${criterioNome}`,
      `90 DIAS_${criterioNome}`
    ];
    moduloSheet.getRange(1, startColumn, 1, headersResumo.length).setValues([headersResumo]).setFontWeight("bold");

    var analiseResultados = [
      ["Engajado",
        contarPorCriterio("força de vendas", criterio, colAnaliseUtilizado, 2),
        contarPorCriterio("força de vendas", criterio, colPosOngUtilizado, 2),
        contarPorCriterio("força de vendas", criterio, col30Dias, 2),
        contarPorCriterio("força de vendas", criterio, col90Dias, 2)
      ],
      ["Não engajado",
        contarPorCriterio("força de vendas", criterio, colAnaliseUtilizado, 1),
        contarPorCriterio("força de vendas", criterio, colPosOngUtilizado, 1),
        contarPorCriterio("força de vendas", criterio, col30Dias, 1),
        contarPorCriterio("força de vendas", criterio, col90Dias, 1)
      ]
    ];

    moduloSheet.getRange(2, startColumn, analiseResultados.length, headersResumo.length).setValues(analiseResultados);
    startColumn += headersResumo.length;
  });

  moduloSheet.autoResizeColumns(10, startColumn - 10);
}

function criarAbaCRM() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var analiseSheet = ss.getSheetByName(SHEETS.ANALYSIS);

  if (!analiseSheet) {
    SpreadsheetApp.getUi().alert("A aba 'Analise' não foi encontrada.");
    return;
  }

  var data = analiseSheet.getDataRange().getValues();
  if (data.length <= 1) {
    SpreadsheetApp.getUi().alert("A aba 'Analise' não possui dados suficientes.");
    return;
  }

  var headers = data[0].map(h => (h || "").toString().trim().toLowerCase());
  var colAnaliseUtilizado = headers.indexOf("análise - utilizado");
  var colPosOngUtilizado = headers.indexOf("pós-ong utilizado");
  var col30Dias = headers.indexOf("30 dias");
  var col90Dias = headers.indexOf("90 dias");

  if ([colAnaliseUtilizado, colPosOngUtilizado, col30Dias, col90Dias].some(idx => idx === -1)) {
    SpreadsheetApp.getUi().alert("Uma ou mais colunas essenciais não foram encontradas na aba 'Analise'.");
    return;
  }

  var criterios = ["dashboards de agendamentos", "painel de relacionamento", "dashboards movimentações", "relatórios"];
  var abaNome = "CRM";
  var moduloSheet = ss.getSheetByName(abaNome);

  if (!moduloSheet) {
    SpreadsheetApp.getUi().alert("A aba '" + abaNome + "' não foi encontrada.");
    return;
  }

  var startColumn = 8;

  criterios.forEach(criterio => {
    var criterioNome = `crm_${criterio}`;
    var headersResumo = [
      criterioNome,
      `ANÁLISE - UTILIZADO_${criterioNome}`,
      `PÓS-ONG UTILIZADO_${criterioNome}`,
      `30 DIAS_${criterioNome}`,
      `90 DIAS_${criterioNome}`
    ];
    moduloSheet.getRange(1, startColumn, 1, headersResumo.length).setValues([headersResumo]).setFontWeight("bold");

    var analiseResultados = [
      ["Engajado",
        contarPorCriterio("crm", criterio, colAnaliseUtilizado, 2),
        contarPorCriterio("crm", criterio, colPosOngUtilizado, 2),
        contarPorCriterio("crm", criterio, col30Dias, 2),
        contarPorCriterio("crm", criterio, col90Dias, 2)
      ],
      ["Não engajado",
        contarPorCriterio("crm", criterio, colAnaliseUtilizado, 1),
        contarPorCriterio("crm", criterio, colPosOngUtilizado, 1),
        contarPorCriterio("crm", criterio, col30Dias, 1),
        contarPorCriterio("crm", criterio, col90Dias, 1)
      ]
    ];

    moduloSheet.getRange(2, startColumn, analiseResultados.length, headersResumo.length).setValues(analiseResultados);
    startColumn += headersResumo.length;
  });

  moduloSheet.autoResizeColumns(8, startColumn - 8);
}

function criarAbaPortalB2B() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var analiseSheet = ss.getSheetByName(SHEETS.ANALYSIS);

  if (!analiseSheet) {
    SpreadsheetApp.getUi().alert("A aba 'Analise' não foi encontrada.");
    return;
  }

  var data = analiseSheet.getDataRange().getValues();
  if (data.length <= 1) {
    SpreadsheetApp.getUi().alert("A aba 'Analise' não possui dados suficientes.");
    return;
  }

  var headers = data[0].map(h => (h || "").toString().trim().toLowerCase());
  var colAnaliseUtilizado = headers.indexOf("análise - utilizado");
  var colPosOngUtilizado = headers.indexOf("pós-ong utilizado");
  var col30Dias = headers.indexOf("30 dias");
  var col90Dias = headers.indexOf("90 dias");

  if ([colAnaliseUtilizado, colPosOngUtilizado, col30Dias, col90Dias].some(idx => idx === -1)) {
    SpreadsheetApp.getUi().alert("Uma ou mais colunas essenciais não foram encontradas na aba 'Analise'.");
    return;
  }

  var criterios = ["entrada de pedido", "clientes cadastrados", "dash", "banner", "imagens", "menu/galerias", "institucionais", "rede social", "política", "cupom de desconto", "portal do cliente"];
  var abaNome = "Portal B2B";
  var moduloSheet = ss.getSheetByName(abaNome);

  if (!moduloSheet) {
    SpreadsheetApp.getUi().alert("A aba '" + abaNome + "' não foi encontrada.");
    return;
  }

  var startColumn = 14;

  criterios.forEach(criterio => {
    var criterioNome = `portal_b2b_${criterio}`;
    var headersResumo = [
      criterioNome,
      `ANÁLISE - UTILIZADO_${criterioNome}`,
      `PÓS-ONG UTILIZADO_${criterioNome}`,
      `30 DIAS_${criterioNome}`,
      `90 DIAS_${criterioNome}`
    ];
    moduloSheet.getRange(1, startColumn, 1, headersResumo.length).setValues([headersResumo]).setFontWeight("bold");

    var analiseResultados = [
      ["Engajado",
        contarPorCriterio("portal b2b", criterio, colAnaliseUtilizado, 2),
        contarPorCriterio("portal b2b", criterio, colPosOngUtilizado, 2),
        contarPorCriterio("portal b2b", criterio, col30Dias, 2),
        contarPorCriterio("portal b2b", criterio, col90Dias, 2)
      ],
      ["Não engajado",
        contarPorCriterio("portal b2b", criterio, colAnaliseUtilizado, 1),
        contarPorCriterio("portal b2b", criterio, colPosOngUtilizado, 1),
        contarPorCriterio("portal b2b", criterio, col30Dias, 1),
        contarPorCriterio("portal b2b", criterio, col90Dias, 1)
      ]
    ];

    moduloSheet.getRange(2, startColumn, analiseResultados.length, headersResumo.length).setValues(analiseResultados);
    startColumn += headersResumo.length;
  });

  moduloSheet.autoResizeColumns(14, startColumn - 14);
}

function contarPorCriterio(modulo, criterio, coluna, valor) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var analiseSheet = ss.getSheetByName(SHEETS.ANALYSIS);
  var data = analiseSheet.getDataRange().getValues();
  var headers = data[0].map(h => (h || "").toString().trim().toLowerCase());
  var colModulo = headers.indexOf("módulo");
  var colCriterio = headers.indexOf("critérios");

  return data.filter(row =>
    (row[colModulo] || "").toString().toLowerCase() === modulo &&
    (row[colCriterio] || "").toString().toLowerCase() === criterio &&
    row[coluna] === valor
  ).length;
}

function attlooker() {
  criarAbasPorFerramenta();
  criarAbaAnalytics();
  criarAbaForcaDeVendas();
  criarAbaCRM();
  criarAbaPortalB2B();
}
