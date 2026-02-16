function getSheet(name) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(name);
  if (!sheet) throw new Error(`Sheet "${name}" not found.`);
  return sheet;
}

function normalizeText(text) {
  if (!text) return "";
  return text.toString().trim().toUpperCase();
}

/**
 * Formata CNPJ/CPF para comparação (mantém só números).
 * Retorna "00000000000000" se vazio/ inválido.
 */
function formatarCNPJ(valor) {
  if (valor === null || valor === undefined) return "00000000000000";
  const digits = valor.toString().replace(/\D/g, "");
  if (!digits) return "00000000000000";
  // Mantém tamanho “max” de CNPJ (14) mas não força padding (depende do seu dataset)
  return digits;
}

