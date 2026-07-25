/**
 * DESPROTEGER — aba PROGRAMAÇÃO
 * Operador Logístico × Cliente Exportador
 *
 * Remove a proteção aplicada pelo script "Proteger_Formulas_PROGRAMACAO".
 * Use quando precisar editar fórmulas/cabeçalho ou reconfigurar a aba.
 *
 * COMO USAR (Excel Online):
 *   Aba "Automatizar" → "Novo Script" → cole este código → "Executar".
 */

function main(workbook: ExcelScript.Workbook) {
  const NOME_ABA = "PROGRAMAÇÃO";
  const sheet = workbook.getWorksheet(NOME_ABA);
  if (!sheet) {
    console.log(`Aba "${NOME_ABA}" não encontrada.`);
    return;
  }
  const protection = sheet.getProtection();
  if (protection.getIsProtected()) {
    protection.unprotect();
    console.log(`OK — proteção removida de "${NOME_ABA}".`);
  } else {
    console.log(`"${NOME_ABA}" já estava sem proteção.`);
  }
}
