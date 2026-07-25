/**
 * CORREÇÃO — formato de data da aba ESTUFAGEM
 * Ajusta "Data estufagem" (C14) e "D.L Carga" (C22) para dd/mm/aaaa.
 * Não altera mais nada.
 *
 * Aprendizado embutido: usar setNumberFormat (API neutra, string em inglês),
 * NUNCA setNumberFormatLocal (que depende do idioma e quebra em sessões EN).
 */
function main(workbook: ExcelScript.Workbook) {
  const sh = workbook.getWorksheet("ESTUFAGEM");
  if (!sh) throw new Error('Aba "ESTUFAGEM" nao encontrada.');

  sh.getRange("C14").setNumberFormat("dd/mm/yyyy");
  sh.getRange("C22").setNumberFormat("dd/mm/yyyy");

  console.log("Formato de data corrigido em C14 (Data estufagem) e C22 (D.L Carga).");
}
