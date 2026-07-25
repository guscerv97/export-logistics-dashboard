/**
 * PROTEGER FÓRMULAS — aba PROGRAMAÇÃO
 * Operador Logístico × Cliente Exportador
 *
 * O QUE FAZ:
 *   - Libera (destrava) todas as células da aba para edição normal.
 *   - Trava APENAS as colunas de fórmula automática e o cabeçalho (linhas 1–4).
 *   - Ativa a proteção da planilha, mantendo filtro, ordenação, formatação e
 *     inserção/exclusão de linhas liberados — só bloqueia SOBRESCREVER fórmula.
 *
 * COMO USAR (Excel Online):
 *   Aba "Automatizar" → "Novo Script" → cole este código → "Executar".
 *   Reversível a qualquer momento com o script "Desproteger_PROGRAMACAO".
 *
 * OBS.: A coluna A (STATUS) NÃO é travada de propósito — ela tem fórmula, mas
 *   "Cancelado" e "Depositado" são marcados manualmente por cima. Se quiser
 *   travar A também, adicione "A" na lista COLUNAS_FORMULA abaixo.
 */

function main(workbook: ExcelScript.Workbook) {
  const NOME_ABA = "PROGRAMAÇÃO";
  const COLUNAS_FORMULA = ["K", "O", "S", "AD", "AE", "AF", "AM"]; // fórmulas 🔄
  const LINHAS_CABECALHO = "1:4"; // banner + títulos

  const sheet = workbook.getWorksheet(NOME_ABA);
  if (!sheet) {
    console.log(`Aba "${NOME_ABA}" não encontrada. Verifique o nome exato.`);
    return;
  }

  const protection = sheet.getProtection();

  // 1) remove proteção anterior (se houver) para poder reconfigurar
  if (protection.getIsProtected()) {
    protection.unprotect();
  }

  // 2) destrava tudo o que está em uso (área editável por padrão)
  const used = sheet.getUsedRange();
  if (used) {
    used.getFormat().getProtection().locked = false;
  }

  // 3) trava as colunas de fórmula (coluna inteira, cobre linhas futuras)
  for (const col of COLUNAS_FORMULA) {
    sheet.getRange(`${col}:${col}`).getFormat().getProtection().locked = true;
  }

  // 4) trava o cabeçalho (linhas 1–4)
  sheet.getRange(LINHAS_CABECALHO).getFormat().getProtection().locked = true;

  // 5) protege a planilha deixando o resto do fluxo liberado
  protection.protect({
    allowAutoFilter: true,
    allowSort: true,
    allowFormatCells: true,
    allowFormatColumns: true,
    allowFormatRows: true,
    allowInsertRows: true,
    allowInsertColumns: false,
    allowDeleteRows: true,
    allowDeleteColumns: false,
    allowEditObjects: true,
    allowPivotTables: true
  });

  console.log(
    `OK — "${NOME_ABA}" protegida. ` +
    `Colunas travadas: ${COLUNAS_FORMULA.join(", ")} + cabeçalho (1–4). ` +
    `Demais colunas continuam editáveis.`
  );
}
