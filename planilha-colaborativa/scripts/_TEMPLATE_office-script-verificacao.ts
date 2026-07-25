/**
 * TEMPLATE — Office Script (Excel Online) com padrão VERIFICAÇÃO-PRIMEIRO.
 *
 * Anatomia-padrão do projeto. Copie este esqueleto ao criar um script novo.
 *
 * Regras do projeto embutidas aqui:
 *  - MODO_VERIFICACAO = true por padrão: não altera nada, só relata.
 *  - Coluna localizada por NOME DE CABEÇALHO, nunca por letra fixa.
 *  - Log acumulado num array e impresso UMA vez (sem console.log em laço).
 *  - Aborto de segurança se o volume de mudanças fugir do esperado.
 *  - Sem acentos/emoji em strings (montar via String.fromCharCode se preciso).
 */
function main(workbook: ExcelScript.Workbook) {

  // 1) Trava de segurança — comece SEMPRE em true.
  const MODO_VERIFICACAO = true;

  // Limite de aborto: se o script fosse alterar mais que isto, algo está errado.
  const LIMITE_MUDANCAS = 50;

  // 2) Nome da aba sem acento no fonte (ex.: "PROGRAMA" + Ç + Ã + "O").
  const nomeAba = "PROGRAMA" + String.fromCharCode(199) + String.fromCharCode(195) + "O";
  const aba = workbook.getWorksheet(nomeAba);
  if (!aba) {
    console.log("ERRO: aba nao encontrada.");
    return;
  }

  // 3) Localizar coluna por cabeçalho (linha do cabeçalho = 4 neste projeto).
  const LINHA_CABECALHO = 4;
  const cabecalho = aba.getRange(LINHA_CABECALHO + ":" + LINHA_CABECALHO).getValues()[0];

  function colunaPorNome(nome: string): number {
    for (let c = 0; c < cabecalho.length; c++) {
      if (String(cabecalho[c]).trim().toUpperCase() === nome.toUpperCase()) {
        return c; // índice 0-based
      }
    }
    throw new Error("Coluna nao encontrada: " + nome);
  }

  // Exemplo de uso:
  // const idxStatus = colunaPorNome("STATUS");

  // 4) Percorrer os dados acumulando o relatório (NÃO logar dentro do laço).
  const relatorio: string[] = [];
  let mudancas = 0;

  const usado = aba.getUsedRange();
  const primeiraLinhaDados = LINHA_CABECALHO + 1; // dados começam na linha 5
  const ultimaLinha = usado.getRowIndex() + usado.getRowCount();

  for (let linha = primeiraLinhaDados; linha <= ultimaLinha; linha++) {
    // ---- lógica da tarefa aqui ----
    // Exemplo genérico:
    // const cel = aba.getCell(linha - 1, idxAlgumaColuna);
    // const atual = cel.getValue();
    // const novo = /* transformar atual */;
    // if (atual !== novo) {
    //   relatorio.push("L" + linha + ": ANTES(" + atual + ") DEPOIS(" + novo + ")");
    //   mudancas++;
    //   if (!MODO_VERIFICACAO) cel.setValue(novo);
    // }
  }

  // 5) Aborto de segurança.
  if (mudancas > LIMITE_MUDANCAS) {
    console.log("ABORTADO: " + mudancas + " mudancas excedem o limite de " + LIMITE_MUDANCAS + ".");
    return;
  }

  // 6) Impressão única.
  console.log("Mudancas previstas: " + mudancas);
  console.log(relatorio.join(" | "));
  console.log(MODO_VERIFICACAO ? "*** MODO VERIFICACAO — nada alterado ***" : "*** EXECUCAO CONCLUIDA ***");
}
