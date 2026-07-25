/**
 * FASE 2 — Reposicionamento de colunas + fórmulas + proteção
 * Cliente Exportador × Operador Logístico
 *
 * O QUE FAZ (nesta ordem):
 *   1. Remove a coluna O (COLETADO 🔄)
 *   2. Move "Alterações"        -> para antes de REF. INTERNA (OPERADOR)
 *   3. Move "DATA DE COLETA"    -> para depois de CONTAINER
 *   4. Move "AUT. EMBARQUE"     -> para entre CAVALO/CARRETA e DATA/HORA SAÍDA PLANTA
 *   5. Reescreve as 6 fórmulas da PROGRAMAÇÃO nas novas posições
 *      (incluindo o novo STATUS, que passa a usar DATA DE COLETA)
 *   6. Reescreve as 9 fórmulas da ESTUFAGEM que mudaram de coluna-alvo
 *   7. Reaplica as 4 listas suspensas nas novas posições
 *   8. Protege as colunas de fórmula contra digitação
 *
 * O QUE NÃO FAZ:
 *   - Não mexe em DASHBOARD nem DASHBOARD COLETA (nenhuma fórmula deles mudou de coluna)
 *   - Não mexe na aba ARQUIVO
 *   - Não reconstrói a formatação condicional -> isso é a FASE 3
 *
 * ⚠️ ATENÇÃO: entre a Fase 2 e a Fase 3 as CORES da planilha ficarão desalinhadas,
 *    porque as regras de formatação condicional ainda apontam para as posições antigas.
 *    Isso é esperado. Rode a Fase 3 na sequência.
 *
 * COMO USAR:
 *   1) MODO_VERIFICACAO = true  -> confere o layout e imprime o plano, sem alterar nada
 *   2) Confira o relatório
 *   3) Troque para false e rode de novo
 */

function main(workbook: ExcelScript.Workbook) {

  // ====== CONFIGURAÇÃO ======
  const MODO_VERIFICACAO = true;      // <<< troque para false só depois de conferir
  const APLICAR_PROTECAO = true;      // trava as colunas de fórmula ao final
  const PRIMEIRA_LINHA = 5;           // dados começam na linha 5
  const ULTIMA_LINHA = 517;           // buffer: fórmulas e trava vão até aqui
  const LINHA_CABECALHO = 4;
  const LINHAS_COPIA = 600;           // faixa usada ao mover colunas (dados + folga)
  // ==========================

  const prog = workbook.getWorksheet("PROGRAMAÇÃO");
  const est = workbook.getWorksheet("ESTUFAGEM");

  if (!prog) { console.log("ERRO: aba PROGRAMAÇÃO não encontrada."); return; }
  if (!est) { console.log("ERRO: aba ESTUFAGEM não encontrada."); return; }

  console.log("=== FASE 2 — REPOSICIONAMENTO DE COLUNAS ===");
  console.log("Modo: " + (MODO_VERIFICACAO ? "VERIFICAÇÃO (nada será alterado)" : "EXECUÇÃO"));
  console.log("");

  // ================================================================
  // PRÉ-VOO — confere se o layout atual é o esperado
  // ================================================================
  const esperado: [string, string][] = [
    ["O", "COLETADO"],
    ["P", "COLETA"],
    ["Q", "REF. INTERNA"],
    ["AI", "ALTERA"],
    ["AJ", "CONTAINER"],
    ["AO", "MOTORISTA"],
    ["AP", "CAVALO"],
    ["AQ", "SAÍDA PLANTA"],
    ["AY", "AUT. EMBARQUE"]
  ];

  // lê a linha de cabeçalho inteira de uma vez (evita leitura dentro do laço)
  const cabecalhos = prog.getRange("A" + LINHA_CABECALHO + ":BH" + LINHA_CABECALHO).getValues()[0];

  const problemas: string[] = [];
  for (const par of esperado) {
    const col = par[0];
    const trecho = par[1];
    const titulo = String(cabecalhos[letraParaIndice(col) - 1] || "")
      .toUpperCase().replace(/\n/g, " ");
    if (titulo.indexOf(trecho.toUpperCase()) === -1) {
      problemas.push("  " + col + LINHA_CABECALHO + " deveria conter \"" + trecho + "\" mas contém \"" + titulo + "\"");
    }
  }

  console.log("--- PRÉ-VOO: conferência do layout atual ---");
  if (problemas.length > 0) {
    console.log("  ❌ LAYOUT INESPERADO — script ABORTADO por segurança:");
    console.log(problemas.join("\n"));
    console.log("  Nada foi alterado. Me envie este log.");
    return;
  }
  console.log("  ✅ Todas as 9 colunas de referência estão onde deveriam.");
  console.log("");

  if (MODO_VERIFICACAO) {
    console.log("--- PLANO DE ALTERAÇÃO (nada será executado agora) ---");
    console.log("  1. EXCLUIR coluna O (COLETADO 🔄)");
    console.log("  2. MOVER  Alterações      : AI -> O");
    console.log("  3. MOVER  DATA DE COLETA  : P  -> AI (depois de CONTAINER, que vai para AH)");
    console.log("  4. MOVER  AUT. EMBARQUE   : AY -> AP");
    console.log("  5. Layout final: 59 colunas (A -> BG)");
    console.log("");
    console.log("  FÓRMULAS PROGRAMAÇÃO (linhas " + PRIMEIRA_LINHA + " a " + ULTIMA_LINHA + "):");
    console.log("     A  STATUS            -> passa a usar DATA DE COLETA (AI) em vez de CONTAINER");
    console.log("     K  SEMANA            -> inalterada");
    console.log("     R  TIPO SOLICITACAO  -> era S");
    console.log("     AC DATA IDEAL        -> era AD");
    console.log("     AD DATA LIMITE       -> era AE");
    console.log("     AE ALERTA DETENTION  -> era AF");
    console.log("     AL QTDE CARREGAR     -> era AM");
    console.log("");
    console.log("  FÓRMULAS ESTUFAGEM: 9 células (C5, E5, C13, C15, C17, C19, C20, C22, B36, B40)");
    console.log("  VALIDAÇÕES: A, F, U (era V), O (era AI)");
    console.log("  PROTEÇÃO: " + (APLICAR_PROTECAO ? "K, R, AC, AD, AE, AL + cabeçalho travados" : "desativada"));
    console.log("");
    console.log("=== MODO VERIFICAÇÃO — nenhuma alteração foi feita. ===");
    return;
  }

  // ================================================================
  // ETAPA 1 — Excluir coluna O (COLETADO)
  // ================================================================
  prog.getRange("O:O").delete(ExcelScript.DeleteShiftDirection.left);
  console.log("1/8  Coluna O (COLETADO) excluída.");

  // ================================================================
  // ETAPA 2 — Mover "Alterações" (agora AH) para P, antes de REF. INTERNA
  // ================================================================
  moverColuna(prog, "AH", "P", LINHAS_COPIA);
  console.log("2/8  Alterações movida para antes de REF. INTERNA.");

  // ================================================================
  // ETAPA 3 — Mover "DATA DE COLETA" (agora O) para depois de CONTAINER (AI)
  // ================================================================
  moverColuna(prog, "O", "AJ", LINHAS_COPIA);
  console.log("3/8  DATA DE COLETA movida para depois de CONTAINER.");

  // ================================================================
  // ETAPA 4 — Mover "AUT. EMBARQUE" (agora AX) para AP
  // ================================================================
  moverColuna(prog, "AX", "AP", LINHAS_COPIA);
  console.log("4/8  AUT. EMBARQUE movida para depois de CAVALO/CARRETA.");

  // ================================================================
  // ETAPA 5 — Fórmulas da PROGRAMAÇÃO
  // ================================================================
  const nLinhas = ULTIMA_LINHA - PRIMEIRA_LINHA + 1;

  escreverColuna(prog, "A", PRIMEIRA_LINHA, nLinhas,
    (L) => '=IF($B' + L + '="","",IF($AI' + L + '<>"","Container retirado","Nova programação"))');

  escreverColuna(prog, "K", PRIMEIRA_LINHA, nLinhas,
    (L) => '=IF(L' + L + '="","",ISOWEEKNUM(L' + L + '))');

  escreverColuna(prog, "R", PRIMEIRA_LINHA, nLinhas,
    (L) => '=IF(OR($Q' + L + '="",$K' + L + '=""),"",IF($Q' + L + '=MINIFS($Q$' + PRIMEIRA_LINHA + ':$Q$' + ULTIMA_LINHA + ',$K$' + PRIMEIRA_LINHA + ':$K$' + ULTIMA_LINHA + ',$K' + L + '),"Programacao","Extra"))');

  escreverColuna(prog, "AC", PRIMEIRA_LINHA, nLinhas,
    (L) => '=IFERROR(IF(L' + L + '="","",L' + L + '-8),"")');

  escreverColuna(prog, "AD", PRIMEIRA_LINHA, nLinhas,
    (L) => '=IF(OR($AI' + L + '="",$AA' + L + '=""),"",$AI' + L + '+$AA' + L + '-1)');

  escreverColuna(prog, "AE", PRIMEIRA_LINHA, nLinhas,
    (L) => '=IF(OR($AD' + L + '="",$A' + L + '="Cancelado"),"",IF($BD' + L + '<>"",$AD' + L + '-$BD' + L + ',$AD' + L + '-TODAY()))');

  escreverColuna(prog, "AL", PRIMEIRA_LINHA, nLinhas,
    (L) => '=IFERROR(IF(OR(AJ' + L + '="",AK' + L + '=""),"",AJ' + L + '-AK' + L + '),"")');

  console.log("5/8  7 colunas de fórmula reescritas na PROGRAMAÇÃO (linhas " + PRIMEIRA_LINHA + "-" + ULTIMA_LINHA + ").");

  // Formatos numéricos das colunas recalculadas
  prog.getRange("K" + PRIMEIRA_LINHA + ":K" + ULTIMA_LINHA).setNumberFormat("0");
  prog.getRange("AC" + PRIMEIRA_LINHA + ":AD" + ULTIMA_LINHA).setNumberFormat("dd/mm/yyyy");
  prog.getRange("AE" + PRIMEIRA_LINHA + ":AE" + ULTIMA_LINHA).setNumberFormat("0");

  // ================================================================
  // ETAPA 6 — Fórmulas da ESTUFAGEM (9 células que mudaram de alvo)
  // ================================================================
  const X = (colAlvo: string) =>
    '=IF($C$11="","",XLOOKUP($C$11,PROGRAMAÇÃO!H:H,PROGRAMAÇÃO!' + colAlvo + ':' + colAlvo + ',"NÃO ENCONTRADO"))';

  est.getRange("C5").setFormula(X("T"));    // LOCAL DE RETIRADA   (era U)
  est.getRange("E5").setFormula(X("AH"));   // CONTAINER           (era AJ)
  est.getRange("C13").setFormula(X("V"));   // BOOKING             (era W)
  est.getRange("C15").setFormula(X("X"));   // ARMADOR             (era Y)
  est.getRange("C17").setFormula(X("W"));   // NAVIO               (era X)
  est.getRange("C19").setFormula(X("U"));   // TIPO DE COLETA      (era V)
  est.getRange("C20").setFormula(X("S"));   // PORTO DE COLETA     (era T)
  est.getRange("C22").setFormula(X("Z"));   // D.L CARGA           (era AA)

  est.getRange("B36").setFormula(
    '=IF($C$11="","","SOLICITAÇÃO DE ESTUFAGEM SEMANA " & XLOOKUP($C$11,PROGRAMAÇÃO!H:H,PROGRAMAÇÃO!K:K,"NÃO ENCONTRADO")' +
    ' & " - CONTRATO " & $C$11 & " / BOOKING " & XLOOKUP($C$11,PROGRAMAÇÃO!H:H,PROGRAMAÇÃO!V:V,"NÃO ENCONTRADO")' +
    ' & " - " & TEXT(XLOOKUP($C$11,PROGRAMAÇÃO!H:H,PROGRAMAÇÃO!L:L,"NÃO ENCONTRADO"),"dd/mm/yyyy")' +
    ' & " EM PLANTA " & XLOOKUP($C$11,PROGRAMAÇÃO!H:H,PROGRAMAÇÃO!N:N,"NÃO ENCONTRADO")' +
    ' & " / ARMADOR " & XLOOKUP($C$11,PROGRAMAÇÃO!H:H,PROGRAMAÇÃO!X:X,"NÃO ENCONTRADO"))'
  );

  est.getRange("B40").setFormula(
    '=IF($C$11="","","Liberação de unidade - " & XLOOKUP($C$11,PROGRAMAÇÃO!H:H,PROGRAMAÇÃO!M:M,"NÃO ENCONTRADO")' +
    ' & " - (PÁTIO) - CLIENTE EXPORTADOR - " & $C$11 & " / " & XLOOKUP($C$11,PROGRAMAÇÃO!H:H,PROGRAMAÇÃO!V:V,"NÃO ENCONTRADO")' +
    ' & " / " & XLOOKUP($C$11,PROGRAMAÇÃO!H:H,PROGRAMAÇÃO!AH:AH,"NÃO ENCONTRADO")' +
    ' & " - PROGRAMAÇÃO SEMANA " & XLOOKUP($C$11,PROGRAMAÇÃO!H:H,PROGRAMAÇÃO!K:K,"NÃO ENCONTRADO"))'
  );

  console.log("6/8  10 fórmulas da ESTUFAGEM atualizadas.");

  // ================================================================
  // ETAPA 7 — Listas suspensas nas novas posições
  // ================================================================
  aplicarLista(prog, "A" + PRIMEIRA_LINHA + ":A" + ULTIMA_LINHA,
    "Nova programação,Container retirado,Container depositado,Cancelado");
  aplicarLista(prog, "F" + PRIMEIRA_LINHA + ":F" + ULTIMA_LINHA,
    "Resfriada,Congelada");
  aplicarLista(prog, "U" + PRIMEIRA_LINHA + ":U" + ULTIMA_LINHA,
    "Pulmão Operador Logístico,Terminal do Armador,Direto");
  aplicarLista(prog, "O" + PRIMEIRA_LINHA + ":O" + ULTIMA_LINHA,
    "🔔");

  console.log("7/8  4 listas suspensas reaplicadas (A, F, U, O).");

  // ================================================================
  // ETAPA 8 — Proteção das colunas de fórmula
  // ================================================================
  if (APLICAR_PROTECAO) {
    const prot = prog.getProtection();
    prot.unprotect();

    // 1) libera tudo na faixa útil
    prog.getRange("A1:BG" + ULTIMA_LINHA).getFormat().getProtection().setLocked(false);

    // 2) trava o cabeçalho
    prog.getRange("A1:BG" + LINHA_CABECALHO).getFormat().getProtection().setLocked(true);

    // 3) trava as colunas de fórmula pura
    //    (coluna A fica DE FORA de propósito: precisa aceitar o override
    //     manual "Cancelado"/"Container depositado" pelo dropdown)
    const colunasTravadas = ["K", "R", "AC", "AD", "AE", "AL"];
    for (const c of colunasTravadas) {
      prog.getRange(c + PRIMEIRA_LINHA + ":" + c + ULTIMA_LINHA).getFormat().getProtection().setLocked(true);
    }

    prot.protect({
      allowAutoFilter: true,
      allowDeleteColumns: false,
      allowDeleteRows: true,
      allowEditObjects: true,
      allowEditScenarios: false,
      allowFormatCells: true,
      allowFormatColumns: true,
      allowFormatRows: true,
      allowInsertColumns: false,
      allowInsertHyperlinks: true,
      allowInsertRows: true,
      allowPivotTables: false,
      allowSort: true
    });

    console.log("8/8  Proteção aplicada: K, R, AC, AD, AE, AL + cabeçalho travados.");
    console.log("     (coluna A permanece editável para o override manual de STATUS)");
  } else {
    console.log("8/8  Proteção NÃO aplicada (APLICAR_PROTECAO = false).");
  }

  // ================================================================
  // RELATÓRIO FINAL
  // ================================================================
  console.log("");
  console.log("=== FASE 2 CONCLUÍDA ===");
  console.log("");
  console.log("CONFIRA AGORA:");
  console.log("  1. Cabeçalho vai de A até BG (59 colunas, sem COLETADO)");
  console.log("  2. Ordem: ... N UNIDADE | O Alterações | P REF. INTERNA ...");
  console.log("  3. Ordem: ... AH CONTAINER | AI DATA DE COLETA | AJ MAX GROSS ...");
  console.log("  4. Ordem: ... AO CAVALO/CARRETA | AP AUT. EMBARQUE | AQ DATA/HORA SAÍDA ...");
  console.log("  5. STATUS mudando conforme DATA DE COLETA (não mais pelo CONTAINER)");
  console.log("  6. ESTUFAGEM: digite um contrato em C11 e veja se todos os campos batem");
  console.log("  7. Tente digitar na coluna K -> deve bloquear");
  console.log("");
  console.log("⚠️ As CORES ainda estarão desalinhadas — isso é esperado. Rode a FASE 3.");
}


/**
 * Move uma coluna inteira de origem para a posição destino.
 * O destino é a letra da coluna ANTES da qual a coluna será inserida.
 */
function moverColuna(
  ws: ExcelScript.Worksheet,
  origem: string,
  destino: string,
  linhas: number
): void {
  // largura original, para preservar
  const largura = ws.getRange(origem + "1").getFormat().getColumnWidth();

  // 1) insere coluna em branco no destino (tudo desloca para a direita)
  ws.getRange(destino + ":" + destino).insert(ExcelScript.InsertShiftDirection.right);

  // 2) descobre onde a origem ficou depois da inserção
  const idxOrigem = letraParaIndice(origem);
  const idxDestino = letraParaIndice(destino);
  const origemAtual = idxOrigem >= idxDestino ? indiceParaLetra(idxOrigem + 1) : origem;

  // 3) copia conteúdo + formatos para a nova coluna
  const alvo = ws.getRange(destino + "1:" + destino + linhas);
  const fonte = ws.getRange(origemAtual + "1:" + origemAtual + linhas);
  alvo.copyFrom(fonte, ExcelScript.RangeCopyType.all, false, false);
  alvo.getFormat().setColumnWidth(largura);

  // 4) apaga a coluna de origem
  ws.getRange(origemAtual + ":" + origemAtual).delete(ExcelScript.DeleteShiftDirection.left);
}


/** Escreve uma fórmula em toda a faixa de uma coluna, montando linha a linha. */
function escreverColuna(
  ws: ExcelScript.Worksheet,
  coluna: string,
  primeiraLinha: number,
  qtdLinhas: number,
  montar: (linha: number) => string
): void {
  const dados: string[][] = [];
  for (let i = 0; i < qtdLinhas; i++) {
    dados.push([montar(primeiraLinha + i)]);
  }
  ws.getRange(coluna + primeiraLinha + ":" + coluna + (primeiraLinha + qtdLinhas - 1)).setFormulas(dados);
}


/** Aplica uma lista suspensa simples a uma faixa. */
function aplicarLista(ws: ExcelScript.Worksheet, faixa: string, itens: string): void {
  const rng = ws.getRange(faixa);
  rng.getDataValidation().clear();
  rng.getDataValidation().setRule({
    list: {
      inCellDropDown: true,
      source: itens
    }
  });
}


/** "A" -> 1, "AB" -> 28 */
function letraParaIndice(letra: string): number {
  let n = 0;
  const L = letra.toUpperCase();
  for (let i = 0; i < L.length; i++) {
    n = n * 26 + (L.charCodeAt(i) - 64);
  }
  return n;
}


/** 1 -> "A", 28 -> "AB" */
function indiceParaLetra(indice: number): string {
  let s = "";
  let n = indice;
  while (n > 0) {
    const resto = (n - 1) % 26;
    s = String.fromCharCode(65 + resto) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}
