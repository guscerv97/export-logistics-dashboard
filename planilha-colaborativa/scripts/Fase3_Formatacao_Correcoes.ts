/**
 * FASE 3 — Formatação, cabeçalho e correções finais
 * Cliente Exportador × Operador Logístico
 *
 * O QUE FAZ:
 *   1. Desprotege a aba (a Fase 2 deixou protegida)
 *   2. Congela a coluna K (SEMANA) como valor fixo e a destrava para digitação
 *   3. Restaura as mesclagens das linhas 1 e 2
 *   4. Refaz as 4 tarjas da linha 3 alinhadas ao novo layout
 *   5. Reconstrói TODA a formatação condicional da PROGRAMAÇÃO de forma unificada
 *      (fim da fragmentação em blocos + conserta o sino, que estava #REF!)
 *   6. Reconstrói a tabela "Próximas estufagens" do DASHBOARD (A12:D21)
 *   7. Reaplica a proteção — agora SEM a coluna K
 *
 * CORES: todas extraídas do arquivo original, nenhuma inventada.
 *
 * COMO USAR:
 *   1) MODO_VERIFICACAO = true  -> confere o layout e imprime o plano
 *   2) Troque para false e rode de novo
 */

function main(workbook: ExcelScript.Workbook) {

  // ====== CONFIGURAÇÃO ======
  const MODO_VERIFICACAO = true;
  const APLICAR_PROTECAO = true;
  const PRIMEIRA_LINHA = 5;
  const ULTIMA_LINHA = 517;
  const ULTIMA_COLUNA = "BG";
  // ==========================

  // ====== CORES (extraídas do arquivo original) ======
  const VERDE_FRIGO = "#1E5631";
  const AZUL_OPERADOR = "#1F4E79";
  const BRANCO = "#FFFFFF";

  const CF_SINO = "#FFEB00";
  const CF_CANCELADO = "#FF0000";
  const CF_RETIRADO = "#FFEB9C";
  const CF_DEPOSITADO = "#C6EFCE";
  const CF_NOVA = "#DDEBF7";
  const CF_RESFRIADA_FUNDO = "#FF99CC";
  const CF_RESFRIADA_FONTE = "#880044";
  const CF_DET_ATENCAO_FUNDO = "#FFEB9C";
  const CF_DET_ATENCAO_FONTE = "#9C6500";
  const CF_DET_OK_FUNDO = "#C6EFCE";
  const CF_DET_OK_FONTE = "#006100";
  const CF_DET_CRITICO_FUNDO = "#FFC7CE";
  const CF_DET_CRITICO_FONTE = "#9C0006";
  const CF_DUPLICADO = "#FF878B";
  const CF_AG_PUXAR = "#FF00E1";
  const CF_AG_APROVACAO = "#FF0000";
  // ===================================================

  const prog = workbook.getWorksheet("PROGRAMAÇÃO");
  const dash = workbook.getWorksheet("DASHBOARD");
  if (!prog) { console.log("ERRO: aba PROGRAMAÇÃO não encontrada."); return; }

  console.log("=== FASE 3 — FORMATAÇÃO E CORREÇÕES ===");
  console.log("Modo: " + (MODO_VERIFICACAO ? "VERIFICAÇÃO" : "EXECUÇÃO"));
  console.log("");

  // ================================================================
  // PRÉ-VOO
  // ================================================================
  const cab = prog.getRange("A4:" + ULTIMA_COLUNA + "4").getValues()[0];
  const checagens: [number, string][] = [
    [15, "ALTERA"],      // O
    [16, "REF. INTERNA"], // P
    [34, "CONTAINER"],   // AH
    [35, "COLETA"],      // AI
    [42, "AUT. EMBARQUE"] // AP
  ];

  const erros: string[] = [];
  for (const c of checagens) {
    const titulo = String(cab[c[0] - 1] || "").toUpperCase().replace(/\n/g, " ");
    if (titulo.indexOf(c[1]) === -1) {
      erros.push("  coluna " + c[0] + " deveria conter \"" + c[1] + "\" e contém \"" + titulo + "\"");
    }
  }

  if (erros.length > 0) {
    console.log("❌ LAYOUT INESPERADO — a Fase 2 rodou? Script ABORTADO:");
    console.log(erros.join("\n"));
    return;
  }
  console.log("--- PRÉ-VOO: ✅ layout da Fase 2 confirmado (59 colunas, A→BG) ---");
  console.log("");

  if (MODO_VERIFICACAO) {
    console.log("--- PLANO (nada será executado agora) ---");
    console.log("  1. Congelar coluna K como valor fixo + destravar para digitação");
    console.log("  2. Remesclar A1:" + ULTIMA_COLUNA + "1 e A2:" + ULTIMA_COLUNA + "2");
    console.log("  3. Refazer tarjas da linha 3:");
    console.log("       B3:N3    🟢 FRIGO preenche");
    console.log("       O3:AG3   🔵 OPERADOR preenche");
    console.log("       AH3:BA3  🟢 FRIGO preenche — Container / Transporte / Gate / Docs");
    console.log("       BB3:BG3  🔵 OPERADOR preenche — Controle / Status / Saving");
    console.log("  4. Reconstruir formatação condicional (12 regras, faixas unificadas até " + ULTIMA_LINHA + "):");
    console.log("       prio 1  sino          A5:" + ULTIMA_COLUNA + " -> LEN($O5)>0        [CORRIGE O #REF!]");
    console.log("       prio 2-5 status       A5:" + ULTIMA_COLUNA);
    console.log("       prio 6  resfriada     F5:F");
    console.log("       prio 7-9 detention    AE5:AE");
    console.log("       prio 10-12 motorista  AN5:AN + duplicados H");
    console.log("  5. Reconstruir DASHBOARD A12:D21 (próximas estufagens)");
    console.log("  6. Proteger: R, AC, AD, AE, AL + cabeçalho  (K fica LIVRE agora)");
    console.log("");
    console.log("=== MODO VERIFICAÇÃO — nada foi alterado. ===");
    return;
  }

  // ================================================================
  // 1 — Desproteger
  // ================================================================
  prog.getProtection().unprotect();
  console.log("1/7  Aba desprotegida.");

  // ================================================================
  // 2 — Congelar coluna K como valor fixo
  // ================================================================
  const rngK = prog.getRange("K" + PRIMEIRA_LINHA + ":K" + ULTIMA_LINHA);
  const valoresK = rngK.getValues();   // lê o resultado calculado
  rngK.setValues(valoresK);            // regrava como valor puro
  rngK.setNumberFormat("0");
  console.log("2/7  Coluna K congelada como valor fixo (" + valoresK.length + " células).");

  // ================================================================
  // 3 — Mesclagens das linhas 1 e 2 + tarjas da linha 3
  // ================================================================
  prog.getRange("A1:" + ULTIMA_COLUNA + "3").unmerge();

  prog.getRange("A1:" + ULTIMA_COLUNA + "1").merge(true);
  prog.getRange("A2:" + ULTIMA_COLUNA + "2").merge(true);

  // limpa a linha 3 inteira antes de refazer
  const linha3 = prog.getRange("A3:" + ULTIMA_COLUNA + "3");
  linha3.clear(ExcelScript.ClearApplyTo.contents);

  criarTarja(prog, "B3:N3", "🟢  FRIGO preenche", VERDE_FRIGO, BRANCO);
  criarTarja(prog, "O3:AG3", "🔵  OPERADOR preenche", AZUL_OPERADOR, BRANCO);
  criarTarja(prog, "AH3:BA3", "🟢  FRIGO preenche — Container / Transporte / Gate / Docs", VERDE_FRIGO, BRANCO);
  criarTarja(prog, "BB3:BG3", "🔵  OPERADOR preenche — Controle / Status / Saving", AZUL_OPERADOR, BRANCO);

  console.log("3/7  Mesclagens restauradas e 4 tarjas refeitas.");

  // ================================================================
  // 4 — Formatação condicional unificada
  // ================================================================
  const faixaToda = "A" + PRIMEIRA_LINHA + ":" + ULTIMA_COLUNA + ULTIMA_LINHA;

  // limpa tudo que existe hoje
  prog.getRange("A1:" + ULTIMA_COLUNA + "1048576").clearAllConditionalFormats();

  let prioridade = 0;

  // --- SINO: linha inteira em amarelo (prioridade máxima) ---
  addRegra(prog, faixaToda, '=LEN($O' + PRIMEIRA_LINHA + ')>0', CF_SINO, null, false, prioridade++);

  // --- STATUS: 4 regras na linha inteira ---
  addRegra(prog, faixaToda, '=$A' + PRIMEIRA_LINHA + '="Cancelado"', CF_CANCELADO, null, false, prioridade++);
  addRegra(prog, faixaToda, '=$A' + PRIMEIRA_LINHA + '="Container retirado"', CF_RETIRADO, null, false, prioridade++);
  addRegra(prog, faixaToda, '=$A' + PRIMEIRA_LINHA + '="Container depositado"', CF_DEPOSITADO, null, false, prioridade++);
  addRegra(prog, faixaToda, '=$A' + PRIMEIRA_LINHA + '="Nova programação"', CF_NOVA, null, false, prioridade++);

  // --- CARGA RESFRIADA (coluna F) ---
  const faixaF = "F" + PRIMEIRA_LINHA + ":F" + ULTIMA_LINHA;
  addRegra(prog, faixaF, '=$F' + PRIMEIRA_LINHA + '="Resfriada"',
    CF_RESFRIADA_FUNDO, CF_RESFRIADA_FONTE, true, prioridade++);

  // --- ALERTA DETENTION (coluna AE) ---
  const faixaAE = "AE" + PRIMEIRA_LINHA + ":AE" + ULTIMA_LINHA;
  addRegra(prog, faixaAE, '=AND($AE' + PRIMEIRA_LINHA + '<>"",$AE' + PRIMEIRA_LINHA + '<=0)',
    CF_DET_CRITICO_FUNDO, CF_DET_CRITICO_FONTE, false, prioridade++);
  addRegra(prog, faixaAE, '=AND($AE' + PRIMEIRA_LINHA + '<>"",$AE' + PRIMEIRA_LINHA + '>=1,$AE' + PRIMEIRA_LINHA + '<=3)',
    CF_DET_ATENCAO_FUNDO, CF_DET_ATENCAO_FONTE, false, prioridade++);
  addRegra(prog, faixaAE, '=AND($AE' + PRIMEIRA_LINHA + '<>"",$AE' + PRIMEIRA_LINHA + '>3)',
    CF_DET_OK_FUNDO, CF_DET_OK_FONTE, false, prioridade++);

  // --- MOTORISTA (coluna AN): avisos de pendência ---
  const faixaAN = "AN" + PRIMEIRA_LINHA + ":AN" + ULTIMA_LINHA;
  addRegra(prog, faixaAN, '=NOT(ISERROR(SEARCH("AG PUXAR",AN' + PRIMEIRA_LINHA + ')))',
    null, CF_AG_PUXAR, false, prioridade++);
  addRegra(prog, faixaAN, '=NOT(ISERROR(SEARCH("AG RETORNO",AN' + PRIMEIRA_LINHA + ')))',
    null, CF_AG_PUXAR, true, prioridade++);
  addRegra(prog, faixaAN, '=NOT(ISERROR(SEARCH("AG APROVAÇÃO",AN' + PRIMEIRA_LINHA + ')))',
    null, CF_AG_APROVACAO, true, prioridade++);

  // --- REF. CLIENTE duplicada (coluna H) ---
  const cfDup = prog.getRange("H" + PRIMEIRA_LINHA + ":H" + ULTIMA_LINHA)
    .addConditionalFormat(ExcelScript.ConditionalFormatType.presetCriteria);
  cfDup.getPreset().getFormat().getFill().setColor(CF_DUPLICADO);
  cfDup.getPreset().setRule({
    criterion: ExcelScript.ConditionalFormatPresetCriterion.duplicateValues
  });

  console.log("4/7  Formatação condicional reconstruída: 12 regras, faixas unificadas até a linha " + ULTIMA_LINHA + ".");
  console.log("     Sino corrigido: LEN($O5)>0 (estava LEN(#REF!)>0).");

  // ================================================================
  // 5 — DASHBOARD: tabela "Próximas estufagens"
  // ================================================================
  if (dash) {
    dash.getRange("A12:D21").clear(ExcelScript.ClearApplyTo.contents);
    dash.getRange("A12").setFormula(
      '=IFERROR(TAKE(SORT(FILTER(HSTACK(' +
      'PROGRAMAÇÃO!P' + PRIMEIRA_LINHA + ':P' + ULTIMA_LINHA + ',' +
      'PROGRAMAÇÃO!C' + PRIMEIRA_LINHA + ':C' + ULTIMA_LINHA + ',' +
      'PROGRAMAÇÃO!L' + PRIMEIRA_LINHA + ':L' + ULTIMA_LINHA + ',' +
      'PROGRAMAÇÃO!A' + PRIMEIRA_LINHA + ':A' + ULTIMA_LINHA + '),' +
      '(PROGRAMAÇÃO!L' + PRIMEIRA_LINHA + ':L' + ULTIMA_LINHA + '>=TODAY())*' +
      '(PROGRAMAÇÃO!L' + PRIMEIRA_LINHA + ':L' + ULTIMA_LINHA + '<=TODAY()+7)*' +
      '(PROGRAMAÇÃO!A' + PRIMEIRA_LINHA + ':A' + ULTIMA_LINHA + '<>"Cancelado")' +
      '),3,1),10),"")'
    );
    dash.getRange("C12:C21").setNumberFormat("dd/mm/yyyy");
    console.log("5/7  DASHBOARD A12:D21 reconstruído (10 próximas estufagens, ordenadas por data).");
  } else {
    console.log("5/7  Aba DASHBOARD não encontrada — etapa pulada.");
  }

  // ================================================================
  // 6 — Ajustes finais de formato
  // ================================================================
  prog.getRange("AC" + PRIMEIRA_LINHA + ":AD" + ULTIMA_LINHA).setNumberFormat("dd/mm/yyyy");
  prog.getRange("AE" + PRIMEIRA_LINHA + ":AE" + ULTIMA_LINHA).setNumberFormat("0");
  console.log("6/7  Formatos numéricos reaplicados.");

  // ================================================================
  // 7 — Proteção (sem a coluna K)
  // ================================================================
  if (APLICAR_PROTECAO) {
    const prot = prog.getProtection();
    prot.unprotect();

    prog.getRange("A1:" + ULTIMA_COLUNA + ULTIMA_LINHA).getFormat().getProtection().setLocked(false);
    prog.getRange("A1:" + ULTIMA_COLUNA + "4").getFormat().getProtection().setLocked(true);

    // K saiu da lista: agora é digitação manual
    const travadas = ["R", "AC", "AD", "AE", "AL"];
    for (const c of travadas) {
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
    console.log("7/7  Proteção reaplicada: R, AC, AD, AE, AL + cabeçalho. K e A livres.");
  } else {
    console.log("7/7  Proteção NÃO aplicada.");
  }

  console.log("");
  console.log("=== FASE 3 CONCLUÍDA ===");
  console.log("");
  console.log("CONFIRA AGORA:");
  console.log("  1. Linhas 1 e 2 mescladas de A até " + ULTIMA_COLUNA);
  console.log("  2. Linha 3 com as 4 tarjas verde/azul alinhadas aos blocos");
  console.log("  3. Marque um sino na coluna O -> a LINHA INTEIRA deve ficar amarela");
  console.log("  4. Cores de STATUS voltando ao normal nas demais linhas");
  console.log("  5. Coluna AE (detention) com vermelho/amarelo/verde");
  console.log("  6. Coluna K editável; colunas R, AC, AD, AE, AL bloqueadas");
  console.log("  7. DASHBOARD linha 12 em diante listando as próximas estufagens");
}


/** Cria uma tarja mesclada e colorida na linha 3. */
function criarTarja(
  ws: ExcelScript.Worksheet,
  faixa: string,
  texto: string,
  fundo: string,
  fonte: string
): void {
  const r = ws.getRange(faixa);
  r.merge(true);
  r.setValue(texto);
  const f = r.getFormat();
  f.getFill().setColor(fundo);
  f.getFont().setColor(fonte);
  f.getFont().setBold(true);
  f.getFont().setSize(9);
  f.getFont().setName("Arial");
  f.setHorizontalAlignment(ExcelScript.HorizontalAlignment.center);
  f.setVerticalAlignment(ExcelScript.VerticalAlignment.center);
}


/** Adiciona uma regra condicional de fórmula personalizada. */
function addRegra(
  ws: ExcelScript.Worksheet,
  faixa: string,
  formula: string,
  fundo: string | null,
  fonte: string | null,
  negrito: boolean,
  prioridade: number
): void {
  const cf = ws.getRange(faixa).addConditionalFormat(ExcelScript.ConditionalFormatType.custom);
  const custom = cf.getCustom();
  custom.getRule().setFormula(formula);

  const fmt = custom.getFormat();
  if (fundo !== null) {
    fmt.getFill().setColor(fundo);
  }
  if (fonte !== null) {
    fmt.getFont().setColor(fonte);
  }
  if (negrito) {
    fmt.getFont().setBold(true);
  }
  cf.setPriority(prioridade);
}
