/**
 * ITEM 1 — Alteração de colunas + migração de dados (carga única)
 * Projeto: Controle Cliente Exportador x Operador Logístico  |  Excel Online (Office Script)
 *
 * COMO USAR:
 *   1. Abra a CÓPIA DE TESTE no Excel Online
 *   2. Guia Automate -> New Script
 *   3. Apague o código de exemplo e cole ESTE arquivo inteiro
 *   4. Rode primeiro com MODO_VERIFICACAO = true  (não altera nada, só confere)
 *   5. Se o log disser "TUDO CERTO", troque para false e rode de novo
 *
 * SEGURANÇA:
 *   - Localiza as colunas pelo NOME do cabeçalho, não pela letra.
 *   - Se qualquer coluna esperada não existir, PARA sem alterar nada.
 *   - NÃO toca na aba DASHBOARD.
 *   - Ao colar a migração, célula vazia NÃO apaga dado existente.
 */

// ================================================================ CONFIGURAÇÃO

const MODO_VERIFICACAO = true;   // <<< troque para false para executar de verdade

const ABA_PROG = "PROGRAMAÇÃO";
const ABA_GERAL = "GERAL";

const LIN_CAB_PROG = 4;      // linha do cabeçalho na PROGRAMAÇÃO
const LIN_CAB_GERAL = 3;     // linha do cabeçalho na GERAL
const LIN_INI_PROG = 5;      // 1ª linha de dados da PROGRAMAÇÃO
const LIN_FIM_GERAL = 503;   // até onde vão as fórmulas da GERAL

const PROG_APAGAR = [
  "TIPO DE CONTRATO", "UNIDADE (TRANSPORTE)", "CARRETA", "NO REDEX?",
  "OBSERVAÇÕES FRIGO", "SEMANA REAL", "PULMÃO / DEPOT SAÍDA",
  "GATE ANTECIPADO?", "STATUS GERAL", "DATA ÚLTIMA ALTERAÇÃO"
];

const GERAL_APAGAR = [
  "TIPO DE CONTRATO", "SEMANA REAL", "UNIDADE", "CARRETA",
  "GATE ANTECIPADO?", "NO REDEX?", "OBSERVAÇÕES FRIGO", "STATUS GERAL"
];

const NOVAS = ["AUT. EMBARQUE", "Nº FORMULÁRIO", "Nº CTE VAZIO", "VALOR CTE VAZIO"];

const TXT_FRIGO = "🟢  FRIGO preenche — Container / Transporte / Gate / Docs";
const TXT_OPERADOR = "🔵  OPERADOR preenche — Controle / Status / Saving";

// Carga única: 199 linhas x 6 campos
// [CAVALO/CARRETA, POSIÇÃO, AUT. EMBARQUE, Nº FORMULÁRIO, Nº CTE VAZIO, VALOR CTE VAZIO]
const MIGRACAO: (string | number)[][] = [
  ["QTH8B47/FGV4B87", "CARREGADO", "OK", "", "", ""],
  ["UDO4H42/FSW5C97", "CARREGADO", "OK", "", "", ""],
  ["SVG4A59/FPG5J67", "CARREGADO", "OK", "", "", ""],
  ["EUZ4460/FUP9C58", "CARREGADO", "OK", "", "", ""],
  ["FPU8G58/FYJ6G02", "CARREGADO", 36513939, "", "", ""],
  ["FIN3E97/SSZ7E55", "CARREGADO", "OK", "", "", ""],
  ["TBP9J02/STV8E33", "CARREGADO", "", "", "", ""],
  ["SVZ1D45/STQ5F14", "CARREGADO", "OK", "", "", ""],
  ["UBD7J47/SSW6H63", "CARREGADO", "OK", "", "", ""],
  ["FXN1I27/EHW9D75", "CARREGADO", "OK", "", "", ""],
  ["UFJ5F80/FKY9F68", "CARREGADO", "OK", "", "", ""],
  ["UDO4H42/FSW5C97", "CARREGADO", "OK", "", "", ""],
  ["TKP3F56/TKT9F28", "CARREGADO", "OK", "", "", ""],
  ["QSU8C32/CUI7B08", "CARREGADO", "OK", "", "", ""],
  ["SHC4D29/SHC4D30", "CARREGADO", 36524643, "", "", ""],
  ["TKH8F07/GGW0E92", "CARREGADO", 36521304, "", "", ""],
  ["FUR8B41/RNZ5G47", "CARREGADO", "OK", "", "", ""],
  ["TJH1I41/STB3D79", "CARREGADO", "OK", "", "", ""],
  ["TKY0D78/SUA6C43", "CARREGADO", "", "", "", ""],
  ["TBQ1B28/SWH8I62", "CARREGADO", "OK", "", "", ""],
  ["FGX0C64/QST5C99", "CARREGADO", "OK", "", "", ""],
  ["TBU8I03/UGO8E27", "CARREGADO", "OK", "", "", ""],
  ["FCF6J22/UGR9H90", "CARREGADO", 36542931, "", "", ""],
  ["BDM8B68/FRO7C07", "CARREGADO", 36548291, 160, "", 6500],
  ["TBI8E56/FTB8F91", "CARREGADO", "", "", "", ""],
  ["SVX5G98/TJZ1C60", "CARREGADO", "OK", 161, 963, 6500],
  ["TBQ1B57/UGQ1A08", "CARREGADO", "", "", "", ""],
  ["SWS5D93FNO9F83", "CARREGADO", 36496102, "", "", ""],
  ["SWJ3B86/SYP9C22", "CARREGADO", "OK", "", "", ""],
  ["TCB2G64/TXR1B00", "CARREGADO", "OK", 159, 10209, 6500],
  ["SUI5H52/GID5H42", "CARREGADO", "", "", "", ""],
  ["TKO5F18/STB9G74", "CARREGADO", "OK", "", "", ""],
  ["EMA6I55/SWJ1A90", "CARREGADO", "OK", 158, 72, 6500],
  ["GCY5C62/TMG3A06", "CARREGADO", "OK", "", "", ""],
  ["BKU5E84/FWN4D56", "CARREGADO", "OK", "", "", ""],
  ["ESP2990/FOT3F78", "CARREGADO", "OK", 164, 100, 6500],
  ["ETC8G10/UFL8J80", "CARREGADO", "OK", "", "", ""],
  ["TKA1B85/QSQ0G11", "CARREGADO", 36570277, "", "", ""],
  ["TIP5D09/TKP7I80", "CARREGADO", "OK", "", "", ""],
  ["FNP6C93/EGN5B35", "CARREGADO", 36571763, "", "", ""],
  ["RXX8F76", "MATUPA/MT/BRASIL (VAZIO)", "", "", "", ""],
  ["ECU7D83/GCP3J17", "CARREGADO", "OK", "", "", ""],
  ["SSV6C59/FWP1H31", "CARREGADO", "OK", "", "", ""],
  ["TCK0I62/SYN7B08", "CARREGADO", "OK", "", "", ""],
  ["GDA6C54/GHE6F15", "CARREGADO", "OK", "", "", ""],
  ["TBQ3B09/QSS3F43", "CARREGADO", "OK", "", "", ""],
  ["TBL5D61/QSR8E14", "CARREGADO", 36574602, "", "", ""],
  ["TYH8C74/TYN1C86", "CARREGADO", 36562266, "", "", ""],
  ["RNT6D31/RNU4H68", "CARREGADO", 36563113, "", "", ""],
  ["FXJ0H21/GBM6E11", "CARREGADO", 36557292, "", "", ""],
  ["GBQ8D69/TLJ9E27", "CARREGADO", "OK", 165, 967, 6500],
  ["KAJ7H19/RQR6H41", "CARREGADO", 36589866, 166, 968, 6500],
  ["TMB9I57/TJY8B36", "CARREGADO", 36595975, 167, 973, 6500],
  ["UAX5D20/FLZ7F91", "CARREGADO", 36575120, "", "", ""],
  ["FJZ7C23", "CARREGADO AG M", "", "", "", ""],
  ["FPG7D92/FNL4C27", "CARREGADO", "OK", "", "", ""],
  ["FVP6J75/DFI6B91", "CARREGADO", "OK", "", "", ""],
  ["GGL5A74/BQU9C13", "CARREGADO AG LAUDO", "OK", "", "", ""],
  ["EWZ1C47/FOA7H57", "CARREGADO", "OK", "", "", ""],
  ["GAV8F19/EVP7G19", "CARREGADO", "OK", "", "", ""],
  ["GGY0I65/EMT7D87", "CARREGADO", 36563343, "", "", ""],
  ["FVZ3G62/FVC3F84", "CARREGADO", 36602064, 168, 974, 6500],
  ["UEB5E62/GBT2I87", "CARREGADO", 36601942, "", "", ""],
  ["FXO5G65/FRZ1F96", "ALTO GARÇAS/MT/BRASIL (CARREGADO)", "CHECKLIST", "", "", ""],
  ["UBD7J47/RXW1I57", "CARREGADO", 36617192, "", "", ""],
  ["TKG4H10/FRV3C12", "CARREGADO", 36612382, 169, 974, 6500],
  ["GBM9A21/END9I07", "CARREGADO", 36605439, "", "", ""],
  ["ETC0J22/UDC6A90", "CARREGADO", "OK", "", "", ""],
  ["SUG7E03/UGQ1F29", "CARREGADO", 36616972, "", "", ""],
  ["UAX5H15/UGJ0C35", "CARREGADO", 36605756, "", "", ""],
  ["GHZ8G14/EUJ1C37", "CARREGADO", "OK", "", "", ""],
  ["SYP4J70/EFO5J08", "CARREGADO", "OK", 170, 17310, 6500],
  ["EVW3F28/FOT7A24", "JI-PARANA/RO/BRASIL (VAZIO)", "OK", 171, "", 6500],
  ["", "", "", "", "", ""],
  ["GCD8A64/TMD4G24", "JI-PARANA/RO/BRASIL (VAZIO)", "OK", "", "", ""],
  ["GGW0E71/GGW0E91", "CARREGADO AG LAUDO", "OK", "", "", ""],
  ["QWP4E92/QWQ5I63", "CARREGADO", "OK", "", "", ""],
  ["RYQ3G83/RYY2D48", "CARREGADO", "OK", "", "", ""],
  ["QLY6D06/QLY4D37", "CARREGADO", 36587193, 163, "", 6500],
  ["TCB0C01/SYP9B93", "SINOP/MT/BRASIL (VAZIO)", "OK", "", "", ""],
  ["RYM9B46/RKW5F83", "CARREGADO", 36610220, "", "", ""],
  ["SXU0A37/RYP8A16", "CARREGADO", "OK", "", "", ""],
  ["SXU0B17/TPV4D26", "CARREGADO PARCIAL", "OK", "", "", ""],
  ["RYK6I62/RYP5G96", "CARREGADO", 36613128, "", "", ""],
  ["TBQ1B57/QSR1C23", "CARREGADO", 36617544, "", "", ""],
  ["TPU2I67/BCB9I93", "CARREGADO", 36575765, 162, 45949, 6500],
  ["RWA7G85/FPE3C84", "CARREGADO", 36598590, "", "", ""],
  ["STN5B06/ENZ7J98", "SINOP/MT/BRASIL (VAZIO)", "OK", "", "", ""],
  ["QSZ4A89/FCL1D77", "SINOP/MT/BRASIL (VAZIO)", "OK", "", "", ""],
  ["GHT0G81/GIF0F58", "SINOP/MT/BRASIL (VAZIO)", "OK", "", "", ""],
  ["SHC4D33/SHC9B69", "SINOP/MT/BRASIL (VAZIO)", "OK", "", "", ""],
  ["DMF7C32/SWK6F12", "SINOP/MT/BRASIL (VAZIO)", "OK", "", "", ""],
  ["FXO5G65/FRZ1F96", "ALTO GARÇAS/MT/BRASIL (CARREGADO)", "CHECKLIST", "", "", ""],
  ["GBA5F44/GDV8C87", "VARZEA GRANDE/MT/BRASIL (CARREGADO)", "OK", "", "", ""],
  ["", "", "", "", "", ""],
  ["SWA8E05", "SINOP/MT/BRASIL (CARREGADO)", "", "", "", ""],
  ["", "CACOAL/RO/BRASIL (CARREGADO)", "", "", "", ""],
  ["DAY1B89/CRO9A49", "VILHENA/RO/BRASIL (CARREGADO)", "OK", "", "", ""],
  ["TJH1I41/STB3D79", "CUIABA/MT/BRASIL (CARREGADO)", "OK", "", "", 6500],
  ["FPY5729/EVP7G97", "NOVA MUTUM/MT/BRASIL (CARREGADO)", "OK", "", "", ""],
  ["FUK7B10/GFS6F87", "RONDONOPOLIS/MT/BRASIL (CARREGADO)", "OK", "", "", ""],
  ["GBK6J35/FPL7J62", "SINOP/MT/BRASIL (VAZIO)", "OK", "", "", ""],
  ["EKA3J63/CUM2F19", "GUARULHOS/SP/BRASIL (CARREGADO", "OK", "", "", ""],
  ["FGB7J13/FMP9J32", "RONDONOPOLIS/MT/BRASIL (CARREGADO", "OK", "", "", ""],
  ["FGX0C64/QST5C99", "ANDRADINA/SP/BRASIL (PAPELAO)", "OK", "", "", ""],
  ["", "", "", "", "", ""],
  ["", "", "", "", "", ""],
  ["FCI9G56/BPQ5C21", "CAMPO GRANDE/MS/BRASIL (CARREGADO", "PESQUISA CNH", "", "", ""],
  ["BDM8B68/FRO7C07", "CAMPO GRANDE/MS/BRASIL (CARREGADO)", "OK", "", "", ""],
  ["SSW6H63", "MATUPA/MT/BRASIL (VAZIO)", "", "", "", ""],
  ["STL9H45", "MATUPA/MT/BRASIL (VAZIO)", "", "", "", ""],
  ["", "", "", "", "", ""],
  ["STV8E33", "CUIABA/MT/BRASIL (CARREGADO)", "", "", "", ""],
  ["STQ5F14", "VOTUPORANGA/SP/BRASIL (CARREGADO)", "", "", "", ""],
  ["", "", "", "", "", ""],
  ["", "", "", "", "", ""],
  ["UGO8E27", "MATUPA/MT/BRASIL (VAZIO)", "", "", "", ""],
  ["TBQ1B57", "MATUPA/MT/BRASIL (VAZIO)", "", "", "", ""],
  ["FNO9F83", "RONDONOPOLIS/MT/BRASIL (CARREGADO)", "", "", "", ""],
  ["FJH9J83", "MATUPA/MT/BRASIL (VAZIO)", "", "", "", ""],
  ["GID5H42", "MATUPA/MT/BRASIL (VAZIO)", "", "", "", ""],
  ["STB9G74", "ANDRADINA/SP/BRASIL (PAPELAO)", "", "", "", ""],
  ["TLM3D29/FHR6J85", "CUIABA/MT/BRASIL (CARREGADO)", "CHECKLIST", "", "", ""],
  ["UGV3F86", "DIAMANTINO/MT/BRASIL (VAZIO)", "", "", "", ""],
  ["QSR8E14", "ALTO TAQUARI/MT/BRASIL (CARREGADO)", "", "", "", ""],
  ["QSS3F43", "ANDRADINA/SP/BRASIL (PAPELAO)", "", "", "", ""],
  ["FTB8F91", "ALTO TAQUARI/MT/BRASIL (CARREGADO)", "", "", "", ""],
  ["FEA5D61/GDF4B71", "CUIABA/MT/BRASIL (CARREGADO)", "OK", "", "", ""],
  ["QSU8C32/CUI7B08", "SÃO JOSE DO RIO PRETO/SP/BRASIL (CARREGADO)", "OK", "", "", ""],
  ["GGR2F92/CCU0F78", "CACERES/MT/BRASIL (CARREGADO)", "OK", "", "", ""],
  ["UFI3A08/FPO8C08", "PEDRA PRETA/MT/BRASIL (CARREGADO)", "OK", "", "", ""],
  ["SVG4A59/FPG5J67", "CHAPADAO DO SUL/MS/BRASIL (CARREGADO)", "OK", "", "", ""],
  ["", "", "", "", "", ""],
  ["", "", "", "", "", ""],
  ["UFJ5F80/FKY9F68", "ALTO GARÇAS/MT/BRASIL (CARREGADO)", "OK", "", "", ""],
  ["UGD6F06/FWY4F28", "ALTO ARAGUAIA/MT/BRASIL (CARREGADO)", "CHECKLIST", "", "", ""],
  ["", "", "", "", "", ""],
  ["", "", "", "", "", 6500],
  ["", "", "", "", "", 6500],
  ["", "", "", "", "", ""],
  ["", "", "", "", "", ""],
  ["", "", "", "", "", ""],
  ["", "", "", "", "", ""],
  ["", "", "", "", "", ""],
  ["", "", "", "", "", ""],
  ["", "", "", "", "", ""],
  ["", "", "", "", "", ""],
  ["", "", "", "", "", ""],
  ["", "", "", "", "", ""],
  ["", "", "", "", "", ""],
  ["", "", "", "", "", 6500],
  ["", "", "", "", "", ""],
  ["", "", "", "", "", ""],
  ["", "", "", "", "", ""],
  ["", "", "", "", "", ""],
  ["", "", "", "", "", ""],
  ["", "", "", "", "", 6500],
  ["", "", "", "", "", ""],
  ["", "", "", "", "", ""],
  ["GGN2990/SWJ1A90", "", "", "", "", ""],
  ["", "", "", "", "", ""],
  ["", "", "", "", "", ""],
  ["", "", "", "", "", ""],
  ["", "", "", "", "", ""],
  ["", "", "", "", "", ""],
  ["", "", "", "", "", 6500],
  ["", "", "", "", "", ""],
  ["", "", "", "", "", ""],
  ["", "", "", "", "", ""],
  ["", "", "", "", "", ""],
  ["", "", "", "", "", ""],
  ["", "", "", "", "", 6500],
  ["", "", "", "", "", ""],
  ["", "", "", "", "", ""],
  ["", "", "", "", "", ""],
  ["", "", "", "", "", ""],
  ["", "", "", "", "", ""],
  ["", "", "", "", "", ""],
  ["", "", "", "", "", ""],
  ["", "", "", "", "", ""],
  ["", "", "", "", "", ""],
  ["", "", "", "", "", ""],
  ["", "", "", "", "", ""],
  ["", "SEM SINAL", "", "", "", ""],
  ["", "", "", "", "", ""],
  ["", "", "", "", "", ""],
  ["", "", "", "", "", ""],
  ["", "", "", "", "", ""],
  ["", "", "", "", "", ""],
  ["", "", "", "", "", ""],
  ["", "", "", "", "", ""],
  ["", "", "", "", "", ""],
  ["GHZ8G14/EUJ1C37", "CARREGADO", "OK", "", "", ""],
  ["", "", "", "", "", ""],
  ["", "", "", "", "", ""],
  ["", "", "", "", "", ""],
  ["", "", "", "", "", ""],
  ["", "", "", "", "", ""],
  ["", "", "", "", "", ""]
];

// ================================================================ PRINCIPAL

// buffer de log: acumula tudo e imprime UMA vez só (evita os avisos de performance)
let LOG: string[] = [];
function say(msg: string) { LOG.push(msg); }
function flush() { console.log(LOG.join("\n")); LOG = []; }

function main(workbook: ExcelScript.Workbook) {
  LOG = [];
  const prog = workbook.getWorksheet(ABA_PROG);
  const geral = workbook.getWorksheet(ABA_GERAL);

  if (!prog) throw new Error('Aba "' + ABA_PROG + '" não encontrada.');
  if (!geral) throw new Error('Aba "' + ABA_GERAL + '" não encontrada.');

  // ---------- VERIFICAÇÃO (sempre roda)
  const erros: string[] = [];

  let cabProg = lerCabecalhos(prog, LIN_CAB_PROG);
  let cabGeral = lerCabecalhos(geral, LIN_CAB_GERAL);

  say("== PROGRAMAÇÃO: " + cabProg.length + " colunas ==");
  for (const nome of PROG_APAGAR.concat(["CAVALO", "POSIÇÃO ATUAL", "LOCAL DE DEPÓSITO", "DOCUMENTAÇÃO", "CONTAINER", "COLETADO"])) {
    const c = achar(cabProg, nome);
    const real = c ? ' ["' + String(cabProg[c - 1]).replace(/\n/g, " ") + '"]' : "";
    say((c ? "  OK  " : "  !!  ") + nome + (c ? " -> col " + letra(c) + real : " -> NÃO ENCONTRADA"));
    if (!c) erros.push('PROGRAMAÇÃO: coluna "' + nome + '" não encontrada.');
  }

  say("== GERAL: " + cabGeral.length + " colunas ==");
  for (const nome of GERAL_APAGAR) {
    const c = achar(cabGeral, nome);
    say((c ? "  OK  " : "  !!  ") + nome + (c ? " -> col " + letra(c) : " -> NÃO ENCONTRADA"));
    if (!c) erros.push('GERAL: coluna "' + nome + '" não encontrada.');
  }

  // conta as linhas de dados (coluna B = PROFORMA)
  const colB = prog.getRange("B" + LIN_INI_PROG + ":B" + (LIN_INI_PROG + 600)).getValues();
  let nLinhas = 0;
  for (const l of colB) {
    if (l[0] === "" || l[0] === null) break;
    nLinhas++;
  }
  say("Linhas de dados na PROGRAMAÇÃO: " + nLinhas + " (migração espera " + MIGRACAO.length + ")");
  if (nLinhas !== MIGRACAO.length) {
    erros.push("Nº de linhas (" + nLinhas + ") diferente da migração (" + MIGRACAO.length +
      "). A planilha mudou desde a análise — avise antes de executar.");
  }

  if (erros.length > 0) {
    say("\n*** NÃO PODE EXECUTAR ***");
    for (const e of erros) say("  - " + e);
    flush();
    throw new Error("Verificação falhou. Nada foi alterado.");
  }

  say("\n*** VERIFICAÇÃO OK ***");
  if (MODO_VERIFICACAO) {
    say("MODO_VERIFICACAO = true -> nada foi alterado.");
    say("Troque para false no topo do script e execute de novo.");
    flush();
    return;
  }

  // ---------- ETAPA 1: apagar 10 colunas da PROGRAMAÇÃO (direita -> esquerda)
  let alvos = PROG_APAGAR.map(n => ({ nome: n, col: achar(cabProg, n) }));
  alvos.sort((a, b) => b.col - a.col);
  for (const a of alvos) {
    prog.getRange(letra(a.col) + ":" + letra(a.col)).delete(ExcelScript.DeleteShiftDirection.left);
    say("PROGRAMAÇÃO: apagada " + a.nome + " (col " + letra(a.col) + ")");
  }

  // ---------- ETAPA 2: renomear cabeçalhos
  cabProg = lerCabecalhos(prog, LIN_CAB_PROG);
  const cCav = achar(cabProg, "CAVALO");
  prog.getRange(letra(cCav) + LIN_CAB_PROG).setValue("CAVALO/CARRETA");
  const cPos = achar(cabProg, "POSIÇÃO ATUAL");
  prog.getRange(letra(cPos) + LIN_CAB_PROG).setValue("POSIÇÃO");
  say("PROGRAMAÇÃO: renomeadas CAVALO/CARRETA e POSIÇÃO");

  // ---------- ETAPA 3: inserir 4 colunas antes de LOCAL DE DEPÓSITO
  cabProg = lerCabecalhos(prog, LIN_CAB_PROG);
  const cLocDep = achar(cabProg, "LOCAL DE DEPÓSITO");
  const cDoc = achar(cabProg, "DOCUMENTAÇÃO");

  prog.getRange(letra(cLocDep) + ":" + letra(cLocDep + NOVAS.length - 1))
    .insert(ExcelScript.InsertShiftDirection.right);

  const alvoNovas = prog.getRange(letra(cLocDep) + LIN_CAB_PROG + ":" + letra(cLocDep + NOVAS.length - 1) + LIN_CAB_PROG);
  alvoNovas.copyFrom(prog.getRange(letra(cDoc) + LIN_CAB_PROG), ExcelScript.RangeCopyType.formats);
  alvoNovas.setValues([NOVAS]);
  for (let i = 0; i < NOVAS.length; i++) {
    prog.getRange(letra(cLocDep + i) + "1").getFormat().setColumnWidth(110);
  }
  say("PROGRAMAÇÃO: 4 colunas novas inseridas em " + letra(cLocDep) + ".." + letra(cLocDep + 3));

  // ---------- ETAPA 3b: refazer as faixas da linha 3
  cabProg = lerCabecalhos(prog, LIN_CAB_PROG);
  const cContainer = achar(cabProg, "CONTAINER");
  const cValorCte = achar(cabProg, "VALOR CTE VAZIO");
  const cLocDep2 = achar(cabProg, "LOCAL DE DEPÓSITO");
  const ultCol = cabProg.length;

  const faixaToda = prog.getRange(letra(cContainer) + "3:" + letra(ultCol) + "3");
  faixaToda.unmerge();
  faixaToda.clear(ExcelScript.ClearApplyTo.contents);

  const fFrigo = prog.getRange(letra(cContainer) + "3:" + letra(cValorCte) + "3");
  fFrigo.merge(false);
  prog.getRange(letra(cContainer) + "3").setValue(TXT_FRIGO);

  const fIT = prog.getRange(letra(cLocDep2) + "3:" + letra(ultCol) + "3");
  fIT.merge(false);
  prog.getRange(letra(cLocDep2) + "3").setValue(TXT_OPERADOR);
  say("PROGRAMAÇÃO: faixas da linha 3 refeitas (FRIGO até " + letra(cValorCte) +
    ", OPERADOR de " + letra(cLocDep2) + " a " + letra(ultCol) + ")");

  // ---------- ETAPA 4: limpar a GERAL
  cabGeral = lerCabecalhos(geral, LIN_CAB_GERAL);
  let alvosG = GERAL_APAGAR.map(n => ({ nome: n, col: achar(cabGeral, n) }));
  alvosG.sort((a, b) => b.col - a.col);
  for (const a of alvosG) {
    geral.getRange(letra(a.col) + ":" + letra(a.col)).delete(ExcelScript.DeleteShiftDirection.left);
    say("GERAL: apagada " + a.nome + " (col " + letra(a.col) + ")");
  }

  cabGeral = lerCabecalhos(geral, LIN_CAB_GERAL);
  const gCav = achar(cabGeral, "CAVALO");
  if (gCav) geral.getRange(letra(gCav) + LIN_CAB_GERAL).setValue("CAVALO/CARRETA");
  const gPos = achar(cabGeral, "POSIÇÃO ATUAL");
  if (gPos) geral.getRange(letra(gPos) + LIN_CAB_GERAL).setValue("POSIÇÃO");

  // corrige o #REF! pré-existente da coluna COLETADO da GERAL
  cabProg = lerCabecalhos(prog, LIN_CAB_PROG);
  cabGeral = lerCabecalhos(geral, LIN_CAB_GERAL);
  const gColetado = achar(cabGeral, "COLETADO");
  const lColetado = letra(achar(cabProg, "COLETADO"));
  if (gColetado) {
    geral.getRange(letra(gColetado) + "4:" + letra(gColetado) + LIN_FIM_GERAL)
      .setFormulas(formulasGeral(lColetado, 4, LIN_FIM_GERAL));
    say("GERAL: COLETADO corrigido (#REF! -> PROGRAMAÇÃO!" + lColetado + ")");
  }

  // ---------- ETAPA 4b: 4 colunas novas no fim da GERAL
  const gUlt = cabGeral.length;
  const alvoG = geral.getRange(letra(gUlt + 1) + LIN_CAB_GERAL + ":" + letra(gUlt + NOVAS.length) + LIN_CAB_GERAL);
  alvoG.copyFrom(geral.getRange(letra(gUlt) + LIN_CAB_GERAL), ExcelScript.RangeCopyType.formats);
  alvoG.setValues([NOVAS]);

  for (let k = 0; k < NOVAS.length; k++) {
    const lp = letra(achar(cabProg, NOVAS[k]));
    geral.getRange(letra(gUlt + 1 + k) + "4:" + letra(gUlt + 1 + k) + LIN_FIM_GERAL)
      .setFormulas(formulasGeral(lp, 4, LIN_FIM_GERAL));
    geral.getRange(letra(gUlt + 1 + k) + "1").getFormat().setColumnWidth(110);
    say('GERAL: coluna "' + NOVAS[k] + '" criada -> PROGRAMAÇÃO!' + lp);
  }

  // ---------- ETAPA 5: colar os dados da carga única
  const cCavF = achar(cabProg, "CAVALO/CARRETA");
  const cPosF = achar(cabProg, "POSIÇÃO");
  const cAutF = achar(cabProg, "AUT. EMBARQUE");

  colarPreservando(prog, LIN_INI_PROG, cCavF, MIGRACAO.map(d => [d[0]]));
  colarPreservando(prog, LIN_INI_PROG, cPosF, MIGRACAO.map(d => [d[1]]));
  colarPreservando(prog, LIN_INI_PROG, cAutF, MIGRACAO.map(d => [d[2], d[3], d[4], d[5]]));

  let comDados = 0;
  for (const d of MIGRACAO) if (d.join("") !== "") comDados++;
  say("MIGRAÇÃO: " + comDados + " de " + MIGRACAO.length + " linhas com algum dado");

  // ---------- RELATÓRIO
  const finalProg = lerCabecalhos(prog, LIN_CAB_PROG);
  const finalGeral = lerCabecalhos(geral, LIN_CAB_GERAL);
  say("\n===== RESULTADO =====");
  say("PROGRAMAÇÃO: " + finalProg.length + " colunas (esperado 58) — última: " + finalProg[finalProg.length - 1]);
  say("GERAL: " + finalGeral.length + " colunas (esperado 58) — última: " + finalGeral[finalGeral.length - 1]);
  say("Item 1 concluído.");
  flush();
}

// ================================================================ AUXILIARES

function lerCabecalhos(sh: ExcelScript.Worksheet, linha: number): string[] {
  const vals = sh.getRange("A" + linha + ":CZ" + linha).getValues()[0];
  const out: string[] = [];
  for (const v of vals) {
    if (v === "" || v === null) break;
    out.push(String(v));
  }
  return out;
}

/** normaliza: tira espaços, quebras de linha e emojis */
function norm(s: string): string {
  return String(s == null ? "" : s)
    .replace(/[\u{1F300}-\u{1FAFF}\u{2600}-\u{27BF}\u{FE0F}]/gu, "")
    .replace(/\s+/g, "")
    .toUpperCase();
}

/** índice 1-based da coluna com aquele cabeçalho; 0 se não achar.
 *  1º tenta igualdade exata; se não achar, aceita "começa com" — mas só se
 *  houver UM único candidato (evita confundir UNIDADE com UNIDADE DETALHADA). */
function achar(cabs: string[], nome: string): number {
  const alvo = norm(nome);
  for (let i = 0; i < cabs.length; i++) {
    if (norm(cabs[i]) === alvo) return i + 1;
  }
  const candidatos: number[] = [];
  for (let i = 0; i < cabs.length; i++) {
    if (norm(cabs[i]).indexOf(alvo) === 0) candidatos.push(i + 1);
  }
  return candidatos.length === 1 ? candidatos[0] : 0;
}

/** cola sem apagar o que já existe (vazio na migração = mantém o atual) */
function colarPreservando(sh: ExcelScript.Worksheet, linIni: number, col: number, dados: (string | number)[][]) {
  const nLin = dados.length;
  const nCol = dados[0].length;
  const end = letra(col + nCol - 1) + (linIni + nLin - 1);
  const rng = sh.getRange(letra(col) + linIni + ":" + end);
  const atual = rng.getValues();
  for (let i = 0; i < nLin; i++) {
    for (let j = 0; j < nCol; j++) {
      if (dados[i][j] === "" || dados[i][j] === null) dados[i][j] = atual[i][j] as (string | number);
    }
  }
  rng.setValues(dados);
}

/** fórmula padrão da GERAL (busca pela PROFORMA na coluna C) */
function formulasGeral(letraProg: string, ini: number, fim: number): string[][] {
  const out: string[][] = [];
  for (let r = ini; r <= fim; r++) {
    const x = 'XLOOKUP($C' + r + ',PROGRAMAÇÃO!B:B,PROGRAMAÇÃO!' + letraProg + ':' + letraProg + ',"")';
    out.push(['=IFERROR(IF(' + x + '="","",' + x + '),"")']);
  }
  return out;
}

/** 1 => A, 27 => AA */
function letra(n: number): string {
  let s = "";
  while (n > 0) {
    const m = (n - 1) % 26;
    s = String.fromCharCode(65 + m) + s;
    n = Math.floor((n - m) / 26);
  }
  return s;
}
