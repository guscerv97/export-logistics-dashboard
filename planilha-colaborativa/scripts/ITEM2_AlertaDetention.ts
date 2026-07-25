/**
 * ITEM 2 — ALERTA DETENTION (contagem regressiva do Free Time)
 * Projeto: Controle Cliente Exportador x Operador Logístico  |  Excel Online (Office Script)
 *
 * O QUE FAZ:
 *   1. Substitui a fórmula da coluna ALERTA DETENTION (AE5:AE517)
 *   2. Substitui a formatação condicional DESSA COLUNA por 3 faixas (🔴 🟡 🟢)
 *
 * REGRA:
 *   Dias restantes = DATA LIMITE DEPÓSITO - HOJE
 *   (DATA LIMITE DEPÓSITO já é DATA DE COLETA + FREE TIME - 1, ou seja,
 *    o último dia antes de estourar. O dia da coleta conta como dia 1.)
 *   - Depositado  -> congela em (LIMITE - DATA DE DEPÓSITO); sem data -> vazio
 *   - Cancelado   -> vazio
 *   - Sem coleta / sem free time -> vazio
 *
 * SEGURANÇA:
 *   - Localiza as colunas pelo NOME do cabeçalho.
 *   - Só apaga regras de formatação cuja faixa esteja INTEIRAMENTE dentro da
 *     coluna ALERTA DETENTION. A regra de cor por STATUS (que pinta a linha
 *     toda) NÃO é tocada.
 *   - MODO_VERIFICACAO = true não altera nada.
 */

const MODO_VERIFICACAO = true;   // <<< troque para false para executar de verdade

const ABA_PROG = "PROGRAMAÇÃO";
const LIN_CAB = 4;
const LIN_INI = 5;
const LIN_FIM = 517;   // mesma faixa das demais fórmulas/validações da aba

// cores das 3 faixas
const VERMELHO_FUNDO = "#FFC7CE", VERMELHO_FONTE = "#9C0006";
const AMARELO_FUNDO = "#FFEB9C", AMARELO_FONTE = "#9C6500";
const VERDE_FUNDO = "#C6EFCE", VERDE_FONTE = "#006100";

let LOG: string[] = [];
function say(m: string) { LOG.push(m); }
function flush() { console.log(LOG.join("\n")); LOG = []; }

function main(workbook: ExcelScript.Workbook) {
  LOG = [];
  const sh = workbook.getWorksheet(ABA_PROG);
  if (!sh) throw new Error('Aba "' + ABA_PROG + '" não encontrada.');

  const cabs = lerCabecalhos(sh, LIN_CAB);
  const erros: string[] = [];

  const cStatus = achar(cabs, "STATUS");
  const cLimite = achar(cabs, "DATA LIMITE DEPÓSITO");
  const cAlerta = achar(cabs, "ALERTA DETENTION");
  const cDeposit = achar(cabs, "DEPOSITADO");
  const cDataDep = achar(cabs, "DATA DE DEPÓSITO");

  const checar = [
    { n: "STATUS", c: cStatus }, { n: "DATA LIMITE DEPÓSITO", c: cLimite },
    { n: "ALERTA DETENTION", c: cAlerta }, { n: "DEPOSITADO", c: cDeposit },
    { n: "DATA DE DEPÓSITO", c: cDataDep }
  ];
  for (const x of checar) {
    const real = x.c ? ' ["' + String(cabs[x.c - 1]).replace(/\n/g, " ") + '"]' : "";
    say((x.c ? "  OK  " : "  !!  ") + x.n + (x.c ? " -> col " + letra(x.c) + real : " -> NÃO ENCONTRADA"));
    if (!x.c) erros.push('Coluna "' + x.n + '" não encontrada.');
  }
  if (cabs.length !== 58) {
    say("  !!  A aba tem " + cabs.length + " colunas (esperado 58 — o item 1 foi aplicado?)");
    erros.push("Nº de colunas inesperado: " + cabs.length);
  }

  if (erros.length > 0) {
    say("\n*** NÃO PODE EXECUTAR ***");
    for (const e of erros) say("  - " + e);
    flush();
    throw new Error("Verificação falhou. Nada foi alterado.");
  }

  const lA = letra(cStatus), lLim = letra(cLimite), lAl = letra(cAlerta);
  const lDep = letra(cDeposit), lDta = letra(cDataDep);
  const faixa = lAl + LIN_INI + ":" + lAl + LIN_FIM;

  say("\nFórmula que será aplicada em " + faixa + ":");
  say('=IF(OR($' + lLim + LIN_INI + '="",$' + lA + LIN_INI + '="Cancelado"),"",' +
      'IF($' + lDep + LIN_INI + '<>"",IF($' + lDta + LIN_INI + '="","",$' + lLim + LIN_INI + '-$' + lDta + LIN_INI + '),' +
      '$' + lLim + LIN_INI + '-TODAY()))');

  // ---------- inventário da formatação condicional da coluna
  const rngAlerta = sh.getRange(faixa);
  const cfs = rngAlerta.getConditionalFormats();
  const apagar: ExcelScript.ConditionalFormat[] = [];
  const manter: string[] = [];

  for (const cf of cfs) {
    const addr = cf.getRanges().getAddress();
    if (soDaColuna(addr, lAl)) apagar.push(cf);
    else manter.push(addr);
  }
  say("\nFormatação condicional que toca a coluna " + lAl + ":");
  say("  - " + apagar.length + " regra(s) exclusiva(s) da coluna -> serão SUBSTITUÍDAS");
  say("  - " + manter.length + " regra(s) que pegam outras colunas -> PRESERVADAS");
  for (const m of manter) say("      preservada: " + m);

  if (MODO_VERIFICACAO) {
    say("\n*** VERIFICAÇÃO OK — MODO_VERIFICACAO = true, nada foi alterado. ***");
    say("Troque para false no topo do script e execute de novo.");
    flush();
    return;
  }

  // ---------- 1) fórmula
  const fs: string[][] = [];
  for (let r = LIN_INI; r <= LIN_FIM; r++) {
    fs.push(['=IF(OR($' + lLim + r + '="",$' + lA + r + '="Cancelado"),"",' +
             'IF($' + lDep + r + '<>"",IF($' + lDta + r + '="","",$' + lLim + r + '-$' + lDta + r + '),' +
             '$' + lLim + r + '-TODAY()))']);
  }
  rngAlerta.setFormulas(fs);
  rngAlerta.setNumberFormatLocal("0");
  say("\nFórmula aplicada em " + (LIN_FIM - LIN_INI + 1) + " linhas.");

  // ---------- 2) formatação condicional
  for (const cf of apagar) cf.delete();
  say("Regras antigas da coluna removidas: " + apagar.length);

  criarRegra(rngAlerta, '=AND($' + lAl + LIN_INI + '<>"",$' + lAl + LIN_INI + '>3)', VERDE_FUNDO, VERDE_FONTE);
  criarRegra(rngAlerta, '=AND($' + lAl + LIN_INI + '<>"",$' + lAl + LIN_INI + '>=1,$' + lAl + LIN_INI + '<=3)', AMARELO_FUNDO, AMARELO_FONTE);
  criarRegra(rngAlerta, '=AND($' + lAl + LIN_INI + '<>"",$' + lAl + LIN_INI + '<=0)', VERMELHO_FUNDO, VERMELHO_FONTE);
  say("Novas regras criadas: 🔴 <=0  |  🟡 1 a 3  |  🟢 >3");

  say("\n===== ITEM 2 CONCLUÍDO =====");
  flush();
}

/** cria uma regra de fórmula personalizada na prioridade mais alta */
function criarRegra(rng: ExcelScript.Range, formula: string, fundo: string, fonte: string) {
  const cf = rng.addConditionalFormat(ExcelScript.ConditionalFormatType.custom);
  const custom = cf.getCustom();
  custom.getRule().setFormula(formula);
  custom.getFormat().getFill().setColor(fundo);
  custom.getFormat().getFont().setColor(fonte);
  cf.setPriority(0);
}

/** true se TODAS as áreas do endereço estiverem dentro da coluna informada */
function soDaColuna(endereco: string, col: string): boolean {
  const semAba = endereco.indexOf("!") >= 0 ? endereco.split("!")[1] : endereco;
  const areas = semAba.split(",");
  for (const a of areas) {
    const partes = a.replace(/\$/g, "").split(":");
    for (const p of partes) {
      const letras = p.replace(/[0-9]/g, "").toUpperCase();
      if (letras !== col) return false;
    }
  }
  return true;
}

function lerCabecalhos(sh: ExcelScript.Worksheet, linha: number): string[] {
  const vals = sh.getRange("A" + linha + ":CZ" + linha).getValues()[0];
  const out: string[] = [];
  for (const v of vals) {
    if (v === "" || v === null) break;
    out.push(String(v));
  }
  return out;
}

function norm(s: string): string {
  return String(s == null ? "" : s)
    .replace(/[\u{1F300}-\u{1FAFF}\u{2600}-\u{27BF}\u{FE0F}]/gu, "")
    .replace(/\s+/g, "")
    .toUpperCase();
}

function achar(cabs: string[], nome: string): number {
  const alvo = norm(nome);
  for (let i = 0; i < cabs.length; i++) if (norm(cabs[i]) === alvo) return i + 1;
  const cand: number[] = [];
  for (let i = 0; i < cabs.length; i++) if (norm(cabs[i]).indexOf(alvo) === 0) cand.push(i + 1);
  return cand.length === 1 ? cand[0] : 0;
}

function letra(n: number): string {
  let s = "";
  while (n > 0) {
    const m = (n - 1) % 26;
    s = String.fromCharCode(65 + m) + s;
    n = Math.floor((n - m) / 26);
  }
  return s;
}
