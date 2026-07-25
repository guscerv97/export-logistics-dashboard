/**
 * ITEM 4 (v2) — Aba DASHBOARD COLETA com FILTRO POR SEMANA
 * Projeto: Controle Cliente Exportador x Operador Logístico  |  Excel Online (Office Script)
 *
 * NOVIDADE DESTA VERSÃO:
 *   Célula de filtro (B3) com lista suspensa: TODAS ou uma semana específica.
 *   - Cartões, PLANTA e PAÍS -> filtram pela semana escolhida
 *   - Gráfico por semana     -> NÃO filtra (fica como contexto/comparação)
 *   - A semana de referência do filtro é a SEMANA PROGRAMADA (col. K),
 *     única semana que toda linha possui.
 *
 * INDICADORES:
 *   - Cartões: Retirados | Pendentes | Depositados | Cancelados | Total ativo
 *   - Retirados x Pendentes por SEMANA (janela móvel de 14 semanas)
 *        Retirados -> semana da DATA DE COLETA
 *        Pendentes -> SEMANA PROGRAMADA
 *   - Carregamentos por PLANTA / por PAÍS
 *   Cancelados ficam fora de todas as contagens (exceto o cartão próprio).
 *
 * SEGURANÇA:
 *   - Só cria/recria a aba DASHBOARD COLETA.
 *   - NÃO toca em PROGRAMAÇÃO, GERAL, DASHBOARD ou ESTUFAGEM.
 *   - MODO_VERIFICACAO = true não altera nada.
 */

const MODO_VERIFICACAO = true;   // <<< troque para false para executar de verdade

const ABA_PROG = "PROGRAMAÇÃO";
const ABA_DASH = "DASHBOARD COLETA";
const LIN_CAB = 4;
const LIN_INI = 5;
const LIN_FIM = 517;

const N_SEMANAS = 12;   // janela: 4 semanas atrás até 7 à frente
const RECUO = 4;
const FILTRO = "$B$3";  // célula do filtro

const AZUL_ESC = "#1F4E78", AZUL_CLARO = "#DDEBF7", AMARELO = "#FFEB9C";
const BRANCO = "#FFFFFF";

let LOG: string[] = [];
function say(m: string) { LOG.push(m); }
function flush() { console.log(LOG.join("\n")); LOG = []; }

function main(workbook: ExcelScript.Workbook) {
  LOG = [];
  const prog = workbook.getWorksheet(ABA_PROG);
  if (!prog) throw new Error('Aba "' + ABA_PROG + '" não encontrada.');

  const cabs = lerCabecalhos(prog, LIN_CAB);
  const erros: string[] = [];
  const col: { [k: string]: string } = {};
  for (const nome of ["STATUS", "PROFORMA", "PAÍS DE DESTINO", "SEMANA PROGRAMADA",
                      "DATA DE COLETA", "UNIDADE DETALHADA"]) {
    const c = achar(cabs, nome);
    if (c) { col[nome] = letra(c); say("  OK  " + nome + " -> col " + col[nome]); }
    else { say("  !!  " + nome + " -> NÃO ENCONTRADA"); erros.push('Coluna "' + nome + '" não encontrada.'); }
  }

  const nCol = achar(cabs, "UNIDADE DETALHADA");
  const dCol = achar(cabs, "PAÍS DE DESTINO");
  const aCol = achar(cabs, "STATUS");
  const dados = prog.getRange("A" + LIN_INI + ":" + letra(cabs.length) + LIN_FIM).getValues();

  const plantas: string[] = [];
  const paises: string[] = [];
  for (const lin of dados) {
    const st = String(lin[aCol - 1] == null ? "" : lin[aCol - 1]);
    if (st === "Cancelado" || st === "") continue;
    const pl = String(lin[nCol - 1] == null ? "" : lin[nCol - 1]).trim();
    const pa = String(lin[dCol - 1] == null ? "" : lin[dCol - 1]).trim();
    if (pl !== "" && plantas.indexOf(pl) < 0) plantas.push(pl);
    if (pa !== "" && paises.indexOf(pa) < 0) paises.push(pa);
  }
  plantas.sort();
  paises.sort();
  say("\nPlantas (" + plantas.length + "): " + plantas.join(", "));
  say("Países (" + paises.length + "): " + paises.join(", "));
  say("\nAba DASHBOARD COLETA já existe? " + (workbook.getWorksheet(ABA_DASH) != null ? "SIM (será recriada)" : "não"));

  if (erros.length > 0) {
    say("\n*** NÃO PODE EXECUTAR ***");
    for (const e of erros) say("  - " + e);
    flush();
    throw new Error("Verificação falhou. Nada foi alterado.");
  }

  say("\n*** VERIFICAÇÃO OK ***");
  if (MODO_VERIFICACAO) {
    say("MODO_VERIFICACAO = true -> nada foi alterado.");
    flush();
    return;
  }

  // ---------- recria a aba
  const antiga = workbook.getWorksheet(ABA_DASH);
  if (antiga) antiga.delete();
  const sh = workbook.addWorksheet(ABA_DASH);
  sh.activate();

  const P = (c: string) => ABA_PROG + "!$" + c + "$" + LIN_INI + ":$" + c + "$" + LIN_FIM;
  const rStatus = P(col["STATUS"]);
  const rProf = P(col["PROFORMA"]);
  const rPais = P(col["PAÍS DE DESTINO"]);
  const rSem = P(col["SEMANA PROGRAMADA"]);
  const rColeta = P(col["DATA DE COLETA"]);
  const rPlanta = P(col["UNIDADE DETALHADA"]);

  /** monta "=IF(filtro=TODAS; semFiltro; comFiltro)" */
  const filtrado = (semFiltro: string, comFiltro: string) =>
    '=IF(' + FILTRO + '="TODAS",' + semFiltro + ',' + comFiltro + ')';

  // ---------- título
  sh.getRange("A1").setValue("📊  DASHBOARD COLETA  —  Operador Logístico × Cliente Exportador");
  sh.getRange("A1:I1").merge(false);
  sh.getRange("A1").getFormat().getFill().setColor(AZUL_ESC);
  sh.getRange("A1").getFormat().getFont().setColor(BRANCO);
  sh.getRange("A1").getFormat().getFont().setBold(true);
  sh.getRange("A1").getFormat().getFont().setSize(14);
  sh.getRange("A1").getFormat().setRowHeight(28);
  sh.getRange("A2").setFormula('=CONCATENATE("Atualizado em: ",TEXT(TODAY(),"dd/mm/yyyy"))');
  sh.getRange("A2").getFormat().getFont().setItalic(true);

  const lSem = 9, fSem = lSem + N_SEMANAS - 1;

  // ---------- seção 1: por semana (NÃO filtra)
  sh.getRange("A8").setValue("📅  RETIRADOS × PENDENTES POR SEMANA  (visão completa — não filtra)");
  sh.getRange("A8").getFormat().getFont().setBold(true);
  sh.getRange("A" + (lSem - 1) + ":C" + (lSem - 1)).setValues([["SEMANA", "RETIRADOS", "PENDENTES"]]);

  const sem: string[][] = [], ret: string[][] = [], pen: string[][] = [];
  for (let i = 0; i < N_SEMANAS; i++) {
    const r = lSem + i;
    sem.push(i === 0
      ? ['=MOD(ISOWEEKNUM(TODAY())-' + (RECUO + 1) + ',52)+1']
      : ['=MOD(A' + (r - 1) + ',52)+1']);
    ret.push(['=SUMPRODUCT((' + rStatus + '="Container retirado")*(' + rColeta + '<>"")*' +
              '(IFERROR(ISOWEEKNUM(' + rColeta + '),0)=$A' + r + '))']);
    pen.push(['=COUNTIFS(' + rStatus + ',"Nova programação",' + rSem + ',$A' + r + ')']);
  }
  sh.getRange("A" + lSem + ":A" + fSem).setFormulas(sem);
  sh.getRange("B" + lSem + ":B" + fSem).setFormulas(ret);
  sh.getRange("C" + lSem + ":C" + fSem).setFormulas(pen);

  const hS = sh.getRange("A" + (lSem - 1) + ":C" + (lSem - 1));
  hS.getFormat().getFill().setColor(AZUL_ESC);
  hS.getFormat().getFont().setColor(BRANCO);
  hS.getFormat().getFont().setBold(true);
  hS.getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.center);
  sh.getRange("A" + lSem + ":C" + fSem).getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.center);
  borda(sh.getRange("A" + (lSem - 1) + ":C" + fSem));

  // ---------- cartões
  sh.getRange("A5:E5").setValues([["🟡 RETIRADOS", "🔵 PENDENTES", "🟢 DEPOSITADOS", "🔴 CANCELADOS", "📦 TOTAL ATIVO"]]);
  sh.getRange("A6:E6").setFormulas([[
    filtrado('COUNTIF(' + rStatus + ',"Container retirado")',
             'COUNTIFS(' + rStatus + ',"Container retirado",' + rSem + ',' + FILTRO + ')'),
    filtrado('COUNTIF(' + rStatus + ',"Nova programação")',
             'COUNTIFS(' + rStatus + ',"Nova programação",' + rSem + ',' + FILTRO + ')'),
    filtrado('COUNTIF(' + rStatus + ',"Container depositado")',
             'COUNTIFS(' + rStatus + ',"Container depositado",' + rSem + ',' + FILTRO + ')'),
    filtrado('COUNTIF(' + rStatus + ',"Cancelado")',
             'COUNTIFS(' + rStatus + ',"Cancelado",' + rSem + ',' + FILTRO + ')'),
    filtrado('COUNTA(' + rProf + ')-COUNTIF(' + rStatus + ',"Cancelado")',
             'COUNTIFS(' + rSem + ',' + FILTRO + ',' + rStatus + ',"<>Cancelado")')
  ]]);
  const hC = sh.getRange("A5:E5");
  hC.getFormat().getFill().setColor(AZUL_CLARO);
  hC.getFormat().getFont().setBold(true);
  hC.getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.center);
  const vC = sh.getRange("A6:E6");
  vC.getFormat().getFont().setBold(true);
  vC.getFormat().getFont().setSize(20);
  vC.getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.center);
  vC.getFormat().setRowHeight(34);
  borda(sh.getRange("A5:E6"));

  // ---------- filtro
  sh.getRange("A3").setValue("🔎  SEMANA:");
  sh.getRange("A3").getFormat().getFont().setBold(true);
  sh.getRange("B3").setValue("TODAS");
  const cel = sh.getRange("B3");
  cel.getFormat().getFill().setColor(AMARELO);
  cel.getFormat().getFont().setBold(true);
  cel.getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.center);
  borda(cel);
  sh.getRange("C3").setValue("← escolha uma semana ou TODAS");
  sh.getRange("C3").getFormat().getFont().setItalic(true);

  // lista da suspensa (coluna AA, oculta)
  sh.getRange("AA1").setValue("TODAS");
  const listaSem: string[][] = [];
  for (let i = 0; i < N_SEMANAS; i++) listaSem.push(['=A' + (lSem + i)]);
  sh.getRange("AA2:AA" + (N_SEMANAS + 1)).setFormulas(listaSem);
  say("Lista de semanas montada em AA1:AA" + (N_SEMANAS + 1));

  try {
    cel.getDataValidation().setRule({
      list: { inCellDropDown: true, source: "=$AA$1:$AA$" + (N_SEMANAS + 1) }
    });
    say("Lista suspensa aplicada em B3.");
  } catch (e) {
    say("!! Falhou ao aplicar a lista suspensa em B3: " + e);
  }

  try {
    sh.getRange("AA:AA").getFormat().setColumnWidth(0);   // esconde a coluna auxiliar
    say("Coluna AA ocultada (largura 0).");
  } catch (e) {
    say("!! Não consegui ocultar a coluna AA (não é crítico): " + e);
  }

  // destaca a semana selecionada na tabela
  const cf = sh.getRange("A" + lSem + ":C" + fSem)
    .addConditionalFormat(ExcelScript.ConditionalFormatType.custom);
  cf.getCustom().getRule().setFormula('=AND(' + FILTRO + '<>"TODAS",$A' + lSem + '=' + FILTRO + ')');
  cf.getCustom().getFormat().getFill().setColor(AMARELO);
  cf.getCustom().getFormat().getFont().setBold(true);

  // ---------- seção 2: por planta (FILTRA)
  sh.getRange("E8").setValue("🏭  POR PLANTA");
  sh.getRange("E8").getFormat().getFont().setBold(true);
  sh.getRange("E" + (lSem - 1) + ":F" + (lSem - 1)).setValues([["PLANTA", "QTDE"]]);

  const fPl = lSem + plantas.length;
  sh.getRange("E" + lSem + ":E" + (fPl - 1)).setValues(plantas.map(p => [p]));
  const cPl: string[][] = [];
  for (let i = 0; i < plantas.length; i++) {
    const r = lSem + i;
    const p = plantas[i].replace(/"/g, '""');
    cPl.push([filtrado(
      'COUNTIFS(' + rPlanta + ',"' + p + '",' + rStatus + ',"<>Cancelado")',
      'COUNTIFS(' + rPlanta + ',"' + p + '",' + rStatus + ',"<>Cancelado",' + rSem + ',' + FILTRO + ')')]);
  }
  sh.getRange("F" + lSem + ":F" + (fPl - 1)).setFormulas(cPl);
  sh.getRange("E" + fPl).setValue("(sem planta / outros)");
  sh.getRange("F" + fPl).setFormula('=$E$6-SUM(F' + lSem + ':F' + (fPl - 1) + ')');
  cabecalho(sh, "E" + (lSem - 1) + ":F" + (lSem - 1));
  borda(sh.getRange("E" + (lSem - 1) + ":F" + fPl));

  // ---------- seção 3: por país (FILTRA)
  sh.getRange("H8").setValue("🌎  POR PAÍS");
  sh.getRange("H8").getFormat().getFont().setBold(true);
  sh.getRange("H" + (lSem - 1) + ":I" + (lSem - 1)).setValues([["PAÍS", "QTDE"]]);

  const fPa = lSem + paises.length;
  sh.getRange("H" + lSem + ":H" + (fPa - 1)).setValues(paises.map(p => [p]));
  const cPa: string[][] = [];
  for (let i = 0; i < paises.length; i++) {
    const p = paises[i].replace(/"/g, '""');
    cPa.push([filtrado(
      'COUNTIFS(' + rPais + ',"' + p + '",' + rStatus + ',"<>Cancelado")',
      'COUNTIFS(' + rPais + ',"' + p + '",' + rStatus + ',"<>Cancelado",' + rSem + ',' + FILTRO + ')')]);
  }
  sh.getRange("I" + lSem + ":I" + (fPa - 1)).setFormulas(cPa);
  sh.getRange("H" + fPa).setValue("(sem país / outros)");
  sh.getRange("I" + fPa).setFormula('=$E$6-SUM(I' + lSem + ':I' + (fPa - 1) + ')');
  cabecalho(sh, "H" + (lSem - 1) + ":I" + (lSem - 1));
  borda(sh.getRange("H" + (lSem - 1) + ":I" + fPa));

  // ---------- larguras
  const larg: [string, number][] = [["A", 115], ["B", 100], ["C", 105], ["D", 110],
                                    ["E", 170], ["F", 70], ["G", 20], ["H", 150], ["I", 70]];
  for (const [c, w] of larg) sh.getRange(c + "1").getFormat().setColumnWidth(w);

  // ---------- gráficos
  const linG = Math.max(fSem, fPl, fPa) + 2;

  const g1 = sh.addChart(ExcelScript.ChartType.columnClustered,
    sh.getRange("A" + (lSem - 1) + ":C" + fSem), ExcelScript.ChartSeriesBy.columns);
  g1.setName("PorSemana");
  g1.getTitle().setText("RETIRADOS × PENDENTES POR SEMANA");
  g1.setPosition(sh.getRange("A" + linG), sh.getRange("F" + (linG + 17)));

  const g2 = sh.addChart(ExcelScript.ChartType.barClustered,
    sh.getRange("E" + (lSem - 1) + ":F" + fPl), ExcelScript.ChartSeriesBy.columns);
  g2.setName("PorPlanta");
  g2.getTitle().setText("POR PLANTA (respeita o filtro)");
  g2.setPosition(sh.getRange("A" + (linG + 19)), sh.getRange("E" + (linG + 36)));

  const g3 = sh.addChart(ExcelScript.ChartType.pie,
    sh.getRange("H" + (lSem - 1) + ":I" + fPa), ExcelScript.ChartSeriesBy.columns);
  g3.setName("PorPais");
  g3.getTitle().setText("POR PAÍS (respeita o filtro)");
  g3.setPosition(sh.getRange("G" + (linG + 19)), sh.getRange("K" + (linG + 36)));

  sh.setShowGridlines(false);

  say("\n===== ITEM 4 (v2) CONCLUÍDO =====");
  say("Filtro em B3 (lista: TODAS + " + N_SEMANAS + " semanas)");
  say("Cartões/PLANTA/PAÍS filtram pela SEMANA PROGRAMADA. Gráfico de semanas não filtra.");
  flush();
}

function cabecalho(sh: ExcelScript.Worksheet, addr: string) {
  const r = sh.getRange(addr);
  r.getFormat().getFill().setColor(AZUL_ESC);
  r.getFormat().getFont().setColor(BRANCO);
  r.getFormat().getFont().setBold(true);
}

function borda(r: ExcelScript.Range) {
  const idx = [
    ExcelScript.BorderIndex.edgeTop, ExcelScript.BorderIndex.edgeBottom,
    ExcelScript.BorderIndex.edgeLeft, ExcelScript.BorderIndex.edgeRight,
    ExcelScript.BorderIndex.insideHorizontal, ExcelScript.BorderIndex.insideVertical
  ];
  for (const i of idx) {
    const b = r.getFormat().getRangeBorder(i);
    b.setStyle(ExcelScript.BorderLineStyle.continuous);
    b.setColor("#BFBFBF");
    b.setWeight(ExcelScript.BorderWeight.thin);
  }
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
