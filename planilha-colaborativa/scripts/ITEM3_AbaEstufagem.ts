/**
 * ITEM 3 — Aba ESTUFAGEM
 * Projeto: Controle Cliente Exportador x Operador Logístico  |  Excel Online (Office Script)
 *
 * O QUE FAZ:
 *   Cria a aba ESTUFAGEM reproduzindo o layout da planilha do cliente.
 *   O usuário digita o CONTRATO em C11 e todo o resto é preenchido sozinho,
 *   buscando na PROGRAMAÇÃO pela coluna REF. CLIENTE (H).
 *
 * COMO USAR (depois de criada):
 *   1. Digite o contrato em C11
 *   2. Selecione B1:H30 -> Ctrl+C
 *   3. No Outlook, Ctrl+V (cola com cores e tabelas)
 *   4. Assunto: copie de B36 (Estufagem) ou B40 (Minutas)
 *
 * SEGURANÇA:
 *   - Só cria/recria a aba ESTUFAGEM. NÃO toca em PROGRAMAÇÃO, GERAL ou DASHBOARD.
 *   - MODO_VERIFICACAO = true não altera nada.
 */

const MODO_VERIFICACAO = true;   // <<< troque para false para executar de verdade

const ABA_PROG = "PROGRAMAÇÃO";
const ABA_EST = "ESTUFAGEM";
const LIN_CAB = 4;

// cores
const VERMELHO = "#FF0000", BRANCO = "#FFFFFF", PRETO = "#000000";
const AMARELO = "#FFFF00", AZUL_CLARO = "#DDEBF7", VERDE_CLARO = "#C6EFCE";
const CINZA_BORDA = "#000000";

let LOG: string[] = [];
function say(m: string) { LOG.push(m); }
function flush() { console.log(LOG.join("\n")); LOG = []; }

function main(workbook: ExcelScript.Workbook) {
  LOG = [];
  const prog = workbook.getWorksheet(ABA_PROG);
  if (!prog) throw new Error('Aba "' + ABA_PROG + '" não encontrada.');

  const cabs = lerCabecalhos(prog, LIN_CAB);
  const erros: string[] = [];

  // colunas de origem na PROGRAMAÇÃO
  const need: { [k: string]: string } = {
    "REF. CLIENTE": "", "LOCAL DE RETIRADA": "", "CONTAINER": "", "UNIDADE DETALHADA": "",
    "TIPO DE CARGA": "", "TEMPERATURA": "", "BOOKING": "", "DATA DE ESTUFAGEM": "",
    "ARMADOR": "", "NAVIO": "", "PAÍS DE DESTINO": "", "TIPO DE RETIRADA": "",
    "ORIGEM": "", "DESTINO": "", "DEADLINE DE CARGA": "", "TRANSPORTADORA": "",
    "SEMANA PROGRAMADA": ""
  };
  for (const nome of Object.keys(need)) {
    const c = achar(cabs, nome);
    if (c) {
      need[nome] = letra(c);
      say("  OK  " + nome + " -> col " + need[nome]);
    } else {
      say("  !!  " + nome + " -> NÃO ENCONTRADA");
      erros.push('PROGRAMAÇÃO: coluna "' + nome + '" não encontrada.');
    }
  }

  const jaExiste = workbook.getWorksheet(ABA_EST) != null;
  say("\nAba ESTUFAGEM já existe? " + (jaExiste ? "SIM (será recriada do zero)" : "não"));

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
  const antiga = workbook.getWorksheet(ABA_EST);
  if (antiga) antiga.delete();
  const sh = workbook.addWorksheet(ABA_EST);
  sh.activate();

  // busca padrão: XLOOKUP pelo contrato digitado em C11
  const bk = (col: string) =>
    'XLOOKUP($C$11,' + ABA_PROG + '!' + need["REF. CLIENTE"] + ':' + need["REF. CLIENTE"] +
    ',' + ABA_PROG + '!' + col + ':' + col + ',"NÃO ENCONTRADO")';
  const campo = (col: string) => '=IF($C$11="","",' + bk(col) + ')';

  // ---------- larguras
  sh.getRange("A1").getFormat().setColumnWidth(20);
  sh.getRange("B1").getFormat().setColumnWidth(170);
  sh.getRange("C1").getFormat().setColumnWidth(330);
  sh.getRange("D1").getFormat().setColumnWidth(90);
  sh.getRange("E1").getFormat().setColumnWidth(170);
  sh.getRange("F1").getFormat().setColumnWidth(60);
  sh.getRange("G1").getFormat().setColumnWidth(60);
  sh.getRange("H1").getFormat().setColumnWidth(60);

  // ---------- linha 1: faixa vermelha
  sh.getRange("B1").setValue("Caros,");
  sh.getRange("B1").getFormat().getFont().setBold(true);
  const banner = sh.getRange("C1:H1");
  banner.merge(false);
  sh.getRange("C1").setValue("ANEXAR O COMPROVANTE DO AGENDAMENTO AO RESPONDER ESTE E-MAIL");
  banner.getFormat().getFill().setColor(VERMELHO);
  banner.getFormat().getFont().setColor(BRANCO);
  banner.getFormat().getFont().setBold(true);
  banner.getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.center);

  sh.getRange("B3").setValue("Segue dados do nosso carregamento. ");

  // ---------- bloco LOCAL DE RETIRADA / CONTAINER
  sh.getRange("C4").setValue("LOCAL DE RETIRADA");
  sh.getRange("E4").setValue("CONTAINER");
  const tit = sh.getRange("C4:C4");
  for (const addr of ["C4", "E4"]) {
    const r = sh.getRange(addr);
    r.getFormat().getFill().setColor(AMARELO);
    r.getFormat().getFont().setBold(true);
    r.getFormat().getFont().setSize(14);
    r.getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.center);
    borda(r);
  }
  sh.getRange("C5").setFormula(campo(need["LOCAL DE RETIRADA"]));
  sh.getRange("E5").setFormula(campo(need["CONTAINER"]));
  for (const addr of ["C5", "E5"]) {
    const r = sh.getRange(addr);
    r.getFormat().getFont().setBold(true);
    r.getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.center);
    borda(r);
  }
  sh.getRange("C5").getFormat().getFont().setColor(VERMELHO);
  sh.getRange("C5").getFormat().getFont().setSize(9);

  // ---------- avisos
  sh.getRange("B7").setValue("·    ATENÇÃO: NÃO RETIRAR EQUIPAMENTO COM MAX. GROSS ABAIXO DE 34.000 KG.");
  sh.getRange("B7").getFormat().getFont().setColor(VERMELHO);
  sh.getRange("B7").getFormat().getFont().setBold(true);
  sh.getRange("B8").setValue("·      Favor retornar com a ordem de coleta neste e-mail para que todos os interessados recebam.");
  sh.getRange("B8").getFormat().getFont().setBold(true);

  // ---------- tabela de dados (B10:C23)
  const rotulos = [
    "Unidade Detalhada:", "Contrato/Order:", "01X40` REEFER ", "Booking: ",
    "Data estufagem: ", "Armador: ", "Agente: ", "Navio: ", "Destino: ",
    "Tipo de Coleta:", "Porto de Coleta: ", "Porto de Destino", "D.L Carga: ", "Transportador"
  ];
  sh.getRange("B10:B23").setValues(rotulos.map(r => [r]));

  sh.getRange("C10").setFormula(campo(need["UNIDADE DETALHADA"]));
  // C11 = ENTRADA DO USUÁRIO (fica vazia)
  sh.getRange("C12").setFormula(
    '=IF($C$11="","",UPPER(IF(' + bk(need["TIPO DE CARGA"]) + '="Congelada","CONGELADOS","RESFRIADOS"))' +
    ' & " : " & TEXT(' + bk(need["TEMPERATURA"]) + ',"0") & "°C")');
  sh.getRange("C13").setFormula(campo(need["BOOKING"]));
  sh.getRange("C14").setFormula(campo(need["DATA DE ESTUFAGEM"]));
  sh.getRange("C15").setFormula(campo(need["ARMADOR"]));
  sh.getRange("C16").setFormula('=IF($C$11="","","OPERADOR")');
  sh.getRange("C17").setFormula(campo(need["NAVIO"]));
  sh.getRange("C18").setFormula(campo(need["PAÍS DE DESTINO"]));      // Destino = PAÍS
  sh.getRange("C19").setFormula(campo(need["TIPO DE RETIRADA"]));
  sh.getRange("C20").setFormula(campo(need["ORIGEM"]));
  sh.getRange("C21").setFormula(campo(need["DESTINO"]));              // Porto de Destino = DESTINO
  sh.getRange("C22").setFormula(campo(need["DEADLINE DE CARGA"]));
  sh.getRange("C23").setFormula(campo(need["TRANSPORTADORA"]));

  sh.getRange("C14").setNumberFormat("dd/mm/yyyy");
  sh.getRange("C22").setNumberFormat("dd/mm/yyyy");

  const tabela = sh.getRange("B10:C23");
  borda(tabela);
  sh.getRange("B10:B23").getFormat().getFont().setBold(true);
  const valores = sh.getRange("C10:C23");
  valores.getFormat().getFont().setBold(true);
  valores.getFormat().getFont().setColor(VERMELHO);
  valores.getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.center);

  // Contrato (entrada) e Booking: fundo azul, texto preto
  for (const addr of ["C11", "C13"]) {
    const r = sh.getRange(addr);
    r.getFormat().getFill().setColor(AZUL_CLARO);
    r.getFormat().getFont().setColor(PRETO);
  }
  // linha do REEFER: texto preto
  sh.getRange("C12").getFormat().getFont().setColor(PRETO);
  // Transportador: destaque amarelo
  const transp = sh.getRange("B23:C23");
  transp.getFormat().getFill().setColor(AMARELO);
  transp.getFormat().getFont().setColor(PRETO);
  sh.getRange("C23").getFormat().getFont().setSize(12);

  sh.getRange("B24").setValue("RETIRADA DO VAZIO A PARTIR DE: HOJE");
  sh.getRange("B24").getFormat().getFont().setBold(true);

  // ---------- rodapé do e-mail
  const faixa1 = sh.getRange("B28:H28");
  faixa1.merge(false);
  sh.getRange("B28").setValue("FAVOR ENVIAR A ORDEM DE COLETA NESTE E-MAIL, INFORMANDO A DATA DA COLETA DO CTR VAZIO.");
  faixa1.getFormat().getFill().setColor(AMARELO);

  const faixa2 = sh.getRange("B29:H29");
  faixa2.merge(false);
  sh.getRange("B29").setValue("FAVOR ENVIAR FOTOS DA COLETA DO CTR VAZIO PARA CONFERÊNCIA DE AVARIAS.");
  faixa2.getFormat().getFill().setColor(AZUL_CLARO);

  const faixa3 = sh.getRange("B30:H30");
  faixa3.merge(false);
  sh.getRange("B30").setValue("FAVOR JÁ ESPELHAR SUA FROTA PARA A PLATAFORMA DE RASTREAMENTO.");
  faixa3.getFormat().getFill().setColor(VERDE_CLARO);

  for (const f of [faixa1, faixa2, faixa3]) {
    f.getFormat().getFont().setBold(true);
    f.getFormat().getFont().setColor(PRETO);
  }

  // ---------- ASSUNTOS (fora da faixa copiada B1:H30)
  sh.getRange("B33").setValue("ASSUNTO — SOLICITAÇÃO DE ESTUFAGEM");
  sh.getRange("B33").getFormat().getFont().setBold(true);
  sh.getRange("B34").setValue("Utilizado quando o agendamento é entre transportador e armador.");

  const semana = bk(need["SEMANA PROGRAMADA"]);
  const contrato = "$C$11";
  const booking = bk(need["BOOKING"]);
  const dataEst = bk(need["DATA DE ESTUFAGEM"]);
  const planta = bk(need["UNIDADE DETALHADA"]);
  const armador = bk(need["ARMADOR"]);

  sh.getRange("B36").setFormula(
    '=IF($C$11="","","SOLICITAÇÃO DE ESTUFAGEM SEMANA " & ' + semana +
    ' & " - CONTRATO " & ' + contrato +
    ' & " / BOOKING " & ' + booking +
    ' & " - " & TEXT(' + dataEst + ',"dd/mm/yyyy")' +
    ' & " EM PLANTA " & ' + planta +
    ' & " / ARMADOR " & ' + armador + ')');
  sh.getRange("B36:H36").merge(false);
  sh.getRange("B36").getFormat().getFill().setColor(AMARELO);
  sh.getRange("B36").getFormat().getFont().setBold(true);

  sh.getRange("B38").setValue("ASSUNTO — MINUTAS");
  sh.getRange("B38").getFormat().getFont().setBold(true);
  sh.getRange("B39").setValue("Utilizado para retirada em Pátio.");

  const transportadora = bk(need["TRANSPORTADORA"]);
  const container = bk(need["CONTAINER"]);
  sh.getRange("B40").setFormula(
    '=IF($C$11="","","Liberação de unidade - " & ' + transportadora +
    ' & " - (PÁTIO) - CLIENTE EXPORTADOR - " & ' + contrato +
    ' & " / " & ' + booking +
    ' & " / " & ' + container +
    ' & " - PROGRAMAÇÃO SEMANA " & ' + semana + ')');
  sh.getRange("B40:H40").merge(false);
  sh.getRange("B40").getFormat().getFill().setColor(AMARELO);
  sh.getRange("B40").getFormat().getFont().setBold(true);

  // ---------- CORPO DAS MINUTAS (B43:B55)
  const minutas = [
    "Segue anexo a minuta de liberação.", ".",
    "Favor preencher com os dados e responder este e-mail com o anexo preenchido e data prevista de coleta. ", ".",
    "Notar que o motorista precisa de 2 vias impressas da minuta quando for retirar o vazio no pátio, ok.", ".", ".", ".",
    "FAVOR ENVIAR A ORDEM DE COLETA NESTE E-MAIL, INFORMANDO A DATA DA COLETA DO CTR VAZIO.", ".",
    "FAVOR ENVIAR FOTOS DA COLETA DO CTR VAZIO PARA CONFERÊNCIA DE AVARIAS.", ".",
    "FAVOR JÁ ESPELHAR SUA FROTA PARA A PLATAFORMA DE RASTREAMENTO."
  ];
  sh.getRange("B43:B55").setValues(minutas.map(m => [m]));

  // ---------- instruções (coluna J, fora da faixa copiada)
  const inst = [
    ["COMO USAR"],
    ["1) Digite o CONTRATO na célula C11 (é a única célula de digitação)"],
    ["2) Selecione B1:H30 e copie (Ctrl+C)"],
    ["3) No Outlook, cole (Ctrl+V) — cores e tabelas vão junto"],
    ["4) Assunto: copie B36 (Estufagem) ou B40 (Minutas)"],
    ["5) Corpo das Minutas: copie B43:B55"],
    [""],
    ["O CONTAINER aparece como 0 quando ainda não foi informado."],
    ["Se aparecer NÃO ENCONTRADO, o contrato não existe na aba PROGRAMAÇÃO."],
    [""],
    ["RETIRADA DO VAZIO — coordenar com: contato@operadorlogistico.com.br"]
  ];
  sh.getRange("J1:J11").setValues(inst);
  sh.getRange("J1").getFormat().getFont().setBold(true);
  sh.getRange("J1:J11").getFormat().getFont().setColor("#808080");
  sh.getRange("J1").getFormat().setColumnWidth(420);

  // destaque da célula de entrada
  const entrada = sh.getRange("C11");
  entrada.getFormat().getFill().setColor(AZUL_CLARO);
  const bordas = entrada.getFormat().getRangeBorder(ExcelScript.BorderIndex.edgeBottom);
  bordas.setColor(VERMELHO);
  bordas.setWeight(ExcelScript.BorderWeight.thick);
  bordas.setStyle(ExcelScript.BorderLineStyle.continuous);

  sh.getRange("A1").setValue("");
  sh.setShowGridlines(false);

  say("\n===== ITEM 3 CONCLUÍDO =====");
  say("Aba ESTUFAGEM criada. Digite um contrato em C11 para testar.");
  flush();
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
    b.setColor(CINZA_BORDA);
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
