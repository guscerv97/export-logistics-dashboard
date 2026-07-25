/**
 * FASE 1.1 — Correção pós-migração (PROGRAMAÇÃO)
 * Cliente Exportador × Operador Logístico
 *
 * O QUE FAZ:
 *   1. Converte a coluna L (DATA DE ESTUFAGEM) de TEXTO "dd/mm/aaaa" para DATA real
 *   2. Restaura a fórmula da coluna K (SEMANA PROGRAMADA) = ISOWEEKNUM(L)
 *
 * O QUE NÃO FAZ:
 *   - Não toca em nenhuma outra coluna
 *   - Não toca na aba ARQUIVO (valores fixos ali são o comportamento correto)
 *   - Não mexe em linhas, formatação condicional ou validações
 *
 * COMO USAR:
 *   1) Rode com MODO_VERIFICACAO = true  -> só imprime o relatório, não altera nada
 *   2) Confira o relatório
 *   3) Troque para false e rode de novo para aplicar
 */

function main(workbook: ExcelScript.Workbook) {

  // ====== CONFIGURAÇÃO ======
  const MODO_VERIFICACAO = true;   // <<< troque para false só depois de conferir
  const NOME_ABA = "PROGRAMAÇÃO";
  const PRIMEIRA_LINHA = 5;        // dados começam na linha 5 (cabeçalho na 4)
  const COL_STATUS = "A";
  const COL_PROFORMA = "B";        // usada para saber se a linha tem dado
  const COL_SEMANA = "K";
  const COL_DATA_ESTUFAGEM = "L";
  // ==========================

  const ws = workbook.getWorksheet(NOME_ABA);
  if (!ws) {
    console.log("ERRO: aba " + NOME_ABA + " não encontrada.");
    return;
  }

  // --- Descobrir a última linha com dados (baseado na coluna PROFORMA) ---
  const usado = ws.getUsedRange();
  const ultimaLinhaPlanilha = usado ? usado.getLastRow().getRowIndex() + 1 : PRIMEIRA_LINHA;

  const colB = ws.getRange(COL_PROFORMA + PRIMEIRA_LINHA + ":" + COL_PROFORMA + ultimaLinhaPlanilha);
  const valoresB = colB.getValues();

  let ultimaLinha = PRIMEIRA_LINHA - 1;
  for (let i = 0; i < valoresB.length; i++) {
    const v = valoresB[i][0];
    if (v !== null && v !== undefined && String(v).trim() !== "") {
      ultimaLinha = PRIMEIRA_LINHA + i;
    }
  }

  if (ultimaLinha < PRIMEIRA_LINHA) {
    console.log("Nenhuma linha com dados encontrada. Nada a fazer.");
    return;
  }

  const totalLinhas = ultimaLinha - PRIMEIRA_LINHA + 1;
  console.log("=== FASE 1.1 — CORREÇÃO K e L ===");
  console.log("Modo: " + (MODO_VERIFICACAO ? "VERIFICAÇÃO (nada será alterado)" : "EXECUÇÃO (alterações serão aplicadas)"));
  console.log("Faixa analisada: linhas " + PRIMEIRA_LINHA + " a " + ultimaLinha + "  (" + totalLinhas + " linhas)");
  console.log("");

  // ================================================================
  // ETAPA 1 — Coluna L: texto -> data
  // ================================================================
  const rngL = ws.getRange(COL_DATA_ESTUFAGEM + PRIMEIRA_LINHA + ":" + COL_DATA_ESTUFAGEM + ultimaLinha);
  const valoresL = rngL.getValues();

  const novosL: (string | number | boolean)[][] = [];
  let convertidas = 0;
  let jaEramData = 0;
  let vazias = 0;
  const naoReconhecidas: string[] = [];

  for (let i = 0; i < valoresL.length; i++) {
    const linha = PRIMEIRA_LINHA + i;
    const bruto = valoresL[i][0];

    // vazio -> mantém vazio
    if (bruto === null || bruto === undefined || String(bruto).trim() === "") {
      novosL.push([""]);
      vazias++;
      continue;
    }

    // já é número (data real no Excel) -> mantém
    if (typeof bruto === "number") {
      novosL.push([bruto]);
      jaEramData++;
      continue;
    }

    // é texto -> tentar converter
    const serial = textoParaSerialExcel(String(bruto));
    if (serial === null) {
      novosL.push([bruto]);              // não mexe no que não entendeu
      naoReconhecidas.push("linha " + linha + ": \"" + String(bruto) + "\"");
    } else {
      novosL.push([serial]);
      convertidas++;
    }
  }

  console.log("--- ETAPA 1: coluna L (DATA DE ESTUFAGEM) ---");
  console.log("  Textos a converter para data : " + convertidas);
  console.log("  Já eram data (mantidas)      : " + jaEramData);
  console.log("  Vazias                       : " + vazias);
  console.log("  NÃO reconhecidas             : " + naoReconhecidas.length);
  if (naoReconhecidas.length > 0) {
    console.log("  >>> ATENÇÃO — estas ficarão como estão, confira manualmente:");
    for (const item of naoReconhecidas) {
      console.log("      " + item);
    }
  }
  console.log("");

  // ================================================================
  // ETAPA 2 — Coluna K: restaurar fórmula
  // ================================================================
  const rngK = ws.getRange(COL_SEMANA + PRIMEIRA_LINHA + ":" + COL_SEMANA + ultimaLinha);
  const formulasK = rngK.getFormulas();

  const novasK: string[][] = [];
  let comFormula = 0;
  let semFormula = 0;

  for (let i = 0; i < formulasK.length; i++) {
    const linha = PRIMEIRA_LINHA + i;
    const f = String(formulasK[i][0]);
    if (f.indexOf("=") === 0) {
      comFormula++;
    } else {
      semFormula++;
    }
    // reescreve todas para garantir padrão único
    novasK.push(['=IF(L' + linha + '="","",ISOWEEKNUM(L' + linha + '))']);
  }

  console.log("--- ETAPA 2: coluna K (SEMANA PROGRAMADA) ---");
  console.log("  Já tinham fórmula   : " + comFormula);
  console.log("  Estavam como valor  : " + semFormula);
  console.log("  Serão padronizadas  : " + novasK.length);
  console.log("");

  // ================================================================
  // APLICAÇÃO
  // ================================================================
  if (MODO_VERIFICACAO) {
    console.log("=== MODO VERIFICAÇÃO — nenhuma alteração foi feita. ===");
    console.log("Se o relatório acima estiver correto, troque MODO_VERIFICACAO para false e rode novamente.");
    return;
  }

  // L: grava valores + formato de data
  rngL.setValues(novosL);
  rngL.setNumberFormat("dd/mm/yyyy");   // formato em inglês, funciona em qualquer idioma

  // K: grava fórmulas + formato numérico
  rngK.setFormulas(novasK);
  rngK.setNumberFormat("0");

  console.log("=== EXECUÇÃO CONCLUÍDA ===");
  console.log("  Coluna L: " + convertidas + " datas convertidas, formato dd/mm/yyyy aplicado.");
  console.log("  Coluna K: " + novasK.length + " fórmulas ISOWEEKNUM restauradas.");
  console.log("");
  console.log("CONFIRA AGORA:");
  console.log("  1. Coluna L mostrando datas alinhadas à direita (não à esquerda)");
  console.log("  2. Coluna K mostrando 30 e 31");
  console.log("  3. Coluna AD (DATA IDEAL RETIRADA) voltou a preencher");
  console.log("  4. DASHBOARD D8 (estufagens próximos 7 dias) com número coerente");
}


/**
 * Converte texto "dd/mm/aaaa" (ou dd-mm-aa) para o número de série do Excel.
 * Retorna null se não reconhecer o formato.
 */
function textoParaSerialExcel(txt: string): number | null {
  const limpo = txt.trim();
  const m = limpo.match(/^(\d{1,2})[\/\-\.](\d{1,2})[\/\-\.](\d{2,4})$/);
  if (!m) {
    return null;
  }

  const dia = parseInt(m[1], 10);
  const mes = parseInt(m[2], 10);
  let ano = parseInt(m[3], 10);
  if (ano < 100) {
    ano += 2000;
  }

  if (mes < 1 || mes > 12 || dia < 1 || dia > 31) {
    return null;
  }

  const d = new Date(Date.UTC(ano, mes - 1, dia));

  // valida se a data existe de fato (ex: 31/02 seria rejeitada)
  if (d.getUTCFullYear() !== ano || d.getUTCMonth() !== mes - 1 || d.getUTCDate() !== dia) {
    return null;
  }

  // 25569 = dias entre 30/12/1899 (base do Excel) e 01/01/1970 (base do JS)
  return Math.round(d.getTime() / 86400000 + 25569);
}
