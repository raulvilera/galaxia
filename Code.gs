/**
 * Apps Script — Atividade de Ciências — 8º Ano — 3º Bimestre.
 * Registra as respostas na planilha vinculada ao projeto.
 */
const CONFIG = Object.freeze({
  SPREADSHEET_ID: '1RN5dP-xVKzq6GshpLFwtS7362UfgX6oUF4-usVykmOM',
  SHEET_NAME: 'Respostas_8Ano_3Bim',
  ESCOLA: 'E.E. PROFª WANDA MASCAGNI DE SÁ',
  TURMA: '8º Ano B'
});

// Índices das alternativas: A=0, B=1, C=2, D=3.
const GABARITO = Object.freeze({
  Q01: 1,
  Q02: 0,
  Q03: 1,
  Q04: 2,
  Q05: 0,
  Q06: 1,
  Q07: 0,
  Q08: 0,
  Q09: 0
});

function doGet() {
  return respostaJson_({
    success: true,
    service: 'atividade-ciencias-8ano-3bimestre',
    turma: CONFIG.TURMA
  });
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);

  try {
    const payload = lerPayload_(e);
    validarPayload_(payload);

    const questaoIds = normalizarIds_(payload);
    const resultado = corrigirObjetivas_(questaoIds, payload.respostas);
    const temDissertativas = questaoIds.some(id => !Object.prototype.hasOwnProperty.call(GABARITO, id));
    const respostasTexto = questaoIds
      .map((id, i) => `${id}: ${payload.respostas[i] == null ? '' : String(payload.respostas[i])}`)
      .join(' | ');

    const planilha = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const aba = obterAba_(planilha);
    aba.appendRow([
      new Date(),
      payload.escola || CONFIG.ESCOLA,
      payload.turma,
      payload.nome,
      payload.ra || '',
      payload.email || '',
      resultado.acertos,
      resultado.total,
      resultado.nota,
      temDissertativas ? 'Aguardando correção das dissertativas' : 'Concluída',
      respostasTexto
    ]);

    return respostaJson_({
      success: true,
      acertos: resultado.acertos,
      totalObjetivas: resultado.total,
      notaObjetivas: resultado.nota
    });
  } catch (erro) {
    console.error(erro);
    return respostaJson_({
      success: false,
      error: erro && erro.message ? erro.message : String(erro)
    });
  } finally {
    lock.releaseLock();
  }
}

function lerPayload_(e) {
  const corpo = e && e.postData && e.postData.contents;
  if (!corpo) throw new Error('O corpo da requisição está vazio.');
  try {
    return JSON.parse(corpo);
  } catch (erro) {
    throw new Error('O corpo da requisição não contém JSON válido.');
  }
}

function validarPayload_(payload) {
  if (!payload || typeof payload !== 'object') throw new Error('Payload inválido.');
  if (!payload.turma || !payload.nome || !Array.isArray(payload.respostas)) {
    throw new Error('Informe turma, nome e respostas.');
  }
  if (payload.turma !== CONFIG.TURMA) {
    throw new Error(`Turma não autorizada: ${payload.turma}.`);
  }
}

function normalizarIds_(payload) {
  const ids = Array.isArray(payload.questaoIds) && payload.questaoIds.length
    ? payload.questaoIds
    : payload.respostas.map((_, i) => `Q${String(i + 1).padStart(2, '0')}`);
  if (ids.length !== payload.respostas.length) {
    throw new Error('A quantidade de IDs das questões não coincide com a quantidade de respostas.');
  }
  return ids.map(id => String(id).trim());
}

function corrigirObjetivas_(questaoIds, respostas) {
  let acertos = 0;
  let total = 0;
  questaoIds.forEach((id, i) => {
    if (!Object.prototype.hasOwnProperty.call(GABARITO, id)) return;
    total++;
    if (String(respostas[i]) === String(GABARITO[id])) acertos++;
  });
  return {
    acertos,
    total,
    nota: total ? Math.round((acertos / total) * 10 * 100) / 100 : 0
  };
}

function obterAba_(planilha) {
  let aba = planilha.getSheetByName(CONFIG.SHEET_NAME);
  if (!aba) aba = planilha.insertSheet(CONFIG.SHEET_NAME);
  if (aba.getLastRow() === 0) {
    aba.appendRow([
      'Data/hora', 'Escola', 'Turma', 'Nome', 'RA', 'E-mail',
      'Acertos objetivas', 'Total objetivas', 'Nota objetivas (0-10)',
      'Status', 'Respostas'
    ]);
    aba.setFrozenRows(1);
  }
  return aba;
}

function respostaJson_(objeto) {
  return ContentService
    .createTextOutput(JSON.stringify(objeto))
    .setMimeType(ContentService.MimeType.JSON);
}

function testarAcessoPlanilha() {
  const planilha = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  console.log(`${planilha.getName()} / ${obterAba_(planilha).getName()}`);
}
