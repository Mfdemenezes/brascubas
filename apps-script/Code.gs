/**
 * Backend do ranking do Quiz Brás Cubas (Google Apps Script vinculado à planilha).
 *
 * Contrato usado pelo index.html (tudo via GET, sem preflight de CORS):
 *   ?                      → lista
 *   ?action=add&cid=..&nome=..&acertos=..&total=..&tempoMs=..
 *   ?action=clear          → move tudo para a aba "Arquivo" (nada é apagado de verdade)
 * Resposta: { ok: true, id, scores: [{ id, nome, acertos, total, tempoMs }] }
 *
 * Implantar: Implantar → Nova implantação → App da Web
 *   Executar como: Eu · Quem pode acessar: Qualquer pessoa
 */

const SHEET_NAME = "Ranking";
const ARCHIVE_NAME = "Arquivo";
const HEADER = ["id", "nome", "acertos", "total", "tempoMs", "quando"];

function doGet(e) {
  const p = (e && e.parameter) || {};
  const action = p.action || "list";
  // Dois iPads gravando ao mesmo tempo: o lock serializa as escritas.
  const lock = LockService.getScriptLock();
  try {
    if (action !== "list") lock.waitLock(20000);
    const sh = sheet_(SHEET_NAME);
    let id = null;
    if (action === "add") id = add_(sh, p);
    else if (action === "clear") archive_(sh);
    else if (action !== "list") throw new Error("ação inválida");
    return json_({ ok: true, id: id, scores: read_(sh) });
  } catch (err) {
    return json_({ ok: false, erro: String((err && err.message) || err) });
  } finally {
    lock.releaseLock();
  }
}

function add_(sh, p) {
  const nome = String(p.nome || "").trim().slice(0, 80);
  const acertos = int_(p.acertos, 0, 1000);
  const total = int_(p.total, 1, 1000);
  const tempoMs = int_(p.tempoMs, 0, 86400000);
  if (!nome) throw new Error("nome vazio");
  if (acertos > total) throw new Error("acertos > total");

  // cid vem do aparelho: se o envio for repetido (Wi-Fi caiu na resposta), não duplica.
  const id = String(p.cid || Utilities.getUuid()).slice(0, 64);
  if (sh.getLastRow() > 1) {
    const hit = sh.getRange(2, 1, sh.getLastRow() - 1, 1).createTextFinder(id).matchEntireCell(true).findNext();
    if (hit) return id;
  }
  // Prefixo ' evita que um nome começando com = + - @ vire fórmula.
  const safeNome = /^[=+\-@]/.test(nome) ? "'" + nome : nome;
  sh.appendRow([id, safeNome, acertos, total, tempoMs, new Date()]);
  return id;
}

function read_(sh) {
  const n = sh.getLastRow() - 1;
  if (n < 1) return [];
  return sh.getRange(2, 1, n, HEADER.length).getValues()
    .filter(r => r[0] !== "")
    .map(r => ({ id: String(r[0]), nome: String(r[1]), acertos: Number(r[2]), total: Number(r[3]), tempoMs: Number(r[4]) }));
}

function archive_(sh) {
  const n = sh.getLastRow() - 1;
  if (n < 1) return;
  const rows = sh.getRange(2, 1, n, HEADER.length).getValues();
  const arq = sheet_(ARCHIVE_NAME);
  arq.getRange(arq.getLastRow() + 1, 1, rows.length, HEADER.length).setValues(rows);
  sh.deleteRows(2, n);
}

function sheet_(name) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  if (sh.getLastRow() === 0) {
    sh.appendRow(HEADER);
    sh.setFrozenRows(1);
  }
  // Não mistura com uma aba de mesmo nome criada em outro formato.
  const head = sh.getRange(1, 1, 1, HEADER.length).getValues()[0].map(String);
  if (head.join() !== HEADER.join()) throw new Error('aba "' + name + '" com cabeçalho diferente: ' + head.join(", "));
  return sh;
}

function int_(v, min, max) {
  const n = Math.round(Number(v));
  if (!isFinite(n) || n < min || n > max) throw new Error("valor inválido: " + v);
  return n;
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
