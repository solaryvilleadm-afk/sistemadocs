/**
 * Mensalidades e Contatos - Santa Felicidade
 * Recebe os dados da Central de Documentos Santa Felicidade e lança na planilha.
 *
 * Abas usadas (estrutura atual da planilha):
 *   Residentes          A nome | B CPF | C nascimento | D idade (fórmula) | E meio de pagamento | F entrada
 *                       G vencimento | H mês do reajuste (fórmula) | I mensalidade inicial | J atual (fórmula)
 *                       K a N reajustes | O observações
 *   Contatos            A nome (fórmula) | B responsável | C parentesco | D e-mail | E WhatsApp | F endereço | G obs
 *   Centro Dia          A nome | B CPF | C nascimento | D idade | E meio | F entrada | G vencimento | H dias da semana
 *                       I mês reajuste | J diária inicial | K diária atual | L diárias no mês | M mensalidade do mês
 *                       N a Q diárias 2027 a 2030 | R observações
 *   Contatos Centro Dia igual à aba Contatos
 *
 * Linha 4 é a primeira linha de dados. A linha do cadastro e a linha dos contatos são sempre a mesma pessoa.
 */

const CHAVE = 'solary-docs-2026';
const LINHA1 = 4;

const ABAS = {
  residentes: {
    nome: 'Residentes', contatos: 'Contatos', cols: 15,
    entradas: [1, 2, 3, 5, 6, 7, 9, 11, 12, 13, 14, 15],  // colunas digitáveis (as demais são fórmula)
    obs: 15, valor: 9
  },
  centroDia: {
    nome: 'Centro Dia', contatos: 'Contatos Centro Dia', cols: 18,
    entradas: [1, 2, 3, 5, 6, 7, 8, 10, 12, 14, 15, 16, 17, 18],
    obs: 18, valor: 10
  }
};

const AMARELO = '#fff3b0';

// ---------------------------------------------------------------- web app
function doGet() {
  return json_({ ok: true, planilha: 'Mensalidades e Contatos - Santa Felicidade' });
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const d = JSON.parse(e.postData.contents);
    if (d.chave !== CHAVE) return json_({ ok: false, erro: 'chave inválida' });
    if (!d.nome) return json_({ ok: false, erro: 'nome vazio' });
    const linha = lancar_(d);
    return json_({ ok: true, linha: linha });
  } catch (err) {
    return json_({ ok: false, erro: String(err) });
  } finally {
    lock.releaseLock();
  }
}

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

// ---------------------------------------------------------------- lançamento
function lancar_(d) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const tipo = String(d.tipo || 'Permanente');
  const cd = tipo === 'Centro Dia';
  const hot = tipo === 'Hotelaria' || tipo === 'Pós-Operatório';
  const cfg = cd ? ABAS.centroDia : ABAS.residentes;
  const sh = ss.getSheetByName(cfg.nome);
  const ct = ss.getSheetByName(cfg.contatos);

  let nome = String(d.nome).trim().toUpperCase();
  if (hot && nome.indexOf('HOTELARIA') < 0) nome += ' - HOTELARIA';

  let linha = acharLinha_(sh, d.cpf, nome);
  const nova = !linha;
  if (nova) linha = primeiraLinhaVazia_(sh, ct, cfg);

  const set = function (col, v) { if (v !== '' && v !== null && v !== undefined) sh.getRange(linha, col).setValue(v); };
  if (nova) set(1, nome);
  if (nova || !sh.getRange(linha, 2).getValue()) set(2, d.cpf || '');
  set(3, data_(d.nascimento));
  if (nova || !sh.getRange(linha, 5).getValue()) set(5, 'BOLETO');
  set(6, data_(d.entrada));

  let venc = '';
  if (d.vencimento) { const dv = data_(d.vencimento); venc = dv ? dv.getDate() : ''; }
  if (!venc && (nova || !sh.getRange(linha, 7).getValue())) venc = 1;
  set(7, venc);

  if (cd) {
    set(8, d.regime || '');
    set(10, numero_(d.diaria));
  } else {
    set(9, numero_(d.valor));
  }

  // observações de hotelaria e pós-operatório
  if (hot) {
    const periodo = [curta_(d.entrada), curta_(d.saida)].filter(Boolean).join(' a ');
    const txt = (tipo === 'Pós-Operatório' ? 'Hotelaria pós-operatório' : 'Hotelaria') + (periodo ? ': ' + periodo : '');
    const atual = String(sh.getRange(linha, cfg.obs).getValue() || '');
    if (atual.indexOf(txt) < 0) sh.getRange(linha, cfg.obs).setValue(atual ? atual + ' | ' + txt : txt);
  }

  // contatos: responsável 1 nas colunas, demais em observações
  const resps = (d.responsaveis || []).filter(function (p) { return p && p.nome; });
  if (resps.length) {
    const r = resps[0];
    const email = r.email || d.emailContratante || '';
    const cset = function (col, v) { if (v) ct.getRange(linha, col).setValue(v); };
    cset(2, String(r.nome).toUpperCase());
    cset(3, r.parentesco);
    cset(4, email);
    cset(5, r.telefone);
    cset(6, r.endereco);
    if (resps.length > 1) {
      const outros = resps.slice(1).map(function (p) {
        return '2º responsável: ' + [p.nome, p.parentesco, p.telefone, p.email].filter(Boolean).join(', ');
      }).join(' | ');
      const atual = String(ct.getRange(linha, 7).getValue() || '');
      if (atual.indexOf(outros) < 0) ct.getRange(linha, 7).setValue(atual ? atual + ' | ' + outros : outros);
    }
  }

  SpreadsheetApp.flush();
  ordenar_(sh, ct, cfg);
  return linha;
}

function normal_(s) {
  return String(s || '').toUpperCase().normalize('NFD').replace(/[̀-ͯ]/g, '')
    .replace(/\s*-\s*(HOTELARIA|CENTRO DIA)\s*$/, '').replace(/\s+/g, ' ').trim();
}
function digitos_(s) { return String(s || '').replace(/\D/g, ''); }

function acharLinha_(sh, cpf, nome) {
  const n = sh.getLastRow() - LINHA1 + 1;
  if (n < 1) return 0;
  const v = sh.getRange(LINHA1, 1, n, 2).getValues();
  const c = digitos_(cpf), alvo = normal_(nome);
  for (let i = 0; i < v.length; i++) if (c && c.length >= 11 && digitos_(v[i][1]) === c) return LINHA1 + i;
  for (let i = 0; i < v.length; i++) if (v[i][0] && normal_(v[i][0]) === alvo) return LINHA1 + i;
  return 0;
}

function primeiraLinhaVazia_(sh, ct, cfg) {
  const ultima = sh.getLastRow();
  const n = ultima - LINHA1 + 1;
  const v = sh.getRange(LINHA1, 1, n, 1).getValues();
  for (let i = 0; i < v.length; i++) if (!String(v[i][0]).trim()) return LINHA1 + i;
  // sem linha livre: cria uma nova nas duas abas, copiando fórmulas e formatos da última
  [[sh, cfg.cols], [ct, 7]].forEach(function (p) {
    const s = p[0], last = s.getLastRow();
    s.insertRowAfter(last);
    s.getRange(last, 1, 1, p[1]).copyTo(s.getRange(last + 1, 1, 1, p[1]));
  });
  const nova = ultima + 1;
  cfg.entradas.forEach(function (col) { sh.getRange(nova, col).clearContent(); });
  ct.getRange(nova, 2, 1, 6).clearContent();
  return nova;
}

function data_(iso) {
  if (!iso) return '';
  const m = String(iso).match(/^(\d{4})-(\d{2})-(\d{2})/);
  return m ? new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]), 12) : '';
}
function curta_(iso) {
  const m = String(iso || '').match(/^(\d{4})-(\d{2})-(\d{2})/);
  return m ? m[3] + '/' + m[2] + '/' + m[1] : '';
}
function numero_(txt) {
  if (txt === '' || txt === null || txt === undefined) return '';
  const n = Number(String(txt).replace(/[^\d,.-]/g, '').replace(/\./g, '').replace(',', '.'));
  return isNaN(n) ? '' : n;
}

// ---------------------------------------------------------------- ordem alfabética
// Ordena o cadastro e leva junto a linha correspondente dos contatos. Só mexe nas colunas digitáveis,
// então as fórmulas continuam no lugar.
function ordenar_(sh, ct, cfg) {
  const n = sh.getLastRow() - LINHA1 + 1;
  if (n < 2) return;
  const cad = sh.getRange(LINHA1, 1, n, cfg.cols).getValues();
  const con = ct.getRange(LINHA1, 2, n, 6).getValues();
  const reg = cad.map(function (l, i) { return { l: l, c: con[i] }; });
  reg.sort(function (a, b) {
    const x = String(a.l[0]).trim(), y = String(b.l[0]).trim();
    if (!x && !y) return 0; if (!x) return 1; if (!y) return -1;
    return normal_(x).localeCompare(normal_(y), 'pt-BR');
  });
  cfg.entradas.forEach(function (col) {
    sh.getRange(LINHA1, col, n, 1).setValues(reg.map(function (r) { return [r.l[col - 1]]; }));
  });
  ct.getRange(LINHA1, 2, n, 6).setValues(reg.map(function (r) { return r.c; }));
}

// ---------------------------------------------------------------- menu e configuração
function onOpen() {
  SpreadsheetApp.getUi().createMenu('Solary')
    .addItem('Organizar ordem alfabética', 'organizarTudo')
    .addItem('Configurar cores (rodar uma vez)', 'configurar')
    .addToUi();
}

function organizarTudo() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  [ABAS.residentes, ABAS.centroDia].forEach(function (cfg) {
    ordenar_(ss.getSheetByName(cfg.nome), ss.getSheetByName(cfg.contatos), cfg);
  });
}

// Hotelaria em amarelo claro na aba Residentes (pelo nome com "- HOTELARIA").
function configurar() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ABAS.residentes.nome);
  const faixa = sh.getRange(LINHA1, 1, Math.max(sh.getMaxRows() - LINHA1 + 1, 1), ABAS.residentes.cols);
  const regras = sh.getConditionalFormatRules().filter(function (r) {
    const c = r.getBooleanCondition();
    return !(c && String(c.getCriteriaValues()[0]).indexOf('HOTELARIA') >= 0);
  });
  // fórmulas do Apps Script usam sempre a sintaxe em inglês, com vírgula
  regras.push(SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=ISNUMBER(SEARCH("HOTELARIA",$A' + LINHA1 + '))')
    .setBackground(AMARELO).setRanges([faixa]).build());
  sh.setConditionalFormatRules(regras);
}
