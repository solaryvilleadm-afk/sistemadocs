/**
 * Solary Ville - Mensalidades e contato
 * Recebe os dados dos contratos gerados no Sistema de Documentos
 * e organiza as abas em ordem alfabética, com as cores por tipo.
 *
 * Tipos (identificados pelo final do nome do residente):
 *   Permanente  -> nome normal          -> linhas mescladas azul claro / branco
 *   Centro Dia  -> "Nome - CENTRO DIA"  -> linha vermelho claro
 *   Hotelaria   -> "Nome - HOTELARIA"   -> linha amarelo claro
 */

var CHAVE = 'solary-docs-2026';

var COR = {
  azul: '#deeaf1', branco: '#ffffff',
  centroDia: '#f8d7da', hotelaria: '#fff3b0',
  // cores fixas dos blocos da aba de residentes (colunas F a O)
  financeiro: '#fff2cc', mensalidade: '#e2efda', reajustes: '#fce4d6'
};

// ---------------- RECEBER DO SISTEMA DE DOCUMENTOS ----------------
function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var d = JSON.parse(e.postData.contents);
    if (d.chave !== CHAVE) return resposta({ ok: false, erro: 'chave' });
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var u = String(d.unidade) === '2' ? '2' : '1';
    gravarResidente(ss.getSheetByName('Solary ' + u), d);
    gravarContato(ss.getSheetByName('Contatos Solary ' + u), d);
    return resposta({ ok: true });
  } catch (err) {
    return resposta({ ok: false, erro: String(err) });
  } finally {
    lock.releaseLock();
  }
}

function doGet() { return resposta({ ok: true, status: 'ativo' }); }

function resposta(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

// ---------------- AUXILIARES ----------------
function sufixo(tipo) {
  if (tipo === 'Centro Dia') return ' - CENTRO DIA';
  if (tipo === 'Hotelaria') return ' - HOTELARIA';
  return '';
}

function tipoDoNome(nome) {
  var n = String(nome || '').toUpperCase();
  if (n.indexOf('CENTRO DIA') >= 0) return 'Centro Dia';
  if (n.indexOf('HOTELARIA') >= 0) return 'Hotelaria';
  return 'Permanente';
}

function nomeBase(nome) {
  return String(nome || '').replace(/\s*-\s*(CENTRO DIA|HOTELARIA)\s*$/i, '').trim();
}

function chaveNome(nome) {
  return nomeBase(nome).normalize('NFD').replace(/[̀-ͯ]/g, '').replace(/\s+/g, ' ').toLowerCase();
}

function soDigitos(s) { return String(s || '').replace(/\D/g, ''); }

function formatarCpf(s) {
  var d = soDigitos(s);
  if (d.length !== 11) return String(s || '').trim();
  return d.slice(0, 3) + '.' + d.slice(3, 6) + '.' + d.slice(6, 9) + '-' + d.slice(9);
}

function dataIso(iso) {
  if (!iso) return '';
  var p = String(iso).split('-');
  if (p.length !== 3) return '';
  return new Date(Number(p[0]), Number(p[1]) - 1, Number(p[2]));
}

function dataCurta(iso) {
  var p = String(iso || '').split('-');
  return p.length === 3 ? p[2] + '/' + p[1] + '/' + p[0] : '';
}

function valorNumero(txt) {
  var s = String(txt || '').replace(/[^\d,.-]/g, '');
  if (!s) return '';
  if (s.indexOf(',') >= 0) s = s.replace(/\./g, '').replace(',', '.');
  var n = Number(s);
  return isNaN(n) ? '' : n;
}

// ---------------- RESIDENTES ----------------
function gravarResidente(sh, d) {
  var INICIO = 4, NCOL = 16;
  var fim = linhaTotalResidentes(sh) - 1;
  var nome = nomeBase(d.nome) + sufixo(d.tipo);
  var cpf = formatarCpf(d.cpf);
  var linha = -1;

  if (fim >= INICIO) {
    var vals = sh.getRange(INICIO, 1, fim - INICIO + 1, 2).getValues();
    for (var i = 0; i < vals.length; i++) {
      var mesmoCpf = soDigitos(cpf) && soDigitos(vals[i][1]) === soDigitos(cpf);
      var mesmoNome = chaveNome(vals[i][0]) && chaveNome(vals[i][0]) === chaveNome(nome);
      if (mesmoCpf || mesmoNome) { linha = INICIO + i; break; }
    }
  }
  if (linha < 0) {
    sh.insertRowBefore(fim + 1);
    linha = fim + 1;
  }

  var entrada = dataIso(d.entrada);
  var mesReajuste = '';
  if (d.tipo === 'Permanente' && entrada) {
    var m = entrada.getMonth() + 1;
    mesReajuste = 'mês ' + m + ' a ' + (m === 1 ? 12 : m - 1);
  }
  var venc = d.vencimento ? Number(String(d.vencimento).split('-')[2]) : '';

  var obs = [];
  if (d.tipo === 'Centro Dia') {
    if (d.diaria) obs.push('Diária R$ ' + d.diaria);
    if (d.regime) obs.push(d.regime);
  }
  if (d.tipo === 'Hotelaria') {
    obs.push('Hotelaria de ' + dataCurta(d.entrada) + (d.saida ? ' a ' + dataCurta(d.saida) : ''));
  }

  var atual = sh.getRange(linha, 1, 1, NCOL).getValues()[0];
  function manter(novo, velho) { return (novo === '' || novo === null || novo === undefined) ? velho : novo; }

  sh.getRange(linha, 1, 1, 3).setValues([[nome, manter(cpf, atual[1]), manter(dataIso(d.nascimento), atual[2])]]);
  sh.getRange(linha, 6, 1, 2).setValues([[manter(entrada, atual[5]), manter(venc, atual[6])]]);
  sh.getRange(linha, 9, 1, 2).setValues([[manter(mesReajuste, atual[8]), manter(valorNumero(d.valor), atual[9])]]);
  if (obs.length) sh.getRange(linha, 16).setValue(obs.join(' | '));

  organizarResidentes(sh);
}

function linhaTotalResidentes(sh) {
  var col = sh.getRange(1, 1, sh.getMaxRows(), 2).getValues();
  for (var i = 3; i < col.length; i++) {
    if (String(col[i][1]).toUpperCase().indexOf('TOTAL') >= 0 || String(col[i][0]).toUpperCase().indexOf('TOTAL') >= 0) return i + 1;
  }
  return sh.getLastRow() + 1;
}

function organizarResidentes(sh) {
  var INICIO = 4, NCOL = 16;
  var total = linhaTotalResidentes(sh);
  var n = total - INICIO;
  if (n <= 0) return;
  var rg = sh.getRange(INICIO, 1, n, NCOL);
  var vals = rg.getValues();
  var forms = rg.getFormulas();

  var linhas = vals.map(function (v, i) {
    var r = v.map(function (c, j) { return forms[i][j] ? forms[i][j] : c; });
    r[0] = String(r[0] || '').trim();
    return r;
  });
  linhas.sort(function (a, b) {
    if (!a[0] && b[0]) return 1;
    if (a[0] && !b[0]) return -1;
    return chaveNome(a[0]).localeCompare(chaveNome(b[0]), 'pt-BR');
  });
  // a idade é sempre recalculada pela data de nascimento da própria linha
  linhas.forEach(function (r, i) {
    var lr = INICIO + i;
    r[3] = '=IFERROR(DATEDIF(C' + lr + ',TODAY(),"Y"),"")';
  });
  rg.setValues(linhas);
  // setValues grava fórmulas iniciadas por "=" como fórmulas
  sh.getRange(INICIO, 4, n, 1).setFormulas(linhas.map(function (r) { return [r[3]]; }));
  sh.getRange(total, 10).setFormula('=SUM(J' + INICIO + ':J' + (total - 1) + ')');

  colorirResidentes(sh, INICIO, n, linhas);
}

function colorirResidentes(sh, inicio, n, linhas) {
  var cores = linhas.map(function (r, i) {
    var t = tipoDoNome(r[0]);
    var base = i % 2 === 0 ? COR.azul : COR.branco;
    var linha = [];
    for (var c = 0; c < 16; c++) {
      if (t === 'Centro Dia') linha.push(COR.centroDia);
      else if (t === 'Hotelaria') linha.push(COR.hotelaria);
      else if (c >= 5 && c <= 8) linha.push(COR.financeiro);
      else if (c === 9) linha.push(COR.mensalidade);
      else if (c >= 10 && c <= 14) linha.push(COR.reajustes);
      else linha.push(base);
    }
    return linha;
  });
  sh.getRange(inicio, 1, n, 16).setBackgrounds(cores);
}

// ---------------- CONTATOS ----------------
function gravarContato(sh, d) {
  var INICIO = 3, NCOL = 6;
  var nome = nomeBase(d.nome) + sufixo(d.tipo);
  var fim = sh.getLastRow();
  var linha = -1;
  if (fim >= INICIO) {
    var vals = sh.getRange(INICIO, 1, fim - INICIO + 1, 1).getValues();
    for (var i = 0; i < vals.length; i++) {
      if (chaveNome(vals[i][0]) && chaveNome(vals[i][0]) === chaveNome(nome)) { linha = INICIO + i; break; }
    }
  }
  if (linha < 0) linha = Math.max(fim, INICIO - 1) + 1;

  var resp = (d.responsaveis || []).filter(function (r) { return r && r.nome; });
  var nomes = resp.map(function (r) { return r.nome + (r.parentesco ? ' (' + r.parentesco + ')' : ''); }).join(' / ');
  var emails = resp.map(function (r) { return r.email; }).filter(Boolean).join(' / ') || (d.emailContratante || '');
  var fones = resp.map(function (r) { return r.telefone; }).filter(Boolean).join(' / ');
  var end = resp.length ? resp[0].endereco : '';

  var atual = sh.getRange(linha, 1, 1, NCOL).getValues()[0];
  function manter(novo, velho) { return novo ? novo : velho; }
  sh.getRange(linha, 1, 1, 5).setValues([[nome, manter(nomes, atual[1]), manter(emails, atual[2]), manter(fones, atual[3]), manter(end, atual[4])]]);

  organizarContatos(sh);
}

function organizarContatos(sh) {
  var INICIO = 3, NCOL = 6;
  var fim = sh.getLastRow();
  var n = fim - INICIO + 1;
  if (n <= 0) return;
  var rg = sh.getRange(INICIO, 1, n, NCOL);
  var linhas = rg.getValues().map(function (r) { r[0] = String(r[0] || '').trim(); return r; })
    .filter(function (r) { return r.join('').trim() !== ''; });
  linhas.sort(function (a, b) { return chaveNome(a[0]).localeCompare(chaveNome(b[0]), 'pt-BR'); });
  rg.clearContent();
  if (!linhas.length) return;
  sh.getRange(INICIO, 1, linhas.length, NCOL).setValues(linhas);
  var cores = linhas.map(function (r, i) {
    var t = tipoDoNome(r[0]);
    var c = t === 'Centro Dia' ? COR.centroDia : t === 'Hotelaria' ? COR.hotelaria : (i % 2 === 0 ? COR.azul : COR.branco);
    return [c, c, c, c, c, c];
  });
  sh.getRange(INICIO, 1, linhas.length, NCOL).setBackgrounds(cores);
  // linhas que ficaram vazias no fim voltam ao padrão branco
  if (n > linhas.length) sh.getRange(INICIO + linhas.length, 1, n - linhas.length, NCOL).setBackground(COR.branco);
}

// ---------------- MENU MANUAL ----------------
function onOpen() {
  SpreadsheetApp.getUi().createMenu('Solary')
    .addItem('Organizar ordem e cores', 'organizarTudo')
    .addToUi();
}

function organizarTudo() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  ['1', '2'].forEach(function (u) {
    organizarResidentes(ss.getSheetByName('Solary ' + u));
    organizarContatos(ss.getSheetByName('Contatos Solary ' + u));
  });
}
