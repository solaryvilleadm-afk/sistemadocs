/**
 * Solary Ville - Mensalidades e contato
 * Recebe os dados dos contratos gerados no Sistema de Documentos
 * e mantém as abas em ordem alfabética, com as cores por tipo.
 *
 * Coluna Q (TIPO) nas abas Solary 1 e Solary 2:
 *   Permanente -> linhas mescladas azul claro / branco
 *   Centro Dia -> linha vermelho claro
 *   Hotelaria  -> linha amarelo claro
 * A cor muda na hora ao escolher o tipo na coluna Q.
 * Se o nome terminar em "- CENTRO DIA" ou "- HOTELARIA", o tipo é preenchido sozinho.
 */

var CHAVE = 'solary-docs-2026';
var TIPOS = ['Permanente', 'Centro Dia', 'Hotelaria'];
var INI_RES = 4, NCOL_RES = 17, COL_TIPO = 17;
var INI_CON = 3, NCOL_CON = 6;

var COR = {
  azul: '#deeaf1', branco: '#ffffff',
  centroDia: '#f8d7da', hotelaria: '#fff3b0',
  financeiro: '#fff2cc', mensalidade: '#e2efda', reajustes: '#fce4d6'
};

// ================= RECEBER DO SISTEMA DE DOCUMENTOS =================
function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var d = JSON.parse(e.postData.contents);
    if (d.chave !== CHAVE) return resposta({ ok: false, erro: 'chave' });
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var u = String(d.unidade) === '2' ? '2' : '1';
    var shR = ss.getSheetByName('Solary ' + u);
    var shC = ss.getSheetByName('Contatos Solary ' + u);
    prepararAba(shR);
    gravarResidente(shR, d);
    gravarContato(shC, d);
    organizarResidentes(shR);
    organizarContatos(shC, shR);
    return resposta({ ok: true, residente: d.nome });
  } catch (err) {
    console.error(err && err.stack ? err.stack : err);
    return resposta({ ok: false, erro: String(err) });
  } finally {
    lock.releaseLock();
  }
}

function doGet() { return resposta({ ok: true, status: 'ativo' }); }

function resposta(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

// ================= AUXILIARES =================
function tipoDoNome(nome) {
  var n = String(nome || '').toUpperCase();
  if (n.indexOf('CENTRO DIA') >= 0) return 'Centro Dia';
  if (n.indexOf('HOTELARIA') >= 0) return 'Hotelaria';
  return '';
}

function nomeBase(nome) {
  return String(nome || '').replace(/\s*-\s*(CENTRO DIA|HOTELARIA)\s*$/i, '').replace(/\s+/g, ' ').trim();
}

function chaveNome(nome) {
  return nomeBase(nome).normalize('NFD').replace(/[̀-ͯ]/g, '').toLowerCase();
}

function soDigitos(s) { return String(s || '').replace(/\D/g, ''); }

function formatarCpf(s) {
  var d = soDigitos(s);
  if (d.length !== 11) return String(s || '').trim();
  return d.slice(0, 3) + '.' + d.slice(3, 6) + '.' + d.slice(6, 9) + '-' + d.slice(9);
}

function dataIso(iso) {
  var p = String(iso || '').split('-');
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

function linhaTotal(sh) {
  var col = sh.getRange(1, 1, sh.getMaxRows(), 2).getValues();
  for (var i = INI_RES - 1; i < col.length; i++) {
    var t = (String(col[i][0]) + ' ' + String(col[i][1])).toUpperCase();
    if (t.indexOf('TOTAL') >= 0) return i + 1;
  }
  return sh.getLastRow() + 1;
}

// A fórmula da idade é gravada no padrão de separador da planilha (vírgula ou ponto e vírgula)
function gravarIdade(sh, ini, n) {
  if (n <= 0) return;
  var sep = PropertiesService.getDocumentProperties().getProperty('sep') || ',';
  var monta = function (s) {
    var f = [];
    for (var i = 0; i < n; i++) {
      var c = 'C' + (ini + i);
      f.push(['=IF(' + c + '=""' + s + '""' + s + 'IFERROR(DATEDIF(' + c + s + 'TODAY()' + s + '"Y")' + s + '""))']);
    }
    return f;
  };
  var rg = sh.getRange(ini, 4, n, 1);
  rg.setFormulas(monta(sep));
  SpreadsheetApp.flush();
  if (String(sh.getRange(ini, 4).getDisplayValue()).indexOf('#') === 0) {
    sep = sep === ',' ? ';' : ',';
    rg.setFormulas(monta(sep));
    PropertiesService.getDocumentProperties().setProperty('sep', sep);
  }
}

// ================= ESTRUTURA DA COLUNA TIPO =================
function prepararAba(sh) {
  var cab = sh.getRange(2, COL_TIPO, 2, 1);
  var v = cab.getValues();
  if (String(v[1][0]).toUpperCase() === 'TIPO' && sh.getRange(3, COL_TIPO).getBackground() !== '#ffffff') return;
  var ref = sh.getRange(3, 16);
  cab.setValues([['TIPO'], ['TIPO']])
    .setBackground(ref.getBackground())
    .setFontColor(ref.getFontColor())
    .setFontWeight('bold')
    .setFontFamily(ref.getFontFamily())
    .setFontSize(ref.getFontSize())
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');
  sh.setColumnWidth(COL_TIPO, 110);
  aplicarRegras(sh);
}

function aplicarRegras(sh) {
  var total = linhaTotal(sh);
  var n = Math.max(total - INI_RES, 1);
  var dados = sh.getRange(INI_RES, 1, n, NCOL_RES);
  var tipo = sh.getRange(INI_RES, COL_TIPO, n, 1);
  tipo.setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(TIPOS, true).setAllowInvalid(false).build());
  tipo.setHorizontalAlignment('center');
  var outras = sh.getConditionalFormatRules().filter(function (r) {
    var c = r.getBooleanCondition();
    var v = c ? String(c.getCriteriaValues()[0]) : '';
    return v.indexOf('$Q' + INI_RES) < 0;
  });
  outras.push(SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=$Q' + INI_RES + '="Centro Dia"').setBackground(COR.centroDia).setRanges([dados]).build());
  outras.push(SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=$Q' + INI_RES + '="Hotelaria"').setBackground(COR.hotelaria).setRanges([dados]).build());
  sh.setConditionalFormatRules(outras);
}

// ================= RESIDENTES =================
function gravarResidente(sh, d) {
  var fim = linhaTotal(sh) - 1;
  var nome = nomeBase(d.nome);
  var cpf = formatarCpf(d.cpf);
  var linha = -1;
  if (fim >= INI_RES) {
    var vals = sh.getRange(INI_RES, 1, fim - INI_RES + 1, 2).getValues();
    for (var i = 0; i < vals.length; i++) {
      var mesmoCpf = soDigitos(cpf).length === 11 && soDigitos(vals[i][1]) === soDigitos(cpf);
      var mesmoNome = chaveNome(vals[i][0]) !== '' && chaveNome(vals[i][0]) === chaveNome(nome);
      if (mesmoCpf || mesmoNome) { linha = INI_RES + i; break; }
    }
  }
  if (linha < 0) { sh.insertRowBefore(fim + 1); linha = fim + 1; }

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
  if (d.tipo === 'Hotelaria') obs.push('Hotelaria de ' + dataCurta(d.entrada) + (d.saida ? ' a ' + dataCurta(d.saida) : ''));

  var atual = sh.getRange(linha, 1, 1, NCOL_RES).getValues()[0];
  var manter = function (novo, velho) { return (novo === '' || novo === null || novo === undefined) ? velho : novo; };
  // nome com o sufixo do tipo, como a planilha já usava
  var nomeFinal = nome + (d.tipo === 'Centro Dia' ? ' - CENTRO DIA' : d.tipo === 'Hotelaria' ? ' - HOTELARIA' : '');
  sh.getRange(linha, 1, 1, 3).setValues([[nomeFinal, manter(cpf, atual[1]), manter(dataIso(d.nascimento), atual[2])]]);
  sh.getRange(linha, 6, 1, 2).setValues([[manter(entrada, atual[5]), manter(venc, atual[6])]]);
  sh.getRange(linha, 9, 1, 2).setValues([[manter(mesReajuste, atual[8]), manter(valorNumero(d.valor), atual[9])]]);
  if (obs.length) sh.getRange(linha, 16).setValue(obs.join(' | '));
  sh.getRange(linha, COL_TIPO).setValue(d.tipo || 'Permanente');
}

function organizarResidentes(sh) {
  prepararAba(sh);
  var total = linhaTotal(sh);
  var n = total - INI_RES;
  if (n <= 0) return;
  // nomes sem espaços sobrando e tipo preenchido a partir do nome, quando houver sufixo
  var rgNT = sh.getRange(INI_RES, 1, n, 1);
  var nomes = rgNT.getValues();
  var tipos = sh.getRange(INI_RES, COL_TIPO, n, 1).getValues();
  for (var i = 0; i < n; i++) {
    nomes[i][0] = String(nomes[i][0] || '').replace(/\s+/g, ' ').trim();
    var t = tipoDoNome(nomes[i][0]);
    if (t) tipos[i][0] = t;
    else if (nomes[i][0] && !tipos[i][0]) tipos[i][0] = 'Permanente';
  }
  rgNT.setValues(nomes);
  sh.getRange(INI_RES, COL_TIPO, n, 1).setValues(tipos);

  // ordem alfabética nativa da planilha (move linhas inteiras com fórmulas e formatos)
  sh.getRange(INI_RES, 1, n, NCOL_RES).sort({ column: 1, ascending: true });

  gravarIdade(sh, INI_RES, n);
  sh.getRange(total, 10).setFormula('=SUM(J' + INI_RES + ':J' + (total - 1) + ')');
  pintarResidentes(sh, n);
  aplicarRegras(sh);
}

function pintarResidentes(sh, n) {
  var tipos = sh.getRange(INI_RES, COL_TIPO, n, 1).getValues();
  var cores = tipos.map(function (t, i) {
    var base = i % 2 === 0 ? COR.azul : COR.branco;
    var linha = [];
    for (var c = 0; c < NCOL_RES; c++) {
      // Centro Dia e Hotelaria ganham a cor pela formatação condicional da coluna TIPO
      if (c >= 5 && c <= 8) linha.push(COR.financeiro);
      else if (c === 9) linha.push(COR.mensalidade);
      else if (c >= 10 && c <= 14) linha.push(COR.reajustes);
      else linha.push(base);
    }
    return linha;
  });
  sh.getRange(INI_RES, 1, n, NCOL_RES).setBackgrounds(cores);
}

function mapaTipos(shR) {
  var n = linhaTotal(shR) - INI_RES;
  var mapa = {};
  if (n <= 0) return mapa;
  var v = shR.getRange(INI_RES, 1, n, NCOL_RES).getValues();
  v.forEach(function (r) { if (r[0]) mapa[chaveNome(r[0])] = r[COL_TIPO - 1] || tipoDoNome(r[0]) || 'Permanente'; });
  return mapa;
}

// ================= CONTATOS =================
function gravarContato(sh, d) {
  var nome = nomeBase(d.nome) + (d.tipo === 'Centro Dia' ? ' - CENTRO DIA' : d.tipo === 'Hotelaria' ? ' - HOTELARIA' : '');
  var fim = sh.getLastRow();
  var linha = -1;
  if (fim >= INI_CON) {
    var vals = sh.getRange(INI_CON, 1, fim - INI_CON + 1, 1).getValues();
    for (var i = 0; i < vals.length; i++) {
      if (chaveNome(vals[i][0]) !== '' && chaveNome(vals[i][0]) === chaveNome(nome)) { linha = INI_CON + i; break; }
    }
  }
  if (linha < 0) linha = Math.max(fim, INI_CON - 1) + 1;

  var resp = (d.responsaveis || []).filter(function (r) { return r && r.nome; });
  var nomes = resp.map(function (r) { return r.nome + (r.parentesco ? ' (' + r.parentesco + ')' : ''); }).join(' / ');
  var emails = resp.map(function (r) { return r.email; }).filter(Boolean).join(' / ') || (d.emailContratante || '');
  var fones = resp.map(function (r) { return r.telefone; }).filter(Boolean).join(' / ');
  var end = resp.length ? resp[0].endereco : '';
  var atual = sh.getRange(linha, 1, 1, NCOL_CON).getValues()[0];
  var manter = function (novo, velho) { return novo ? novo : velho; };
  sh.getRange(linha, 1, 1, 5).setValues([[nome, manter(nomes, atual[1]), manter(emails, atual[2]), manter(fones, atual[3]), manter(end, atual[4])]]);
}

function organizarContatos(sh, shR) {
  var fim = sh.getLastRow();
  var n = fim - INI_CON + 1;
  if (n <= 0) return;
  var rg = sh.getRange(INI_CON, 1, n, NCOL_CON);
  var linhas = rg.getValues().map(function (r) { r[0] = String(r[0] || '').replace(/\s+/g, ' ').trim(); return r; })
    .filter(function (r) { return r.join('').trim() !== ''; });
  linhas.sort(function (a, b) { return chaveNome(a[0]).localeCompare(chaveNome(b[0]), 'pt-BR'); });
  rg.clearContent();
  rg.setBackground(COR.branco);
  if (!linhas.length) return;
  sh.getRange(INI_CON, 1, linhas.length, NCOL_CON).setValues(linhas);
  pintarContatos(sh, shR);
}

function pintarContatos(sh, shR) {
  var fim = sh.getLastRow();
  var n = fim - INI_CON + 1;
  if (n <= 0) return;
  var mapa = mapaTipos(shR);
  var nomes = sh.getRange(INI_CON, 1, n, 1).getValues();
  var cores = nomes.map(function (r, i) {
    var t = mapa[chaveNome(r[0])] || tipoDoNome(r[0]);
    var c = t === 'Centro Dia' ? COR.centroDia : t === 'Hotelaria' ? COR.hotelaria : (i % 2 === 0 ? COR.azul : COR.branco);
    var l = []; for (var k = 0; k < NCOL_CON; k++) l.push(c); return l;
  });
  sh.getRange(INI_CON, 1, n, NCOL_CON).setBackgrounds(cores);
}

// ================= EDIÇÃO À MÃO =================
// Ao digitar o nome com "- CENTRO DIA" / "- HOTELARIA", o tipo é marcado sozinho.
// Ao mudar o tipo, a aba de contatos acompanha a cor.
function onEdit(e) {
  try {
    var sh = e.range.getSheet();
    var nomeAba = sh.getName();
    if (!/^Solary [12]$/.test(nomeAba)) return;
    var r0 = e.range.getRow(), c0 = e.range.getColumn();
    if (r0 < INI_RES) return;
    var nr = e.range.getNumRows();
    if (c0 <= 1 && c0 + e.range.getNumColumns() - 1 >= 1) {
      var nomes = sh.getRange(r0, 1, nr, 1).getValues();
      var tipos = sh.getRange(r0, COL_TIPO, nr, 1).getValues();
      for (var i = 0; i < nr; i++) {
        var t = tipoDoNome(nomes[i][0]);
        if (t) tipos[i][0] = t;
        else if (nomes[i][0] && !tipos[i][0]) tipos[i][0] = 'Permanente';
      }
      sh.getRange(r0, COL_TIPO, nr, 1).setValues(tipos);
    }
    var u = nomeAba.slice(-1);
    pintarContatos(e.source.getSheetByName('Contatos Solary ' + u), sh);
  } catch (err) {}
}

// ================= MENU =================
function onOpen() {
  SpreadsheetApp.getUi().createMenu('Solary')
    .addItem('Organizar ordem e cores', 'organizarTudo')
    .addToUi();
}

function organizarTudo() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  ['1', '2'].forEach(function (u) {
    var shR = ss.getSheetByName('Solary ' + u);
    organizarResidentes(shR);
    organizarContatos(ss.getSheetByName('Contatos Solary ' + u), shR);
  });
}
