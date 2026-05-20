// ═══════════════════════════════════════════════════════════════════════════════
// CHAMADA AM — SRJ9 — Backend Google Apps Script
// ═══════════════════════════════════════════════════════════════════════════════
// MAPEAMENTO DE COLUNAS (0-based para arrays, 1-based para getRange):
//   A(0/1)=CHECK IN   B(1/2)=CHECK OUT   C(2/3)=TMC      D(3/4)=TIPO
//   E(4/5)=ETA        F(5/6)=VAGA        G(6/7)=Onda      I(8/9)=Rota
//   L(11/12)=Transp   N(13/14)=Modal     S(18/19)=Motorista
//   V(21/22)=OOT_status  W(22/23)=HORA_CHAMADO  X(23/24)=OOT_perdido  Y(24/25)=TMC_perdido  Z(25/26)=JUSTIF  AA(26/27)=OBS
// ═══════════════════════════════════════════════════════════════════════════════

var SHEET_ID = '1iJX4sIMvWnh-XdrdZFK3mljXB-hOQentQ9j2sn8OGsg';

// ── doGet ─────────────────────────────────────────────────────────────────────
function doOptions(e) {
  return ContentService.createTextOutput('')
    .setMimeType(ContentService.MimeType.TEXT);
}

function doGet(e) {
  var p = e && e.parameter ? e.parameter : {};

  // ── API mode: ?fn=functionName&param1=val1... ──────────────────
  if (p.fn) {
    var result, error;
    try {
      switch(p.fn) {
        case 'getUsuario':
          result = getUsuario(); break;
        case 'getTurnosServer':
          result = getTurnosServer(); break;
        case 'getDadosServer':
          result = getDadosServer(p.turno); break;
        case 'getDashboardServer':
          result = getDashboardServer(p.turno); break;
        case 'salvarCheckServer':
          result = salvarCheckServer(Number(p.rowIndex), p.turno, p.hora, p.justif, p.operador); break;
        case 'removerCheckServer':
          result = removerCheckServer(Number(p.rowIndex), p.turno); break;
        case 'salvarObsMeliServer':
          result = salvarObsMeliServer(Number(p.rowIndex), p.turno, p.obs); break;
        case 'salvarJustifServer':
          result = salvarJustifServer(Number(p.rowIndex), p.turno, p.justif); break;
        case 'registrarAutoCheckHistorico':
          result = registrarAutoCheckHistorico(Number(p.rowIndex), p.turno, p.horaCheck); break;
        case 'atualizarTMCStats':
          result = atualizarTMCStats(Number(p.rowIndex), p.turno); break;
        case 'ping':
          result = {ok:true, ts: new Date().toISOString()}; break;
        default:
          error = 'Função não encontrada: ' + p.fn;
      }
    } catch(err) {
      error = err.message;
    }
    var json = error
      ? JSON.stringify({error: error})
      : JSON.stringify({result: result});
    return ContentService
      .createTextOutput(json)
      .setMimeType(ContentService.MimeType.JSON);
  }

  // ── Web App mode: serve HTML ───────────────────────────────────
  return HtmlService
    .createHtmlOutputFromFile('index')
    .setTitle('Chamada AM')
    .addMetaTag('viewport','width=device-width,initial-scale=1.0,maximum-scale=1.0')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ── Helpers ───────────────────────────────────────────────────────────────────
function getUsuario() {
  try {
    var email = Session.getActiveUser().getEmail() || Session.getEffectiveUser().getEmail() || 'anonimo';
    var nome  = email.split('@')[0].replace(/[._]/g,' ').replace(/\b\w/g,function(c){return c.toUpperCase();});
    return {email:email, nome:nome};
  } catch(e) { return {email:'anonimo', nome:'Usuario'}; }
}

function getTurnosServer() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  return ss.getSheets().map(function(s){return s.getName();}).filter(function(n){return /ciclo/i.test(n);});
}

function primeiraAba(ss) {
  var list = ss.getSheets().map(function(s){return s.getName();}).filter(function(n){return /ciclo/i.test(n);});
  return list.length ? list[0] : '';
}

// Encontra a linha onde os dados começam (primeiro "Onda X" em col G=índice 6)
// Dados sempre a partir da linha 5 da planilha (indice 4, base-0)
function findDataStart(rows) {
  var minIdx = 4; // linha 5 = indice 4
  for (var i = minIdx; i < Math.min(20, rows.length); i++) {
    var g = (rows[i][6]||'').toString().trim();
    if (g.indexOf('Onda') === 0) return i;
  }
  return minIdx; // fallback: linha 5
}

// Formata Date/número/string → "HH:MM:SS"
function fmtHora(val) {
  if (!val && val !== 0) return '';
  var s = val.toString().trim();
  if (!s || s.indexOf('#') >= 0) return '';
  if (val instanceof Date) {
    return ('0'+val.getHours()).slice(-2)+':'+('0'+val.getMinutes()).slice(-2)+':'+('0'+val.getSeconds()).slice(-2);
  }
  if (typeof val === 'number') {
    var tot = Math.round(val * 86400);
    return ('0'+Math.floor(tot/3600)).slice(-2)+':'+('0'+Math.floor((tot%3600)/60)).slice(-2)+':'+('0'+(tot%60)).slice(-2);
  }
  return s;
}

// "HH:MM:SS" ou "HH:MM" → segundos totais
function toSeconds(t) {
  if (!t) return 0;
  var p = t.toString().trim().split(':').map(function(x){return parseInt(x)||0;});
  return (p[0]||0)*3600 + (p[1]||0)*60 + (p[2]||0);
}

// Segundos → "MM:SS" (ou "H:MM:SS" se >= 1h)
function secsToMMSS(sec) {
  if (sec < 0) sec = 0;
  var h = Math.floor(sec/3600), m = Math.floor((sec%3600)/60), s = Math.floor(sec%60);
  if (h > 0) return h+':'+('0'+m).slice(-2)+':'+('0'+s).slice(-2);
  return ('0'+m).slice(-2)+':'+('0'+s).slice(-2);
}

// "HH:MM" → minutos totais (para calcStatus)
function toMins(t) {
  if (!t) return null;
  var clean = t.toString().replace(/[^0-9:]/g,'');
  var p = clean.split(':').map(Number);
  if (isNaN(p[0])) return null;
  return p[0]*60 + (p[1]||0);
}

function calcStatus(hora, eta) {
  if (!hora || !eta) return 'on time';
  var parts = eta.split(' - ');
  if (parts.length < 2) return 'on time';
  var ini = toMins(parts[0].trim()), fim = toMins(parts[1].trim()), chk = toMins(hora);
  if (chk===null || ini===null || fim===null) return 'on time';
  if (chk < ini) return 'advance';
  if (chk > fim) return 'delay';
  return 'on time';
}

// ── calcWX: calcula TMC perdido (W) e atraso OOT (X) ────────────────────────
function calcWX(checkin, checkout, eta) {
  var w = '', x = '';
  if (!checkin || !checkout) return {w:w, x:x};

  var cinSec  = toSeconds(checkin);
  var coutSec = toSeconds(checkout);
  if (coutSec < cinSec) coutSec += 86400; // virada de meia-noite
  var tmcSec  = coutSec - cinSec;

  // W: minutos acima de 30min em MM:SS
  if (tmcSec > 1800) w = secsToMMSS(tmcSec - 1800);

  // X: quanto o checkout passou do fim da janela ETA
  if (eta) {
    var parts = eta.split(' - ');
    if (parts.length >= 2) {
      var etaFimSec = toSeconds(parts[1].trim());
      if (etaFimSec > 0) {
        // Ajuste virada de meia-noite
        if (coutSec < etaFimSec - 3600) coutSec += 86400;
        if (coutSec > etaFimSec) x = secsToMMSS(coutSec - etaFimSec);
      }
    }
  }
  return {w:w, x:x};
}

// ── getDadosServer ────────────────────────────────────────────────────────────
// Retorna para o frontend:
// [0]=onda  [1]=rota   [2]=transp  [3]=modal   [4]=tipoRota [5]=tipoM
// [6]=eta   [7]=vaga   [8]=motorista [9]=horaChk [10]=status [11]=justif
// [12]=rowIndex [13]=obsMeli(Y) [14]=checkin(A) [15]=checkout(B)
// [16]=tmcPerdidoW    [17]=atrasoOndaX
function getDadosServer(turno) {
  var ss      = SpreadsheetApp.openById(SHEET_ID);
  var nomeAba = turno || primeiraAba(ss);
  var aba     = ss.getSheetByName(nomeAba);
  if (!aba) throw new Error('Aba nao encontrada: ' + nomeAba);

  var rows      = aba.getDataRange().getValues();
  var ds        = findDataStart(rows);
  var result    = [];
  var wxUpdates = []; // {row, w, x} — linhas para gravar W/X em lote

  for (var i = ds; i < rows.length; i++) {
    var r = rows[i];

    // Colunas fixas da planilha
    var onda     = (r[6] ||'').toString().trim();  // G
    var rota     = (r[8] ||'').toString().trim();  // I
    var transp   = (r[11]||'').toString().trim();  // L
    var modal    = (r[13]||'').toString().trim();  // N
    var tipoRota = (r[15]||'').toString().trim();  // P
    var tipoM    = (r[3] ||'').toString().trim();  // D
    var eta      = (r[4] ||'').toString().trim();  // E
    var vaga     = (r[5] ||'').toString().trim();  // F
    var motRaw   = r[18];                          // S
    var motorista = (motRaw instanceof Date || typeof motRaw === 'number') ? '' : (motRaw||'').toString().trim();
    var horaChk  = fmtHora(r[22]);                // W(23) = HORA CHAMADO
    var status   = (r[21]||'').toString().trim();  // V(22) = OOT status
    var justif   = (r[25]||'').toString().trim();  // Z(26) = JUSTIF
    var wSheet   = (r[24]||'').toString().trim();  // Y(25) = TMC_perdido
    var xSheet   = (r[23]||'').toString().trim();  // X(24) = OOT_perdido
    var obsMeli  = (r[26]||'').toString().trim();  // AA(27) = Obs Meli

    // CHECK IN (A) e CHECK OUT (B)
    var checkin  = '', checkout = '';
    var _cin = fmtHora(r[0]);
    if (_cin && _cin !== '00:00:00') checkin = _cin;
    var _cout = fmtHora(r[1]);
    if (_cout && _cout !== '00:00:00') checkout = _cout;

    if (!rota || !onda) continue;
    if (tipoM !== '1 Volumosos' && tipoM !== '2 Leves') tipoM = '2 Leves';

    // ── Garante W e X para toda rota com checkin + checkout ──────────────────
    var wFinal = wSheet, xFinal = xSheet;
    if (checkin && checkout) {
      // Recalcula sempre — garante consistência mesmo se estava vazio ou errado
      var wx = calcWX(checkin, checkout, eta);
      wFinal = wx.w;
      xFinal = wx.x;
      // Só grava na planilha se mudou (evita writes desnecessários)
      if (wFinal !== wSheet || xFinal !== xSheet) {
        wxUpdates.push({row: i + 1, w: wFinal, x: xFinal}); // row é 1-based
      }
    }

    result.push([
      onda, rota, transp, modal, tipoRota, tipoM, eta, vaga, motorista,
      horaChk, status, justif, i + 1, obsMeli, checkin, checkout,
      wFinal, xFinal
    ]);
  }

  // ── Gravar W/X em lote (uma chamada por linha, em paralelo onde possível) ──
  if (wxUpdates.length > 0) {
    try {
      wxUpdates.forEach(function(item) {
        var wR = aba.getRange(item.row, 25); // Y(25) = TMC_perdido
        wR.setNumberFormat('@STRING@');
        wR.setValue(item.w);
        var yR = aba.getRange(item.row, 25); // Y = col 25 = OOT_perdido
        yR.setNumberFormat('@STRING@');
        yR.setValue(item.x);
      });
      SpreadsheetApp.flush();
      Logger.log('getDadosServer: W/X atualizados em ' + wxUpdates.length + ' linhas');
    } catch(e) {
      Logger.log('getDadosServer ERRO ao gravar W/X: ' + e.message);
    }
  }

  return result;
}

// ── salvarCheckServer ─────────────────────────────────────────────────────────
function salvarCheckServer(rowIndex, turno, horaCheck, justif, operador) {
  var ss      = SpreadsheetApp.openById(SHEET_ID);
  var nomeAba = turno || primeiraAba(ss);
  var aba     = ss.getSheetByName(nomeAba);
  if (!aba) return {ok:false, erro:'Aba nao encontrada'};

  var eta    = (aba.getRange(rowIndex, 5).getValue()||'').toString().trim(); // E
  var status = horaCheck ? calcStatus(horaCheck, eta) : '';

  // Grava V(22)=OOT_status, W(23)=CHAMADO, Z(26)=JUSTIF
  aba.getRange(rowIndex, 22).setValue(status);
  var cel = aba.getRange(rowIndex, 23);
  cel.setNumberFormat('@STRING@');
  cel.setValue(horaCheck || '');
  // Justif: só sobrescreve se vier valor novo; preserva o que já está na planilha
  var justifAtual = (aba.getRange(rowIndex, 26).getValue()||'').toString().trim();
  var justifFinal = (justif && justif.trim()) ? justif.trim() : justifAtual;
  aba.getRange(rowIndex, 26).setValue(justifFinal); // Z(26) = JUSTIF

  // Recalcula e grava W(23)=TMC_perdido e Y(25)=OOT_perdido (X=24 é JUSTIF, não sobrescrever)
  try {
    var checkin  = fmtHora(aba.getRange(rowIndex, 1).getValue()); // A
    var checkout = fmtHora(aba.getRange(rowIndex, 2).getValue()); // B
    if (checkin && checkin !== '00:00:00' && checkout && checkout !== '00:00:00') {
      var wx = calcWX(checkin, checkout, eta);
      var yR = aba.getRange(rowIndex, 25); // Y(25) = TMC_perdido
      yR.setNumberFormat('@STRING@');
      yR.setValue(wx.w);
      var xR = aba.getRange(rowIndex, 24); // X(24) = OOT_perdido
      xR.setNumberFormat('@STRING@');
      xR.setValue(wx.x);
    }
  } catch(e) { Logger.log('salvarCheckServer - erro W/X: ' + e.message); }

  SpreadsheetApp.flush();
  try { logHistorico(ss, aba, rowIndex, nomeAba, horaCheck, status, justif, operador); } catch(e) {}
  var rota = (aba.getRange(rowIndex, 9).getValue()||'').toString().trim();
  return {ok:true, status:status, rota:rota};
}

// ── removerCheckServer ────────────────────────────────────────────────────────
function removerCheckServer(rowIndex, turno) {
  var ss      = SpreadsheetApp.openById(SHEET_ID);
  var nomeAba = turno || primeiraAba(ss);
  var aba     = ss.getSheetByName(nomeAba);
  if (!aba) return {ok:false};
  aba.getRange(rowIndex, 22, 1, 4).clearContent(); // limpa V,W,X,Y (preserva Z=justif)
  SpreadsheetApp.flush();
  return {ok:true};
}


// ── salvarJustifServer — salva só justificativa (col X=24), sem tocar hora/check ─
function salvarJustifServer(rowIndex, turno, justif) {
  var ss      = SpreadsheetApp.openById(SHEET_ID);
  var nomeAba = turno || primeiraAba(ss);
  var aba     = ss.getSheetByName(nomeAba);
  if (!aba) return {ok:false};
  aba.getRange(rowIndex, 26).setValue(justif || ''); // Z(26) = JUSTIF
  SpreadsheetApp.flush();
  Logger.log('salvarJustifServer: row=' + rowIndex + ' justif=' + justif);
  return {ok:true};
}

// ── salvarObsMeliServer — grava na coluna Y (col 25) ─────────────────────────
function salvarObsMeliServer(rowIndex, turno, obs) {
  var ss      = SpreadsheetApp.openById(SHEET_ID);
  var nomeAba = turno || primeiraAba(ss);
  var aba     = ss.getSheetByName(nomeAba);
  if (!aba) return {ok:false};
  var cel = aba.getRange(rowIndex, 27); // AA = col 27 = Obs Meli
  cel.setNumberFormat('@STRING@');
  cel.setValue(obs || '');
  SpreadsheetApp.flush();
  Logger.log('salvarObsMeliServer: row=' + rowIndex + ' obs=' + obs);
  return {ok:true, row:rowIndex, obs:obs};
}

// ── atualizarTMCStats — grava W/X para uma linha específica ──────────────────
function atualizarTMCStats(rowIndex, turno) {
  var ss      = SpreadsheetApp.openById(SHEET_ID);
  var nomeAba = turno || primeiraAba(ss);
  var aba     = ss.getSheetByName(nomeAba);
  if (!aba) return {ok:false};

  var checkin  = fmtHora(aba.getRange(rowIndex, 1).getValue()); // A
  var checkout = fmtHora(aba.getRange(rowIndex, 2).getValue()); // B
  var eta      = (aba.getRange(rowIndex, 5).getValue()||'').toString().trim(); // E

  if (!checkin || checkin === '00:00:00' || !checkout || checkout === '00:00:00')
    return {ok:false, msg:'sem checkin/checkout'};

  var wx = calcWX(checkin, checkout, eta);
  var yR = aba.getRange(rowIndex, 25); // Y(25) = TMC_perdido
  yR.setNumberFormat('@STRING@');
  yR.setValue(wx.w);
  var xR = aba.getRange(rowIndex, 24); // X(24) = OOT_perdido
  xR.setNumberFormat('@STRING@');
  xR.setValue(wx.x);
  SpreadsheetApp.flush();
  return {ok:true, w:wx.w, x:wx.x};
}

// ── registrarAutoCheckHistorico ───────────────────────────────────────────────
function registrarAutoCheckHistorico(rowIndex, turno, horaCheck) {
  var ss      = SpreadsheetApp.openById(SHEET_ID);
  var nomeAba = turno || primeiraAba(ss);
  var aba     = ss.getSheetByName(nomeAba);
  if (!aba) return {ok:false};
  var eta    = (aba.getRange(rowIndex, 5).getValue()||'').toString().trim();
  var status = calcStatus(horaCheck, eta);
  try { logHistorico(ss, aba, rowIndex, nomeAba, horaCheck, status, '', 'automatico'); } catch(e) {}
  return {ok:true, status:status};
}

// ── logHistorico ──────────────────────────────────────────────────────────────
function logHistorico(ss, aba, rowIndex, turno, hora, status, justif, operador) {
  try {
    var hist = ss.getSheetByName('_Historico');
    if (!hist) {
      hist = ss.insertSheet('_Historico');
      hist.appendRow(['Data','Turno','Onda','Rota','Transportadora','Modal','ETA','Hora','Status','Ocorrencia','Operador','Timestamp']);
      hist.getRange(1,1,1,12).setFontWeight('bold').setBackground('#3483FA').setFontColor('#FFE600');
      hist.setFrozenRows(1);
    }
    var r   = aba.getRange(rowIndex, 1, 1, 27).getValues()[0];
    var now = new Date();
    hist.appendRow([
      Utilities.formatDate(now, Session.getScriptTimeZone(), 'dd/MM/yyyy'),
      turno,
      (r[6]||'').toString().trim(),   // G=Onda
      (r[8]||'').toString().trim(),   // I=Rota
      (r[11]||'').toString().trim(),  // L=Transp
      (r[13]||'').toString().trim(),  // N=Modal
      (r[4]||'').toString().trim(),   // E=ETA
      hora||'', status||'', justif||'', operador||'', now
    ]);
  } catch(e) { Logger.log('logHistorico: ' + e.message); }
}

// ── recalcularTMCTodas — backfill W/X para todo o ciclo ──────────────────────
// Executar manualmente no Apps Script para preencher dados históricos
function recalcularTMCTodas(turno) {
  var ss      = SpreadsheetApp.openById(SHEET_ID);
  var nomeAba = turno || primeiraAba(ss);
  var aba     = ss.getSheetByName(nomeAba);
  if (!aba) { Logger.log('Aba nao encontrada: ' + nomeAba); return; }

  var rows = aba.getDataRange().getValues();
  var ds   = findDataStart(rows);
  var wVals = [], xVals = [];
  var updated = 0;

  for (var i = ds; i < rows.length; i++) {
    var r        = rows[i];
    var checkin  = fmtHora(r[0]); // A
    var checkout = fmtHora(r[1]); // B
    var eta      = (r[4]||'').toString().trim(); // E
    var w = '', x = '';
    if (checkin && checkin !== '00:00:00' && checkout && checkout !== '00:00:00') {
      var wx = calcWX(checkin, checkout, eta);
      w = wx.w; x = wx.x;
      updated++;
    }
    wVals.push([w]);
    xVals.push([x]);
  }

  if (wVals.length > 0) {
    var wRange = aba.getRange(ds + 1, 25, wVals.length, 1); // Y=col25=TMC_perdido
    wRange.setNumberFormat('@STRING@');
    wRange.setValues(wVals);
    var xRange = aba.getRange(ds + 1, 24, xVals.length, 1); // X=col24=OOT_perdido
    xRange.setNumberFormat('@STRING@');
    xRange.setValues(xVals);
    SpreadsheetApp.flush();
  }

  Logger.log('recalcularTMCTodas [' + nomeAba + ']: ' + updated + '/' + wVals.length + ' linhas com W/X');
  return {ok:true, updated:updated, total:wVals.length};
}

// Roda para TODOS os ciclos de uma vez
function recalcularTMCTodasAbas() {
  var ss   = SpreadsheetApp.openById(SHEET_ID);
  var abas = ss.getSheets().map(function(s){return s.getName();}).filter(function(n){return /ciclo/i.test(n);});
  abas.forEach(function(nome) {
    var r = recalcularTMCTodas(nome);
    Logger.log(nome + ': ' + JSON.stringify(r));
  });
}

// ── getDashboardServer ────────────────────────────────────────────────────────
function getDashboardServer(turno) {
  var ss      = SpreadsheetApp.openById(SHEET_ID);
  var allAbas = ss.getSheets().map(function(s){return s.getName();}).filter(function(n){return /ciclo/i.test(n);});
  if (!turno) turno = allAbas[0] || '';

  var por_transp = {}, por_onda = {}, por_modal = {}, por_ciclo = {}, transp_geral = {};
  var sprTot = 0, sprCnt = 0;

  allAbas.forEach(function(nomeAba) {
    var isAtual = (nomeAba === turno);
    var aba = ss.getSheetByName(nomeAba);
    if (!aba) return;
    var rows = aba.getDataRange().getValues();
    var ds   = findDataStart(rows);
    if (!por_ciclo[nomeAba]) por_ciclo[nomeAba] = {nome:nomeAba,total:0,called:0,ontime:0,advance:0,delay:0,por_transp:{},por_onda:{}};

    for (var i = ds; i < rows.length; i++) {
      var r         = rows[i];
      var onda      = (r[6] ||'').toString().trim();
      var rota      = (r[8] ||'').toString().trim();
      var transp    = (r[11]||'').toString().trim();
      var modal     = (r[13]||'').toString().trim();
      var tipoM     = (r[3] ||'').toString().trim();
      var eta       = (r[4] ||'').toString().trim();
      var vaga      = (r[5] ||'').toString().trim();
      var spr       = parseFloat(r[16])||0;
      var motRaw    = r[18];
      var motorista = (motRaw instanceof Date || typeof motRaw === 'number') ? '' : (motRaw||'').toString().trim();
      var horaChk   = fmtHora(r[22]);               // W(23) = HORA CHAMADO
      var checkinA  = fmtHora(r[0]); if (checkinA === '00:00:00') checkinA = '';
      var efetivo   = horaChk || checkinA;
      var status    = (r[21]||'').toString().trim();  // V(22) = OOT status
      var justif    = (r[25]||'').toString().trim();  // Z(26) = JUSTIF
      var tipoRota  = (r[15]||'').toString().trim();
      if (!rota || !onda) continue;
      if (tipoM !== '1 Volumosos' && tipoM !== '2 Leves') continue;

      por_ciclo[nomeAba].total++;
      if (efetivo) {
        por_ciclo[nomeAba].called++;
        if (status==='on time') por_ciclo[nomeAba].ontime++;
        if (status==='advance') por_ciclo[nomeAba].advance++;
        if (status==='delay')   por_ciclo[nomeAba].delay++;
      }
      if (!por_ciclo[nomeAba].por_transp[transp]) por_ciclo[nomeAba].por_transp[transp]={total:0,called:0,delay:0,advance:0,ontime:0};
      por_ciclo[nomeAba].por_transp[transp].total++;
      if (efetivo) {
        por_ciclo[nomeAba].por_transp[transp].called++;
        if (status==='delay')   por_ciclo[nomeAba].por_transp[transp].delay++;
        if (status==='advance') por_ciclo[nomeAba].por_transp[transp].advance++;
        if (status==='on time') por_ciclo[nomeAba].por_transp[transp].ontime++;
      }
      if (!por_ciclo[nomeAba].por_onda[onda]) por_ciclo[nomeAba].por_onda[onda]={total:0,called:0,ontime:0,advance:0,delay:0,eta:eta};
      por_ciclo[nomeAba].por_onda[onda].total++;
      if (efetivo) {
        por_ciclo[nomeAba].por_onda[onda].called++;
        if (status==='on time') por_ciclo[nomeAba].por_onda[onda].ontime++;
        if (status==='advance') por_ciclo[nomeAba].por_onda[onda].advance++;
        if (status==='delay')   por_ciclo[nomeAba].por_onda[onda].delay++;
      }

      if (!transp_geral[transp]) transp_geral[transp]={total:0,called:0,delay:0,delayNoJustif:0,advance:0,ontime:0,sprTot:0,sprCnt:0};
      transp_geral[transp].total++;
      if (spr>0){transp_geral[transp].sprTot+=spr;transp_geral[transp].sprCnt++;}
      if (efetivo) {
        transp_geral[transp].called++;
        if (status==='delay'){transp_geral[transp].delay++;if(!justif)transp_geral[transp].delayNoJustif++;}
        if (status==='advance') transp_geral[transp].advance++;
        if (status==='on time') transp_geral[transp].ontime++;
      }

      if (!isAtual) continue;

      if (spr>0){sprTot+=spr;sprCnt++;}
      if (!por_transp[transp]) por_transp[transp]={total:0,called:0,delay:0,delayNoJustif:0,advance:0,ontime:0,sprTot:0,sprCnt:0,rotas:[]};
      por_transp[transp].total++;
      if (spr>0){por_transp[transp].sprTot+=spr;por_transp[transp].sprCnt++;}
      por_transp[transp].rotas.push({rota:rota,onda:onda,eta:eta,vaga:vaga,motorista:motorista,modal:modal,tipoM:tipoM,horaCheck:efetivo,status:status,justif:justif});
      if (efetivo) {
        por_transp[transp].called++;
        if (status==='delay'){por_transp[transp].delay++;if(!justif)por_transp[transp].delayNoJustif++;}
        if (status==='advance') por_transp[transp].advance++;
        if (status==='on time') por_transp[transp].ontime++;
      }

      if (!por_onda[onda]) por_onda[onda]={total:0,called:0,eta:eta,ontime:0,advance:0,delay:0,delayNoJustif:0,rotas:[]};
      por_onda[onda].total++;
      por_onda[onda].rotas.push({rota:rota,transp:transp,vaga:vaga,modal:modal,horaCheck:efetivo,status:status,justif:justif});
      if (efetivo) {
        por_onda[onda].called++;
        if (status==='delay'){por_onda[onda].delay++;if(!justif)por_onda[onda].delayNoJustif++;}
        if (status==='advance') por_onda[onda].advance++;
        if (status==='on time') por_onda[onda].ontime++;
      }

      var mk = modal.split(' ').slice(0,2).join(' ') || 'Outros';
      if (!por_modal[mk]) por_modal[mk]={total:0,called:0,delay:0,advance:0,ontime:0};
      por_modal[mk].total++;
      if (efetivo) {
        por_modal[mk].called++;
        if (status==='delay')   por_modal[mk].delay++;
        if (status==='advance') por_modal[mk].advance++;
        if (status==='on time') por_modal[mk].ontime++;
      }
    }
  });

  Object.keys(por_transp).forEach(function(k){var d=por_transp[k];d.sprMedia=d.sprCnt?Math.round(d.sprTot/d.sprCnt):0;});
  Object.keys(transp_geral).forEach(function(k){var d=transp_geral[k];d.sprMedia=d.sprCnt?Math.round(d.sprTot/d.sprCnt):0;});

  return {por_transp:por_transp,por_onda:por_onda,por_modal:por_modal,por_ciclo:por_ciclo,transp_geral:transp_geral,turnoAtual:turno,sprMedia:sprCnt?Math.round(sprTot/sprCnt):0};
}

// ── diagnostico — rodar no Apps Script para debugar ──────────────────────────
function diagnostico() {
  var r = {};
  try {
    var ss = SpreadsheetApp.openById(SHEET_ID);
    r.planilha   = ss.getName();
    r.abas       = ss.getSheets().map(function(s){return s.getName();});
    r.abasCiclo  = r.abas.filter(function(n){return /ciclo/i.test(n);});
    r.usuario    = Session.getActiveUser().getEmail();
    if (r.abasCiclo.length > 0) {
      var dados = getDadosServer(r.abasCiclo[0]);
      r.totalRotas = dados.length;
      r.comCheckin  = dados.filter(function(d){return d[14];}).length;
      r.comCheckout = dados.filter(function(d){return d[15];}).length;
      r.comW        = dados.filter(function(d){return d[16];}).length;
      r.comX        = dados.filter(function(d){return d[17];}).length;
      if (dados.length > 0) r.exemplo = {
        onda:dados[0][0], rota:dados[0][1], checkin:dados[0][14],
        checkout:dados[0][15], w:dados[0][16], x:dados[0][17], obs:dados[0][13]
      };
    }
  } catch(e) { r.erro = e.message; }
  Logger.log(JSON.stringify(r, null, 2));
  return r;
}
