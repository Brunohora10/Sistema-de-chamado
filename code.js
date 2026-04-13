/** ==============================
 *  Sistema de Chamados CadÚnico — BACKEND (Code.gs)
 * ===============================*/

/** ✅ ID DA PLANILHA */
const SPREADSHEET_ID = '1B5XniHj5gvR9oa7TffviVVhgVkeOCURsJuJRLSEfDtI';

/** Configurações gerais */
const TZ = 'America/Sao_Paulo';
const SHEET_PRIMARY   = 'Chamados';
const SHEET_ALTERNATE = 'Senhas';
const HEADERS = [
  'id','numero','nome','cpf','tipo','servico','bairro','setor','status',
  'timestamp','chamadoEm','iniciadoEm','finalizadoEm',
  'guiche','tempoEspera','tempoAtendimento','data',
  'calls'
];

/** Tempo máximo de atendimento/chamada em minutos */
const MAX_ATEND_MIN = 40;

/** Prefixos por tipo (persistidos em "numero") */
const PREFIX = { normal: 'N', prioritario: 'P', agendamento: 'G' };

/** Chaves (ordem de chamada) */
const CALL_ORDER_KEY = 'CALL_ORDER';
const CALL_POS_KEY   = 'CALL_POS';
const CALL_DAY_KEY   = 'CALL_DAY';

/** NOVO: chave das configurações do Admin no Script Properties */
const ADMIN_CFG_KEY  = 'ADMIN_CFG_JSON';

/** Cache curto dos dados (melhora polling sem perder consistência) */
const CACHE_TTL_SEC = 8;
const CACHE_ALL_PREFIX = 'ALL_ROWS_';
const CACHE_DAY_PREFIX = 'DAY_ROWS_';

/** NOVO: tipos de setor/local de atendimento */
const SETORES_VALIDOS = ['guiche', 'servico_social', 'coordenacao'];

/* ================= WEB APP ================== */
function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('Sistema de Chamados CadÚnico')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}
function __auth() { const s = _sheet(); s.getLastRow(); return 'ok'; }

/* ============== HELPERS PLANILHA ============== */
function _ss() {
  try { return SpreadsheetApp.openById(SPREADSHEET_ID); } catch (_) {}
  try { return SpreadsheetApp.openByUrl('https://docs.google.com/spreadsheets/d/'+SPREADSHEET_ID+'/edit'); } catch (_) {}
  throw new Error('PLANILHA_INACESSIVEL');
}
function _sheet() {
  const ss = _ss();
  let s = ss.getSheetByName(SHEET_PRIMARY)
        || ss.getSheetByName(SHEET_ALTERNATE);
  if (!s) s = ss.insertSheet(SHEET_PRIMARY);
  _ensureHeaders(s);
  return s;
}
function _ensureHeaders(sheet) {
  const lastCol = Math.max(1, sheet.getLastColumn());
  const row = sheet.getRange(1,1,1,lastCol).getValues()[0].map(v=>String(v||'').trim());
  const vazia = (row.join('')==='' && sheet.getLastRow()===0);
  if (vazia) {
    sheet.getRange(1,1,1,HEADERS.length).setValues([HEADERS]);
    sheet.setFrozenRows(1);
    return;
  }
  const cur = row.slice(); let changed=false;
  HEADERS.forEach(h=>{ if(!cur.includes(h)){ cur.push(h); changed=true; }});
  if (changed) {
    sheet.getRange(1,1,1,cur.length).setValues([cur]);
    sheet.setFrozenRows(1);
  }
}
function _headerMap(sheet){
  const lastCol = Math.max(1, sheet.getLastColumn());
  const row = sheet.getRange(1,1,1,lastCol).getValues()[0];
  const m={}; row.forEach((h,i)=> m[String(h).trim()] = i);
  return m;
}
function _nowIso(){ return Utilities.formatDate(new Date(), TZ, "yyyy-MM-dd'T'HH:mm:ssXXX"); }
function _today(){  return Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd'); }
function _cacheKeyAll(day){ return CACHE_ALL_PREFIX + String(day||_today()); }
function _cacheKeyDay(day){ return CACHE_DAY_PREFIX + String(day||_today()); }
function _cache(){ return CacheService.getScriptCache(); }
function _invalidateDayCache(day){
  try{ _cache().remove(_cacheKeyAll(day)); }catch(_){ }
  try{ _cache().remove(_cacheKeyDay(day)); }catch(_){ }
}
function _invalidateTodayCache(){ _invalidateDayCache(_today()); }
function _saveAllCache(day, rows){
  try{ _cache().put(_cacheKeyAll(day), JSON.stringify(rows||[]), CACHE_TTL_SEC); }catch(_){ }
}
function _saveDayCache(day, rows){
  try{ _cache().put(_cacheKeyDay(day), JSON.stringify(rows||[]), CACHE_TTL_SEC); }catch(_){ }
}
function _loadAllCache(day){
  try{
    const raw = _cache().get(_cacheKeyAll(day));
    if (!raw) return null;
    const arr = JSON.parse(raw);
    return Array.isArray(arr) ? arr : null;
  }catch(_){ return null; }
}
function _loadDayCache(day){
  try{
    const raw = _cache().get(_cacheKeyDay(day));
    if (!raw) return null;
    const arr = JSON.parse(raw);
    return Array.isArray(arr) ? arr : null;
  }catch(_){ return null; }
}
function _digits(s){ return String(s||'').replace(/\D/g,''); }
function _pad(n,w=3){ return String(n).padStart(w,'0'); }
function _toDateKey(v){
  if (!v) return '';
  if (Object.prototype.toString.call(v)==='[object Date]' && !isNaN(v)) {
    return Utilities.formatDate(v, TZ, 'yyyy-MM-dd');
  }
  const s = String(v).trim();
  const mIso = s.match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (mIso) return `${mIso[1]}-${mIso[2]}-${mIso[3]}`;
  const mBr = s.match(/^(\d{2})\/(\d{2})\/(\d{4})$/);
  if (mBr) return `${mBr[3]}-${mBr[2]}-${mBr[1]}`;
  try { const d = new Date(s); if (!isNaN(d)) return Utilities.formatDate(d, TZ, 'yyyy-MM-dd'); } catch(_){}
  return '';
}

/** Normaliza o setor/local de atendimento */
function _normalizeSetor(raw){
  const s = String(raw||'').toLowerCase().trim();
  if (!s || s === 'guiche' || s === 'guichê') return 'guiche';
  if (s === 'servico_social' || s === 'serviço_social' || s === 'servico social' || s === 'serviço social') {
    return 'servico_social';
  }
  if (s === 'coordenacao' || s === 'coordenação' || s === 'coordenadora' || s === 'sala da coordenadora') {
    return 'coordenacao';
  }
  return SETORES_VALIDOS.indexOf(s) >= 0 ? s : 'guiche';
}

function _rowToObj(row, hm, rowNumber){
  const o={}; HEADERS.forEach(h => o[h] = (hm[h]!=null ? (row[hm[h]]||'') : ''));
  ['timestamp','chamadoEm','iniciadoEm','finalizadoEm'].forEach(k=>{
    const v=o[k];
    if (v && Object.prototype.toString.call(v)==='[object Date]' && !isNaN(v)){
      o[k]=Utilities.formatDate(v,TZ,"yyyy-MM-dd'T'HH:mm:ssXXX");
    } else { o[k]=String(v||''); }
  });
  o.data = _toDateKey(o.data) || _toDateKey(o.timestamp);
  o.calls = Number(o.calls||0);
  o._row = rowNumber;
  return o;
}

function _readDay(day){
  const targetDay = String(day||_today());
  const cached = _loadDayCache(targetDay);
  if (cached) return cached;

  const sheet = _sheet();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  const lastCol = sheet.getLastColumn();
  const hm = _headerMap(sheet);
  const dataIdx = hm['data'];

  if (dataIdx == null){
    const all = _readAll();
    const outFallback = all.filter(o => o.data === targetDay);
    _saveDayCache(targetDay, outFallback);
    return outFallback;
  }

  const n = lastRow - 1;
  const dataVals = sheet.getRange(2, dataIdx+1, n, 1).getValues();
  const matchRows = [];
  for (let i = 0; i < dataVals.length; i++){
    if (_toDateKey(dataVals[i][0]) === targetDay) matchRows.push(i + 2);
  }

  if (!matchRows.length){
    _saveDayCache(targetDay, []);
    return [];
  }

  const ranges = [];
  let start = matchRows[0];
  let prev = matchRows[0];
  for (let i = 1; i < matchRows.length; i++){
    const cur = matchRows[i];
    if (cur === prev + 1){
      prev = cur;
      continue;
    }
    ranges.push([start, prev]);
    start = cur;
    prev = cur;
  }
  ranges.push([start, prev]);

  const out = [];
  ranges.forEach(([rStart, rEnd])=>{
    const rows = sheet.getRange(rStart, 1, (rEnd-rStart+1), lastCol).getValues();
    for (let i = 0; i < rows.length; i++){
      const obj = _rowToObj(rows[i], hm, rStart + i);
      if (obj.data === targetDay) out.push(obj);
    }
  });

  _saveDayCache(targetDay, out);
  return out;
}

function _readAll(){
  const day = _today();
  const cached = _loadAllCache(day);
  if (cached) return cached;

  const sheet = _sheet();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  const lastCol = sheet.getLastColumn();
  const hm = _headerMap(sheet);
  const values = sheet.getRange(2,1,lastRow-1,lastCol).getValues();
  const out = values.map((row,i)=> _rowToObj(row, hm, i+2));
  _saveAllCache(day, out);
  return out;
}
function _writeRow(rowNumber, obj){
  const sheet=_sheet(); const hm=_headerMap(sheet);
  const lastCol = sheet.getLastColumn();
  const row = sheet.getRange(rowNumber,1,1,lastCol).getValues()[0];
  Object.keys(obj).forEach(k=>{ if(hm[k]!=null) row[hm[k]] = obj[k]; });
  sheet.getRange(rowNumber,1,1,lastCol).setValues([row]);
  _invalidateTodayCache();
}
function _append(obj){
  const sheet=_sheet(); const hm=_headerMap(sheet);
  const lastCol = sheet.getLastColumn();
  const row = new Array(lastCol).fill('');
  Object.keys(obj).forEach(k=>{ if(hm[k]!=null) row[hm[k]] = obj[k]; });
  sheet.appendRow(row);
  _invalidateTodayCache();
  return sheet.getLastRow();
}
function _minDiff(aIso,bIso){
  try{
    const a=new Date(aIso).getTime(), b=new Date(bIso).getTime();
    if(!isFinite(a)||!isFinite(b)) return 0;
    return Math.max(0, Math.round((b-a)/60000));
  }catch(_){ return 0; }
}

/** Aplica timeouts automáticos (40 min) em "atendendo" e "chamando" */
function _applyTimeouts(all){
  try{
    const agora = _nowIso();
    all.forEach(o=>{
      if (!o || !o.status) return;

      // Se está em atendimento há mais de MAX_ATEND_MIN → encerra como atendida
      if (o.status === 'atendendo'){
        const inicio = o.iniciadoEm || o.chamadoEm || o.timestamp;
        const dur = _minDiff(inicio, agora);
        if (dur >= MAX_ATEND_MIN){
          const tAt = _minDiff(inicio, agora);
          _writeRow(o._row, {
            status: 'atendida',
            finalizadoEm: agora,
            tempoAtendimento: tAt
          });
          o.status = 'atendida';
          o.finalizadoEm = agora;
          o.tempoAtendimento = tAt;
        }
      }
      // Se está chamando há mais de MAX_ATEND_MIN → marca como não compareceu
      else if (o.status === 'chamando'){
        const base = o.chamadoEm || o.timestamp;
        const dur = _minDiff(base, agora);
        if (dur >= MAX_ATEND_MIN){
          const esp = _minDiff(o.timestamp, o.chamadoEm || agora);
          _writeRow(o._row, {
            status: 'nao_compareceu',
            finalizadoEm: agora,
            tempoEspera: (o.tempoEspera || esp)
          });
          o.status = 'nao_compareceu';
          o.finalizadoEm = agora;
          o.tempoEspera = o.tempoEspera || esp;
        }
      }
    });
    return all;
  }catch(e){
    return all;
  }
}

/** Localiza atendimento/chamada ativa por guichê (para travar fluxo) */
function _findAtivoPorGuiche(all, guiche){
  const hoje = _today();
  const gStr = String(guiche||'');
  let emAtendimento = null;
  let emChamada = null;

  all.forEach(o=>{
    if (!o || o.data !== hoje) return;
    if (String(o.guiche||'') !== gStr) return;

    if (o.status === 'atendendo'){
      emAtendimento = o;
    } else if (o.status === 'chamando' && !emAtendimento){
      emChamada = o;
    }
  });

  return { emAtendimento: emAtendimento, emChamada: emChamada };
}

function _recalcVolatil(o){
  if(!o) return o;
  const agora = _nowIso();
  if (!o.tempoEspera) {
    if (o.status==='aguardando') o.tempoEspera = _minDiff(o.timestamp, agora);
    else if (o.status==='atendendo' || o.status==='chamando'){
      const ate = o.iniciadoEm || o.chamadoEm || agora;
      o.tempoEspera = _minDiff(o.timestamp, ate);
    }
  }
  if (!o.tempoAtendimento && o.status==='atendendo') {
    o.tempoAtendimento = _minDiff(o.iniciadoEm, agora);
  }
  return o;
}
function _codeDisplay_(row){
  const pref = PREFIX[row.tipo] || 'N';
  const m = String(row.numero||'').match(/[A-Z]?(\d+)/i);
  const digits = m && m[1] ? m[1] : '000';
  return pref + _pad(parseInt(digits,10)||0,3);
}
function _nextNumero(tipo){
  const pref = PREFIX[tipo] || 'N';
  const p = PropertiesService.getScriptProperties();
  const key = `SEQ_${pref}_${_today()}`;
  const next = (Number(p.getProperty(key))||0) + 1;
  p.setProperty(key, String(next));
  return pref + _pad(next,3);
}

/* ===== Ordem de Chamada ===== */
function _normalizeOrderStr(str){
  const s = String(str||'').toUpperCase().replace(/\s+/g,'');
  const parts = s.split(/[,\-;]+/).filter(Boolean);
  const out=[]; parts.forEach(x=>{ if(['N','P','G'].includes(x) && !out.includes(x)) out.push(x); });
  return out.length ? out.join(',') : 'P,N,G';
}
function _lettersToType(letter){
  return letter==='N' ? 'normal' : letter==='P' ? 'prioritario' : 'agendamento';
}
function _getOrderParts(){
  const props = PropertiesService.getScriptProperties();
  let ord = props.getProperty(CALL_ORDER_KEY) || 'P,N,G';
  ord = _normalizeOrderStr(ord);
  return ord.split(',');
}
function getCallOrder(){
  try{
    const props = PropertiesService.getScriptProperties();
    const ord  = _normalizeOrderStr(props.getProperty(CALL_ORDER_KEY) || 'P,N,G');
    const pos  = Number(props.getProperty(CALL_POS_KEY) || 0) || 0;
    const day  = props.getProperty(CALL_DAY_KEY) || _today();
    return { ok:true, order: ord, pos: pos, day: day };
  }catch(e){ return { ok:false, msg:String(e) }; }
}
function setCallOrder(orderStr){
  try{
    const props = PropertiesService.getScriptProperties();
    const norm = _normalizeOrderStr(orderStr);
    props.setProperty(CALL_ORDER_KEY, norm);
    props.setProperty(CALL_POS_KEY, '0');
    props.setProperty(CALL_DAY_KEY, _today());
    return { ok:true, order:norm, pos:0, day:_today() };
  }catch(e){ return { ok:false, msg:String(e) }; }
}

/* ===== Configurações do Admin (persistência servidor) ===== */
function _defaultAdminCfg(){
  return {
    titulo_sistema: 'Sistema de Chamados CadÚnico',
    mensagem_aguarde: 'Aguarde ser chamado no painel',
    audio: { volume:0.7, velocidade:1, repeticoes:4 },
    guiches: { numero:6, tempoChamada:15, max:20 },
    print:  {
      widthMm:72,
      marginTopMm:2,
      marginBottomMm:3,
      fontPx:12,
      gapPx:1,
      orgao:'Prefeitura Municipal',
      tempoChamada:15
    },
    services: [
      "Cadastro CadÚnico","Atualização de Dados CadÚnico","Emissão/Comprovante do CadÚnico",
      "Inclusão/Exclusão de Membros","Transferência de Domicílio/CRAS","Correção de Dados/Entrevista",
      "Programa Bolsa Família","Auxílio Gás dos Brasileiros",
      "BPC/LOAS — Pessoa Idosa","BPC/LOAS — Pessoa com Deficiência",
      "Tarifa Social de Energia Elétrica","Tarifa Social de Água/Esgoto (onde houver)",
      "Minha Casa, Minha Vida / Cadastro Habitacional",
      "ID Jovem","Carteira do Idoso (Gratuidade Intermunicipal/Interestadual)","Passe Livre Pessoa com Deficiência (Interestadual)",
      "Isenção de Taxa de Concurso Público","Isenção 2ª Via RG/Certidões (onde houver convênio)",
      "Acesso a Programas Sociais Municipais"
    ],
    salas: {
      coordenadora: 'Sala da Coordenadora',
      servico_social: 'Sala do Serviço Social'
    },
    ordem: 'P,N,G',
    /** NOVO: lista configurável de bairros (usada no combo do cadastro) */
    bairros: [
      "Centro",
      "Jardim",
      "Conjunto João Alves",
      "Conjunto Marcos Freire",
      "Outros"
    ],
    /** NOVO: PIN do Painel TV */
    tv_pin: '12345'
  };
}
function getAdminConfig(){
  try{
    const p = PropertiesService.getScriptProperties();
    const raw = p.getProperty(ADMIN_CFG_KEY);
    if(!raw){
      const def = _defaultAdminCfg();
      p.setProperty(ADMIN_CFG_KEY, JSON.stringify(def));
      p.setProperty(CALL_ORDER_KEY, def.ordem);
      return { ok:true, cfg:def };
    }
    const base = _defaultAdminCfg();
    const cfg = JSON.parse(raw);
    cfg.titulo_sistema   = String(cfg.titulo_sistema || base.titulo_sistema);
    cfg.mensagem_aguarde = String(cfg.mensagem_aguarde || base.mensagem_aguarde);
    cfg.audio   = Object.assign(base.audio,   cfg.audio||{});
    cfg.guiches = Object.assign(base.guiches, cfg.guiches||{});
    cfg.print   = Object.assign(base.print,   cfg.print||{});
    cfg.services= Array.isArray(cfg.services) ? cfg.services : base.services;
    cfg.salas   = Object.assign(base.salas,   cfg.salas||{});
    cfg.ordem   = _normalizeOrderStr(cfg.ordem || base.ordem);
    cfg.bairros = Array.isArray(cfg.bairros) ? cfg.bairros : base.bairros;
    cfg.tv_pin  = String(cfg.tv_pin || base.tv_pin);
    return { ok:true, cfg };
  }catch(e){ return { ok:false, msg:String(e) }; }
}
function saveAdminConfig(cfg){
  try{
    const base=_defaultAdminCfg();
    const out = {
      titulo_sistema: String(cfg?.titulo_sistema || base.titulo_sistema),
      mensagem_aguarde: String(cfg?.mensagem_aguarde || base.mensagem_aguarde),
      audio: {
        volume: Math.min(1, Math.max(0, Number(cfg?.audio?.volume)||base.audio.volume)),
        velocidade: Math.max(0.5, Math.min(2, Number(cfg?.audio?.velocidade)||base.audio.velocidade)),
        repeticoes: Math.max(1, Math.min(8, parseInt(cfg?.audio?.repeticoes,10)||base.audio.repeticoes)),
      },
      guiches:{
        numero: Math.max(1, parseInt(cfg?.guiches?.numero,10)||base.guiches.numero),
        tempoChamada: Math.max(8, Math.min(60, parseInt(cfg?.guiches?.tempoChamada,10)||base.guiches.tempoChamada)),
        max: Math.max(
          parseInt(cfg?.guiches?.max||base.guiches.max,10),
          parseInt(cfg?.guiches?.numero||base.guiches.numero,10)
        )
      },
      print:{
        widthMm: Number(cfg?.print?.widthMm)||base.print.widthMm,
        marginTopMm: Number(cfg?.print?.marginTopMm)||base.print.marginTopMm,
        marginBottomMm: Number(cfg?.print?.marginBottomMm)||base.print.marginBottomMm,
        fontPx: Number(cfg?.print?.fontPx)||base.print.fontPx,
        gapPx: Number(cfg?.print?.gapPx)||base.print.gapPx,
        orgao: String(cfg?.print?.orgao||base.print.orgao),
        tempoChamada: Math.max(8, Math.min(60, parseInt(cfg?.print?.tempoChamada,10)||base.print.tempoChamada)),
      },
      services: Array.isArray(cfg?.services)
        ? cfg.services.filter(s=>String(s).trim()).slice(0,200)
        : base.services,
      salas: {
        coordenadora: String(cfg?.salas?.coordenadora || base.salas.coordenadora),
        servico_social: String(cfg?.salas?.servico_social || base.salas.servico_social)
      },
      ordem: _normalizeOrderStr(cfg?.ordem || base.ordem),
      bairros: Array.isArray(cfg?.bairros)
        ? cfg.bairros.map(b=>String(b).trim()).filter(b=>b).slice(0,300)
        : base.bairros,
      tv_pin: String(cfg?.tv_pin || base.tv_pin)
    };
    const p = PropertiesService.getScriptProperties();
    p.setProperty(ADMIN_CFG_KEY, JSON.stringify(out));
    // mantém ordem também nas chaves de ordem
    p.setProperty(CALL_ORDER_KEY, out.ordem);
    p.setProperty(CALL_POS_KEY, '0');
    p.setProperty(CALL_DAY_KEY, _today());
    return { ok:true, cfg:out };
  }catch(e){
    return { ok:false, msg:String(e) };
  }
}

/* =========== APIs do fluxo =========== */
function getAllSenhas(){
  try{
    const dia=_today();
    let all=_readDay(dia);
    all = _applyTimeouts(all);
    const hoje = all.filter(o=> (o.data && o.data===dia));
    return hoje.map(_recalcVolatil).sort((a,b)=> new Date(a.timestamp)-new Date(b.timestamp));
  }catch(e){ return []; }
}
function createSenha(input){
  const lock = LockService.getScriptLock();
  try{
    lock.waitLock(5000);
    const nome = String(input?.nome||'').trim();
    const cpf  = _digits(input?.cpf||'');
    const tipo = String(input?.tipo||'normal').trim().toLowerCase();
    const servico = String(input?.servico||'').trim();
    const bairro  = String(input?.bairro||'').trim();
    const setor   = _normalizeSetor(input?.setor || input?.setorAtendimento || 'guiche');

    if(!nome)    return { ok:false, msg:'Nome obrigatório.' };
    if(!bairro)  return { ok:false, msg:'Bairro obrigatório.' };
    if(!servico) return { ok:false, msg:'Serviço obrigatório.' };
    if(!['normal','prioritario','agendamento'].includes(tipo))
      return { ok:false, msg:'Tipo inválido.' };
    if(!SETORES_VALIDOS.includes(setor))
      return { ok:false, msg:'Setor de atendimento inválido.' };

    const hoje=_today(), agora=_nowIso();
    const all=_readDay(hoje);
    _applyTimeouts(all);
    const recent = all
      .filter(o => o.data===hoje && o.nome===nome && _digits(o.cpf)===cpf && o.servico===servico)
      .sort((a,b)=> new Date(b.timestamp)-new Date(a.timestamp));
    if (recent.length){
      const last = recent[0];
      if (_minDiff(last.timestamp, agora) <= 1 && last.status==='aguardando'){
        return { ok:true, data:last, reused:true };
      }
    }
    const obj = {
      id: Utilities.getUuid(),
      numero: _nextNumero(tipo),
      nome, cpf, tipo, servico, bairro,
      setor,
      status: 'aguardando',
      timestamp: agora,
      chamadoEm:'', iniciadoEm:'', finalizadoEm:'',
      guiche:'', tempoEspera:'', tempoAtendimento:'',
      data: hoje,
      calls: 0
    };
    const row=_append(obj); obj._row=row;
    return { ok:true, data:obj, reused:false };
  }catch(e){ return _sheetError(e); }
  finally { try{ lock.releaseLock(); }catch(_){ } }
}
function updateSenha(obj){
  try{
    const id = String(obj?.id||''); if(!id) return { ok:false, msg:'ID ausente' };
    let all=_readDay(_today()); all=_applyTimeouts(all);
    const f=all.find(x=>x.id===id);
    if(!f) return { ok:false, msg:'Registro não encontrado' };
    const novoTipo = obj.tipo ?? f.tipo;
    const novoSetor = _normalizeSetor(
      obj.setor ?? obj.setorAtendimento ?? (f.setor || 'guiche')
    );
    const up = {
      nome: obj.nome ?? f.nome,
      cpf:  _digits(obj.cpf ?? f.cpf),
      servico: obj.servico ?? f.servico,
      bairro: obj.bairro ?? f.bairro,
      tipo: novoTipo,
      setor: novoSetor
    };
    if (novoTipo !== f.tipo){
      const m = String(f.numero||'').match(/[A-Z]?(\d+)/i);
      const dig = m && m[1] ? m[1] : '000';
      up.numero = (PREFIX[novoTipo]||'N') + _pad(parseInt(dig,10)||0,3);
    }
    _writeRow(f._row, up);
    return { ok:true, data:Object.assign(f, up) };
  }catch(e){ return _sheetError(e); }
}
function deleteSenha(id){
  try{
    const s=_sheet(); const all=_readDay(_today());
    const f=all.find(x=>x.id===String(id));
    if(!f) return { ok:false, msg:'Registro não encontrado' };
    s.deleteRow(f._row);
    _invalidateTodayCache();
    return { ok:true };
  }catch(e){ return _sheetError(e); }
}

/** Chamada obedecendo ordem configurada, com trava por guichê e separação por setor */
function callNext(guiche, setorRaw){
  const lock=LockService.getScriptLock();
  try{
    lock.waitLock(5000);
    const hoje=_today();
    const setor = _normalizeSetor(setorRaw || 'guiche');

    let all=_readDay(hoje);
    all = _applyTimeouts(all);

    // trava: não permitir nova chamada se este guichê já tem algo em andamento
    const ativos = _findAtivoPorGuiche(all, guiche);
    if (ativos.emAtendimento){
      return {
        ok:false,
        code:'GUICHE_EM_ATENDIMENTO',
        msg:'Já existe um atendimento em andamento neste guichê. Finalize o atendimento antes de chamar outra senha.',
        data:ativos.emAtendimento
      };
    }
    if (ativos.emChamada){
      return {
        ok:false,
        code:'GUICHE_COM_CHAMADA_PENDENTE',
        msg:'Há uma senha em chamada neste guichê. Marque "Não compareceu" ou inicie/finalize o atendimento antes de chamar outra.',
        data:ativos.emChamada
      };
    }

    // apenas senhas do mesmo setor/local (guichês, serviço social ou coordenação)
    const allHoje = all.filter(o=> o.data===hoje && _normalizeSetor(o.setor || 'guiche') === setor);

    const filas = {
      normal:       allHoje.filter(r=>r.status==='aguardando' && r.tipo==='normal'     ).sort((a,b)=> new Date(a.timestamp)-new Date(b.timestamp)),
      prioritario:  allHoje.filter(r=>r.status==='aguardando' && r.tipo==='prioritario' ).sort((a,b)=> new Date(a.timestamp)-new Date(b.timestamp)),
      agendamento:  allHoje.filter(r=>r.status==='aguardando' && r.tipo==='agendamento').sort((a,b)=> new Date(a.timestamp)-new Date(b.timestamp))
    };
    if(!filas.normal.length && !filas.prioritario.length && !filas.agendamento.length){
      return { ok:false, msg:'Sem senhas aguardando para este setor.' };
    }
    const props = PropertiesService.getScriptProperties();
    const parts = _getOrderParts();
    let pos = Number(props.getProperty(CALL_POS_KEY) || 0) || 0;
    const lastDay = props.getProperty(CALL_DAY_KEY) || '';
    if (lastDay !== hoje) { pos = 0; props.setProperty(CALL_DAY_KEY, hoje); }

    let chosen=null, chosenIdx=pos;
    for (let i=0; i<parts.length; i++){
      const idx = (pos + i) % parts.length;
      const tipoLetra = parts[idx];
      const tipo = _lettersToType(tipoLetra);
      const fila = filas[tipo];
      if (fila && fila.length){ chosen = fila[0]; chosenIdx = idx; break; }
    }
    if(!chosen) return { ok:false, msg:'Sem candidatos' };

    const agora=_nowIso();
    const espera=_minDiff(chosen.timestamp, agora);
    const novoNumero = _codeDisplay_(chosen);
    const novasChamadas = Number(chosen.calls||0) + 1;

    _writeRow(chosen._row, {
      status:'chamando', chamadoEm:agora, guiche:String(guiche||''),
      tempoEspera: espera, numero: novoNumero, calls: novasChamadas
    });

    chosen.status='chamando';
    chosen.chamadoEm=agora; chosen.guiche=String(guiche||'');
    chosen.tempoEspera=espera; chosen.numero=novoNumero; chosen.calls=novasChamadas;

    const nextPos = (chosenIdx + 1) % parts.length;
    props.setProperty(CALL_POS_KEY, String(nextPos));
    props.setProperty(CALL_DAY_KEY, hoje);

    // ✅ Inclui ordem e próxima posição para o front sincronizar a fila previsível
    return { ok:true, data:chosen, order: parts.join(','), nextPos: nextPos };
  }catch(e){ return _sheetError(e); }
  finally { try{ lock.releaseLock(); }catch(_){ } }
}

/** Nova função: chamar manualmente uma senha específica (por ID), respeitando setor/local */
function callSenhaManual(idOrPayload, guiche, setorRaw){
  const lock = LockService.getScriptLock();
  try{
    lock.waitLock(5000);

    let id = '';
    let codigo = '';

    if (idOrPayload && typeof idOrPayload === 'object'){
      id = String(idOrPayload.id || '').trim();
      codigo = String(idOrPayload.codigo || idOrPayload.numero || '').trim().toUpperCase();
      guiche = (idOrPayload.guiche != null) ? idOrPayload.guiche : guiche;
      setorRaw = (idOrPayload.setor != null)
        ? idOrPayload.setor
        : ((idOrPayload.setorAtendimento != null) ? idOrPayload.setorAtendimento : setorRaw);
    } else {
      id = String(idOrPayload || '').trim();
    }

    if (!id && !codigo) return { ok:false, msg:'Informe ID ou código da senha.' };

    const setor = _normalizeSetor(setorRaw || 'guiche');

    let all = _readDay(_today());
    all = _applyTimeouts(all);
    const hoje = _today();

    // trava por guichê
    const ativos = _findAtivoPorGuiche(all, guiche);
    if (ativos.emAtendimento){
      return {
        ok:false,
        code:'GUICHE_EM_ATENDIMENTO',
        msg:'Já existe um atendimento em andamento neste guichê. Finalize o atendimento antes de chamar outra senha.',
        data:ativos.emAtendimento
      };
    }
    if (ativos.emChamada){
      return {
        ok:false,
        code:'GUICHE_COM_CHAMADA_PENDENTE',
        msg:'Há uma senha em chamada neste guichê. Marque "Não compareceu" ou inicie/finalize o atendimento antes de chamar outra.',
        data:ativos.emChamada
      };
    }

    let alvo = null;
    if (id){
      alvo = all.find(o=>o.id===id && o.data===hoje) || null;
    }
    if (!alvo && codigo){
      alvo = all.find(o=>{
        if (o.data !== hoje) return false;
        const n = String(o.numero || '').toUpperCase().trim();
        const d = String(_codeDisplay_(o) || '').toUpperCase().trim();
        return n === codigo || d === codigo;
      }) || null;
    }

    if (!alvo) return { ok:false, msg:'Senha não encontrada para hoje.' };

    const setorSenha = _normalizeSetor(alvo.setor || 'guiche');
    if (setorSenha !== setor){
      return {
        ok:false,
        code:'SETOR_DIFERENTE',
        msg:'Esta senha pertence a outro local de atendimento.',
        data: alvo
      };
    }

    if (alvo.status !== 'aguardando'){
      return {
        ok:false,
        msg:'Só é possível chamar manualmente senhas que estão aguardando.',
        data: alvo
      };
    }

    const agora = _nowIso();
    const espera = _minDiff(alvo.timestamp, agora);
    const novoNumero = _codeDisplay_(alvo);
    const novasChamadas = Number(alvo.calls||0) + 1;

    _writeRow(alvo._row, {
      status:'chamando',
      chamadoEm:agora,
      guiche:String(guiche||''),
      tempoEspera:espera,
      numero:novoNumero,
      calls:novasChamadas
    });

    alvo.status='chamando';
    alvo.chamadoEm=agora;
    alvo.guiche=String(guiche||'');
    alvo.tempoEspera=espera;
    alvo.numero=novoNumero;
    alvo.calls=novasChamadas;

    return { ok:true, data:alvo };
  }catch(e){
    return { ok:false, msg:String(e) };
  }finally{
    try{ lock.releaseLock(); }catch(_){}
  }
}

function repetirChamada(id){
  try{
    id=String(id||'').trim();
    let all=_readDay(_today()); all=_applyTimeouts(all);
    const o=all.find(x=>x.id===id);
    if(!o) return { ok:false, msg:'Registro não encontrado.' };
    const agora=_nowIso();
    const chamadas = Number(o.calls||0) + 1;
    _writeRow(o._row, { chamadoEm:agora, calls:chamadas, status:(o.status||'chamando') });
    o.chamadoEm = agora; o.calls = chamadas;
    if(o.status==='aguardando') o.status='chamando';
    return { ok:true, data:o };
  }catch(e){ return { ok:false, msg:String(e) }; }
}
function iniciarAtendimento(id){
  try{
    id=String(id||'').trim();
    let all=_readDay(_today()); all=_applyTimeouts(all);
    const o=all.find(x=>x.id===id);
    if(!o) return { ok:false, msg:'Registro não encontrado.' };

    // se já foi encerrado por timeout ou manualmente
    if (o.status === 'atendida' || o.status === 'nao_compareceu'){
      return { ok:false, msg:'Atendimento já encerrado para esta senha.', data:o };
    }

    const agora=_nowIso(), espera=_minDiff(o.timestamp, agora);
    _writeRow(o._row, { status:'atendendo', iniciadoEm:agora, tempoEspera:espera });
    o.status='atendendo'; o.iniciadoEm=agora; o.tempoEspera=espera;
    return { ok:true, data:o };
  }catch(e){ return { ok:false, msg:String(e) }; }
}
function finalizarAtendimento(id){
  try{
    id=String(id||'').trim();
    let all=_readDay(_today()); all=_applyTimeouts(all);
    const o=all.find(x=>x.id===id);
    if(!o) return { ok:false, msg:'Registro não encontrado.' };
    const agora=_nowIso(), inicio=o.iniciadoEm || o.chamadoEm || o.timestamp || agora;
    const t=_minDiff(inicio, agora);
    _writeRow(o._row, { status:'atendida', finalizadoEm:agora, tempoAtendimento:t });
    o.status='atendida'; o.finalizadoEm=agora; o.tempoAtendimento=t;
    return { ok:true, data:o };
  }catch(e){ return { ok:false, msg:String(e) }; }
}
function marcarNoShow(id){
  try{
    id=String(id||'').trim();
    let all=_readDay(_today()); all=_applyTimeouts(all);
    const o=all.find(x=>x.id===id);
    if(!o) return { ok:false, msg:'Registro não encontrado' };
    const agora=_nowIso(), esp=_minDiff(o.timestamp, o.chamadoEm || agora);
    _writeRow(o._row, { status:'nao_compareceu', finalizadoEm:agora, tempoEspera:(o.tempoEspera || esp) });
    o.status='nao_compareceu'; o.finalizadoEm=agora; o.tempoEspera=o.tempoEspera || esp;
    return { ok:true, data:o };
  }catch(e){ return { ok:false, msg:String(e) }; }
}

/* ===== Consultas/Estatísticas ===== */
function getSenhasPeriodo(de, ate){
  try{
    const inicio=(de && String(de).trim()) || '0000-00-00';
    const fim   =(ate && String(ate).trim()) || '9999-12-31';
    let all=_readAll(); all=_applyTimeouts(all);
    return all.filter(o=>{
      const d = o.data || _toDateKey(o.timestamp);
      return d>=inicio && d<=fim;
    }).map(_recalcVolatil);
  }catch(e){ return []; }
}
function searchCPF(cpf, de, ate){
  try{
    const dig=_digits(cpf); if(!dig) return [];
    const inicio=(de && String(de).trim()) || '0000-00-00';
    const fim   =(ate && String(ate).trim()) || '9999-12-31';
    let all=_readAll(); all=_applyTimeouts(all);
    return all.filter(o=>{
      const d = o.data || _toDateKey(o.timestamp);
      return d>=inicio && d<=fim && _digits(o.cpf)===dig;
    }).map(_recalcVolatil);
  }catch(e){ return []; }
}
function searchNome(nome, de, ate){
  try{
    const q=String(nome||'').trim().toLowerCase(); if(!q) return [];
    const inicio=(de && String(de).trim()) || '0000-00-00';
    const fim   =(ate && String(ate).trim()) || '9999-12-31';
    let all=_readAll(); all=_applyTimeouts(all);
    return all.filter(o=>{
      const d = o.data || _toDateKey(o.timestamp);
      return d>=inicio && d<=fim && String(o.nome||'').toLowerCase().indexOf(q) >= 0;
    }).map(_recalcVolatil);
  }catch(e){ return []; }
}

/** Estatística por bairros (quantidade de atendimentos por bairro) */
function getStatsBairros(de, ate){
  try{
    const inicio=(de && String(de).trim()) || '0000-00-00';
    const fim   =(ate && String(ate).trim()) || '9999-12-31';
    let all=_readAll(); all=_applyTimeouts(all);
    const map = {};
    all.forEach(o=>{
      const d = o.data || _toDateKey(o.timestamp);
      if (!d || d<inicio || d>fim) return;
      const b = String(o.bairro||'').trim() || 'SEM_INFORMACAO';
      if (!map[b]){
        map[b] = {
          bairro: b,
          total: 0,
          aguardando:0,
          chamando:0,
          atendendo:0,
          atendida:0,
          nao_compareceu:0
        };
      }
      const m = map[b];
      m.total++;
      if (o.status && m[o.status]!=null) m[o.status]++;
    });
    return { ok:true, bairros:Object.values(map) };
  }catch(e){
    return { ok:false, msg:String(e) };
  }
}

/* ===== Painel TV: últimos chamados sem duplicar ===== */
function getUltimosChamados(limit, dedupBy){
  try{
    const L = Math.max(1, parseInt(limit,10) || 8);
    const by = (dedupBy || 'id');
    const hoje = _today();

    let all = _readDay(hoje); all=_applyTimeouts(all);
    all = all
      .filter(o => o.data === hoje && o.chamadoEm)
      .sort((a,b) => new Date(b.chamadoEm) - new Date(a.chamadoEm));

    const principal = all.find(o => o.status === 'chamando') || all[0] || null;

    const vistos = new Set();
    const recentes = [];
    for (const o of all){
      const key = (by === 'numero') ? String(o.numero) : String(o.id);
      if (vistos.has(key)) continue;
      vistos.add(key);
      recentes.push(o);
      if (recentes.length >= L) break;
    }

    return { ok:true, principal, recentes };
  }catch(e){
    return { ok:false, msg:String(e) };
  }
}

/** NOVO: estado atual do guichê (para recuperar após recarregar a página) */
function getEstadoGuiche(guiche){
  try{
    let all = _readDay(_today()); all = _applyTimeouts(all);
    const hoje = _today();
    const gStr = String(guiche||'');
    let emAtendimento = null;
    let emChamada = null;

    all.forEach(o=>{
      if (o.data !== hoje) return;
      if (String(o.guiche||'') !== gStr) return;
      if (o.status === 'atendendo') emAtendimento = o;
      else if (o.status === 'chamando' && !emAtendimento) emChamada = o;
    });

    return {
      ok:true,
      guiche: gStr,
      emAtendimento: emAtendimento,
      emChamada: emChamada
    };
  }catch(e){
    return { ok:false, msg:String(e) };
  }
}

/* ===== Admin: login simples ===== */
function validateAdmin(u,p){
  try{
    const U = String(u||'').trim();
    const P = String(p||'').trim();
    // aceita admin/admin12345 (principal) e admin/12345 (contingência)
    if (U==='admin' && (P==='admin12345' || P==='12345')){
      return { ok:true, token:'ok', usuario:'admin' };
    }
    return { ok:false, msg:'Usuário ou senha inválidos.' };
  }catch(e){ return { ok:false, msg:String(e) }; }
}

/** NOVO: validação do PIN do Painel TV (trava de segurança) */
function validateTvPin(pin){
  try{
    const r = getAdminConfig();
    if (!r || !r.ok) return { ok:false, msg: (r && r.msg) || 'Falha ao carregar configuração.' };
    const cfg = r.cfg || {};
    const cfgPin = String(cfg.tv_pin || '12345').trim();
    const inPin  = String(pin || '').trim();
    if (!inPin) return { ok:false, msg:'Informe o PIN.' };
    if (inPin === cfgPin) return { ok:true };
    return { ok:false, msg:'PIN inválido.' };
  }catch(e){
    return { ok:false, msg:String(e) };
  }
}

/* ===== Ações administrativas ===== */
function clearToday(){
  const lock = LockService.getScriptLock();
  try{
    lock.waitLock(5000);
    const s=_sheet(); const hm=_headerMap(s);
    const lastRow=s.getLastRow(); if(lastRow<2) return { ok:true, removidos:0 };
    const dataHoje=_today();
    let removidos=0;
    for(let row=lastRow; row>=2; row--){
      const vData = s.getRange(row, (hm['data']||0)+1).getValue();
      const vTs   = s.getRange(row, (hm['timestamp']||0)+1).getValue();
      const key = _toDateKey(vData) || _toDateKey(vTs);
      if(key === dataHoje){ s.deleteRow(row); removidos++; }
    }
    _invalidateTodayCache();
    return { ok:true, removidos };
  }catch(e){ return { ok:false, msg:String(e) }; }
  finally{ try{ lock.releaseLock(); }catch(_){ } }
}
function resetCounters(){
  try{
    const p = PropertiesService.getScriptProperties();
    const d = _today();
    ['N','P','G'].forEach(L=>{ p.deleteProperty(`SEQ_${L}_${d}`); });
    p.setProperty(CALL_POS_KEY, '0');
    p.setProperty(CALL_DAY_KEY, d);
    return { ok:true };
  }catch(e){ return { ok:false, msg:String(e) }; }
}

/* ===== Tratamento de erros ===== */
function _sheetError(e){
  const msg = String(e && e.message || e || 'Erro');
  if (msg === 'PLANILHA_INACESSIVEL'){
    return { ok:false, msg:'Planilha inacessível: confirme o ID e implante o App como “Executar como: VOCÊ”.' };
  }
  return { ok:false, msg: msg };
}

/* ===== Compatibilidade com o front (aliases) ===== */
function getOrderSequence(){ return getCallOrder(); }
function setOrderSequence(seq){ return setCallOrder(seq); }
function repeatCall(id, guiche){ return repetirChamada(id); }

/* ======== SHIMS DO ADMIN PARA O FRONT ======== */
function loadConfig(){
  try{
    var r = getAdminConfig();
    if (r && r.ok && r.cfg){
      var cfg = r.cfg || {};
      var out = {
        audio: cfg.audio || {},
        print: cfg.print || {},
        salas: cfg.salas || {},
        ordem: cfg.ordem || 'P,N,G',
        services: cfg.services || [],
        guiches: (cfg.guiches && cfg.guiches.numero) ? cfg.guiches.numero : (Number(cfg.guiches)||6),
        titulo_sistema: cfg.titulo_sistema || 'Sistema de Chamados CadÚnico',
        nome_orgao: (cfg.print && cfg.print.orgao) ? cfg.print.orgao : (cfg.nome_orgao || 'Prefeitura Municipal'),
        mensagem_aguarde: cfg.mensagem_aguarde || 'Aguarde ser chamado no painel',
        bairros: cfg.bairros || [],
        tv_pin: cfg.tv_pin || ''
      };
      return { ok:true, config: out };
    }
    return { ok:false, msg: (r && r.msg) || 'Falha ao carregar configuração' };
  }catch(e){
    return { ok:false, msg: String(e) };
  }
}

function saveConfig(config){
  try{
    var cfg = config || {};
    var guichesObj = (cfg.guiches && typeof cfg.guiches === 'object')
      ? cfg.guiches
      : {
          numero: parseInt(cfg.guiches,10) || 6,
          tempoChamada: (cfg.print && parseInt(cfg.print.tempoChamada,10)) || 15
        };

    var payload = {
      titulo_sistema: cfg.titulo_sistema || 'Sistema de Chamados CadÚnico',
      mensagem_aguarde: cfg.mensagem_aguarde || 'Aguarde ser chamado no painel',
      audio: cfg.audio || {},
      print: cfg.print || {},
      services: Array.isArray(cfg.services) ? cfg.services : [],
      ordem: cfg.ordem || 'P,N,G',
      guiches: guichesObj,
      salas: cfg.salas || {},
      bairros: Array.isArray(cfg.bairros) ? cfg.bairros : [],
      tv_pin: cfg.tv_pin || ''
    };
    return saveAdminConfig(payload);
  }catch(e){
    return { ok:false, msg: String(e) };
  }
}
