/* ============================================================
   core.js — lógica sin interfaz: leer el Excel, modelo de datos,
   vista previa HTML y generación del PDF del anexo fotográfico.
   ============================================================ */

const STATUS = {
  SI:  {word:'CORRECTO',    color:[0,140,60],   css:'#008c3c'},
  NO:  {word:'NO CONFORME', color:[220,0,0],    css:'#dc0000'},
  NA:  {word:'NA',          color:[0,0,0],      css:'#000000'},
  PTE: {word:'PTE',         color:[0,0,0],      css:'#000000'},
  '':  {word:'SIN VALORAR', color:[120,120,120],css:'#787878'},
};
const ORANGE = [247,148,29];
const SIZE_FRAC = {S:0.48, M:0.72, L:1};
const DEF_SIZE = 'M';

function normStatus(v){
  if(v===null || v===undefined) return '';
  const s = String(v).trim().toUpperCase().replace(/Í/g,'I');
  if(s==='SI' || s==='CORRECTO') return 'SI';
  if(s==='NO' || s==='NO CONFORME' || s==='INCORRECTO') return 'NO';
  if(s==='N/A' || s==='NA') return 'NA';
  if(s==='PTE' || s==='PENDIENTE') return 'PTE';
  return '';
}

function cellText(ws, addr){
  const c = ws[addr];
  if(!c) return '';
  const t = (c.w !== undefined) ? c.w : c.v;
  return (t===null || t===undefined) ? '' : String(t).trim();
}

/* Lee el Excel de la plataforma (ya rellenado) y devuelve cabecera + preguntas */
function parseWorkbook(buf){
  const wb = XLSX.read(buf, {type:'array'});
  const ws = wb.Sheets[wb.SheetNames[0]];
  if(!ws || !ws['!ref']) throw new Error('El archivo no tiene ninguna hoja con datos.');

  const nameRaw = cellText(ws,'A2');
  const name = nameRaw.replace(/^Centro de Coste:\s*/i,'').trim();
  const cPNraw = cellText(ws,'B5');
  const ceco = /^\d+$/.test(cPNraw) ? String(parseInt(cPNraw,10)) : cPNraw;
  const meta = {
    name, ceco, cPN: cPNraw,
    zona: cellText(ws,'B3'),
    fecha: cellText(ws,'G3'),
    codSupervisor: cellText(ws,'G5'),
  };

  const range = XLSX.utils.decode_range(ws['!ref']);
  const questions = [];
  let area = {letra:'', nombre:''};
  let apartado = '';
  let num = 0;
  for(let r=1; r<=range.e.r+1; r++){
    const tipo = cellText(ws,'K'+r);
    if(tipo==='AREA'){
      area = {letra: cellText(ws,'C'+r), nombre: cellText(ws,'D'+r)};
    }else if(tipo==='APARTADO'){
      apartado = cellText(ws,'E'+r);
    }else if(tipo==='PREGUNTA'){
      num++;
      questions.push({
        num, row:r,
        area: area.letra, areaName: area.nombre, apartado,
        text: cellText(ws,'E'+r).replace(/^\d+\s+/,'').trim(),
        status: normStatus(cellText(ws,'G'+r)),
        valor: cellText(ws,'F'+r),
        obs: cellText(ws,'H'+r),
      });
      const qq = questions[questions.length-1];
      qq.orig = {status:qq.status, valor:qq.valor, obs:qq.obs};   // lo que traía el Excel
    }
  }
  if(!questions.length){
    throw new Error('No encuentro las preguntas en este Excel. Debe ser el archivo por estación descargado de la plataforma de Repsol (hoja "Hoja1").');
  }
  if(!meta.ceco && !meta.name){
    throw new Error('El Excel no trae el centro de coste (celdas A2 y B5 vacías).');
  }
  return {meta, questions};
}


/* ============================================================
   VALORES POR DEFECTO AL CARGAR EL EXCEL LIMPIO
   - Toda pregunta sin valorar pasa a SÍ (verde).
   - 13, 37 y 53 pasan a N/A con su observación.
   - Interruptor «no dispone de servicios»: 6, 8, 56, 65 y 71 → N/A + NO DISPONE.
   Cada regla se comprueba también por el texto de la pregunta: si Repsol cambiara la
   numeración, la regla no se aplica y la app avisa en vez de marcar la pregunta equivocada.
   ============================================================ */
const DEFAULT_NA = {
  13: {kw:'RECAUDACION DE TABACO', obs:'RECAUDACIÓN CON TABLET'},
  37: {kw:'SOTANO',                obs:'NO DISPONE'},
  53: {kw:'LEGIONELLA',            obs:'NO APLICA'},
};
const SERVICE_Q = {6:'CONTADORES', 8:'COMPRESOR', 48:'COMPRESOR', 56:'SERVICIOS COMPLEMENTARIOS', 65:'SERVICIOS COMPLEMENTARIOS', 71:'SERVICIOS COMPLEMENTARIOS'};
const SERVICE_OBS = 'NO DISPONE';
function normTxt(s){ return String(s||'').toUpperCase().normalize('NFD').replace(/[̀-ͯ]/g,''); }
function ruleMatches(q, kw){ return normTxt(q.text).includes(kw); }
/* Devuelve la lista de preguntas cuya regla no coincide con su texto */

/* ---- Inventarios 14-17: AJUSTE FISICO / VENTAS en € y % calculado ---- */
const INV_Q = {14:'TABACO', 15:'ONCE', 16:'BUTANO', 17:'BEBIDAS CALIENTES'};
function parseNum(s){
  s=String(s==null?'':s).replace(/[€\s]/g,'');
  if(!s||s==='-'||s==='+') return NaN;
  if(s.includes(',') && s.includes('.')) s=s.replace(/\./g,'').replace(',','.');
  else if(s.includes(',')) s=s.replace(',','.');
  else if(/^[-+]?\d{1,3}(\.\d{3})+$/.test(s)) s=s.replace(/\./g,'');
  const n=Number(s); return isFinite(n)?n:NaN;
}
function fmtPct(p){
  for(const d of [2,1,0]){ const t=p.toFixed(d).replace('.',',')+'%'; if(t.length<=10) return t; }
  return '%';
}
function invBlank(q){
  const v=String(q.valor||'').trim(), o=String(q.obs||'').trim();
  return (!v || v==='%') && (!o || /^AJUSTE\s*F[IÍ]SICO/i.test(o));
}
/* Tolerancias de inventario (en %, positivo o negativo): fuera de rango → NO, dentro → SÍ */
const INV_TOL = {14:0.50, 15:0.85, 16:0, 17:0.85};
function recalcInv(q){
  const aj=parseNum(q.ajuste), ve=parseNum(q.ventas);
  const ok = isFinite(aj)&&isFinite(ve)&&ve!==0;
  q.valor = ok ? fmtPct(aj*100/ve) : '%';
  if(ok && INV_TOL[q.num]!==undefined && (q.status==='SI'||q.status==='NO'||!q.status)){
    const p = Math.round(aj*100/ve*100)/100;                       // el % tal como se ve (2 decimales)
    q.status = Math.abs(p) <= INV_TOL[q.num]+1e-9 ? 'SI' : 'NO';
  }
  const aj2=String(q.ajuste||'').trim(), ve2=String(q.ventas||'').trim();
  q.obs = 'AJUSTE FISICO: '+(aj2?aj2+' ':'')+'€ VENTAS: '+(ve2?ve2+' ':'')+'€';
}
/* Encuestas creadas con una versión anterior: convierte 14-17 sin tocar nada más */
function migrateInv(questions){
  let ch=false;
  questions.forEach(q=>{
    if(INV_Q[q.num] && !q.inv && invBlank(q) && ruleMatches(q, INV_Q[q.num])){ q.inv=true; q.ajuste=''; q.ventas=''; recalcInv(q); ch=true; }
  });
  return ch;
}
function applyDefaults(questions){
  const warn = [];
  questions.forEach(q=>{
    if(INV_Q[q.num] && !q.inv && invBlank(q) && ruleMatches(q, INV_Q[q.num])){ q.inv=true; q.ajuste=''; q.ventas=''; recalcInv(q); }
    if(q.status) return;                              // lo que ya traiga rellenado el Excel se respeta
    const r = DEFAULT_NA[q.num];
    if(r){
      if(ruleMatches(q, r.kw)){ q.status='NA'; if(!q.obs) q.obs=r.obs; return; }
      warn.push(q.num);
    }
    q.status = 'SI';
  });
  return warn;
}
function setSinServicios(audit, on){
  const warn = [];
  Object.keys(SERVICE_Q).forEach(n=>{
    const q = audit.questions.find(x=>String(x.num)===n);
    if(!q || !ruleMatches(q, SERVICE_Q[n])){ warn.push(+n); return; }
    if(on){ q.status='NA'; q.obs=SERVICE_OBS; }
    else if(q.status==='NA' && q.obs===SERVICE_OBS){ q.status='SI'; q.obs=''; }   // solo si no la has cambiado a mano
  });
  audit.sinServicios = !!on;
  return warn;
}


/* ---- Lectura de la foto de inventario: aplica lo leído a 14-17 y 19 ---- */
const INV_KEYS = {tabaco:14, once:15, butano:16, cafe:17};
const FONDOS_KW = 'FONDOS FIJO';
function cleanAmount(s){
  s = String(s==null?'':s).replace(/[−–—]/g,'-').replace(/[€\s]/g,'').replace(/[^0-9.,-]/g,'');
  return /\d/.test(s) ? s : '';
}
function applyInvReading(audit, r){
  const warn = [];
  (r.nodispone||[]).forEach(k=>{
    const n = INV_KEYS[k], q = audit.questions.find(x=>x.num===n);
    if(!q || !ruleMatches(q, INV_Q[n])){ warn.push(n); return; }
    q.inv=false; delete q.ajuste; delete q.ventas; q.status='NA'; q.valor=''; q.obs='NO DISPONE';
  });
  Object.keys(INV_KEYS).forEach(k=>{
    if((r.nodispone||[]).includes(k)) return;
    const n = INV_KEYS[k], e = r[k]; if(!e) return;
    const q = audit.questions.find(x=>x.num===n);
    if(!q || !ruleMatches(q, INV_Q[n])){ warn.push(n); return; }
    q.inv = true;
    q.ajuste = cleanAmount(e.a); q.ventas = cleanAmount(e.v);
    recalcInv(q);
  });
  const ff = cleanAmount(r.ffijo), fm = cleanAmount(r.fmaniobra);
  if(ff || fm){
    const q = audit.questions.find(x=>x.num===19);
    if(!q || !ruleMatches(q, FONDOS_KW)) warn.push(19);
    else q.obs = [ff?('F.FIJO: '+ff+'€.'):'', fm?('F.MANIOBRA: '+fm+'€'):''].filter(Boolean).join(' ').slice(0,100);
  }
  return warn;
}


/* ---- Interpretar el texto escaneado (Texto en directo de iOS) ---- */
function parseInventoryText(raw){
  const t = String(raw||'').toUpperCase().normalize('NFD').replace(/[̀-ͯ]/g,'').replace(/[−–—]/g,'-');
  const NUM = '(-?\\s?\\d[\\d.,]*)';
  const out = {tabaco:{a:'',v:''},cafe:{a:'',v:''},once:{a:'',v:''},butano:{a:'',v:''},ffijo:'',fmaniobra:'',dudas:''};
  const dud = [];
  const after = (kw)=>{                       // primera cifra tras la palabra, antes de una V o de otra palabra clave
    const m = t.match(new RegExp(kw+'[^0-9\\-\\n]{0,8}?(?:\\bA\\b)?[^0-9\\-\\n]{0,4}'+NUM));
    return m ? m[1].replace(/\s/g,'').replace(/[.,]+$/,'') : '';
  };
  out.tabaco.a = after('TABAC\\w*'); out.cafe.a = after('CAFE\\w*'); out.once.a = after('ONCE'); out.butano.a = after('BUTAN\\w*');
  const ff = t.match(/(?:F\.?\s*FIJO|FFIJO|FFIJ\w*|FONDO\s+FIJO)[^0-9\n]{0,6}(\d[\d.,]*)/);
  const fm = t.match(/(?:F\.?\s*MANIOBRA|F\.?\s*MANIO\w*|FONDO\s+(?:DE\s+)?MANIOBRA|MANIOBRA)[^0-9\n]{0,6}(\d[\d.,]*)/);
  out.ffijo = ff ? ff[1].replace(/[.,]+$/,'') : ''; out.fmaniobra = fm ? fm[1].replace(/[.,]+$/,'') : '';
  const vs = []; const re = /(?:^|[^A-Z])V\s*[:=]?\s*(-?\d[\d.,]*)/g; let m;
  while((m = re.exec(t))) vs.push(m[1].replace(/[.,]+$/,''));
  ['tabaco','cafe','once','butano'].forEach((k,i)=>{ out[k].v = vs[i]||''; });
  if(vs.length && vs.length!==4) dud.push('He encontrado '+vs.length+' cifras de ventas (V) y esperaba 4: comprueba que cada una está en su fila.');
  ['tabaco','cafe','once','butano'].forEach(k=>{ if(out[k].a==='') dud.push('Falta el ajuste de '+k+'.'); });
  if(!vs.length && !out.tabaco.a) dud.push('No he reconocido nada. Revisa el texto o escríbelo a mano en la ventana siguiente.');
  out.dudas = dud.join(' ');
  return out;
}


/* ============================================================
   LEER LOS EXCEL DE INVENTARIO (AJUSTE/VENTAS de café, tabaco y once-butano)
   Reglas de Samuel:
   - Café y tabaco: AJUSTE → última cifra de la columna Aj.Físico; VENTAS → última cifra de la columna Ventas.
   - Once y butano (un solo archivo por tipo): bloques que terminan en la fila «MERCANCIA EN DEPÓSI».
       butano = bloque con BUTANO/PROPANO; recargas y «cupón electrónico/super once» se ignoran;
       once = el resto de bloques (se suman). Ventas siempre en positivo.
   - Sin archivo de una familia (o sin su bloque) = la estación no dispone de ella.
   ============================================================ */
function decodeReportRows(ab){
  const u = new Uint8Array(ab); let text=null;
  if(u[0]===0xFF && u[1]===0xFE) text = new TextDecoder('utf-16le').decode(u);
  else if(u[0]===0xFE && u[1]===0xFF) text = new TextDecoder('utf-16be').decode(u);
  else if(u[0]===0xD0 && u[1]===0xCF || (u[0]===0x50 && u[1]===0x4B)){            // .xls binario o .xlsx de verdad
    const wb = XLSX.read(ab,{type:'array'});
    return XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]],{header:1,raw:true,defval:''});
  }else text = new TextDecoder('utf-8').decode(u);
  return text.replace(/\r/g,'').split('\n').map(l=>l.split('\t'));
}
function repNum(v){
  if(typeof v==='number') return isFinite(v)?v:NaN;
  let s=String(v==null?'':v).replace(/\s/g,'').replace(/[€%]/g,'');
  if(!/^-?[\d.,]+$/.test(s) || !/\d/.test(s)) return NaN;
  s = s.includes(',') ? s.replace(/\./g,'').replace(',','.') : (/^-?\d{1,3}(\.\d{3})+$/.test(s) ? s.replace(/\./g,'') : s);
  const n=Number(s); return isFinite(n)?n:NaN;
}
function fmtEs(n){
  n=Math.round(n*100)/100; if(Object.is(n,-0)) n=0;
  return (Number.isInteger(n)?String(n):n.toFixed(2)).replace('.',',');
}
function repCols(rows){
  for(let i=0;i<rows.length;i++){
    const r=rows[i].map(x=>normTxt(x).replace(/\s+/g,''));
    const ajf=r.findIndex(x=>/^AJ\.?FISICO$/.test(x)), ven=r.findIndex(x=>x==='VENTAS');
    if(ajf>=0 && ven>=0){
      const desc=r.findIndex(x=>/^DESCRIPCION$/.test(x));
      return {hdr:i, ajf, ven, desc: desc>=0?desc:-1, mat: r.findIndex(x=>x==='MATERIAL')};
    }
  }
  return null;
}
function lastNum(rows, cols, ci){
  for(let i=rows.length-1;i>cols.hdr;i--){ const n=repNum(rows[i][ci]); if(!isNaN(n) && String(rows[i][ci]).trim()!=='') return n; }
  return NaN;
}
function reportBlocks(rows, cols){
  const blocks=[]; let cur=[];
  for(let i=cols.hdr+1;i<rows.length;i++){
    const r=rows[i], d=normTxt(r[cols.desc]).trim();
    if(/^MERCANCIA EN DEPOSI/.test(d)){ blocks.push({items:cur, ajf:repNum(r[cols.ajf])||0, ven:repNum(r[cols.ven])||0}); cur=[]; }
    else if(/^N.?\s*ART/.test(d)) cur=[];
    else if(d && String(r[cols.mat]||'').trim()) cur.push(d);
  }
  return blocks;
}
function classifyBlock(b){
  const t=b.items;
  if(t.some(d=>/BUTAN|PROPAN/.test(d))) return 'butano';
  if(t.some(d=>/RECARGA/.test(d))) return 'recargas';
  if(t.length && t.every(d=>/CUPON ELECTRONICO|SUPER ONCE/.test(d))) return 'ignorar';
  return t.length ? 'once' : 'ignorar';
}
function reportKind(name, rows){
  const n=normTxt(name);
  const fam = /CAFE/.test(n)?'cafe' : /TABACO/.test(n)?'tabaco' : /ONCE|BUTANO/.test(n)?'oncebutano' : '';
  let kind = /AJUSTE/.test(n)?'ajuste' : /VENTAS/.test(n)?'ventas' : '';
  let f2=fam;
  const txt=rows.slice(0,12).map(r=>r.join(' ')).join(' ');
  if(!f2){ const m=txt.match(/\[\]\s*(\d+)/); if(m) f2 = m[1]==='138'?'cafe' : m[1]==='130'?'tabaco' : m[1]==='109'?'oncebutano' : ''; }
  if(!kind){ kind = /DIFERENCIAS DE INVENTARIO/i.test(txt) ? 'ventas' : ''; }
  return {fam:f2, kind};
}
/* files = [{name, ab}] */
function parseInventoryReports(files){
  const out={tabaco:{a:'',v:''},cafe:{a:'',v:''},once:{a:'',v:''},butano:{a:'',v:''},ffijo:'',fmaniobra:'',nodispone:[],dudas:'',info:''};
  const dud=[], info=[], seen={};
  const data={cafe:{},tabaco:{},oncebutano:{}};
  files.forEach(f=>{
    let rows; try{ rows=decodeReportRows(f.ab); }catch(e){ dud.push('No he podido abrir «'+f.name+'».'); return; }
    const k=reportKind(f.name, rows), cols=repCols(rows);
    if(!k.fam || !k.kind || !cols){ dud.push('No reconozco «'+f.name+'» (esperaba AJUSTE o VENTAS de CAFE, TABACO u ONCE-BUTANO).'); return; }
    data[k.fam][k.kind]={rows,cols}; seen[k.fam]=true;
  });
  ['cafe','tabaco'].forEach(fam=>{
    const d=data[fam];
    if(d.ajuste){ const n=lastNum(d.ajuste.rows,d.ajuste.cols,d.ajuste.cols.ajf); if(isNaN(n)) dud.push('No encuentro el ajuste físico de '+fam+'.'); else out[fam].a=fmtEs(n); }
    if(d.ventas){ const n=lastNum(d.ventas.rows,d.ventas.cols,d.ventas.cols.ven); if(isNaN(n)) dud.push('No encuentro las ventas de '+fam+'.'); else out[fam].v=fmtEs(Math.abs(n)); }
    if(!d.ajuste && !d.ventas) out.nodispone.push(fam);
    else if(!d.ajuste) dud.push('Falta el archivo AJUSTE de '+fam+'.');
    else if(!d.ventas) dud.push('Falta el archivo VENTAS de '+fam+'.');
  });
  const ob=data.oncebutano;
  if(!ob.ajuste && !ob.ventas){ out.nodispone.push('once','butano'); }
  else{
    if(!ob.ajuste) dud.push('Falta el archivo AJUSTE de ONCE-BUTANO.');
    if(!ob.ventas) dud.push('Falta el archivo VENTAS de ONCE-BUTANO.');
    const agg={};
    ['ajuste','ventas'].forEach(kind=>{
      if(!ob[kind]) return;
      const bl=reportBlocks(ob[kind].rows, ob[kind].cols), g={butano:[],once:[]};
      bl.forEach(b=>{ const c=classifyBlock(b); if(g[c]) g[c].push(b); });
      agg[kind]=g;
    });
    ['butano','once'].forEach(fam=>{
      const present = ['ajuste','ventas'].filter(k=>agg[k] && agg[k][fam].length);
      if(!present.length && agg.ajuste && agg.ventas){ out.nodispone.push(fam); return; }
      if(agg.ajuste){ const s=agg.ajuste[fam].reduce((t,b)=>t+b.ajf,0); if(agg.ajuste[fam].length) out[fam].a=fmtEs(s); }
      if(agg.ventas){ const s=agg.ventas[fam].reduce((t,b)=>t+b.ven,0); if(agg.ventas[fam].length) out[fam].v=fmtEs(Math.abs(s)); }
      if(agg.ajuste && agg.ventas && (!agg.ajuste[fam].length || !agg.ventas[fam].length) && present.length)
        dud.push('De '+fam+' solo aparece en uno de los dos archivos: revisa las cifras.');
    });
  }
  if(out.nodispone.length) info.push('No dispone de: '+out.nodispone.join(', ')+' (no hay datos). Al aplicar se pondrá N/A y «NO DISPONE».');
  out.dudas=dud.join(' '); out.info=info.join(' ');
  return out;
}

function sanitizeId(s){ return String(s||'').replace(/[^a-zA-Z0-9]/g,'_'); }

function newAudit(parsed, fileName, settings, buf){
  const m = parsed.meta;
  return {
    id: sanitizeId(m.ceco)+'_'+sanitizeId(m.fecha),
    ceco: m.ceco, name: m.name, cPN: m.cPN, zona: m.zona, fecha: m.fecha,
    da: '',
    auditor: (settings && settings.auditor) || 'SAMUEL VACA',
    fileName, originalXlsx: buf,   // Excel original, intacto, para exportarlo al final
    questions: parsed.questions,
    photos: {},        // {'0': cartel, '1'..'76': fotos de cada pregunta}
    sinServicios: false,
    status: 'en_curso',
    createdAt: Date.now(), updatedAt: Date.now(),
  };
}

/* Actualiza respuestas desde un Excel nuevo conservando fotos y comentarios */
function mergeAudit(audit, parsed, fileName, buf){
  const m = parsed.meta;
  const oldByNum = {};
  (audit.questions||[]).forEach(q=>{ oldByNum[q.num]=q; });
  parsed.questions.forEach(nq=>{
    const o = oldByNum[nq.num];
    if(o && o.orig && (o.status!==o.orig.status || o.valor!==o.orig.valor || o.obs!==o.orig.obs)){
      nq.status=o.status; nq.valor=o.valor; nq.obs=o.obs; if(o.inv){ nq.inv=true; nq.ajuste=o.ajuste; nq.ventas=o.ventas; }     // lo que rellenaste en la app tiene prioridad
    }
  });
  audit.ceco = m.ceco; audit.name = m.name; audit.cPN = m.cPN;
  audit.zona = m.zona; audit.fecha = m.fecha;
  audit.questions = parsed.questions;
  audit.fileName = fileName; audit.originalXlsx = buf;
  audit.updatedAt = Date.now();
  return audit;
}

function auditStats(a){
  const st = {SI:0, NO:0, NA:0, PTE:0, '':0, photos:0, ncSinFoto:0};
  a.questions.forEach(q=>{
    st[q.status]++;
    const n = (a.photos[q.num]||[]).length;
    st.photos += n;
    if(q.status==='NO' && n===0) st.ncSinFoto++;
  });
  st.photos += (a.photos['0']||[]).length;
  return st;
}

function esc(s){
  return String(s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;');
}

function pdfFileName(a){
  const nm = String(a.name||'ESTACION').trim().replace(/\s+/g,'_').replace(/[^A-Za-z0-9ÁÉÍÓÚÑáéíóúñ_-]/g,'');
  return 'INFORME_ANEXO_Fotografico_'+nm+'_'+(a.ceco||'')+'.pdf';
}

/* ---------------- Vista previa HTML (misma estructura que el PDF) ---------------- */
function imgTag(p){
  const pct = Math.round((SIZE_FRAC[p.size||DEF_SIZE]||1)*100);
  return '<img src="'+p.data+'" style="width:calc('+pct+'% - 6px)">';
}
function buildPreviewHTML(a){
  const hdr = [['D.A.',''],['CENTRO DE COSTE',a.ceco],['NOMBRE E.S.',a.name],['FECHA',a.fecha],['AUDITOR',a.auditor]];
  let h = '<div class="pv-page"><table class="pv-hdr">' +
    hdr.map(([k,v])=>'<tr><th>'+esc(k)+'</th><td>'+esc(v||'')+'</td></tr>').join('') + '</table>';
  h += '<div class="pv-cartel-title">CARTEL IDENTIFICACIÓN ES</div>';
  h += '<div class="pv-cartel">'+ (a.photos['0']||[]).map(imgTag).join('') + '</div>';
  a.questions.forEach(q=>{
    const s = STATUS[q.status];
    const photos = a.photos[q.num]||[];
    h += '<table class="pv-q"><tr><th class="pv-l">PREGUNTA '+q.num+'</th><th class="pv-r" style="color:'+s.css+'">'+s.word+'</th></tr>'+
         '<tr><td class="pv-l pv-text">'+esc(q.text)+'</td><td class="pv-r pv-ph">'+photos.map(imgTag).join('')+'</td></tr></table>';
  });
  return h + '</div>';
}

/* ---------------- PDF ---------------- */
const PHOTO_MAXH = 330;   // altura máxima de una foto dentro de la celda (pt)
function photoDims(p, availW, maxH){
  const frac = SIZE_FRAC[p.size||DEF_SIZE] || 1;
  let w = availW*frac;
  let h = w*p.h/p.w;
  if(h>maxH){ h = maxH; w = h*p.w/p.h; }
  return {w,h};
}
/* Coloca las fotos en filas dentro de un ancho disponible */
function layoutPhotos(photos, availW, gap){
  const rows=[]; let row=null;
  photos.forEach(p=>{
    const d = photoDims(p, availW, PHOTO_MAXH);
    if(!row || row.w + gap + d.w > availW + 0.5){ row={items:[],w:0,h:0}; rows.push(row); }
    row.w += (row.items.length? gap:0) + d.w;
    row.items.push({p,w:d.w,h:d.h});
    row.h = Math.max(row.h, d.h);
  });
  return rows;
}

function buildPdfBlob(a){
  const doc = new jspdf.jsPDF({unit:'pt', format:'a4'});
  const PW = doc.internal.pageSize.getWidth();
  const PH = doc.internal.pageSize.getHeight();
  const M = 40, CW = PW-2*M, BOT = PH-M;
  const LW = Math.round(CW*0.30), RW = CW-LW;
  const PAD = 6, GAP = 6, HEAD = 20;
  let y = M;

  doc.setLineWidth(0.8);
  doc.setDrawColor(0,0,0);

  // ---- cabecera (D.A. se deja en blanco) ----
  const hdr = [['D.A.',''],['CENTRO DE COSTE',a.ceco],['NOMBRE E.S.',a.name],['FECHA',a.fecha],['AUDITOR',a.auditor]];
  hdr.forEach(([k,v])=>{
    doc.setFont('helvetica','bold'); doc.setFontSize(10); doc.setTextColor(0,0,0);
    doc.rect(M,y,LW,20); doc.rect(M+LW,y,RW,20);
    doc.text(k, M+5, y+14);
    doc.setFont('helvetica','normal');
    doc.text(String(v||''), M+LW+5, y+14);
    y += 20;
  });
  y += 14;

  // ---- cartel ----
  doc.setFont('helvetica','bold'); doc.setFontSize(10); doc.setTextColor(0,0,0);
  doc.text('CARTEL IDENTIFICACIÓN ES', PW/2, y+10, {align:'center'});
  y += 20;
  (a.photos['0']||[]).forEach(p=>{
    const d = photoDims(p, CW, 260);
    if(y+d.h>BOT){ doc.addPage(); y=M; }
    try{ doc.addImage(p.data,'JPEG',M+(CW-d.w)/2,y,d.w,d.h); }catch(e){}
    y += d.h+10;
  });
  y += 8;

  // ---- preguntas: la foto va DENTRO de la celda derecha, que crece con ella ----
  a.questions.forEach(q=>{
    const s = STATUS[q.status];
    const photos = a.photos[q.num]||[];
    doc.setFont('helvetica','bold'); doc.setFontSize(9.5);
    const qLines = doc.splitTextToSize(q.text, LW-10);
    const qH = Math.max(20, qLines.length*12+8);
    const rows = layoutPhotos(photos, RW-2*PAD, GAP);
    const photosH = rows.reduce((t,r,i)=>t+r.h+(i?GAP:0),0);
    const fullH = HEAD + Math.max(qH, rows.length? photosH+2*PAD : 0);

    // si todo el bloque cabe en una página nueva pero no en el hueco actual, salto de página
    if(y+fullH>BOT && fullH<=BOT-M && y>M){ doc.addPage(); y=M; }

    let i=0;
    do{
      // espacio para fotos en este tramo
      let space = BOT - y - HEAD - 2*PAD;
      if(rows.length && rows[i].h>space && y>M){ doc.addPage(); y=M; space = BOT-y-HEAD-2*PAD; }
      let j=i, hs=0;
      while(j<rows.length){
        const add = rows[j].h+(j>i?GAP:0);
        if(j===i || hs+add<=space){ hs+=add; j++; } else break;
      }
      const row2H = Math.max(qH, (j>i)? hs+2*PAD : 0);

      // fila 1: PREGUNTA N | ESTADO
      doc.setDrawColor(0,0,0); doc.setLineWidth(0.8);
      doc.rect(M,y,LW,HEAD); doc.rect(M+LW,y,RW,HEAD);
      doc.setFont('helvetica','bold'); doc.setFontSize(10); doc.setTextColor(0,0,0);
      doc.text('PREGUNTA '+q.num, M+5, y+14);
      doc.setTextColor(s.color[0],s.color[1],s.color[2]);
      doc.text(s.word, M+LW+5, y+14);
      y += HEAD;

      // fila 2: texto de la pregunta | fotos
      doc.rect(M,y,LW,row2H); doc.rect(M+LW,y,RW,row2H);
      doc.setFont('helvetica','bold'); doc.setFontSize(9.5);
      doc.setTextColor(ORANGE[0],ORANGE[1],ORANGE[2]);
      doc.text(qLines, M+5, y+13);
      let py = y+PAD;
      for(let r=i;r<j;r++){
        const row=rows[r];
        let px = M+LW+PAD + ((RW-2*PAD)-row.w)/2;   // fila centrada en la celda
        row.items.forEach(it=>{
          try{ doc.addImage(it.p.data,'JPEG',px,py,it.w,it.h); }catch(e){}
          px += it.w+GAP;
        });
        py += row.h+GAP;
      }
      y += row2H;
      i=j;
      if(i<rows.length){ doc.addPage(); y=M; }   // continúa en la página siguiente
    }while(i<rows.length);

    y += 14;
    if(y>BOT){ doc.addPage(); y=M; }
  });

  return doc.output('blob');
}


/* ============================================================
   EXPORTAR EL EXCEL ORIGINAL TOCANDO SOLO F (valor), G (valoración) y H (observaciones)
   - No se regenera el libro: se edita el XML de esas celdas dentro del archivo original.
   - Las valoraciones reutilizan las mismas cadenas compartidas que la lista desplegable (K1:K4),
     igual que hace Excel; los textos nuevos se añaden al final de sharedStrings.
   - Funciona tanto con el XML de Repsol (etiquetas con prefijo x:) como con el guardado por Excel.
   ============================================================ */
const XL_LABEL = {SI:'SÍ', NO:'NO', NA:'N/A', PTE:'PTE'};
function xmlEsc(s){ return String(s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;'); }
function xmlUnesc(s){
  return String(s).replace(/&lt;/g,'<').replace(/&gt;/g,'>').replace(/&quot;/g,'"').replace(/&apos;/g,"'")
    .replace(/&#x([0-9a-f]+);/gi,(m,n)=>String.fromCharCode(parseInt(n,16)))
    .replace(/&#(\d+);/g,(m,n)=>String.fromCharCode(+n)).replace(/&amp;/g,'&');
}
function cleanText(s, max){
  let t = String(s===undefined||s===null?'':s).replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F]/g,'').replace(/[\r\n\t]+/g,' ').trim();
  return t.length>max ? t.slice(0,max) : t;
}
function colNum(letters){ let n=0; for(const ch of letters) n=n*26+(ch.charCodeAt(0)-64); return n; }
function detectPrefix(xml, tag){
  const m = xml.match(new RegExp('<(?:([A-Za-z_][\\w.-]*):)?'+tag+'[\\s>/]'));
  if(!m) throw new Error('No encuentro <'+tag+'> en el Excel.');
  return m[1] ? m[1]+':' : '';
}
function parseSst(xml){
  const sp = detectPrefix(xml,'sst');
  const re = new RegExp('<'+sp+'si\\b[^>]*?(?:/>|>([\\s\\S]*?)</'+sp+'si>)','g');
  const tre = new RegExp('<'+sp+'t\\b[^>]*>([\\s\\S]*?)</'+sp+'t>','g');
  const map = new Map(); let n = 0, m;
  while((m = re.exec(xml))){
    let text = '';
    if(m[1]) for(const t of m[1].matchAll(tre)) text += xmlUnesc(t[1]);
    if(!map.has(text)) map.set(text, n);
    n++;
  }
  return {sp, map, n};
}
function setCellString(sheet, sp, ref, idx, stat){
  const rowNum = ref.replace(/^[A-Z]+/,'');
  const rowRe = new RegExp('<'+sp+'row r="'+rowNum+'"[^>]*[^/]>[\\s\\S]*?</'+sp+'row>');
  const rm = sheet.match(rowRe);
  if(!rm) throw new Error('No encuentro la fila '+rowNum+' del Excel.');
  let row = rm[0];
  const cellRe = new RegExp('<'+sp+'c r="'+ref+'"([^>]*?)\\s*(?:/>|>([\\s\\S]*?)</'+sp+'c>)');
  const cm = row.match(cellRe);
  const selfClose = sp ? ' />' : '/>';
  const mk = (sAttr)=> idx===null
      ? '<'+sp+'c r="'+ref+'"'+sAttr+selfClose
      : '<'+sp+'c r="'+ref+'"'+sAttr+' t="s"><'+sp+'v>'+idx+'</'+sp+'v></'+sp+'c>';
  if(cm){
    const attrs = cm[1]||'';
    const sAttr = (attrs.match(/\ss="[^"]*"/)||[''])[0];
    if(/\st="s"/.test(attrs) && /<(?:\w+:)?v>/.test(cm[2]||'')) stat.removed++;
    if(idx!==null) stat.added++;
    row = row.replace(cellRe, ()=>mk(sAttr));
  }else{
    if(idx===null) return sheet;
    stat.added++;
    const target = colNum(ref.replace(/\d+$/,''));
    let at = -1;
    const cre = new RegExp('<'+sp+'c r="([A-Z]+)\\d+"','g'); let m;
    while((m = cre.exec(row))){ if(colNum(m[1])>target){ at = m.index; break; } }
    if(at<0) at = row.lastIndexOf('</'+sp+'row>');
    row = row.slice(0,at) + mk('') + row.slice(at);
  }
  return sheet.replace(rowRe, ()=>row);
}
/* Aplica las respuestas a los XML del hoja y de las cadenas. Solo toca celdas que han cambiado. */
function applyAnswersToXml(sheetXml, sstXml, questions){
  const sp = detectPrefix(sheetXml,'sheetData');
  const sst = parseSst(sstXml);
  const newStrings = []; const stat = {added:0, removed:0};
  let sheet = sheetXml, changed = 0;
  const strIndex = (text)=>{
    if(sst.map.has(text)) return sst.map.get(text);
    const i = sst.n + newStrings.length;
    newStrings.push(text); sst.map.set(text, i);
    return i;
  };
  questions.forEach(q=>{
    const o = q.orig || {status:'',valor:'',obs:''};
    const want = [
      ['F', cleanText(q.valor,10),  cleanText(o.valor,10),  true ],
      ['G', q.status? XL_LABEL[q.status] : '', o.status? XL_LABEL[o.status] : '', false],
      ['H', cleanText(q.obs,100),   cleanText(o.obs,100),   true ],
    ];
    want.forEach(([col, now, before])=>{
      if(now===before) return;                    // sin cambios: no se toca la celda original
      sheet = setCellString(sheet, sp, col+q.row, now? strIndex(now) : null, stat);
      changed++;
    });
  });
  let sstOut = sstXml;
  if(newStrings.length || stat.added!==stat.removed){
    const ssp = sst.sp;
    const add = newStrings.map(t=>'<'+ssp+'si><'+ssp+'t'+(/^\s|\s$/.test(t)?' xml:space="preserve"':'')+'>'+xmlEsc(t)+'</'+ssp+'t></'+ssp+'si>').join('');
    sstOut = sstOut.replace('</'+ssp+'sst>', add+'</'+ssp+'sst>');
    sstOut = sstOut.replace(/(<(?:\w+:)?sst\b[^>]*?\scount=")(\d+)(")/, (m,a,n,b)=>a+Math.max(0,(+n)+stat.added-stat.removed)+b);
    sstOut = sstOut.replace(/(<(?:\w+:)?sst\b[^>]*?\suniqueCount=")(\d+)(")/, (m,a,n,b)=>a+((+n)+newStrings.length)+b);
  }
  return {sheet, sst:sstOut, changed, newStrings:newStrings.length};
}
async function buildXlsxBlob(audit){
  if(!audit.originalXlsx) throw new Error('Este informe se creó con una versión anterior y no guarda el Excel original. Pulsa «Excel nuevo» y vuelve a cargar el Excel de la estación.');
  const zip = await JSZip.loadAsync(audit.originalXlsx);
  let sheetPath = 'xl/worksheets/sheet1.xml';
  if(!zip.file(sheetPath)){
    const cand = Object.keys(zip.files).filter(p=>/^xl\/worksheets\/sheet\d+\.xml$/.test(p)).sort();
    if(!cand.length) throw new Error('El Excel no tiene ninguna hoja.');
    sheetPath = cand[0];
  }
  const sstFile = zip.file('xl/sharedStrings.xml');
  if(!sstFile) throw new Error('El Excel no tiene tabla de textos (sharedStrings).');
  const res = applyAnswersToXml(await zip.file(sheetPath).async('string'), await sstFile.async('string'), audit.questions);
  zip.file(sheetPath, res.sheet, {createFolders:false});
  zip.file('xl/sharedStrings.xml', res.sst, {createFolders:false});
  return zip.generateAsync({type:'blob', compression:'DEFLATE', mimeType:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'});
}
function xlsxFileName(a){ return a.fileName || ('ENCUESTA_SO_'+(a.name||'ESTACION')+'.xlsx'); }
