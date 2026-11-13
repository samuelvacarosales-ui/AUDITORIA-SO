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
const SERVICE_Q = {6:'CONTADORES', 8:'COMPRESOR', 56:'SERVICIOS COMPLEMENTARIOS', 65:'SERVICIOS COMPLEMENTARIOS', 71:'SERVICIOS COMPLEMENTARIOS'};
const SERVICE_OBS = 'NO DISPONE';
function normTxt(s){ return String(s||'').toUpperCase().normalize('NFD').replace(/[̀-ͯ]/g,''); }
function ruleMatches(q, kw){ return normTxt(q.text).includes(kw); }
/* Devuelve la lista de preguntas cuya regla no coincide con su texto */
function applyDefaults(questions){
  const warn = [];
  questions.forEach(q=>{
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
      nq.status=o.status; nq.valor=o.valor; nq.obs=o.obs;     // lo que rellenaste en la app tiene prioridad
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
