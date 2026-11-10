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
const DEF_SIZE = 'L';

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

function sanitizeId(s){ return String(s||'').replace(/[^a-zA-Z0-9]/g,'_'); }

function newAudit(parsed, fileName, settings){
  const m = parsed.meta;
  return {
    id: sanitizeId(m.ceco)+'_'+sanitizeId(m.fecha),
    ceco: m.ceco, name: m.name, cPN: m.cPN, zona: m.zona, fecha: m.fecha,
    da: '',
    auditor: (settings && settings.auditor) || 'SAMUEL VACA',
    fileName,
    questions: parsed.questions,
    photos: {},        // {'0': cartel, '1'..'76': fotos de cada pregunta}
    status: 'en_curso',
    createdAt: Date.now(), updatedAt: Date.now(),
  };
}

/* Actualiza respuestas desde un Excel nuevo conservando fotos y comentarios */
function mergeAudit(audit, parsed, fileName){
  const m = parsed.meta;
  audit.ceco = m.ceco; audit.name = m.name; audit.cPN = m.cPN;
  audit.zona = m.zona; audit.fecha = m.fecha;
  audit.questions = parsed.questions;
  audit.fileName = fileName;
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
