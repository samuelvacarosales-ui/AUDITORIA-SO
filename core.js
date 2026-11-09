/* ============================================================
   core.js — lógica sin interfaz: leer el Excel, modelo de datos,
   vista previa HTML y generación del PDF del anexo fotográfico.
   ============================================================ */

const STATUS = {
  SI:  {word:'CORRECTO',    color:[0,140,60],   css:'#008c3c'},
  NO:  {word:'NO CONFORME', color:[220,0,0],    css:'#dc0000'},
  NA:  {word:'NA',          color:[0,0,0],      css:'#000000'},
  PTE: {word:'PENDIENTE',   color:[200,120,0],  css:'#c87800'},
  '':  {word:'SIN VALORAR', color:[120,120,120],css:'#787878'},
};
const ORANGE = [247,148,29];
const SIZE_FRAC = {S:0.48, M:0.72, L:1};

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
    da: (settings && settings.da) || '',
    auditor: (settings && settings.auditor) || 'SAMUEL VACA',
    fileName,
    questions: parsed.questions,
    comments: {},      // solo comentarios editados a mano: {num: texto}
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

function defaultComment(q){
  const lines = [];
  if(q.obs) lines.push(q.obs);
  if(q.valor) lines.push('VALOR/FECHA: '+q.valor);
  return lines.join('\n');
}
function commentFor(audit, q){
  const c = audit.comments[q.num];
  return (c !== undefined) ? c : defaultComment(q);
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
function buildPreviewHTML(a){
  const hdr = [['D.A.',a.da],['CENTRO DE COSTE',a.ceco],['NOMBRE E.S.',a.name],['FECHA',a.fecha],['AUDITOR',a.auditor]];
  let h = '<div class="pv-page"><table class="pv-hdr">' +
    hdr.map(([k,v])=>'<tr><th>'+esc(k)+'</th><td>'+esc(v||'')+'</td></tr>').join('') + '</table>';
  h += '<div class="pv-cartel-title">CARTEL IDENTIFICACIÓN ES</div>';
  h += '<div class="pv-cartel">'+ (a.photos['0']||[]).map(p=>'<img src="'+p.data+'" style="width:'+Math.round(SIZE_FRAC[p.size||'M']*100)+'%">').join('') + '</div>';
  a.questions.forEach(q=>{
    const s = STATUS[q.status];
    const photos = a.photos[q.num]||[];
    const com = commentFor(a,q);
    h += '<table class="pv-q"><tr><th class="pv-l">PREGUNTA '+q.num+'</th><th class="pv-r" style="color:'+s.css+'">'+s.word+'</th></tr>'+
         '<tr><td class="pv-l pv-text">'+esc(q.text)+'</td><td class="pv-r"></td></tr></table>';
    if(com || photos.length){
      h += '<div class="pv-body"><div class="pv-com">'+esc(com).replace(/\n/g,'<br>')+'</div><div class="pv-ph">'+
           photos.map(p=>'<img src="'+p.data+'" style="width:'+Math.round(SIZE_FRAC[p.size||'M']*100)+'%">').join('')+'</div></div>';
    }
  });
  return h + '</div>';
}

/* ---------------- PDF ---------------- */
function photoDims(p, availW, maxH){
  const frac = SIZE_FRAC[p.size||'M'] || SIZE_FRAC.M;
  let w = availW*frac;
  let h = w*p.h/p.w;
  if(h>maxH){ h = maxH; w = h*p.w/p.h; }
  return {w,h};
}

function buildPdfBlob(a){
  const doc = new jspdf.jsPDF({unit:'pt', format:'a4'});
  const PW = doc.internal.pageSize.getWidth();
  const PH = doc.internal.pageSize.getHeight();
  const M = 40, CW = PW-2*M, BOT = PH-M;
  let y = M;
  const LW = Math.round(CW*0.30), RW = CW-LW;

  doc.setLineWidth(0.8);
  doc.setDrawColor(0,0,0);

  // ---- cabecera ----
  const hdr = [['D.A.',a.da],['CENTRO DE COSTE',a.ceco],['NOMBRE E.S.',a.name],['FECHA',a.fecha],['AUDITOR',a.auditor]];
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

  // ---- preguntas ----
  a.questions.forEach(q=>{
    const s = STATUS[q.status];
    const photos = a.photos[q.num]||[];
    const com = commentFor(a,q);

    doc.setFont('helvetica','bold'); doc.setFontSize(9.5);
    const qLines = doc.splitTextToSize(q.text, LW-10);
    const row2H = Math.max(20, qLines.length*12+8);
    doc.setFont('helvetica','normal'); doc.setFontSize(9);
    const comLines = com ? doc.splitTextToSize(com, LW-6) : [];
    const firstPh = photos.length ? Math.min(photoDims(photos[0], RW-8, 280).h, 120) : 0;
    const need = 20 + row2H + 6 + Math.max(firstPh, Math.min(comLines.length*11, 33));
    if(y+need > BOT){ doc.addPage(); y=M; }

    // fila 1
    doc.setDrawColor(0,0,0); doc.setLineWidth(0.8);
    doc.rect(M,y,LW,20); doc.rect(M+LW,y,RW,20);
    doc.setFont('helvetica','bold'); doc.setFontSize(10); doc.setTextColor(0,0,0);
    doc.text('PREGUNTA '+q.num, M+5, y+14);
    doc.setTextColor(s.color[0],s.color[1],s.color[2]);
    doc.text(s.word, M+LW+5, y+14);
    y += 20;
    // fila 2
    doc.rect(M,y,LW,row2H); doc.rect(M+LW,y,RW,row2H);
    doc.setFont('helvetica','bold'); doc.setFontSize(9.5);
    doc.setTextColor(ORANGE[0],ORANGE[1],ORANGE[2]);
    doc.text(qLines, M+5, y+13);
    y += row2H + 6;

    // comentario (izquierda) + fotos (derecha)
    let commentBottom = y;
    if(comLines.length){
      doc.setFont('helvetica','normal'); doc.setFontSize(9); doc.setTextColor(0,0,0);
      doc.text(comLines, M+3, y+9);
      commentBottom = y + comLines.length*11 + 2;
    }
    let rowY = y, cx = 0, rowH = 0, drewAny = false;
    const x0 = M+LW+8, avail = RW-8;
    photos.forEach(p=>{
      const d = photoDims(p, avail, 280);
      if(cx>0 && cx+d.w>avail+0.5){ rowY += rowH+8; cx=0; rowH=0; }
      if(rowY+d.h>BOT){ doc.addPage(); rowY=M; cx=0; rowH=0; commentBottom=0; }
      try{ doc.addImage(p.data,'JPEG',x0+cx,rowY,d.w,d.h); }catch(e){}
      cx += d.w+8; rowH = Math.max(rowH,d.h); drewAny = true;
    });
    const photosBottom = drewAny ? rowY+rowH : y;
    y = Math.max(commentBottom, photosBottom) + 14;
  });

  return doc.output('blob');
}
