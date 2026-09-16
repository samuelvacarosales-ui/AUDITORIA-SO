/* ============================================================
   Auditor SO — lógica principal
   ============================================================ */

const VAL_OPTIONS = [
  {key:'SI',  label:'SÍ',  pdfWord:'CORRECTO'},
  {key:'NO',  label:'NO',  pdfWord:'NO CONFORME'},
  {key:'NA',  label:'N/A', pdfWord:'N/A'},
  {key:'PTE', label:'PTE', pdfWord:'PENDIENTE'},
];

let db = null;
let currentAudit = null;   // objeto en memoria de la auditoría abierta
let saveTimer = null;

/* ---------------- IndexedDB ---------------- */
function openDB(){
  return new Promise((resolve, reject)=>{
    const req = indexedDB.open('auditorSO', 1);
    req.onupgradeneeded = e=>{
      const d = e.target.result;
      if(!d.objectStoreNames.contains('audits')){
        d.createObjectStore('audits', {keyPath:'id'});
      }
    };
    req.onsuccess = e=>resolve(e.target.result);
    req.onerror = e=>reject(e);
  });
}
function dbPut(audit){
  return new Promise((resolve,reject)=>{
    const tx = db.transaction('audits','readwrite');
    tx.objectStore('audits').put(audit);
    tx.oncomplete = ()=>resolve();
    tx.onerror = e=>reject(e);
  });
}
function dbGetAll(){
  return new Promise((resolve,reject)=>{
    const tx = db.transaction('audits','readonly');
    const req = tx.objectStore('audits').getAll();
    req.onsuccess = ()=>resolve(req.result);
    req.onerror = e=>reject(e);
  });
}
function dbDelete(id){
  return new Promise((resolve,reject)=>{
    const tx = db.transaction('audits','readwrite');
    tx.objectStore('audits').delete(id);
    tx.oncomplete = ()=>resolve();
    tx.onerror = e=>reject(e);
  });
}

/* ---------------- Settings (localStorage) ---------------- */
function getSettings(){
  try{ return JSON.parse(localStorage.getItem('auditorSO_settings')) || {}; }
  catch(e){ return {}; }
}
function saveSettings(s){ localStorage.setItem('auditorSO_settings', JSON.stringify(s)); }

/* ---------------- Utilidades ---------------- */
function toast(msg, ms=2200){
  const t = document.getElementById('toast');
  t.textContent = msg;
  t.classList.add('show');
  clearTimeout(toast._tm);
  toast._tm = setTimeout(()=>t.classList.remove('show'), ms);
}
function sanitizeId(s){ return (s||'').toString().replace(/[^a-zA-Z0-9]/g,'_'); }
function fmtDate(iso){
  if(!iso) return '';
  const d = new Date(iso);
  if(isNaN(d)) return iso;
  return d.toLocaleDateString('es-ES')+' '+d.toLocaleTimeString('es-ES',{hour:'2-digit',minute:'2-digit'});
}

/* ============================================================
   NAVEGACIÓN
   ============================================================ */
function showScreen(id){
  document.querySelectorAll('.screen').forEach(s=>s.classList.add('hidden'));
  document.getElementById(id).classList.remove('hidden');
  document.getElementById('bottombar').classList.toggle('hidden', id !== 'screen-audit');
  document.getElementById('backBtn').classList.toggle('hidden', id === 'screen-home');
  document.getElementById('settingsBtn').classList.toggle('hidden', id !== 'screen-home');
  const titles = {
    'screen-home':['Auditor SO','Repsol · Supervisión Operativa'],
    'screen-audit':[currentAudit?currentAudit.ceco:'Auditoría', currentAudit?('CECO '+currentAudit.ceco):''],
    'screen-settings':['Ajustes','Se guardan en este dispositivo'],
  };
  const [t,s] = titles[id]||['',''];
  document.getElementById('topTitle').textContent = t;
  document.getElementById('topSub').textContent = s;
}

document.getElementById('backBtn').addEventListener('click', ()=>{
  if(currentAudit) saveCurrentAudit(false);
  currentAudit = null;
  renderHome();
  showScreen('screen-home');
});
document.getElementById('settingsBtn').addEventListener('click', ()=>{
  const s = getSettings();
  document.getElementById('setAuditor').value = s.auditor || '';
  document.getElementById('setDA').value = s.da || '';
  showScreen('screen-settings');
});
document.getElementById('saveSettingsBtn').addEventListener('click', ()=>{
  saveSettings({
    auditor: document.getElementById('setAuditor').value.trim(),
    da: document.getElementById('setDA').value.trim(),
  });
  toast('Ajustes guardados');
  showScreen('screen-home');
});

/* ============================================================
   HOME — listado de auditorías
   ============================================================ */
async function renderHome(){
  const audits = await dbGetAll();
  audits.sort((a,b)=> (b.updatedAt||0) - (a.updatedAt||0));
  const ongoing = audits.filter(a=>a.status!=='finalizado');
  const finished = audits.filter(a=>a.status==='finalizado');

  const ongoingList = document.getElementById('ongoingList');
  const finishedList = document.getElementById('finishedList');
  ongoingList.innerHTML = '';
  finishedList.innerHTML = '';

  if(ongoing.length===0){
    ongoingList.innerHTML = '<div class="empty-state">No tienes auditorías en curso.<br>Carga un Excel para empezar una nueva.</div>';
  }
  ongoing.forEach(a=> ongoingList.appendChild(buildAuditCard(a)));

  document.getElementById('finishedTitle').style.display = finished.length? 'block':'none';
  finished.forEach(a=> finishedList.appendChild(buildAuditCard(a)));
}

function buildAuditCard(a){
  const total = 76;
  const answered = Object.values(a.answers||{}).filter(x=>x.valoracion).length;
  const pct = Math.round(answered/total*100);
  const photos = Object.values(a.answers||{}).reduce((s,x)=>s+((x.photos||[]).length),0);

  const card = document.createElement('div');
  card.className = 'audit-card';
  card.innerHTML = `
    <div class="progress-ring" style="--pct:${pct}"><span>${pct}%</span></div>
    <div class="info">
      <div class="ceco">${a.ceco || '(sin CECO)'}</div>
      <div class="meta">${a.fecha||''} · ${answered}/${total} preguntas · ${photos} fotos</div>
      <div class="meta">Actualizado ${fmtDate(a.updatedAt)}</div>
    </div>
    <div class="del-btn">🗑</div>
  `;
  card.querySelector('.info').addEventListener('click', ()=> openAudit(a.id));
  card.querySelector('.progress-ring').addEventListener('click', ()=> openAudit(a.id));
  card.querySelector('.del-btn').addEventListener('click', async (e)=>{
    e.stopPropagation();
    if(confirm('¿Eliminar esta auditoría de este dispositivo? No se puede deshacer.')){
      await dbDelete(a.id);
      renderHome();
    }
  });
  return card;
}

async function openAudit(id){
  const all = await dbGetAll();
  currentAudit = all.find(a=>a.id===id);
  if(!currentAudit){ toast('No se ha encontrado la auditoría'); return; }
  renderAuditScreen();
  showScreen('screen-audit');
}

/* ============================================================
   CARGA DE EXCEL NUEVO
   ============================================================ */
document.getElementById('fileInputNew').addEventListener('change', async (e)=>{
  const file = e.target.files[0];
  e.target.value = '';
  if(!file) return;

  const missingLibs = checkLibsLoaded();
  if(missingLibs.length){
    alert('No se puede leer el Excel porque faltan por cargar: '+missingLibs.join(', ')+'.\nRevisa que esos archivos estén subidos en la raíz del repositorio de GitHub.');
    return;
  }

  try{
    const buf = await file.arrayBuffer();
    const wb = XLSX.read(buf, {type:'array'});
    const ws = wb.Sheets[wb.SheetNames[0]];
    if(!ws) throw new Error('El archivo no tiene ninguna hoja legible (SheetNames vacío).');
    const get = (addr)=> (ws[addr]? ws[addr].v : '') ;

    const zona = get('B3');
    const cecoRaw = (get('A2')||'').toString();
    const ceco = cecoRaw.replace(/^Centro de Coste:\s*/i,'').trim();
    const fecha = get('G3');
    const cPN = get('B5');
    const codSup = get('G5');

    if(!ceco && !zona && !fecha){
      throw new Error('Se ha leído el Excel pero las celdas A2/B3/G3 están vacías. ¿Es este el Excel original descargado de la plataforma (sin filas/columnas movidas)?');
    }

    const settings = getSettings();
    const id = sanitizeId(ceco)+'_'+sanitizeId(fecha)+'_'+Date.now();

    currentAudit = {
      id, ceco, zona, fecha, cPN, codSupervisor: codSup,
      da: settings.da || '',
      auditor: settings.auditor || '',
      fileName: file.name,
      originalXlsx: buf,           // se guarda tal cual para reabrir al exportar
      answers: {},                 // {num: {valor, valoracion, obs, photos:[{data,comment}]}}
      status: 'en_curso',
      createdAt: Date.now(),
      updatedAt: Date.now(),
    };
    await dbPut(currentAudit);
    toast('Excel cargado: '+ (ceco||file.name));
    renderAuditScreen();
    showScreen('screen-audit');
  }catch(err){
    console.error(err);
    alert('No se ha podido leer este Excel.\n\nDetalle técnico: '+ (err && err.message ? err.message : err) +'\n\nComprueba que es el archivo descargado de la plataforma de Repsol sin modificar su estructura.');
  }
});

/* ============================================================
   PANTALLA DE AUDITORÍA
   ============================================================ */
function renderAuditScreen(){
  renderHeaderCard();
  renderAreas();
  updateProgress();
}

function renderHeaderCard(){
  const a = currentAudit;
  const el = document.getElementById('headerCard');
  el.innerHTML = `
    <div class="header-grid">
      <div class="field"><label>Centro de coste</label><div class="readonly">${a.ceco||'—'}</div></div>
      <div class="field"><label>Zona</label><div class="readonly">${a.zona||'—'}</div></div>
      <div class="field"><label>Fecha</label><div class="readonly">${a.fecha||'—'}</div></div>
      <div class="field"><label>Cº PN</label><div class="readonly">${a.cPN||'—'}</div></div>
      <div class="field full"><label>D.A. (Dirección de Área)</label>
        <input type="text" id="fld-da" value="${a.da||''}" placeholder="Ej. NORESTE">
      </div>
      <div class="field full"><label>Auditor</label>
        <input type="text" id="fld-auditor" value="${a.auditor||''}" placeholder="Nombre del auditor">
      </div>
    </div>
  `;
  document.getElementById('fld-da').addEventListener('input', e=>{ currentAudit.da = e.target.value; scheduleSave(); });
  document.getElementById('fld-auditor').addEventListener('input', e=>{ currentAudit.auditor = e.target.value; scheduleSave(); });
}

function renderAreas(){
  const container = document.getElementById('areasContainer');
  container.innerHTML = '';

  const areas = [];
  QUESTIONS.forEach(q=>{
    let area = areas.find(x=>x.letra===q.area);
    if(!area){ area = {letra:q.area, nombre:q.areaName, apartados:[]}; areas.push(area); }
    let ap = area.apartados.find(x=>x.nombre===q.apartado);
    if(!ap){ ap = {nombre:q.apartado, preguntas:[]}; area.apartados.push(ap); }
    ap.preguntas.push(q);
  });

  areas.forEach(area=>{
    const block = document.createElement('div');
    block.className = 'area-block';
    const answeredInArea = area.apartados.flatMap(ap=>ap.preguntas)
      .filter(q=> currentAudit.answers[q.num] && currentAudit.answers[q.num].valoracion).length;
    const totalInArea = area.apartados.flatMap(ap=>ap.preguntas).length;

    const header = document.createElement('div');
    header.className = 'area-header';
    header.innerHTML = `<span>${area.letra} · ${area.nombre}</span><span class="cnt">${answeredInArea}/${totalInArea}</span>`;
    block.appendChild(header);

    area.apartados.forEach(ap=>{
      const apTitle = document.createElement('div');
      apTitle.className = 'apartado-title';
      apTitle.textContent = ap.nombre;
      block.appendChild(apTitle);

      ap.preguntas.forEach(q=>{
        block.appendChild(buildQuestionCard(q));
      });
    });

    container.appendChild(block);
  });
}

function buildQuestionCard(q){
  if(!currentAudit.answers[q.num]) currentAudit.answers[q.num] = {valor:'', valoracion:'', obs:'', photos:[]};
  const ans = currentAudit.answers[q.num];

  const card = document.createElement('div');
  card.className = 'q-card' + (ans.valoracion? ' answered':'');
  card.dataset.num = q.num;

  card.innerHTML = `
    <div class="q-top">
      <div class="q-num ${ans.valoracion?('a-'+ans.valoracion):''}">${q.num}</div>
      <div class="q-text">${q.text}</div>
    </div>
    <div class="val-row">
      ${VAL_OPTIONS.map(v=>`<div class="val-btn ${ans.valoracion===v.key?('sel-'+v.key):''}" data-val="${v.key}">${v.label}</div>`).join('')}
    </div>
    <div class="mini-row">
      <div class="mini-field"><label>Valor / Fecha</label><input type="text" maxlength="10" value="${ans.valor||''}" data-field="valor"></div>
      <div class="mini-field"><label>Observaciones</label><input type="text" maxlength="100" value="${ans.obs||''}" data-field="obs"></div>
    </div>
    <div class="photos-row" data-photos></div>
  `;

  card.querySelectorAll('.val-btn').forEach(btn=>{
    btn.addEventListener('click', ()=>{
      const val = btn.dataset.val;
      ans.valoracion = (ans.valoracion===val)? '' : val;
      card.querySelectorAll('.val-btn').forEach(b=>b.className = 'val-btn' + (b.dataset.val===ans.valoracion? (' sel-'+ans.valoracion):''));
      card.querySelector('.q-num').className = 'q-num' + (ans.valoracion?(' a-'+ans.valoracion):'');
      card.classList.toggle('answered', !!ans.valoracion);
      scheduleSave();
      updateProgress();
      renderAreaCounts();
    });
  });
  card.querySelector('[data-field="valor"]').addEventListener('input', e=>{ ans.valor = e.target.value; scheduleSave(); });
  card.querySelector('[data-field="obs"]').addEventListener('input', e=>{ ans.obs = e.target.value; scheduleSave(); });

  renderPhotos(card, q.num);
  return card;
}

function renderAreaCounts(){
  document.querySelectorAll('.area-block').forEach(block=>{
    const total = block.querySelectorAll('.q-card').length;
    const answered = block.querySelectorAll('.q-card.answered').length;
    block.querySelector('.cnt').textContent = answered+'/'+total;
  });
}

function renderPhotos(card, num){
  const row = card.querySelector('[data-photos]');
  const ans = currentAudit.answers[num];
  row.innerHTML = '';
  (ans.photos||[]).forEach((p, idx)=>{
    const th = document.createElement('div');
    th.className = 'photo-thumb';
    th.innerHTML = `<img src="${p.data}"><div class="rm">✕</div>`;
    th.querySelector('.rm').addEventListener('click', (e)=>{
      e.stopPropagation();
      ans.photos.splice(idx,1);
      renderPhotos(card, num);
      updateProgress();
      scheduleSave();
    });
    row.appendChild(th);
  });
  const addBtn = document.createElement('div');
  addBtn.className = 'add-photo-btn';
  addBtn.textContent = '+';
  addBtn.addEventListener('click', ()=> openPhotoPicker(num, card));
  row.appendChild(addBtn);
}

/* ============================================================
   PROGRESO
   ============================================================ */
function updateProgress(){
  const total = 76;
  const answers = Object.values(currentAudit.answers);
  const answered = answers.filter(a=>a.valoracion).length;
  const photos = answers.reduce((s,a)=>s+((a.photos||[]).length),0);
  document.getElementById('progressFill').style.width = Math.round(answered/total*100)+'%';
  document.getElementById('progressText').textContent = answered+' / '+total+' respondidas';
  document.getElementById('photoCount').textContent = photos+' fotos';
}

/* ============================================================
   GUARDADO
   ============================================================ */
function scheduleSave(){
  clearTimeout(saveTimer);
  saveTimer = setTimeout(()=> saveCurrentAudit(false), 700);
}
async function saveCurrentAudit(showToast){
  if(!currentAudit) return;
  currentAudit.updatedAt = Date.now();
  await dbPut(currentAudit);
  if(showToast) toast('Guardado');
}
document.getElementById('saveBtn').addEventListener('click', ()=> saveCurrentAudit(true));

/* ============================================================
   CÁMARA / RECORTE
   ============================================================ */
let photoTarget = null; // {num, card}
let cropState = null;

function openPhotoPicker(num, card){
  photoTarget = {num, card};
  // En iOS Safari, un <input capture> lanza directamente la cámara.
  // Mostramos elección simple mediante confirm-like: usamos un pequeño menú nativo.
  if(confirm('Pulsa Aceptar para hacer una foto, o Cancelar para elegir de la galería')){
    document.getElementById('cameraInput').click();
  }else{
    document.getElementById('galleryInput').click();
  }
}
document.getElementById('cameraInput').addEventListener('change', handlePickedFile);
document.getElementById('galleryInput').addEventListener('change', handlePickedFile);

function handlePickedFile(e){
  const file = e.target.files[0];
  e.target.value = '';
  if(!file) return;
  const img = new Image();
  const reader = new FileReader();
  reader.onload = ()=>{
    img.onload = ()=> openCropModal(img);
    img.src = reader.result;
  };
  reader.readAsDataURL(file);
}

function openCropModal(img){
  document.getElementById('photoModal').classList.remove('hidden');
  const canvas = document.getElementById('cropCanvas');
  const stage = document.getElementById('cropStage');
  const ctx = canvas.getContext('2d');

  const W = Math.min(stage.clientWidth || 360, 500);
  const H = Math.round(W * 0.75); // salida 4:3
  canvas.width = W; canvas.height = H;
  canvas.style.width = '100%';
  canvas.style.maxWidth = W+'px';

  const scaleFit = Math.max(W/img.width, H/img.height);
  cropState = {
    img, rot:0,
    scale: scaleFit,
    baseScale: scaleFit,
    offX: 0, offY: 0,
    W, H,
    dragging:false, lastX:0, lastY:0,
    pinchDist:0,
  };
  drawCrop();
}

function drawCrop(){
  const c = cropState;
  const canvas = document.getElementById('cropCanvas');
  const ctx = canvas.getContext('2d');
  ctx.save();
  ctx.fillStyle = '#111';
  ctx.fillRect(0,0,c.W,c.H);
  ctx.translate(c.W/2 + c.offX, c.H/2 + c.offY);
  ctx.rotate(c.rot * Math.PI/180);
  ctx.scale(c.scale, c.scale);
  ctx.drawImage(c.img, -c.img.width/2, -c.img.height/2);
  ctx.restore();
}

/* --- interacción: arrastrar y pellizcar --- */
(function bindCropEvents(){
  const canvas = document.getElementById('cropCanvas');
  canvas.addEventListener('touchstart', e=>{
    if(!cropState) return;
    if(e.touches.length===1){
      cropState.dragging = true;
      cropState.lastX = e.touches[0].clientX;
      cropState.lastY = e.touches[0].clientY;
    }else if(e.touches.length===2){
      cropState.dragging = false;
      cropState.pinchDist = dist(e.touches[0], e.touches[1]);
      cropState.pinchScale0 = cropState.scale;
    }
  }, {passive:true});
  canvas.addEventListener('touchmove', e=>{
    if(!cropState) return;
    if(e.touches.length===1 && cropState.dragging){
      const dx = e.touches[0].clientX - cropState.lastX;
      const dy = e.touches[0].clientY - cropState.lastY;
      cropState.offX += dx; cropState.offY += dy;
      cropState.lastX = e.touches[0].clientX;
      cropState.lastY = e.touches[0].clientY;
      drawCrop();
    }else if(e.touches.length===2){
      const d = dist(e.touches[0], e.touches[1]);
      const factor = d / (cropState.pinchDist || d);
      cropState.scale = Math.max(cropState.baseScale*0.5, cropState.pinchScale0 * factor);
      drawCrop();
    }
    e.preventDefault();
  }, {passive:false});
  canvas.addEventListener('touchend', e=>{
    if(!cropState) return;
    cropState.dragging = false;
  });
  // ratón (para pruebas en escritorio)
  canvas.addEventListener('mousedown', e=>{
    if(!cropState) return;
    cropState.dragging = true; cropState.lastX = e.clientX; cropState.lastY = e.clientY;
  });
  window.addEventListener('mousemove', e=>{
    if(!cropState || !cropState.dragging) return;
    const dx = e.clientX - cropState.lastX, dy = e.clientY - cropState.lastY;
    cropState.offX += dx; cropState.offY += dy;
    cropState.lastX = e.clientX; cropState.lastY = e.clientY;
    drawCrop();
  });
  window.addEventListener('mouseup', ()=>{ if(cropState) cropState.dragging=false; });
  canvas.addEventListener('wheel', e=>{
    if(!cropState) return;
    e.preventDefault();
    cropState.scale = Math.max(cropState.baseScale*0.5, cropState.scale * (e.deltaY<0?1.08:0.92));
    drawCrop();
  }, {passive:false});
  function dist(t1,t2){ return Math.hypot(t1.clientX-t2.clientX, t1.clientY-t2.clientY); }
})();

document.getElementById('photoRotateBtn').addEventListener('click', ()=>{
  if(!cropState) return;
  cropState.rot = (cropState.rot + 90) % 360;
  drawCrop();
});
document.getElementById('photoCancelBtn').addEventListener('click', closeCropModal);
document.getElementById('photoRetakeBtn').addEventListener('click', ()=>{
  closeCropModal();
  openPhotoPicker(photoTarget.num, photoTarget.card);
});
document.getElementById('photoUseBtn').addEventListener('click', ()=>{
  const canvas = document.getElementById('cropCanvas');
  const dataUrl = canvas.toDataURL('image/jpeg', 0.78);
  const ans = currentAudit.answers[photoTarget.num];
  ans.photos = ans.photos || [];
  ans.photos.push({data:dataUrl, comment:''});
  renderPhotos(photoTarget.card, photoTarget.num);
  updateProgress();
  scheduleSave();
  closeCropModal();
});
function closeCropModal(){
  document.getElementById('photoModal').classList.add('hidden');
  cropState = null;
}

/* ============================================================
   FINALIZAR / EXPORTAR
   ============================================================ */
document.getElementById('finishBtn').addEventListener('click', async ()=>{
  const answers = Object.values(currentAudit.answers);
  const answered = answers.filter(a=>a.valoracion).length;
  if(answered < 76){
    if(!confirm(`Solo has respondido ${answered} de 76 preguntas. ¿Exportar igualmente?`)) return;
  }
  await saveCurrentAudit(false);
  await exportAudit();
});

async function exportAudit(){
  const overlay = document.getElementById('exportOverlay');
  const status = document.getElementById('exportStatus');
  overlay.classList.remove('hidden');
  try{
    status.textContent = 'Generando Excel…';
    await sleep(50);
    const xlsxBlob = await buildExcelBlob();

    status.textContent = 'Generando informe fotográfico…';
    await sleep(50);
    const pdfBlob = await buildPdfBlob();

    currentAudit.status = 'finalizado';
    await saveCurrentAudit(false);

    status.textContent = 'Descargando…';
    const baseName = sanitizeId(currentAudit.ceco)+'_'+sanitizeId(currentAudit.fecha);
    downloadBlob(xlsxBlob, 'ENCUESTA_SO_'+baseName+'.xlsx');
    await sleep(400);
    downloadBlob(pdfBlob, 'ANEXO_FOTOGRAFICO_'+baseName+'.pdf');

    overlay.classList.add('hidden');
    toast('Archivos exportados correctamente');
    renderHome();
    showScreen('screen-home');
  }catch(err){
    console.error(err);
    overlay.classList.add('hidden');
    alert('Ha ocurrido un error generando los archivos: '+err.message);
  }
}
function sleep(ms){ return new Promise(r=>setTimeout(r,ms)); }
function downloadBlob(blob, filename){
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url; a.download = filename;
  document.body.appendChild(a); a.click();
  setTimeout(()=>{ document.body.removeChild(a); URL.revokeObjectURL(url); }, 1000);
}

/* Escribe valores directamente en el XML interno del .xlsx original,
   sin recomponer el libro entero. Esto conserva el 100% del formato,
   colores, bordes y listas desplegables tal como los generó la
   plataforma de Repsol — solo se tocan las celdas F/G/H de cada
   pregunta. */
function xmlEscape(s){
  return String(s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;');
}
function setCellInline(xml, cellRef, text){
  const value = xmlEscape(text);
  const selfClosing = new RegExp('<c r="'+cellRef+'"([^>]*)/>');
  const withContent  = new RegExp('<c r="'+cellRef+'"([^>]*)>.*?</c>');
  if(selfClosing.test(xml)){
    return xml.replace(selfClosing, (m, attrs)=>{
      const cleanAttrs = attrs.replace(/\st="[^"]*"/,''); // por si ya tuviera un tipo
      return '<c r="'+cellRef+'"'+cleanAttrs+' t="inlineStr"><is><t xml:space="preserve">'+value+'</t></is></c>';
    });
  }
  if(withContent.test(xml)){
    return xml.replace(withContent, (m)=>{
      const sMatch = m.match(/ s="(\d+)"/);
      const sAttr = sMatch ? ' s="'+sMatch[1]+'"' : '';
      return '<c r="'+cellRef+'"'+sAttr+' t="inlineStr"><is><t xml:space="preserve">'+value+'</t></is></c>';
    });
  }
  console.warn('Celda no encontrada en el XML original:', cellRef);
  return xml;
}

async function buildExcelBlob(){
  const zip = await JSZip.loadAsync(currentAudit.originalXlsx);
  // localizar la hoja activa a partir de workbook.xml (normalmente sheet1.xml)
  let sheetPath = 'xl/worksheets/sheet1.xml';
  if(!zip.file(sheetPath)){
    const candidates = Object.keys(zip.files).filter(p=>/^xl\/worksheets\/sheet\d+\.xml$/.test(p));
    if(candidates.length) sheetPath = candidates[0];
  }
  let xml = await zip.file(sheetPath).async('string');

  QUESTIONS.forEach(q=>{
    const ans = currentAudit.answers[q.num] || {};
    const row = q.row;
    if(ans.valor)      xml = setCellInline(xml, 'F'+row, ans.valor);
    if(ans.valoracion) xml = setCellInline(xml, 'G'+row, valLabelForExcel(ans.valoracion));
    if(ans.obs)        xml = setCellInline(xml, 'H'+row, ans.obs);
  });

  zip.file(sheetPath, xml);
  const out = await zip.generateAsync({type:'blob', mimeType:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'});
  return out;
}
function valLabelForExcel(key){
  // El desplegable de la plataforma usa: SÍ, NO, N/A, PTE
  return {SI:'SÍ', NO:'NO', NA:'N/A', PTE:'PTE'}[key] || '';
}

async function buildPdfBlob(){
  const { jsPDF } = window.jspdf;
  const doc = new jsPDF({unit:'pt', format:'a4'});
  const pageW = doc.internal.pageSize.getWidth();
  const margin = 40;
  let y = margin;

  doc.setFont('helvetica','bold');
  doc.setFontSize(11);
  const headerRows = [
    ['D.A.', currentAudit.da || ''],
    ['CENTRO DE COSTE', currentAudit.ceco || ''],
    ['NOMBRE E.S.', currentAudit.ceco || ''],
    ['FECHA', currentAudit.fecha || ''],
    ['AUDITOR', currentAudit.auditor || ''],
  ];
  headerRows.forEach(([k,v])=>{
    doc.setFont('helvetica','bold'); doc.setFontSize(10);
    doc.rect(margin, y, 150, 20);
    doc.rect(margin+150, y, pageW-2*margin-150, 20);
    doc.text(k, margin+5, y+13);
    doc.setFont('helvetica','normal');
    doc.text(String(v), margin+155, y+13);
    y += 20;
  });
  y += 14;

  for(const q of QUESTIONS){
    const ans = currentAudit.answers[q.num] || {};
    const valWord = ans.valoracion ? VAL_OPTIONS.find(v=>v.key===ans.valoracion).pdfWord : 'SIN RESPONDER';
    const photos = ans.photos || [];

    // estimar altura necesaria del bloque
    const textLines = doc.setFont('helvetica','bold').setFontSize(10).splitTextToSize(q.text, pageW-2*margin-10);
    let blockH = 22 + (textLines.length*13) + 6;
    if(ans.obs) blockH += 14;
    if(photos.length) blockH += 145;

    if(y + blockH > doc.internal.pageSize.getHeight() - margin){
      doc.addPage(); y = margin;
    }

    doc.setFont('helvetica','bold'); doc.setFontSize(10);
    doc.rect(margin, y, pageW-2*margin-90, 20);
    doc.rect(pageW-margin-90, y, 90, 20);
    doc.text('PREGUNTA '+q.num, margin+5, y+13);
    doc.text(valWord, pageW-margin-85, y+13);
    y += 20;

    doc.setFont('helvetica','bold'); doc.setFontSize(9.5);
    doc.rect(margin, y, pageW-2*margin, textLines.length*13+8);
    doc.text(textLines, margin+5, y+12);
    y += textLines.length*13 + 8;

    if(ans.obs){
      doc.setFont('helvetica','normal'); doc.setFontSize(9);
      const obsLines = doc.splitTextToSize('Observaciones: '+ans.obs, pageW-2*margin-10);
      doc.text(obsLines, margin+5, y+10);
      y += obsLines.length*11 + 6;
    }
    if(ans.valor){
      doc.setFont('helvetica','normal'); doc.setFontSize(9);
      doc.text('Valor/Fecha: '+ans.valor, margin+5, y+10);
      y += 14;
    }

    if(photos.length){
      let x = margin;
      const imgW = Math.min(150, (pageW-2*margin-10*(photos.length-1))/Math.min(photos.length,3));
      const imgH = imgW*0.75;
      photos.forEach((p, i)=>{
        if(i>0 && i%3===0){ y += imgH+8; x = margin; }
        if(y + imgH > doc.internal.pageSize.getHeight()-margin){ doc.addPage(); y=margin; x=margin; }
        try{ doc.addImage(p.data, 'JPEG', x, y, imgW, imgH); }catch(e){}
        x += imgW+10;
      });
      y += imgH + 10;
    }
    y += 8;
  }

  return doc.output('blob');
}

/* ============================================================
   ARRANQUE
   ============================================================ */
function checkLibsLoaded(){
  const missing = [];
  if(typeof XLSX === 'undefined') missing.push('xlsx.full.min.js');
  if(typeof JSZip === 'undefined') missing.push('jszip.min.js');
  if(typeof window.jspdf === 'undefined') missing.push('jspdf.umd.min.js');
  return missing;
}

(async function init(){
  const missingLibs = checkLibsLoaded();
  if(missingLibs.length){
    alert(
      'La app no ha podido cargar estos archivos:\n\n' +
      missingLibs.join('\n') +
      '\n\nEsto casi siempre significa que esos archivos .js no se subieron a GitHub, o se subieron con otro nombre. ' +
      'Revisa en tu repositorio que estén sueltos en la raíz, junto a index.html.'
    );
    return; // no seguimos inicializando si faltan librerías
  }
  db = await openDB();
  if('serviceWorker' in navigator){
    navigator.serviceWorker.register('sw.js').catch(()=>{});
  }
  renderHome();
  showScreen('screen-home');
})();
