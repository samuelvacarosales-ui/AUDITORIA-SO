/* ============================================================
   app.js — interfaz de la app "Encuesta SO"
   ============================================================ */
const $ = id => document.getElementById(id);
let db = null, cur = null, screen = 'home', filter = 'all', saveTimer = null, pdfBlob = null;
let photoTarget = null;

/* ---------------- almacenamiento (IndexedDB) ---------------- */
function openDB(){
  return new Promise((res,rej)=>{
    const rq = indexedDB.open('anexoSO',1);
    rq.onupgradeneeded = e=>{ const d=e.target.result; if(!d.objectStoreNames.contains('audits')) d.createObjectStore('audits',{keyPath:'id'}); };
    rq.onsuccess = e=>res(e.target.result);
    rq.onerror = e=>rej(e.target.error);
  });
}
const dbPut = a => new Promise((res,rej)=>{ const tx=db.transaction('audits','readwrite'); tx.objectStore('audits').put(a); tx.oncomplete=res; tx.onerror=()=>rej(tx.error); tx.onabort=()=>rej(tx.error); });
const dbAll = () => new Promise((res,rej)=>{ const rq=db.transaction('audits').objectStore('audits').getAll(); rq.onsuccess=()=>res(rq.result); rq.onerror=()=>rej(rq.error); });
const dbDel = id => new Promise((res,rej)=>{ const tx=db.transaction('audits','readwrite'); tx.objectStore('audits').delete(id); tx.oncomplete=res; tx.onerror=()=>rej(tx.error); });

function getSettings(){
  let s = {};
  try{ s = JSON.parse(localStorage.getItem('anexoSO_settings'))||{}; }catch(e){}
  return {auditor: s.auditor || 'SAMUEL VACA', apiKey: s.apiKey || '', geminiKey: s.geminiKey || ''};
}
function toast(msg, ms){
  const t=$('toast'); t.textContent=msg; t.classList.add('show');
  clearTimeout(toast._t); toast._t=setTimeout(()=>t.classList.remove('show'), ms||2400);
}
function busy(msg){ $('busyMsg').textContent=msg||'Procesando…'; $('busy').classList.remove('hidden'); }
function unbusy(){ $('busy').classList.add('hidden'); $('busyCancel').classList.add('hidden'); }
/* petición con tiempo máximo y botón Cancelar (para que nunca se quede colgada) */
let aiCtl=null, aiCancelled=false;
function showCancel(){ aiCancelled=false; $('busyCancel').classList.remove('hidden'); $('busyCancel').onclick=()=>{ aiCancelled=true; if(aiCtl) aiCtl.abort(); unbusy(); }; }
async function fetchT(url, opts, ms){
  aiCtl = new AbortController(); const t=setTimeout(()=>aiCtl.abort(), ms);
  try{ return await fetch(url, Object.assign({signal:aiCtl.signal}, opts)); }
  catch(e){ if(aiCancelled){ const c=new Error('cancelado'); c.cancelled=true; throw c; } const x=new Error(e&&e.name==='AbortError'?'timeout':'red'); x.kind=x.message; throw x; }
  finally{ clearTimeout(t); }
}
const nextFrame = () => new Promise(r=>requestAnimationFrame(()=>r()));

function scheduleSave(){ clearTimeout(saveTimer); saveTimer=setTimeout(saveNow,600); }
async function saveNow(){
  if(!cur) return;
  clearTimeout(saveTimer);
  cur.updatedAt = Date.now();
  try{ await dbPut(cur); }
  catch(e){ console.error(e); toast('⚠ No se pudo guardar (¿almacenamiento lleno?)',4000); }
}

/* ---------------- navegación ---------------- */
function showScreen(name){
  screen = name;
  ['home','audit','preview','settings'].forEach(n=>$('s-'+n).classList.toggle('hidden', n!==name));
  $('bar-audit').classList.toggle('hidden', name!=='audit');
  $('bar-preview').classList.toggle('hidden', name!=='preview');
  $('backBtn').classList.toggle('hidden', name==='home');
  $('settingsBtn').classList.toggle('hidden', name!=='home');
  let t='Encuesta SO', s='Repsol · Excel y anexo fotográfico · v18';
  if(name==='audit' && cur){ t=cur.name||'Auditoría'; s='CECO '+cur.ceco+' · '+cur.fecha; }
  if(name==='preview' && cur){ t='Vista previa'; s=(cur.name||'')+' · CECO '+cur.ceco; }
  if(name==='settings'){ t='Ajustes'; s='Se guardan en este dispositivo'; }
  $('topTitle').textContent=t; $('topSub').textContent=s;
  $('s-'+name).scrollTop = 0;
}
$('backBtn').onclick = async ()=>{
  if(screen==='preview'){ showScreen('audit'); return; }
  if(screen==='audit'){ await saveNow(); cur=null; await renderHome(); }
  showScreen('home');
};
$('settingsBtn').onclick = ()=>{
  const s=getSettings(); $('setAuditor').value=s.auditor; $('setKey').value=s.apiKey; $('setGem').value=s.geminiKey; showScreen('settings');
};
$('saveSet').onclick = ()=>{
  localStorage.setItem('anexoSO_settings', JSON.stringify({auditor:$('setAuditor').value.trim()||'SAMUEL VACA', apiKey:$('setKey').value.trim(), geminiKey:$('setGem').value.trim()}));
  toast('Ajustes guardados'); showScreen('home');
};

/* ---------------- inicio ---------------- */
async function renderHome(){
  const all = (await dbAll()).sort((a,b)=>b.updatedAt-a.updatedAt);
  const on=$('ongoing'), fin=$('finished'); on.innerHTML=''; fin.innerHTML='';
  const ongoing = all.filter(a=>a.status!=='finalizado'), done = all.filter(a=>a.status==='finalizado');
  if(!ongoing.length) on.innerHTML='<div class="empty">No tienes encuestas en curso.<br>Carga el Excel limpio de una estación para empezar.</div>';
  ongoing.forEach(a=>on.appendChild(auditCard(a)));
  $('finTitle').classList.toggle('hidden', !done.length);
  done.forEach(a=>fin.appendChild(auditCard(a)));
}
function auditCard(a){
  const st = auditStats(a);
  const d = document.createElement('div'); d.className='audit-card';
  d.innerHTML = '<div class="info"><div class="nm">'+esc(a.name)+' · '+esc(a.ceco)+'</div>'+
    '<div class="meta">'+esc(a.fecha)+' · '+st.photos+' fotos'+(a.exportedXlsx?' · Excel ✔':'')+(a.exportedPdf?' · PDF ✔':'')+'</div>'+
    '<div class="meta"><span class="pill" style="background:var(--bad)">'+st.NO+' no conformes</span>'+
    (st.NO? '<span class="pill" style="background:'+(st.ncSinFoto?'var(--warn)':'var(--ok)')+'">'+(st.ncSinFoto? st.ncSinFoto+' sin foto':'todas con foto')+'</span>':'')+'</div></div>'+
    '<div class="del">🗑</div>';
  d.querySelector('.info').onclick = ()=>openAudit(a);
  d.querySelector('.del').onclick = async ev=>{
    ev.stopPropagation();
    if(confirm('¿Eliminar este informe y sus fotos de este dispositivo? No se puede deshacer.')){ await dbDel(a.id); renderHome(); }
  };
  return d;
}


/* ---------------- leer foto de inventario con Claude ---------------- */
const INV_MODEL = 'claude-sonnet-5-5';
const INV_PROMPT = `Esta foto es una nota escrita a mano por un auditor de gasolineras con cifras de inventario. Extrae SOLO las cantidades.
Estructura habitual: a la izquierda, líneas "TABACO A <cifra>", "CAFÉ A <cifra>", "ONCE A <cifra>", "BUTANO A <cifra>" (A = ajuste físico, puede ser negativo, p. ej. -6 o -18,79 y también 0). A la derecha, líneas que empiezan por "V" (ventas) en el MISMO orden vertical: la 1ª V es de tabaco, la 2ª de café, la 3ª de once, la 4ª de butano. Debajo puede haber "F.FIJO" (fondo fijo) y "F.MANIOBRA" (fondo de maniobra) con su importe.
Reglas: ignora texto tenue o al revés que se transparente del otro lado del papel y los garabatos o recuadros dibujados; si un número está tachado o escrito sobre otro, toma el que se ve más oscuro. Devuelve los importes en formato español (punto para miles, coma para decimales, signo - si es negativo), sin el símbolo €. Si algo no aparece o no se lee, deja "" y explícalo en "dudas". Si algún dígito es dudoso, dilo en "dudas".
Responde ÚNICAMENTE con JSON válido, sin texto antes ni después, con esta forma exacta:
{"tabaco":{"a":"","v":""},"cafe":{"a":"","v":""},"once":{"a":"","v":""},"butano":{"a":"","v":""},"ffijo":"","fmaniobra":"","dudas":""}`;

async function shrinkToJpeg(file, max){
  const bmp = await createImageBitmap(file, {imageOrientation:'from-image'});
  const k = Math.min(1, max/Math.max(bmp.width,bmp.height));
  const cv = document.createElement('canvas'); cv.width=Math.round(bmp.width*k); cv.height=Math.round(bmp.height*k);
  cv.getContext('2d').drawImage(bmp,0,0,cv.width,cv.height);
  return cv.toDataURL('image/jpeg',0.88).split(',')[1];
}
function aiMode(){ const s=getSettings(); return s.geminiKey ? 'gemini' : s.apiKey ? 'claude' : ''; }
async function askGeminiInventory(file){
  const key = getSettings().geminiKey;
  const b64 = await shrinkToJpeg(file, 1800);
  const body = JSON.stringify({contents:[{parts:[{inline_data:{mime_type:'image/jpeg',data:b64}},{text:INV_PROMPT}]}],
    generationConfig:{temperature:0, responseMimeType:'application/json'}});
  let resp=null, lastStatus=0, lastKind='';
  const tries=['gemini-flash-latest','gemini-flash-latest','gemini-2.5-flash','gemini-flash-lite-latest'];
  for(let i=0;i<tries.length;i++){
    if(i>0) busy('Reintentando la lectura ('+i+'/'+(tries.length-1)+')…'), showCancel();
    try{
      resp = await fetchT('https://generativelanguage.googleapis.com/v1beta/models/'+tries[i]+':generateContent',
        {method:'POST', headers:{'content-type':'application/json','x-goog-api-key':key}, body}, 15000);
    }catch(e){ if(e.cancelled) throw e; lastKind=e.kind; resp=null; continue; }   // sin respuesta a tiempo: reintenta
    lastStatus = resp.status;
    if(resp.status===404 || resp.status===500 || resp.status===503) continue;   // modelo retirado o saturado: reintenta
    break;
  }
  if(!resp) throw new Error(lastKind==='timeout' ? 'Google no responde (la conexión va lenta). Inténtalo de nuevo en un momento.' : 'No hay conexión a internet. Prueba otra vez con cobertura.');
  if(lastStatus===400 || lastStatus===403){
    let m=''; try{ m=(await resp.json()).error.message; }catch(e){}
    throw new Error('Google ha rechazado la petición'+(m?': '+m:'')+'. Revisa que la clave de Gemini está bien copiada en Ajustes.');
  }
  if(lastStatus===429) throw new Error('Has llegado al límite gratuito de Google (por minuto o por día). Espera un poco y repite.');
  if(!resp.ok) throw new Error('Error del servicio de Google ('+lastStatus+'). Inténtalo de nuevo en un momento.');
  const data = await resp.json();
  const txt = ((((data.candidates||[])[0]||{}).content||{}).parts||[]).map(p=>p.text||'').join('');
  const m = txt.match(/\{[\s\S]*\}/);
  if(!m) throw new Error('No he podido entender la respuesta. Haz la foto más de cerca y con buena luz.');
  try{ return JSON.parse(m[0]); }catch(e){ throw new Error('La respuesta llegó incompleta. Inténtalo de nuevo.'); }
}
function askAIInventory(file){ return aiMode()==='gemini' ? askGeminiInventory(file) : askClaudeInventory(file); }
async function askClaudeInventory(file){
  const key = getSettings().apiKey;
  const b64 = await shrinkToJpeg(file, 1800);
  let resp;
  try{
    resp = await fetchT('https://api.anthropic.com/v1/messages',{
      method:'POST',
      headers:{'content-type':'application/json','x-api-key':key,'anthropic-version':'2023-06-01','anthropic-dangerous-direct-browser-access':'true'},
      body: JSON.stringify({model:INV_MODEL, max_tokens:700, messages:[{role:'user',content:[
        {type:'image',source:{type:'base64',media_type:'image/jpeg',data:b64}},
        {type:'text',text:INV_PROMPT}]}]})
    }, 30000);
  }catch(e){ if(e.cancelled) throw e; throw new Error(e.kind==='timeout' ? 'El servicio no responde (la conexión va lenta). Inténtalo de nuevo.' : 'No hay conexión a internet. Prueba otra vez con cobertura.'); }
  if(resp.status===401) throw new Error('La clave de Anthropic no es válida. Revísala en Ajustes.');
  if(resp.status===429) throw new Error('Demasiadas peticiones seguidas. Espera un minuto y repite.');
  if(resp.status===402 || resp.status===400){
    let m=''; try{ m=(await resp.json()).error.message; }catch(e){}
    throw new Error('Anthropic ha rechazado la petición'+(m?': '+m:'')+'. Comprueba que la cuenta tiene saldo.');
  }
  if(!resp.ok) throw new Error('Error del servicio ('+resp.status+'). Inténtalo de nuevo en un momento.');
  const data = await resp.json();
  const txt = (data.content||[]).map(b=>b.text||'').join('');
  const m = txt.match(/\{[\s\S]*\}/);
  if(!m) throw new Error('No he podido entender la respuesta. Haz la foto más de cerca y con buena luz.');
  try{ return JSON.parse(m[0]); }catch(e){ throw new Error('La respuesta llegó incompleta. Inténtalo de nuevo.'); }
}
let invReading = null;
function showInvConfirm(r){
  invReading = r;
  const rows = [['tabaco','Tabaco (14)'],['once','Once (15)'],['butano','Butano (16)'],['cafe','Café (17)']];
  let html = '<div class="invgrid"><span></span><span class="h">Ajuste</span><span class="h">Ventas</span>';
  rows.forEach(([k,l])=>{
    const e=r[k]||{};
    html += '<b>'+l+'</b><input data-k="'+k+'" data-w="a" value="'+esc(e.a||'')+'" autocomplete="off"><input data-k="'+k+'" data-w="v" value="'+esc(e.v||'')+'" autocomplete="off">';
  });
  html += '<b>F. fijo (19)</b><input class="span2" data-k="ffijo" value="'+esc(r.ffijo||'')+'" autocomplete="off">'+
          '<b>F. maniobra</b><input class="span2" data-k="fmaniobra" value="'+esc(r.fmaniobra||'')+'" autocomplete="off"></div>';
  $('invRows').innerHTML = html;
  const inf=$('invInfo'); inf.textContent=r.info||''; inf.classList.toggle('hidden', !r.info);
  const d=$('invDoubts'); d.textContent = r.dudas ? ('⚠ '+r.dudas) : ''; d.classList.toggle('hidden', !r.dudas);
  $('invSheet').classList.remove('hidden');
}
$('fileInv').onchange = async e=>{
  const f=e.target.files[0]; e.target.value=''; if(!f) return;
  $('scanSheet').classList.add('hidden');
  try{
    let r;
    if(aiMode()){ busy('Leyendo la foto con IA…'); showCancel(); r = await askAIInventory(f); }
    else r = await ocrInventory(await createImageBitmap(f,{imageOrientation:'from-image'}));
    unbusy(); showInvConfirm(r);
  }catch(err){ unbusy(); if(err&&err.cancelled) return; console.error(err); alert(err.message||err); }
};
/* ---- lector gratuito (Tesseract, dentro del móvil) ---- */
let tessWorker=null;
function prepForOcr(bmp){
  const k=Math.min(1,2000/Math.max(bmp.width,bmp.height)), w=Math.round(bmp.width*k), h=Math.round(bmp.height*k);
  const cv=document.createElement('canvas'); cv.width=w; cv.height=h; const cx=cv.getContext('2d',{willReadFrequently:true});
  cx.drawImage(bmp,0,0,w,h);
  const img=cx.getImageData(0,0,w,h), d=img.data;
  const sw=Math.max(8,Math.round(w/40)), sh=Math.max(8,Math.round(h/40));            // fondo = imagen muy reducida
  const sm=document.createElement('canvas'); sm.width=sw; sm.height=sh; const sx=sm.getContext('2d',{willReadFrequently:true});
  sx.imageSmoothingEnabled=true; sx.drawImage(cv,0,0,sw,sh);
  const bg=document.createElement('canvas'); bg.width=w; bg.height=h; const bx=bg.getContext('2d',{willReadFrequently:true});
  bx.imageSmoothingEnabled=true; bx.drawImage(sm,0,0,w,h);
  const b=bx.getImageData(0,0,w,h).data;
  for(let i=0;i<d.length;i+=4){
    const g=0.3*d[i]+0.59*d[i+1]+0.11*d[i+2], gb=0.3*b[i]+0.59*b[i+1]+0.11*b[i+2];
    const v=(gb-g)>22?0:255; d[i]=d[i+1]=d[i+2]=v; d[i+3]=255;
  }
  cx.putImageData(img,0,0); return cv;
}
function numFromWord(t){
  t=String(t||'').trim(); if(!/\d/.test(t)) return '';
  t=t.replace(/[Oo]/g,'0').replace(/[lI|]/g,'1').replace(/[−–—]/g,'-');
  t=t.replace(/[^0-9.,-]/g,'');
  return /\d/.test(t) ? t.replace(/[.,]+$/,'') : '';
}
async function ocrInventory(bmp){
  if(typeof Tesseract==='undefined') throw new Error('Falta el lector gratuito en la app. Sube a GitHub también tesseract.min.js, worker.min.js, tesseract-core-simd-lstm.wasm.js y eng.traineddata.gz.');
  busy('Preparando el lector (la primera vez tarda unos segundos)…');
  const abs=u=>new URL(u,location.href).href;
  if(!tessWorker){
    tessWorker = await Tesseract.createWorker('eng',1,{workerPath:abs('worker.min.js'),corePath:abs('tesseract-core-simd-lstm.wasm.js'),langPath:abs('.'),gzip:true});
    await tessWorker.setParameters({tessedit_pageseg_mode:'11', preserve_interword_spaces:'1'});
  }
  busy('Leyendo la foto…');
  const cv=prepForOcr(bmp);
  const res=await tessWorker.recognize(cv,{},{text:true,blocks:true});
  const data=res.data;
  const r=parseInventoryText(data.text||'');
  /* respaldo por posición: cifras sueltas de la columna izquierda (ajustes) y derecha (ventas), de arriba abajo */
  const words=[]; (data.blocks||[]).forEach(b=>(b.paragraphs||[]).forEach(p=>(p.lines||[]).forEach(l=>(l.words||[]).forEach(w=>words.push(w)))));
  const toks=words.map(w=>({n:numFromWord(w.text),x:(w.bbox.x0+w.bbox.x1)/2,y:(w.bbox.y0+w.bbox.y1)/2})).filter(t=>t.n && t.n!=='-');
  const L=toks.filter(t=>t.x<cv.width*0.5).sort((p,q)=>p.y-q.y), R=toks.filter(t=>t.x>=cv.width*0.5).sort((p,q)=>p.y-q.y);
  ['tabaco','cafe','once','butano'].forEach((k,i)=>{ if(!r[k].a && L[i]) r[k].a=L[i].n; if(!r[k].v && R[i]) r[k].v=R[i].n; });
  if(!r.ffijo && L[4]) r.ffijo=L[4].n; if(!r.fmaniobra && R[4]) r.fmaniobra=R[4].n;
  r.dudas='Lectura automática gratuita: con letra a mano falla a menudo. Comprueba cada cifra con tu nota.';
  return r;
}
$('fileRep').onchange = async e=>{
  const fs=[...e.target.files]; e.target.value=''; if(!fs.length) return;
  $('scanSheet').classList.add('hidden');
  try{
    busy('Leyendo los Excel…');
    const files=[]; for(const f of fs) files.push({name:f.name, ab:await f.arrayBuffer()});
    const r=parseInventoryReports(files);
    unbusy(); showInvConfirm(r);
  }catch(err){ unbusy(); console.error(err); alert('No he podido leer los archivos: '+(err.message||err)); }
};
$('scanCancel').onclick = ()=>$('scanSheet').classList.add('hidden');
$('scanBank').onclick = ()=>{
  if(!bank().length){ toast('Esta estación aún no tiene fotos guardadas', 3000); return; }
  $('scanSheet').classList.add('hidden'); openBank('ocr', null);
};
$('scanGo').onclick = ()=>{
  const txt=$('scanTxt').value;
  if(!txt.trim()){ toast('Primero escanea o escribe el texto de la nota', 3200); return; }
  $('scanSheet').classList.add('hidden'); showInvConfirm(parseInventoryText(txt));
};
$('invCancel').onclick = ()=>$('invSheet').classList.add('hidden');
$('invApply').onclick = ()=>{
  const r={tabaco:{},cafe:{},once:{},butano:{}};
  $('invRows').querySelectorAll('input').forEach(i=>{
    const k=i.dataset.k, w=i.dataset.w;
    if(w) r[k][w]=i.value; else r[k]=i.value;
  });
  r.nodispone = (invReading && invReading.nodispone || []).filter(k=>!(r[k].a||r[k].v));   // solo si has dejado sus casillas vacías
  const warn = applyInvReading(cur, r);
  $('invSheet').classList.add('hidden');
  renderList(); refreshTop(); scheduleSave();
  toast(warn.length ? ('⚠ Las preguntas '+warn.join(', ')+' no coinciden con la plantilla y no se han tocado') : 'Inventario y fondos rellenados (14-17 y 19)', 3800);
};

/* ---------------- importar Excel ---------------- */
async function handleExcel(file){
  if(!file) return;
  try{
    busy('Leyendo Excel…');
    const buf = await file.arrayBuffer();
    const parsed = parseWorkbook(buf);
    const id = sanitizeId(parsed.meta.ceco)+'_'+sanitizeId(parsed.meta.fecha);
    const existing = (await dbAll()).find(a=>a.id===id);
    unbusy();
    if(existing){
      if(!confirm('Ya tienes una encuesta de '+parsed.meta.name+' ('+parsed.meta.fecha+').\n\n¿Actualizar sus respuestas con este Excel? Se mantienen tus fotos y lo que hayas rellenado en la app.')) return;
      cur = mergeAudit(existing, parsed, file.name, buf);
    }else{
      cur = newAudit(parsed, file.name, getSettings(), buf);
    }
    const warn = applyDefaults(cur.questions);
    await saveNow();
    filter='all';
    renderAudit(); showScreen('audit');
    toast(warn.length ? ('⚠ Las preguntas '+warn.join(', ')+' ya no coinciden con las reglas por defecto: revísalas')
                      : 'Excel cargado: todo en SÍ, salvo N/A por defecto en 13, 37 y 53', 4200);
  }catch(err){
    unbusy(); console.error(err);
    alert('No se ha podido leer este Excel.\n\n'+(err&&err.message?err.message:err));
  }
}
$('fileNew').onchange = e=>{ const f=e.target.files[0]; e.target.value=''; handleExcel(f); };
$('fileReload').onchange = e=>{ const f=e.target.files[0]; e.target.value=''; handleExcel(f); };

async function openAudit(a){ cur=a; if(migrateInv(cur.questions)) await saveNow(); filter='all'; renderAudit(); showScreen('audit'); }

/* ---------------- pantalla de auditoría ---------------- */
function renderAudit(){
  renderHeaderCard(); renderStats(); renderFilters(); renderList();
}
function renderHeaderCard(){
  const a=cur, nb=bank().length;
  $('hcard').innerHTML =
   '<div class="hgrid">'+
   '<div class="fld"><label>Centro de coste</label><div class="ro">'+esc(a.ceco)+'</div></div>'+
   '<div class="fld"><label>Fecha</label><div class="ro">'+esc(a.fecha)+'</div></div>'+
   '<div class="fld full"><label>Nombre E.S.</label><div class="ro">'+esc(a.name)+'</div></div>'+
   '<div class="fld full"><label>Auditor</label><input id="fAud" value="'+esc(a.auditor||'')+'"></div></div>'+
   '<label class="swrow"><span class="swtxt"><b>Esta estación NO dispone de servicios</b><small>Pone N/A + «NO DISPONE» en las preguntas 6, 8, 48, 56, 65 y 71</small></span>'+
   '<span class="sw"><input type="checkbox" id="sinServ"><i></i></span></label>'+
   (nb? '' : '<div class="hint">Elige de una vez las fotos de esta estación y luego solo las marcas en cada pregunta.</div>')+
   '<div class="hbtns"><label class="linkbtn main" id="invBtn">🧾 Leer foto de inventario</label></div>'+
   '<div class="hbtns"><div class="linkbtn'+(nb?'':' main')+'" id="bankBtn">📷 Fotos de la estación ('+nb+')</div>'+
   '<label class="linkbtn" for="fileReload">↻ Recargar Excel</label></div>';
  $('sinServ').checked = !!a.sinServicios;
  $('sinServ').onchange = e=>{
    const on = e.target.checked, warn = setSinServicios(cur, on);
    renderList(); refreshTop(); scheduleSave();
    toast(warn.length ? ('⚠ Las preguntas '+warn.join(', ')+' no coinciden con la plantilla y no se han tocado')
         : on ? 'N/A + NO DISPONE en las preguntas 6, 8, 48, 56, 65 y 71' : 'Servicios restaurados a SÍ', 3200);
  };
  $('invBtn').onclick = ()=>{
    $('scanTxt').value=''; $('scanMode').textContent = aiMode()==='gemini' ? '✨ Lectura con IA gratuita de Google (Gemini)' : aiMode()==='claude' ? '✨ Lectura con IA de Claude' : 'Lectura gratuita sin IA (poco fiable con letra a mano). Para leer bien, pon tu clave gratuita de Gemini en Ajustes.';
    $('scanSheet').classList.remove('hidden');
  };
  $('bankBtn').onclick = ()=>openBank('manage');
  $('fAud').oninput = e=>{ cur.auditor=e.target.value; scheduleSave(); };
}
function renderStats(){
  const st=auditStats(cur);
  const box=(n,l,c)=>'<div class="stat"><b style="color:'+c+'">'+n+'</b><span>'+l+'</span></div>';
  $('stats').innerHTML = box(st.SI,'Correcto','var(--ok)')+box(st.NO,'No conf.','var(--bad)')+box(st.NA,'NA','#000')+
    (st.PTE?box(st.PTE,'PTE','#000'):'')+(st['']?box(st[''],'Sin valorar','var(--warn)'):'')+box(st.photos,'Fotos','var(--orange)');
}
function renderFilters(){
  const st=auditStats(cur);
  const defs=[['all','Todas ('+cur.questions.length+')'],['sv','Sin valorar ('+st['']+')'],['nc','No conformes ('+st.NO+')'],['ncs','NC sin foto ('+st.ncSinFoto+')'],['ph','Con fotos']];
  $('filters').innerHTML='';
  defs.forEach(([k,l])=>{
    const b=document.createElement('button'); b.className='chip'+(filter===k?' on':''); b.textContent=l;
    b.onclick=()=>{ filter=k; renderFilters(); renderList(); };
    $('filters').appendChild(b);
  });
}
function refreshTop(){ renderStats(); renderFilters(); }
function matches(q){
  const n=(cur.photos[q.num]||[]).length;
  if(filter==='sv') return !q.status;
  if(filter==='nc') return q.status==='NO';
  if(filter==='ncs') return q.status==='NO' && n===0;
  if(filter==='ph') return n>0;
  return true;
}
function renderList(){
  const root=$('list'); root.innerHTML='';
  if(filter==='all' || filter==='ph'){
    const c=document.createElement('div'); c.className='area';
    c.appendChild(photoCard(0)); root.appendChild(c);
  }
  let lastArea=null, lastAp=null, areaEl=null, shown=0;
  cur.questions.forEach(q=>{
    if(!matches(q)) return;
    shown++;
    if(q.area!==lastArea){
      areaEl=document.createElement('div'); areaEl.className='area';
      areaEl.innerHTML='<div class="area-h">'+esc(q.area)+' · '+esc(q.areaName)+'</div>';
      root.appendChild(areaEl); lastArea=q.area; lastAp=null;
    }
    if(q.apartado!==lastAp && q.apartado && q.apartado!==q.areaName){
      const t=document.createElement('div'); t.className='ap-t'; t.textContent=q.apartado; areaEl.appendChild(t);
    }
    lastAp=q.apartado;
    areaEl.appendChild(questionCard(q));
  });
  if(!shown) root.insertAdjacentHTML('beforeend','<div class="empty">No hay preguntas en este filtro.</div>');
}

const VALS=[['SI','SÍ'],['NO','NO'],['NA','N/A'],['PTE','PTE']];
function questionCard(q){
  const el=document.createElement('div'); el.className='q s-'+(q.status||'none'); el.dataset.num=q.num;
  el.innerHTML =
    '<div class="q-head"><span class="q-num">PREGUNTA '+q.num+'</span><span class="q-badge">'+STATUS[q.status].word+'</span></div>'+
    '<div class="q-text">'+esc(q.text)+'</div>'+
    '<div class="vrow">'+VALS.map(([k,l])=>'<button type="button" class="vb'+(q.status===k?' on v-'+k:'')+'" data-v="'+k+'">'+l+'</button>').join('')+'</div>'+
    (q.inv
      ? '<div class="frow"><div class="ff"><label>Ajuste físico (€)</label><div class="sgn"><input type="text" inputmode="decimal" maxlength="14" autocomplete="off" data-f="ajuste"><button type="button" class="sgnb" aria-label="Cambiar signo">±</button></div></div>'+
        '<div class="ff"><label>Ventas (€)</label><input type="text" inputmode="decimal" maxlength="14" autocomplete="off" data-f="ventas"></div></div>'+
        '<div class="invres">Valor: <b data-r="valor"></b> · <span data-r="obs"></span></div>'
      : '<div class="frow"><div class="ff fv"><label>Valor / Fecha</label><input type="text" maxlength="10" autocomplete="off" autocorrect="off" enterkeyhint="done" data-f="valor"></div>'+
        '<div class="ff fo"><label>Observaciones</label><textarea rows="2" maxlength="100" autocomplete="off" enterkeyhint="done" data-f="obs"></textarea></div></div>')+
    '<div class="photos"></div>';
  if(q.inv){
    const upd=()=>{ el.querySelector('[data-r="valor"]').textContent=q.valor; el.querySelector('[data-r="obs"]').textContent=q.obs; };
    ['ajuste','ventas'].forEach(f=>{
      const i=el.querySelector('[data-f="'+f+'"]'); i.value=q[f]||'';
      i.oninput=e=>{
        q[f]=e.target.value; const before=q.status; recalcInv(q); upd();
        if(q.status!==before){
          el.className='q s-'+(q.status||'none'); el.querySelector('.q-badge').textContent=STATUS[q.status].word;
          el.querySelectorAll('.vb').forEach(x=>{ x.className='vb'+(q.status===x.dataset.v?' on v-'+x.dataset.v:''); });
          refreshTop();
        }
        scheduleSave();
      };
    });
    el.querySelector('.sgnb').onclick=()=>{
      const i=el.querySelector('[data-f="ajuste"]'); let v=i.value.trim();
      v = v.startsWith('-') ? v.slice(1) : (v ? '-'+v : '-');
      i.value=v; i.oninput({target:i}); i.focus();
    };
    upd();
  }else{
    el.querySelector('[data-f="valor"]').value = q.valor||'';
    el.querySelector('[data-f="obs"]').value = q.obs||'';
  }
  el.querySelectorAll('.vb').forEach(b=>{
    b.onclick=()=>{
      q.status = (q.status===b.dataset.v) ? '' : b.dataset.v;           // tocar otra vez = quitar
      el.className='q s-'+(q.status||'none');
      el.querySelector('.q-badge').textContent=STATUS[q.status].word;
      el.querySelectorAll('.vb').forEach(x=>{ x.className='vb'+(q.status===x.dataset.v?' on v-'+x.dataset.v:''); });
      refreshTop(); scheduleSave();
    };
  });
  if(!q.inv){
    el.querySelector('[data-f="valor"]').oninput = e=>{ q.valor=e.target.value; scheduleSave(); };
    el.querySelector('[data-f="obs"]').oninput = e=>{
      if(/[\r\n]/.test(e.target.value)) e.target.value=e.target.value.replace(/[\r\n]+/g,' ');
      q.obs=e.target.value; scheduleSave();
    };
  }
  fillPhotos(el.querySelector('.photos'), q.num);
  return el;
}
function photoCard(num){
  const el=document.createElement('div'); el.className='q cartel'; el.dataset.num=num;
  el.innerHTML='<div class="q-head"><span class="q-num">CARTEL IDENTIFICACIÓN ES</span></div>'+
    '<div class="q-text" style="color:var(--soft);font-weight:600;font-size:13px">Foto del cartel de la estación (sale al inicio del informe)</div><div class="photos"></div>';
  fillPhotos(el.querySelector('.photos'), num);
  return el;
}
function fillPhotos(row, num){
  row.innerHTML='';
  (cur.photos[num]||[]).forEach((p,i)=>{
    const t=document.createElement('div'); t.className='th';
    t.innerHTML='<img src="'+(p.thumb||p.data)+'"><div class="sz">'+(p.size||'M')+'</div>';
    t.onclick=()=>editPhoto(num,i);
    row.appendChild(t);
  });
  const add=document.createElement('div'); add.className='add'; add.textContent='+';
  add.onclick=()=>{ photoTarget=num; if(bank().length) openAddSheet(num); else $('filePhoto').click(); };
  row.appendChild(add);
}
function refreshCard(num){
  const el=document.querySelector('[data-num="'+num+'"]');
  if(el) fillPhotos(el.querySelector('.photos'), num);
  refreshTop();
}

/* ---------------- añadir / editar fotos ---------------- */
function loadImage(src){
  return new Promise((res,rej)=>{ const i=new Image(); i.onload=()=>res(i); i.onerror=()=>rej(new Error('No se pudo abrir la imagen')); i.src=src; });
}
function scaledDataURL(img, maxSide, q){
  const w=img.naturalWidth||img.width, h=img.naturalHeight||img.height;
  const s=Math.min(1, maxSide/Math.max(w,h));
  const c=document.createElement('canvas'); c.width=Math.round(w*s); c.height=Math.round(h*s);
  c.getContext('2d').drawImage(img,0,0,c.width,c.height);
  return c.toDataURL('image/jpeg',q);
}
$('filePhoto').onchange = async e=>{
  const files=[...e.target.files]; e.target.value='';
  const num=photoTarget; if(num===null||num===undefined) return;
  for(let i=0;i<files.length;i++){
    try{
      busy('Abriendo foto…');
      const url=URL.createObjectURL(files[i]);
      const img=await loadImage(url);
      const orig=scaledDataURL(img,1600,0.85);
      URL.revokeObjectURL(url);
      unbusy();
      const res=await openEditor({orig, size:'M', canDelete:false, title:files.length>1?('Foto '+(i+1)+' de '+files.length):'Ajustar foto'});
      if(!res || res.deleted) break;
      (cur.photos[num]=cur.photos[num]||[]).push({orig, data:res.data, thumb:res.thumb, w:res.w, h:res.h, size:res.size, edit:res.edit});
      await saveNow(); refreshCard(num);
    }catch(err){ unbusy(); console.error(err); alert('No se pudo procesar la foto: '+err.message); break; }
  }
};
async function editPhoto(num,i){
  const p=cur.photos[num][i]; if(!p) return;
  const res=await openEditor({orig:p.orig||p.data, edit:p.edit, size:p.size, canDelete:true, title:'Editar foto'});
  if(!res) return;
  if(res.deleted) cur.photos[num].splice(i,1);
  else{ if(!p.orig) p.orig=p.data; Object.assign(p,{data:res.data,thumb:res.thumb,w:res.w,h:res.h,size:res.size,edit:res.edit}); }
  await saveNow(); refreshCard(num);
}


/* ---------------- fotos guardadas de la estación ---------------- */
let bankMode='manage', bankSel=[], bankTarget=null;
function bank(){ return cur.bank = cur.bank || []; }
function bankUsedSet(){
  const s=new Set();
  Object.values(cur.photos).forEach(arr=>(arr||[]).forEach(p=>{ if(p.bankId) s.add(p.bankId); }));
  return s;
}
$('fileBank').onchange = async e=>{
  const files=[...e.target.files]; e.target.value='';
  if(!files.length || !cur) return;
  let ok=0;
  for(let i=0;i<files.length;i++){
    try{
      busy('Guardando foto '+(i+1)+' de '+files.length+'…');
      await nextFrame();
      const url=URL.createObjectURL(files[i]);
      const img=await loadImage(url);
      const data=scaledDataURL(img,1600,0.85), thumb=scaledDataURL(img,260,0.7);
      const w=img.naturalWidth, h=img.naturalHeight, s=Math.min(1,1600/Math.max(w,h));
      URL.revokeObjectURL(url);
      bank().push({id:'b'+Date.now().toString(36)+Math.random().toString(36).slice(2,6), data, thumb, w:Math.round(w*s), h:Math.round(h*s)});
      ok++;
      if(ok%4===0) await saveNow();
    }catch(err){ console.error(err); }
  }
  unbusy();
  await saveNow();
  if(ok<files.length) toast('⚠ '+(files.length-ok)+' foto(s) no se pudieron leer', 3500);
  else toast(ok+' foto'+(ok>1?'s':'')+' guardada'+(ok>1?'s':'')+' en la estación');
  renderHeaderCard();
  if(!$('bankModal').classList.contains('hidden')) renderBank();
};
function openBank(mode, target){
  bankMode=mode; bankTarget=target; bankSel=[];
  $('bankModal').classList.remove('hidden');
  renderBank();
}
function closeBank(){ $('bankModal').classList.add('hidden'); renderHeaderCard(); }
function renderBank(){
  const list=bank(), used=bankUsedSet(), grid=$('bankGrid');
  $('bankTitle').textContent = bankMode==='ocr' ? 'Elige la foto del inventario' : bankMode==='pick' ? 'Elegir fotos' : 'Fotos de la estación ('+list.length+')';
  grid.innerHTML='';
  if(!list.length){
    grid.innerHTML='<div class="bank-empty">Aún no hay fotos guardadas.<br><br>Pulsa <b>«Añadir fotos»</b> y marca de una vez todas las de esta estación en tu fototeca.</div>';
  }
  list.forEach(b=>{
    const el=document.createElement('div'); el.className='bk';
    const idx=bankSel.indexOf(b.id);
    if(idx>=0) el.classList.add('sel');
    el.innerHTML='<img src="'+b.thumb+'">'+
      (used.has(b.id)?'<div class="bk-used">en uso</div>':'')+
      ((bankMode==='pick'||bankMode==='ocr') ? (idx>=0?'<div class="bk-n">'+(idx+1)+'</div>':'<div class="bk-n off"></div>') : '<div class="bk-x">✕</div>');
    el.onclick=ev=>{
      if(bankMode==='pick'||bankMode==='ocr'){
        const k=bankSel.indexOf(b.id);
        if(bankMode==='ocr') bankSel=(k>=0?[]:[b.id]);
        else if(k>=0) bankSel.splice(k,1); else bankSel.push(b.id);
        renderBank();
      }else if(ev.target.classList.contains('bk-x')){
        if(confirm('¿Quitar esta foto de la estación?\nLas que ya hayas puesto en preguntas se mantienen.')){
          cur.bank=bank().filter(x=>x.id!==b.id); saveNow(); renderBank();
        }
      }
    };
    grid.appendChild(el);
  });
  const use=$('bankUse');
  use.classList.toggle('hidden', bankMode==='manage');
  use.disabled = !bankSel.length;
  use.textContent = bankMode==='ocr' ? (bankSel.length?'Leer esta foto':'Elige una foto') : bankSel.length ? 'Usar '+bankSel.length+' foto'+(bankSel.length>1?'s':'') : 'Elige fotos';
}
$('bankClose').onclick = closeBank;
$('bankUse').onclick = async ()=>{
  if(bankMode==='ocr'){
    const b=bank().find(x=>x.id===bankSel[0]); if(!b) return;
    $('bankModal').classList.add('hidden');
    try{
      let r;
      if(aiMode()){ busy('Leyendo la foto con IA…'); showCancel(); r=await askAIInventory(await (await fetch(b.data)).blob()); }
      else { const img=await loadImage(b.data); busy('Preparando…'); r=await ocrInventory(img); }
      unbusy(); showInvConfirm(r);
    }catch(err){ unbusy(); if(err&&err.cancelled) return; console.error(err); alert(err.message||err); }
    return;
  }
  const num=bankTarget; if(num===null||num===undefined||!bankSel.length) return;
  const arr = cur.photos[num] = cur.photos[num] || [];
  bankSel.forEach(id=>{
    const b=bank().find(x=>x.id===id); if(!b) return;
    arr.push({bankId:b.id, data:b.data, thumb:b.thumb, w:b.w, h:b.h, size:'M', edit:null});
  });
  const n=bankSel.length;
  $('bankModal').classList.add('hidden');
  await saveNow(); refreshCard(num); renderHeaderCard();
  toast(n+' foto'+(n>1?'s':'')+' añadida'+(n>1?'s':'')+' · toca una foto para recortarla o cambiar su tamaño', 3200);
};
/* hoja "+" de cada pregunta */
function openAddSheet(num){
  $('sheetBankN').textContent = bank().length;
  $('addSheet').classList.remove('hidden');
}
$('sheetBank').onclick = ()=>{ $('addSheet').classList.add('hidden'); openBank('pick', photoTarget); };
$('sheetNew').onclick = ()=>{ $('filePhoto').click(); $('addSheet').classList.add('hidden'); };
$('sheetCancel').onclick = ()=>$('addSheet').classList.add('hidden');
$('addSheet').onclick = e=>{ if(e.target.id==='addSheet') $('addSheet').classList.add('hidden'); };

/* ---------------- editor de foto (recorte libre, rotar, tamaño) ---------------- */
const ED = {rc:null, drag:null};
const clamp=(v,a,b)=>Math.max(a,Math.min(b,v));

function makeRotated(img, rot){
  const w=img.naturalWidth, h=img.naturalHeight, swap=(rot%180!==0);
  const c=document.createElement('canvas'); c.width=swap?h:w; c.height=swap?w:h;
  const x=c.getContext('2d'); x.translate(c.width/2,c.height/2); x.rotate(rot*Math.PI/180); x.drawImage(img,-w/2,-h/2);
  return c;
}
function edRebuild(crop){
  ED.rc=makeRotated(ED.img,ED.rot);
  ED.crop = crop ? {x:crop.x,y:crop.y,w:crop.w,h:crop.h} : {x:0,y:0,w:ED.rc.width,h:ED.rc.height};
}
function edLayout(){
  const st=$('edStage'); const sw=st.clientWidth-16, sh=st.clientHeight-16;
  ED.scale=Math.min(sw/ED.rc.width, sh/ED.rc.height);
  ED.dw=ED.rc.width*ED.scale; ED.dh=ED.rc.height*ED.scale;
}
function edDraw(){
  if(!ED.rc) return;
  const c=$('edCanvas'), dpr=window.devicePixelRatio||1, ctx=c.getContext('2d');
  c.width=Math.round(ED.dw*dpr); c.height=Math.round(ED.dh*dpr); c.style.width=ED.dw+'px'; c.style.height=ED.dh+'px';
  ctx.setTransform(dpr,0,0,dpr,0,0);
  ctx.drawImage(ED.rc,0,0,ED.dw,ED.dh);
  const s=ED.scale, cr=ED.crop, x=cr.x*s, y=cr.y*s, w=cr.w*s, h=cr.h*s;
  ctx.fillStyle='rgba(0,0,0,.58)';
  ctx.fillRect(0,0,ED.dw,y); ctx.fillRect(0,y+h,ED.dw,ED.dh-y-h); ctx.fillRect(0,y,x,h); ctx.fillRect(x+w,y,ED.dw-x-w,h);
  ctx.strokeStyle='rgba(255,255,255,.35)'; ctx.lineWidth=1;
  for(let i=1;i<3;i++){ ctx.beginPath(); ctx.moveTo(x+w*i/3,y); ctx.lineTo(x+w*i/3,y+h); ctx.moveTo(x,y+h*i/3); ctx.lineTo(x+w,y+h*i/3); ctx.stroke(); }
  ctx.strokeStyle='#fff'; ctx.lineWidth=2; ctx.strokeRect(x,y,w,h);
  edHandles().forEach(hd=>{
    ctx.fillStyle='#fff'; ctx.strokeStyle='#f2711c'; ctx.lineWidth=2;
    ctx.beginPath(); ctx.rect(hd.x-8,hd.y-8,16,16); ctx.fill(); ctx.stroke();
  });
  $('edInfo').textContent=Math.round(cr.w)+' × '+Math.round(cr.h)+' px · arrastra los cuadros para ajustar ancho y alto';
}
function edHandles(){
  const s=ED.scale, c=ED.crop, x=c.x*s, y=c.y*s, w=c.w*s, h=c.h*s;
  return [{n:'tl',x:x,y:y},{n:'t',x:x+w/2,y:y},{n:'tr',x:x+w,y:y},{n:'r',x:x+w,y:y+h/2},
          {n:'br',x:x+w,y:y+h},{n:'b',x:x+w/2,y:y+h},{n:'bl',x:x,y:y+h},{n:'l',x:x,y:y+h/2}];
}
function edPos(e){ const r=$('edCanvas').getBoundingClientRect(); return {x:e.clientX-r.left,y:e.clientY-r.top}; }
function edHit(p){
  let best=null, bd=30;
  edHandles().forEach(h=>{ const d=Math.hypot(h.x-p.x,h.y-p.y); if(d<bd){bd=d;best=h.n;} });
  if(best) return best;
  const s=ED.scale,c=ED.crop;
  if(p.x>=c.x*s&&p.x<=(c.x+c.w)*s&&p.y>=c.y*s&&p.y<=(c.y+c.h)*s) return 'move';
  return null;
}
function edApply(dx,dy){
  const d=ED.drag, c0=ED.start, W=ED.rc.width, H=ED.rc.height, MIN=Math.max(30,Math.min(W,H)*0.05);
  let {x,y,w,h}=c0;
  if(d==='move'){ x=clamp(c0.x+dx,0,W-w); y=clamp(c0.y+dy,0,H-h); }
  else{
    if(d.includes('l')){ const nx=clamp(c0.x+dx,0,c0.x+c0.w-MIN); w=c0.x+c0.w-nx; x=nx; }
    if(d.includes('r')){ w=clamp(c0.w+dx,MIN,W-c0.x); }
    if(d.includes('t')){ const ny=clamp(c0.y+dy,0,c0.y+c0.h-MIN); h=c0.y+c0.h-ny; y=ny; }
    if(d.includes('b')){ h=clamp(c0.h+dy,MIN,H-c0.y); }
  }
  ED.crop={x,y,w,h};
}
(function bindEditor(){
  const cv=$('edCanvas');
  cv.addEventListener('pointerdown',e=>{
    if(!ED.rc) return;
    const p=edPos(e), d=edHit(p); if(!d) return;
    ED.drag=d; ED.p0=p; ED.start={...ED.crop};
    try{ cv.setPointerCapture(e.pointerId); }catch(_){}
    e.preventDefault();
  });
  cv.addEventListener('pointermove',e=>{
    if(!ED.drag) return;
    const p=edPos(e); edApply((p.x-ED.p0.x)/ED.scale,(p.y-ED.p0.y)/ED.scale); edDraw(); e.preventDefault();
  });
  const end=()=>{ ED.drag=null; };
  cv.addEventListener('pointerup',end); cv.addEventListener('pointercancel',end);
  window.addEventListener('resize',()=>{ if(ED.rc && !$('editor').classList.contains('hidden')){ edLayout(); edDraw(); } });

  $('edRotL').onclick=()=>edRotate(-90);
  $('edRotR').onclick=()=>edRotate(90);
  $('edReset').onclick=()=>{ ED.crop={x:0,y:0,w:ED.rc.width,h:ED.rc.height}; edDraw(); };
  $('edSize').querySelectorAll('button').forEach(b=>b.onclick=()=>{ ED.size=b.dataset.s; edSizeUI(); });
  $('edCancel').onclick=()=>edClose(null);
  $('edDel').onclick=()=>{ if(confirm('¿Eliminar esta foto?')) edClose({deleted:true}); };
  $('edSave').onclick=edSave;
})();
function edRotate(delta){
  ED.rot=(ED.rot+delta+360)%360; edRebuild(null); edLayout(); edDraw();
}
function edSizeUI(){ $('edSize').querySelectorAll('button').forEach(b=>b.classList.toggle('on', b.dataset.s===ED.size)); }
function edSave(){
  const cr=ED.crop, s=Math.min(1,1400/Math.max(cr.w,cr.h));
  const oc=document.createElement('canvas'); oc.width=Math.max(1,Math.round(cr.w*s)); oc.height=Math.max(1,Math.round(cr.h*s));
  oc.getContext('2d').drawImage(ED.rc,cr.x,cr.y,cr.w,cr.h,0,0,oc.width,oc.height);
  const ts=Math.min(1,200/Math.max(oc.width,oc.height));
  const tc=document.createElement('canvas'); tc.width=Math.max(1,Math.round(oc.width*ts)); tc.height=Math.max(1,Math.round(oc.height*ts));
  tc.getContext('2d').drawImage(oc,0,0,tc.width,tc.height);
  edClose({data:oc.toDataURL('image/jpeg',0.82), thumb:tc.toDataURL('image/jpeg',0.7), w:oc.width, h:oc.height, size:ED.size, edit:{rot:ED.rot, crop:{...cr}}});
}
function edClose(result){
  $('editor').classList.add('hidden'); ED.rc=null; ED.drag=null;
  const r=ED.resolve; ED.resolve=null; if(r) r(result);
}
async function openEditor(opts){
  const img=await loadImage(opts.orig);
  return new Promise(async resolve=>{
    ED.resolve=resolve; ED.img=img; ED.size=opts.size||'M'; ED.rot=opts.edit?opts.edit.rot:0;
    $('edTitle').textContent=opts.title||'Ajustar foto';
    $('edDel').classList.toggle('hidden',!opts.canDelete);
    edSizeUI();
    $('editor').classList.remove('hidden');
    await nextFrame(); await nextFrame();
    edRebuild(opts.edit?opts.edit.crop:null); edLayout(); edDraw();
  });
}

/* ---------------- vista previa y PDF ---------------- */
$('previewBtn').onclick = async ()=>{
  await saveNow();
  $('pvBody').innerHTML = buildPreviewHTML(cur);
  $('pvNote').textContent = 'Generando PDF…';
  $('pvOpen').disabled = true; $('pvSave').disabled = true; pdfBlob = null;
  showScreen('preview');
  const myAudit = cur;
  await new Promise(r=>setTimeout(r,60));
  try{
    const b = buildPdfBlob(myAudit);
    if(cur !== myAudit) return;
    pdfBlob = b;
    $('pvNote').textContent = 'Vista aproximada. El PDF respeta la misma estructura en páginas A4 ('+Math.round(b.size/1024)+' KB).';
    $('pvOpen').disabled = false; $('pvSave').disabled = false;
  }catch(err){
    console.error(err);
    $('pvNote').textContent = 'No se pudo generar el PDF: '+err.message;
  }
};
function downloadBlob(blob, name){
  const url=URL.createObjectURL(blob), a=document.createElement('a');
  a.href=url; a.download=name; document.body.appendChild(a); a.click();
  setTimeout(()=>{ document.body.removeChild(a); URL.revokeObjectURL(url); },1500);
}
$('pvSave').onclick = async ()=>{
  if(!pdfBlob) return;
  const name = pdfFileName(cur);
  const file = new File([pdfBlob], name, {type:'application/pdf'});
  try{
    if(navigator.canShare && navigator.canShare({files:[file]})){
      await navigator.share({files:[file], title:name});
    }else{
      downloadBlob(pdfBlob, name);
    }
    cur.exportedPdf=true; await markDone(); toast('PDF listo ✔');
  }catch(err){
    if(err && err.name==='AbortError') return;
    console.error(err); downloadBlob(pdfBlob, name);
  }
};
async function markDone(){
  if(cur.exportedPdf && cur.exportedXlsx) cur.status='finalizado';
  await saveNow();
}
async function saveExcel(){
  if(!cur) return;
  const sv = cur.questions.filter(q=>!q.status).length;
  if(sv && !confirm('Faltan '+sv+' preguntas sin valorar (SÍ / NO / N/A / PTE).\n\nEl Excel saldría con esas celdas vacías. ¿Guardarlo igualmente?')) return;
  try{
    busy('Preparando Excel…');
    await saveNow();
    const blob = await buildXlsxBlob(cur);
    unbusy();
    const name = xlsxFileName(cur);
    const file = new File([blob], name, {type:blob.type});
    if(navigator.canShare && navigator.canShare({files:[file]})) await navigator.share({files:[file], title:name});
    else downloadBlob(blob, name);
    cur.exportedXlsx=true; await markDone(); toast('Excel listo ✔');
  }catch(err){
    unbusy();
    if(err && err.name==='AbortError') return;
    console.error(err); alert('No se pudo preparar el Excel.\n\n'+(err&&err.message?err.message:err));
  }
}
$('xlsxBtn').onclick = saveExcel;
$('xlsxBtn2').onclick = saveExcel;
$('pvOpen').onclick = ()=>{
  if(!pdfBlob) return;
  window.open(URL.createObjectURL(pdfBlob), '_blank');
};

/* ---------------- arranque ---------------- */
function missingLibs(){
  const m=[];
  if(typeof XLSX==='undefined') m.push('xlsx.full.min.js');
  if(typeof jspdf==='undefined') m.push('jspdf.umd.min.js');
  if(typeof JSZip==='undefined') m.push('jszip.min.js');
  if(typeof parseWorkbook==='undefined') m.push('core.js');
  return m;
}
(async function init(){
  const m=missingLibs();
  if(m.length){
    alert('La app no ha podido cargar estos archivos:\n\n'+m.join('\n')+'\n\nRevisa que estén subidos a GitHub, sueltos en la raíz junto a index.html.');
    return;
  }
  try{ db=await openDB(); }catch(e){ alert('No se puede usar el almacenamiento del navegador: '+e); return; }
  if(navigator.storage && navigator.storage.persist) navigator.storage.persist().catch(()=>{});
  if('serviceWorker' in navigator) navigator.serviceWorker.register('sw.js',{updateViaCache:'none'}).catch(()=>{});
  await renderHome(); showScreen('home');
})();
