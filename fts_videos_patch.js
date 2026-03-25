/* ================================================================
   FTS VIDEO TRAINING MODULE — v2.0 (patch corregido)
   Pegar justo antes de </body> en index.html
   Requiere: videos.json en la raíz del repo GitHub
   IDs verificados contra el index.html real de yinyo1/FTS_DC3
   ================================================================ */
(function () {
'use strict';

const VIDEOS_URL = 'https://raw.githubusercontent.com/yinyo1/FTS_DC3/main/videos.json';

const VM = {
  catalogo:    [],
  orden:       [],
  completados: {},
  empleado:    null,
  indice:      0,
  camStream:   null,
};

/* ── 1. Cargar catálogo ──────────────────────────────────────── */
async function cargarCatalogo() {
  try {
    const r = await fetch(VIDEOS_URL + '?_=' + Date.now());
    if (!r.ok) throw new Error(r.status);
    VM.catalogo = (await r.json()).filter(v => v.activo !== false);
  } catch (e) {
    console.warn('[FTS-Vid]', e.message);
    VM.catalogo = [];
  }
}

/* ── 2. CSS ──────────────────────────────────────────────────── */
function inyectarCSS() {
  if (document.getElementById('fts-vid-css')) return;
  const s = document.createElement('style');
  s.id = 'fts-vid-css';
  s.textContent = `
  #fts-vid-overlay{position:fixed;inset:0;background:rgba(0,0,0,.8);z-index:9000;
    display:none;align-items:flex-start;justify-content:center;
    overflow-y:auto;padding:16px 10px 40px}
  #fts-vid-overlay.on{display:flex}
  #fts-vid-box{background:#fff;border-radius:16px;width:100%;max-width:580px;padding:20px;position:relative}
  .fv-card{background:#f5f5f5;border:1px solid #e0e0e0;border-radius:10px;padding:13px 14px;margin-bottom:9px}
  .fv-title{font-weight:700;font-size:14px;margin-bottom:2px}
  .fv-url{font-size:11px;color:#0078D4;word-break:break-all;background:#f0f6ff;padding:3px 7px;border-radius:4px}
  .fv-meta{font-size:11px;color:#999;margin-top:4px}
  .fv-btns{display:flex;gap:6px;margin-top:8px}
  .fv-btn{font-size:12px;padding:5px 11px;border-radius:6px;border:none;cursor:pointer;font-family:Inter,sans-serif;font-weight:600}
  .fv-edit{background:#e8e8e8;color:#333}.fv-del{background:#ffeaea;color:#c00}.fv-tog{background:#e8f5e9;color:#107C10}
  .fv-form{background:#f0f0f0;border:2px dashed #D83B01;border-radius:10px;padding:15px;margin-bottom:14px}
  .fv-form input{width:100%;box-sizing:border-box;margin-bottom:8px;padding:9px 12px;border:1px solid #ccc;border-radius:7px;font-family:Inter,sans-serif;font-size:13px}
  .fv-form label{font-size:11px;color:#666;display:block;margin-bottom:2px}
  .fv-json{background:#111;color:#00ff88;font-family:monospace;font-size:10px;padding:12px;border-radius:8px;white-space:pre;max-height:250px;overflow-y:auto;overflow-x:auto}
  #fts-order-wrap{background:#fff;border:1px solid #e0e0e0;border-radius:14px;padding:16px;margin-bottom:14px}
  .fo-card{background:#f5f5f5;border:1.5px solid #e0e0e0;border-radius:10px;padding:11px 13px;margin-bottom:8px;display:flex;align-items:center;gap:10px}
  .fo-num{width:28px;height:28px;border-radius:50%;background:#D83B01;color:#fff;font-weight:800;font-size:12px;flex-shrink:0;display:flex;align-items:center;justify-content:center}
  .fo-info{flex:1}.fo-title{font-weight:700;font-size:13px}.fo-sub{font-size:11px;color:#666}
  .fo-arrows{display:flex;flex-direction:column;gap:2px}
  .fo-arr{background:#e8e8e8;border:none;border-radius:4px;width:24px;height:22px;cursor:pointer;font-size:11px;display:flex;align-items:center;justify-content:center}
  .fo-arr:disabled{opacity:.3}
  #fts-player-wrap{background:#fff;border:1px solid #e0e0e0;border-radius:14px;overflow:hidden;margin-bottom:14px}
  #fts-player-wrap iframe,#fts-player-wrap video{width:100%;height:210px;border:none;display:block}
  .fts-vprog-bar{height:4px;background:#e0e0e0}.fts-vprog-fill{height:100%;background:#D83B01;transition:width .4s}
  .fts-hash{background:#0a0a1a;color:#00ff88;font-family:monospace;font-size:10px;padding:7px 10px;border-radius:6px;word-break:break-all;margin:8px 0}
  #fts-cam-modal{position:fixed;inset:0;background:rgba(0,0,0,.88);z-index:9100;display:none;align-items:center;justify-content:center}
  #fts-cam-modal.on{display:flex}
  .fcm-box{background:#fff;border-radius:16px;padding:20px;max-width:340px;width:92%;text-align:center}
  .fcm-box h3{margin:0 0 6px;font-size:16px}.fcm-box p{margin:0 0 14px;font-size:12px;color:#666}
  #fcm-video{width:100%;border-radius:10px;transform:scaleX(-1)}
  #fcm-canvas{display:none}#fcm-preview{width:100%;border-radius:10px;display:none;margin-bottom:10px}
  .fcm-btns{display:flex;gap:8px;margin-top:10px}
  .fcm-btns button{flex:1;padding:10px;border-radius:8px;border:none;font-family:Inter,sans-serif;font-weight:700;font-size:13px;cursor:pointer}
  .fcm-capture{background:#D83B01;color:#fff}.fcm-retake{background:#f0f0f0;color:#333}.fcm-confirm{background:#107C10;color:#fff}
  `;
  document.head.appendChild(s);
}

/* ── 3. Modal cámara ─────────────────────────────────────────── */
function inyectarModalCam() {
  if (document.getElementById('fts-cam-modal')) return;
  document.body.insertAdjacentHTML('beforeend', `
    <div id="fts-cam-modal">
      <div class="fcm-box">
        <h3>📸 Foto de verificación</h3>
        <p>Mira a la cámara — quedará como evidencia de tu curso</p>
        <video id="fcm-video" autoplay playsinline></video>
        <canvas id="fcm-canvas"></canvas>
        <img id="fcm-preview" alt="">
        <div class="fcm-btns" id="fcm-btns">
          <button class="fcm-capture" onclick="FTSVid.tomarFoto()">📷 Tomar foto</button>
        </div>
        <button onclick="FTSVid.saltarFoto()" style="margin-top:8px;width:100%;background:#f0f0f0;color:#666;border:none;border-radius:8px;padding:8px;font-size:12px;cursor:pointer;font-family:Inter,sans-serif">× Continuar sin foto</button>
      </div>
    </div>`);
}

/* ── 4. Botón admin en #settings-panel (ID real del HTML) ───── */
function inyectarBotonAdmin() {
  const panel = document.getElementById('settings-panel');
  if (!panel || document.getElementById('fts-btn-vid-admin')) return;
  panel.insertAdjacentHTML('beforeend', `
    <button id="fts-btn-vid-admin" class="btn btn-s"
      style="padding:11px;margin-top:8px;border-color:#D83B01;color:#D83B01"
      onclick="FTSVid.abrirGestorVideos()">
      🎬 Gestionar Videos de Capacitación
    </button>`);
}

/* ── 5. Overlay gestor de videos ─────────────────────────────── */
function abrirGestorVideos() {
  let ov = document.getElementById('fts-vid-overlay');
  if (!ov) {
    document.body.insertAdjacentHTML('beforeend', `
      <div id="fts-vid-overlay">
        <div id="fts-vid-box">
          <button onclick="document.getElementById('fts-vid-overlay').classList.remove('on')"
            style="position:absolute;top:12px;right:12px;background:#f0f0f0;border:none;border-radius:50%;width:28px;height:28px;cursor:pointer;font-size:15px">✕</button>
          <div id="fts-vid-inner"></div>
        </div>
      </div>`);
    ov = document.getElementById('fts-vid-overlay');
  }
  ov.classList.add('on');
  renderGestor();
}

function renderGestor() {
  const inner = document.getElementById('fts-vid-inner');
  if (!inner) return;
  let h = `
    <h3 style="margin:0 0 4px;font-size:16px">🎬 Catálogo de Videos</h3>
    <p style="font-size:12px;color:#666;margin-bottom:12px">Videos para los empleados (SharePoint / MP4)</p>
    <button onclick="FTSVid.mostrarForm(null)" style="width:100%;background:#D83B01;color:#fff;border:none;border-radius:8px;padding:10px;font-size:13px;font-weight:700;cursor:pointer;font-family:Inter,sans-serif;margin-bottom:12px">➕ Agregar Video</button>
    <div id="fv-form-wrap"></div>`;

  if (!VM.catalogo.length) {
    h += `<p style="color:#999;text-align:center;padding:20px">Sin videos. Agrega el primero.</p>`;
  } else {
    VM.catalogo.forEach(v => {
      h += `<div class="fv-card">
        <div class="fv-title">${v.activo===false?'🔴':'🎬'} ${v.titulo}</div>
        <div class="fv-url">${v.url||'(sin URL)'}</div>
        <div class="fv-meta">⏱ ${v.duracion_min||'?'} min · ${v.obligatorio?'✅ Obligatorio':'⬜ Opcional'} · Orden ${v.orden_sugerido||'?'}</div>
        <div class="fv-btns">
          <button class="fv-btn fv-edit" onclick="FTSVid.mostrarForm('${v.id}')">✏️ Editar</button>
          <button class="fv-btn fv-tog"  onclick="FTSVid.toggleActivo('${v.id}')">${v.activo===false?'▶️ Activar':'⏸ Desactivar'}</button>
          <button class="fv-btn fv-del"  onclick="FTSVid.eliminarVideo('${v.id}')">🗑</button>
        </div>
      </div>`;
    });
  }

  h += `<hr style="margin:14px 0;border-color:#e0e0e0">
    <h4 style="font-size:12px;margin:0 0 6px">📋 JSON → pegar en videos.json en GitHub</h4>
    <div class="fv-json">${JSON.stringify(VM.catalogo,null,2)}</div>
    <button onclick="FTSVid.copiarJSON()" style="margin-top:8px;width:100%;background:#f0f0f0;border:none;border-radius:7px;padding:9px;font-size:12px;font-weight:600;cursor:pointer;font-family:Inter,sans-serif">📋 Copiar JSON</button>`;

  inner.innerHTML = h;
}

function mostrarForm(id) {
  const v = id ? VM.catalogo.find(x=>x.id===id) : null;
  const wrap = document.getElementById('fv-form-wrap');
  if (!wrap) return;
  wrap.innerHTML = `<div class="fv-form">
    <h4 style="margin:0 0 10px;font-size:13px">${v?'✏️ Editar':'➕ Nuevo'} Video</h4>
    <label>ID único</label><input id="fvi-id" value="${v?.id||'v'+String(VM.catalogo.length+1).padStart(3,'0')}" ${v?'readonly style="background:#f0f0f0"':''}>
    <label>Título</label><input id="fvi-titulo" value="${v?.titulo||''}" placeholder="Seguridad en Alturas">
    <label>Descripción</label><input id="fvi-desc" value="${v?.descripcion||''}" placeholder="Breve descripción">
    <label>URL (SharePoint embed o MP4)</label><input id="fvi-url" value="${v?.url||''}" placeholder="https://...sharepoint.com/...">
    <label>Duración (min)</label><input id="fvi-dur" type="number" value="${v?.duracion_min||30}" min="1">
    <label>Orden sugerido</label><input id="fvi-orden" type="number" value="${v?.orden_sugerido||VM.catalogo.length+1}" min="1">
    <label style="display:flex;align-items:center;gap:7px;cursor:pointer"><input type="checkbox" id="fvi-oblig" ${v?.obligatorio!==false?'checked':''}> Obligatorio</label>
    <div style="display:flex;gap:8px;margin-top:10px">
      <button onclick="FTSVid.guardarVideo('${id||''}')" style="flex:1;background:#D83B01;color:#fff;border:none;border-radius:8px;padding:10px;font-size:13px;font-weight:700;cursor:pointer;font-family:Inter,sans-serif">💾 Guardar</button>
      <button onclick="document.getElementById('fv-form-wrap').innerHTML=''" style="flex:1;background:#f0f0f0;color:#333;border:none;border-radius:8px;padding:10px;font-size:13px;cursor:pointer;font-family:Inter,sans-serif">Cancelar</button>
    </div>
    <p style="font-size:10px;color:#aaa;margin-top:8px">💡 SharePoint: abre el video → Compartir → Insertar → copia el src del iframe</p>
  </div>`;
}

function guardarVideo(id) {
  const g = sel => document.getElementById(sel)?.value?.trim();
  const v = {
    id: g('fvi-id'), titulo: g('fvi-titulo'), descripcion: g('fvi-desc'),
    url: g('fvi-url'), tipo: g('fvi-url').includes('sharepoint')?'sharepoint':'video',
    duracion_min: parseInt(g('fvi-dur'))||30,
    obligatorio: document.getElementById('fvi-oblig')?.checked,
    activo: true, fecha_alta: new Date().toISOString().split('T')[0],
    orden_sugerido: parseInt(g('fvi-orden'))||1,
  };
  if (!v.id||!v.titulo||!v.url) { alert('Completa ID, Título y URL'); return; }
  if (id) {
    const i = VM.catalogo.findIndex(x=>x.id===id);
    if (i>=0) VM.catalogo[i]=v; else VM.catalogo.push(v);
  } else {
    if (VM.catalogo.find(x=>x.id===v.id)) { alert('ID ya existe'); return; }
    VM.catalogo.push(v);
  }
  VM.catalogo.sort((a,b)=>(a.orden_sugerido||99)-(b.orden_sugerido||99));
  renderGestor();
}

function toggleActivo(id) { const v=VM.catalogo.find(x=>x.id===id); if(v){v.activo=!v.activo;renderGestor();} }
function eliminarVideo(id) { if(!confirm('¿Eliminar este video?'))return; VM.catalogo=VM.catalogo.filter(x=>x.id!==id); renderGestor(); }
function copiarJSON() { navigator.clipboard.writeText(JSON.stringify(VM.catalogo,null,2)).then(()=>alert('✅ JSON copiado. Pégalo en videos.json en GitHub.')).catch(()=>alert('Selecciona el texto manualmente.')); }

/* ── 6. Hook en #s-edash (ID real del dashboard de empleado) ── */
function hookEdash() {
  const obs = new MutationObserver(() => {
    const edash = document.getElementById('s-edash');
    if (!edash?.classList.contains('on')) return;
    if (!VM.catalogo.length) return;
    if (document.getElementById('fts-order-wrap')) return;

    const content = edash.querySelector('.content');
    if (!content) return;

    // Leer datos del empleado de los inputs reales
    VM.empleado = {
      nombre:  document.getElementById('emp-nom')?.value  || '',
      apaterno:document.getElementById('emp-ap')?.value   || '',
      amaterno:document.getElementById('emp-am')?.value   || '',
      curp:    document.getElementById('emp-curp')?.value || '',
    };
    VM.orden = VM.catalogo.map(v=>v.id);

    // Inyectar selector de orden
    const orderDiv = document.createElement('div');
    orderDiv.id = 'fts-order-wrap';
    orderDiv.innerHTML = renderOrden();
    content.insertBefore(orderDiv, content.firstChild);

    // Inyectar panel del player (oculto)
    const playerDiv = document.createElement('div');
    playerDiv.id = 'fts-player-wrap';
    playerDiv.style.display = 'none';
    content.insertBefore(playerDiv, content.children[1]||null);
  });
  obs.observe(document.body, { attributes:true, subtree:true, attributeFilter:['class'] });
}

/* ── 7. Selector de orden ────────────────────────────────────── */
function renderOrden() {
  let h = `<h4 style="margin:0 0 4px;font-size:14px;font-weight:700">📋 Elige el orden de tus videos</h4>
    <p style="font-size:12px;color:#666;margin-bottom:10px">Usa ▲▼ para reordenar</p>`;
  VM.orden.forEach((id,i) => {
    const v = VM.catalogo.find(x=>x.id===id); if(!v)return;
    h += `<div class="fo-card">
      <div class="fo-num">${i+1}</div>
      <div class="fo-info"><div class="fo-title">${v.titulo}</div>
        <div class="fo-sub">⏱ ${v.duracion_min} min · ${v.obligatorio?'✅ Obligatorio':'⬜ Opcional'}</div></div>
      <div class="fo-arrows">
        <button class="fo-arr" onclick="FTSVid.mover('${id}',-1)" ${i===0?'disabled':''}>▲</button>
        <button class="fo-arr" onclick="FTSVid.mover('${id}',1)"  ${i===VM.orden.length-1?'disabled':''}>▼</button>
      </div></div>`;
  });
  h += `<button onclick="FTSVid.iniciar()" style="width:100%;background:#D83B01;color:#fff;border:none;border-radius:10px;padding:13px;font-size:14px;font-weight:800;cursor:pointer;font-family:Inter,sans-serif;margin-top:6px">🚀 Iniciar videos en este orden →</button>`;
  return h;
}

function moverVideo(id,dir) {
  const i=VM.orden.indexOf(id); if(i<0)return;
  const j=i+dir; if(j<0||j>=VM.orden.length)return;
  [VM.orden[i],VM.orden[j]]=[VM.orden[j],VM.orden[i]];
  const w=document.getElementById('fts-order-wrap'); if(w)w.innerHTML=renderOrden();
}

/* ── 8. Player ───────────────────────────────────────────────── */
function iniciarCursos() { VM.indice=0; siguienteVideo(); }

function siguienteVideo() {
  while(VM.indice<VM.orden.length && VM.completados[VM.orden[VM.indice]]) VM.indice++;
  if(VM.indice>=VM.orden.length){mostrarResumenFinal();return;}
  const v=VM.catalogo.find(x=>x.id===VM.orden[VM.indice]);
  if(!v){VM.indice++;siguienteVideo();return;}
  mostrarPlayer(v);
}

function mostrarPlayer(v) {
  const ow=document.getElementById('fts-order-wrap'); if(ow)ow.style.display='none';
  const wrap=document.getElementById('fts-player-wrap'); if(!wrap)return;
  wrap.style.display='block';
  const esIframe=v.tipo==='sharepoint'||v.url.includes('sharepoint');
  const media=esIframe
    ?`<iframe src="${v.url}" allowfullscreen allow="autoplay"></iframe>`
    :`<video id="fts-vid-tag" controls controlslist="nodownload" preload="metadata"><source src="${v.url}"></video>`;
  wrap.innerHTML=`<div style="padding:14px">
    <p style="font-size:11px;color:#999;margin:0 0 2px">Video ${VM.indice+1} de ${VM.orden.length}</p>
    <h3 style="margin:0 0 10px;font-size:15px">${v.titulo}</h3>
    ${media}
    <div class="fts-vprog-bar"><div class="fts-vprog-fill" id="fts-vfill" style="width:0%"></div></div>
    <p style="font-size:12px;color:#666;margin:6px 0 4px">${v.descripcion||''}</p>
    <p style="font-size:11px;color:#999;margin-bottom:12px">⏱ ${v.duracion_min} min estimados</p>
    <button onclick="FTSVid.pedirFoto()" style="width:100%;background:#107C10;color:#fff;border:none;border-radius:10px;padding:13px;font-size:14px;font-weight:800;cursor:pointer;font-family:Inter,sans-serif">✅ Completado — Tomar foto y continuar →</button>
  </div>`;
  const tag=document.getElementById('fts-vid-tag');
  if(tag)tag.addEventListener('timeupdate',()=>{
    const p=tag.duration?(tag.currentTime/tag.duration*100).toFixed(1):0;
    const f=document.getElementById('fts-vfill');if(f)f.style.width=p+'%';
  });
}

/* ── 9. Foto ─────────────────────────────────────────────────── */
function pedirFoto() {
  const m=document.getElementById('fts-cam-modal');
  if(!m){completarVideo(null);return;}
  m.classList.add('on');
  navigator.mediaDevices.getUserMedia({video:{facingMode:'user'},audio:false})
    .then(s=>{VM.camStream=s;const v=document.getElementById('fcm-video');if(v){v.srcObject=s;v.play();}})
    .catch(()=>{cerrarModalCam();completarVideo(null);});
}

function tomarFoto() {
  const vid=document.getElementById('fcm-video'),cvs=document.getElementById('fcm-canvas'),prev=document.getElementById('fcm-preview');
  if(!vid||!cvs)return;
  cvs.width=vid.videoWidth||320;cvs.height=vid.videoHeight||240;
  cvs.getContext('2d').drawImage(vid,0,0);
  const d=cvs.toDataURL('image/jpeg',.75);
  window._ftsFoto=d;
  if(prev){prev.src=d;prev.style.display='block';}
  vid.style.display='none';
  const b=document.getElementById('fcm-btns');
  if(b)b.innerHTML=`<button class="fcm-retake" onclick="FTSVid.retakeFoto()">🔄 Repetir</button><button class="fcm-confirm" onclick="FTSVid.confirmarFoto()">✅ Confirmar</button>`;
}

function retakeFoto() {
  window._ftsFoto=null;
  const prev=document.getElementById('fcm-preview'),vid=document.getElementById('fcm-video');
  if(prev)prev.style.display='none';if(vid)vid.style.display='block';
  const b=document.getElementById('fcm-btns');
  if(b)b.innerHTML=`<button class="fcm-capture" onclick="FTSVid.tomarFoto()">📷 Tomar foto</button>`;
}

function confirmarFoto() { cerrarModalCam(); completarVideo(window._ftsFoto||null); window._ftsFoto=null; }
function saltarFoto()    { cerrarModalCam(); completarVideo(null); }

function cerrarModalCam() {
  const m=document.getElementById('fts-cam-modal'); if(m)m.classList.remove('on');
  if(VM.camStream){VM.camStream.getTracks().forEach(t=>t.stop());VM.camStream=null;}
  const prev=document.getElementById('fcm-preview'),vid=document.getElementById('fcm-video');
  if(prev){prev.style.display='none';prev.src='';}if(vid)vid.style.display='block';
  const b=document.getElementById('fcm-btns');
  if(b)b.innerHTML=`<button class="fcm-capture" onclick="FTSVid.tomarFoto()">📷 Tomar foto</button>`;
}

/* ── 10. Hash + completar ────────────────────────────────────── */
async function generarHash(t) {
  const buf=await crypto.subtle.digest('SHA-256',new TextEncoder().encode(t));
  return Array.from(new Uint8Array(buf)).map(b=>b.toString(16).padStart(2,'0')).join('');
}

async function completarVideo(foto) {
  const id=VM.orden[VM.indice],v=VM.catalogo.find(x=>x.id===id);
  const ts=new Date().toISOString(),emp=VM.empleado||{};
  const hash=await generarHash([emp.curp||'SIN-CURP',id,ts,navigator.userAgent.slice(0,40),'2.0'].join('|'));
  VM.completados[id]={videoId:id,titulo:v?.titulo||id,hash,foto,ts,empleado:emp};
  mostrarConfirmVideo(v,hash,foto);
}

function mostrarConfirmVideo(v,hash,foto) {
  const wrap=document.getElementById('fts-player-wrap'); if(!wrap)return;
  wrap.innerHTML=`<div style="padding:16px;text-align:center">
    <div style="font-size:36px;margin-bottom:6px">🎓</div>
    <h3 style="color:#107C10;margin:0 0 4px">¡Curso completado!</h3>
    <p style="font-size:13px;color:#333;margin:0 0 10px">${v?.titulo||''}</p>
    ${foto?`<img src="${foto}" style="width:72px;height:72px;object-fit:cover;border-radius:50%;border:3px solid #107C10;display:block;margin:0 auto 10px">`:``}
    <p style="font-size:11px;color:#666;margin:0 0 4px">🔐 Sello SHA-256:</p>
    <div class="fts-hash">${hash}</div>
    <p style="font-size:10px;color:#999;margin:4px 0 12px">${new Date().toLocaleString('es-MX')}</p>
    <button onclick="FTSVid._sig()" style="width:100%;background:#D83B01;color:#fff;border:none;border-radius:10px;padding:13px;font-size:14px;font-weight:800;cursor:pointer;font-family:Inter,sans-serif">
      ${VM.indice+1<VM.orden.length?'▶️ Siguiente video →':'📄 Ver mi constancia →'}
    </button>
  </div>`;
  VM.indice++;
}

/* ── 11. Resumen final + PDF ─────────────────────────────────── */
function mostrarResumenFinal() {
  const wrap=document.getElementById('fts-player-wrap'); if(!wrap)return;
  const total=Object.keys(VM.completados).length;
  wrap.innerHTML=`<div style="padding:16px;text-align:center">
    <div style="font-size:40px;margin-bottom:8px">🏆</div>
    <h2 style="color:#107C10;margin:0 0 4px">¡Capacitación completa!</h2>
    <p style="font-size:13px;color:#666;margin:0 0 14px">${total} video(s) con sello de verificación</p>
    <button onclick="FTSVid.generarPDF()" style="width:100%;background:#D83B01;color:#fff;border:none;border-radius:10px;padding:14px;font-size:15px;font-weight:800;cursor:pointer;font-family:Inter,sans-serif">📄 Descargar Certificados PDF</button>
    <p style="font-size:10px;color:#aaa;margin-top:8px">Incluye foto, sello SHA-256 y fecha exacta por cada video</p>
  </div>`;
}

async function generarPDF() {
  const {jsPDF}=window.jspdf;
  const doc=new jsPDF({unit:'mm',format:'letter',orientation:'portrait'});
  const W=215.9,H=279.4,emp=VM.empleado||{};
  const keys=Object.keys(VM.completados);
  for(let ki=0;ki<keys.length;ki++){
    const c=VM.completados[keys[ki]];
    if(ki>0)doc.addPage();
    doc.setFillColor(0,71,171);doc.rect(0,0,W,16,'F');
    doc.setTextColor(255,255,255);doc.setFont('helvetica','bold');doc.setFontSize(12);
    doc.text('SERVICIOS FTS SA DE CV · CONSTANCIA DE CAPACITACIÓN',W/2,10,{align:'center'});
    doc.setFillColor(245,245,245);doc.rect(0,16,W,H-32,'F');
    doc.setFillColor(255,255,255);doc.roundedRect(10,22,W-20,28,3,3,'F');
    doc.setTextColor(216,59,1);doc.setFontSize(16);
    doc.text('CERTIFICADO INDIVIDUAL DE CAPACITACIÓN',W/2,33,{align:'center'});
    doc.setFontSize(10);doc.setTextColor(100,100,100);
    doc.text(`Video ${ki+1} de ${keys.length}`,W/2,42,{align:'center'});
    if(c.foto){try{doc.addImage(c.foto,'JPEG',12,56,32,32);}catch(e){}}
    const nombre=[emp.nombre,emp.apaterno,emp.amaterno].filter(Boolean).join(' ');
    doc.setFont('helvetica','bold');doc.setFontSize(13);doc.setTextColor(10,10,10);
    doc.text(nombre||'Trabajador',50,62);
    doc.setFont('helvetica','normal');doc.setFontSize(10);doc.setTextColor(100,100,100);
    if(emp.curp)doc.text('CURP: '+emp.curp,50,69);
    doc.setFillColor(16,124,16);doc.roundedRect(10,96,W-20,38,3,3,'F');
    doc.setTextColor(255,255,255);doc.setFont('helvetica','bold');doc.setFontSize(12);
    doc.text('📚 '+c.titulo,W/2,107,{align:'center'});
    doc.setFont('helvetica','normal');doc.setFontSize(9);
    const fecha=new Date(c.ts).toLocaleString('es-MX',{weekday:'long',year:'numeric',month:'long',day:'numeric',hour:'2-digit',minute:'2-digit',second:'2-digit',timeZoneName:'short'});
    doc.text('Completado el: '+fecha,W/2,118,{align:'center'});
    doc.setFillColor(10,10,26);doc.roundedRect(10,142,W-20,28,3,3,'F');
    doc.setTextColor(0,255,136);doc.setFont('courier','bold');doc.setFontSize(7);
    doc.text('🔐 SELLO SHA-256',14,149);
    doc.setFont('courier','normal');doc.setFontSize(6.5);
    doc.text(c.hash.slice(0,32),14,155);doc.text(c.hash.slice(32),14,161);
    doc.setTextColor(120,120,120);doc.setFontSize(6);
    doc.text('Algoritmo: SHA-256 · FTS Training Module v2.0',14,167);
    try{const q=await cargarImg(`https://api.qrserver.com/v1/create-qr-code/?size=80x80&data=${c.hash.slice(0,32)}`);if(q)doc.addImage(q,'PNG',W-30,144,20,20);}catch(e){}
    doc.setFillColor(255,250,240);doc.roundedRect(10,178,W-20,18,2,2,'F');
    doc.setDrawColor(216,59,1);doc.setLineWidth(.4);doc.roundedRect(10,178,W-20,18,2,2,'S');
    doc.setFont('helvetica','italic');doc.setFontSize(7.5);doc.setTextColor(216,59,1);
    doc.text('Este certificado fue generado digitalmente. La foto y el sello SHA-256 son evidencia de participación del trabajador.',W/2,185,{align:'center',maxWidth:W-30});
    doc.setDrawColor(200,200,200);doc.setLineWidth(.3);
    doc.line(25,215,90,215);doc.line(125,215,190,215);
    doc.setFont('helvetica','normal');doc.setFontSize(8);doc.setTextColor(100,100,100);
    doc.text('Trabajador',57,220,{align:'center'});doc.text('Instructor / Supervisor',157,220,{align:'center'});
    doc.setFillColor(0,71,171);doc.rect(0,H-16,W,16,'F');
    doc.setTextColor(255,255,255);doc.setFontSize(8);
    doc.text('SERVICIOS FTS SA DE CV · Sistema de Capacitación Digital · STPS',W/2,H-8,{align:'center'});
  }
  const ap=(emp.apaterno||'empleado').replace(/\s/g,'_');
  doc.save(`FTS_Certificados_${ap}_${Date.now()}.pdf`);
}

async function cargarImg(url) {
  return new Promise(res=>{
    const img=new Image();img.crossOrigin='anonymous';
    img.onload=()=>{const c=document.createElement('canvas');c.width=img.width;c.height=img.height;c.getContext('2d').drawImage(img,0,0);res(c.toDataURL('image/png'));};
    img.onerror=()=>res(null);img.src=url;
  });
}

/* ── 12. Init ────────────────────────────────────────────────── */
async function init() {
  inyectarCSS();
  inyectarModalCam();
  await cargarCatalogo();
  const tryInject=setInterval(()=>{
    if(document.getElementById('settings-panel')){inyectarBotonAdmin();clearInterval(tryInject);}
  },400);
  hookEdash();
  console.log('[FTS-Vid] v2.0 listo —', VM.catalogo.length, 'videos');
}

window.FTSVid = {
  mostrarForm, guardarVideo, toggleActivo, eliminarVideo, copiarJSON, abrirGestorVideos,
  mover: moverVideo, iniciar: iniciarCursos, _sig: ()=>siguienteVideo(),
  tomarFoto, retakeFoto, confirmarFoto, saltarFoto, pedirFoto,
  generarPDF, estado: ()=>VM,
};

if(document.readyState==='loading'){document.addEventListener('DOMContentLoaded',init);}else{init();}
})();
