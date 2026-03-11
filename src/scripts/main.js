/* ══════════════════════════════════════
   DATA & STATE
══════════════════════════════════════ */
let ALL_DATA = [];
let TRM = 4150;

async function fetchTRM() {
  const setTRM = t => {
    TRM = t;
    const inp = document.getElementById('trm-input');
    if(inp) inp.value = Math.round(TRM);
    console.log('[TRM]', TRM);
  };
  // Source 1: frankfurter.app (CORS ok, data del BCE)
  try {
    const r = await fetch('https://api.frankfurter.app/latest?from=USD&to=COP');
    const d = await r.json();
    const t = d.rates && d.rates.COP;
    if(t > 100) { setTRM(t); return; }
  } catch(e) { console.warn('[TRM] frankfurter failed', e.message); }
  // Source 2: open.er-api.com (gratis, CORS ok)
  try {
    const r2 = await fetch('https://open.er-api.com/v6/latest/USD');
    const d2 = await r2.json();
    const t2 = d2.rates && d2.rates.COP;
    if(t2 > 100) { setTRM(t2); return; }
  } catch(e2) { console.warn('[TRM] open.er-api failed', e2.message); }
  // Source 3: cdn.jsdelivr proxy to exchangerate
  try {
    const r3 = await fetch('https://cdn.jsdelivr.net/npm/@fawazahmed0/currency-api@latest/v1/currencies/usd.json');
    const d3 = await r3.json();
    const t3 = d3.usd && d3.usd.cop;
    if(t3 > 100) { setTRM(t3); return; }
  } catch(e3) { console.warn('[TRM] jsdelivr failed', e3.message); }
  console.warn('[TRM] all sources failed, using default', TRM);
}

const COLORS = [
  '#2D4FD6','#8B5FC8','#2ABFDF','#0DBF82','#F0A020',
  '#E84040','#E040A0','#40C8F0','#F06040','#1B2B8C'
];
const MES_LABELS = {'2026-01':'Ene','2026-02':'Feb','2026-03':'Mar'};

/* Format helpers */
function fmtCOP(v){
  if(v===null||v===undefined) return '—';
  return '$ ' + Math.round(v).toLocaleString('es-CO');
}
function fmtUSD(v){
  if(v===null||v===undefined) return '—';
  return 'USD ' + Number(v).toLocaleString('en-US',{minimumFractionDigits:2,maximumFractionDigits:2});
}
function fmtNum(v){return Math.round(v).toLocaleString('es-CO')}
function fmtPct(v){return (v*100).toFixed(1)+'%'}
function abr(v){
  if(v>=1e12) return '$ '+(v/1e12).toFixed(2)+' Bill';
  if(v>=1e9)  return '$ '+(v/1e9).toFixed(2)+' MM';   // miles de millones
  if(v>=1e6)  return '$ '+(v/1e6).toFixed(1)+' M';
  if(v>=1e3)  return '$ '+(v/1e3).toFixed(0)+' K';
  return '$ '+Math.round(v).toLocaleString('es-CO');
}

function normalizePersonName(v){
  return (v||'')
    .toString()
    .replace(/\.(xlsx|xls)$/i,'')
    .normalize('NFD').replace(/[\u0300-\u036f]/g,'')
    .toLowerCase()
    .replace(/[^a-z0-9\s]/g,' ')
    .replace(/\s+/g,' ')
    .trim();
}

function namesMatch(a,b){
  const na = normalizePersonName(a);
  const nb = normalizePersonName(b);
  if(!na || !nb) return false;
  if(na === nb) return true;

  // Permite casos como "dayana chala" vs "dayana chala perez"
  if((na.includes(nb) || nb.includes(na)) && Math.min(na.length, nb.length) >= 8) {
    return true;
  }

  const ta = [...new Set(na.split(' ').filter(Boolean))];
  const tb = [...new Set(nb.split(' ').filter(Boolean))];
  const common = ta.filter(t => tb.includes(t));

  if(common.length >= 2) return true;
  if(common.length === 1 && (ta.length === 1 || tb.length === 1) && common[0].length >= 5) return true;

  return false;
}

/* Get TRM */
function getTRM(){return parseFloat(document.getElementById('trm-input').value)||4150;}

/* Liquidate to COP */
function parseMonto(v){
  if(v===null||v===undefined) return 0;
  if(typeof v==='number') return v;
  let s = String(v).trim().replace(/[$\s]/g,'');
  if(!s) return 0;
  // Detectar formato: colombiano "98.544.430" o americano "98,544.43" o decimal "33262.55"
  const dots = (s.match(/\./g)||[]).length;
  const commas = (s.match(/,/g)||[]).length;
  if(dots > 1){
    // Múltiples puntos = separadores de miles colombianos: "98.544.430" → 98544430
    s = s.replace(/\./g,'').replace(',','.');
  } else if(commas > 1){
    // Múltiples comas = separadores de miles americanos: "98,544,430" → 98544430
    s = s.replace(/,/g,'');
  } else if(dots === 1 && commas === 1){
    // Ambos: puede ser "1.234,56" (EU) o "1,234.56" (US)
    const dotPos = s.indexOf('.');
    const commaPos = s.indexOf(',');
    if(commaPos < dotPos){ s = s.replace(/,/g,''); } // US: 1,234.56
    else { s = s.replace(/\./g,'').replace(',','.'); } // EU: 1.234,56
  } else {
    // Un solo punto o coma — es separador decimal
    s = s.replace(',','.');
  }
  return parseFloat(s)||0;
}

function parsefecha(v){
  if(!v) return '';
  if(v instanceof Date) return v.toISOString().substring(0,7);
  if(typeof v==='number') {
    const d = new Date(Math.round((v - 25569)*86400*1000));
    return d.toISOString().substring(0,7);
  }
  const s = String(v);
  const meses = {ene:'01',feb:'02',mar:'03',abr:'04',may:'05',jun:'06',jul:'07',ago:'08',sep:'09',oct:'10',nov:'11',dic:'12'};
  const m1 = s.match(/(\d{2})[-/](\w{3})\.?[-/](\d{2,4})/i);
  if(m1){
    const mes = meses[m1[2].toLowerCase()]||'01';
    const anio = m1[3].length===2?'20'+m1[3]:m1[3];
    return anio+'-'+mes;
  }
  return s.substring(0,7);
}

function toCOP(row){
  const m = (row['MONEDA 2']||'COP').trim().toUpperCase();
  const val = parseMonto(row['MONTO VENTA CLIENTE']);
  const trm = parseMonto(row['TRM REFERENCIA'])||getTRM();
  if(m==='USD') return val * trm;
  return val;
}

/* Get month from date string */
function getMonth(d){
  return parsefecha(d);
}

/* Today string */
function todayStr(){
  const n = new Date();
  return n.toISOString().substring(0,10);
}

/* ══════════════════════════════════════
   FILE LOADING
══════════════════════════════════════ */
/* ══════════════════════════════════════
   FOLDER & FILE LOADING — detecta estructura director/ejecutivo
══════════════════════════════════════ */

// Parsea un archivo Excel y devuelve array de registros
function parseXlsx(file, directorHint){
  return new Promise((resolve)=>{
    const r = new FileReader();
    r.onload = function(ev){
      try{
        const wb = XLSX.read(ev.target.result,{type:'binary',cellDates:true});
        const wsName = wb.SheetNames.find(s=>s.includes('Gerencia')||s.includes('Comercial'))||wb.SheetNames[0];
        const ws = wb.Sheets[wsName];
        const raw = XLSX.utils.sheet_to_json(ws,{header:1,defval:null});
        
        let hdrIdx = -1;
        for(let i=0;i<raw.length;i++){
          if(raw[i] && raw[i].some(c=>c&&String(c).includes('CLIENTE'))){hdrIdx=i;break;}
        }
        if(hdrIdx<0){resolve([]);return;}
        
        const rawHdrs = raw[hdrIdx].map(h=>h?String(h).trim():'');
        // Normalizar nombres de columnas — distintos ejecutivos usan nombres distintos
        const KEY_MAP = {
          'VENTA CLIENTE':'MONTO VENTA CLIENTE','MONTO VENTA':'MONTO VENTA CLIENTE',
          'VALOR VENTA':'MONTO VENTA CLIENTE','VALOR CLIENTE':'MONTO VENTA CLIENTE',
          'MONEDA':'MONEDA 2','DIVISA':'MONEDA 2','MONEDA2':'MONEDA 2',
          'FECHA':'FECHA DIA/MES/AÑO','FECHA NEGOCIO':'FECHA DIA/MES/AÑO','FECHA VENTA':'FECHA DIA/MES/AÑO',
          'TRM':'TRM REFERENCIA','TRM REF':'TRM REFERENCIA',
          'LINEA':'LINEA DE PRODUCTO','LÍNEA':'LINEA DE PRODUCTO','LÍNEA DE PRODUCTO':'LINEA DE PRODUCTO',
          'COSTO TOTAL':'COSTO','VENDEDOR':'COMERCIAL','EJECUTIVO':'COMERCIAL',
          'MODIFICACION':'MODIFICACION O UPGRADE','MODIFICACIÓN':'MODIFICACION O UPGRADE',
        };
        const hdrs = rawHdrs.map(h => KEY_MAP[h.toUpperCase()] || KEY_MAP[h] || h);
        const recs = [];
        // Debug: log headers for first file
        if(!window._hdrDebugDone){ window._hdrDebugDone=true; console.log('[HDRS]', file.name, hdrs.filter(h=>h)); }
        for(let i=hdrIdx+1;i<raw.length;i++){
          const row=raw[i];
          if(!row[1]) continue;
          const rec={};
          hdrs.forEach((h,j)=>{if(h) rec[h]=row[j]!==undefined?row[j]:null;});
          // Normalizar: carpeta = fuente de verdad para director
          const toTitle = s => {
            if(!s) return '';
            return String(s).replace(/^[' ]+/,'').trim()
              .replace(/\w\S*/g,w=>w.charAt(0).toUpperCase()+w.slice(1).toLowerCase());
          };
          // CARPETA/ARCHIVO tiene prioridad sobre lo que escribió el ejecutivo
          rec['DIRECTOR'] = directorHint ? toTitle(directorHint) : toTitle(rec['DIRECTOR'] || rec['DIRECTOR '] || '');
          // Nombre del ejecutivo = nombre del archivo (sin extensión) — evita typos/abreviaturas
          const fileNameExec = toTitle(file.name.replace(/\.(xlsx|xls)$/i,'').trim());
          rec['COMERCIAL'] = fileNameExec || toTitle(rec['COMERCIAL'] || '');
          if(recs.length===0) console.log('[ROW1]', file.name, {fecha:rec['FECHA DIA/MES/AÑO'], monto:rec['MONTO VENTA CLIENTE'], moneda:rec['MONEDA 2'], cliente:rec['CLIENTE']});
          recs.push(rec);
        }
        // Extraer TRM de hoja Consulta1 si existe
        const consulta = wb.Sheets['Consulta1'];
        if(consulta){
          const cData = XLSX.utils.sheet_to_json(consulta,{header:1,defval:null});
          // Fila 1 = headers (Valor, Unidad...), fila 2 = valor
          if(cData[1] && cData[1][0]){
            const trmVal = parseFloat(cData[1][0]);
            if(trmVal > 100) window._lastTRM = trmVal;
          }
        }
        resolve(recs);
      }catch(err){console.warn('Error parseando',file.name,err);resolve([]);}
    };
    r.readAsBinaryString(file);
  });
}

// Extrae el nombre del director desde el path de la carpeta
// Ej: "FORECAST 2026/Grupo Juan David Novoa/Freddy.xlsx" → "Juan David Novoa"
function directorFromPath(path){
  const parts = path.split('/');
  // Buscar la parte que empiece con "Grupo" o "Gupo" (typo en sharepoint)
  const dirPart = parts.find(p=> /^(Grupo|Gupo)/i.test(p.trim()));
  if(dirPart){
    return dirPart.replace(/^(Grupo|Gupo)\s+/i,'').trim();
  }
  // Si no, usar el folder padre del archivo
  if(parts.length>=2) return parts[parts.length-2].trim();
  return '';
}

// Procesa lista de File objects (de folder o archivos sueltos)
let LOADED_FILES_BY_DIR = {}; // track all loaded files even if empty

async function processFiles(files){
  // Filtrar solo .xlsx/.xls
  const xlsFiles = files.filter(f=>{
    if(!/\.(xlsx|xls)$/i.test(f.name)) return false;
    if(f.name.startsWith('~$')) return false; // Excel temp files
    // Ignorar archivos que claramente no son ejecutivos (bases de datos, plantillas, etc.)
    const nameLower = f.name.toLowerCase();
    if(nameLower.includes('base de datos')) return false;
    if(nameLower.includes('plantilla')) return false;
    if(nameLower.includes('template')) return false;
    if(/^\d/.test(f.name)) return false; // empieza con número
    return true;
  });
  if(!xlsFiles.length){alert('No se encontraron archivos .xlsx de ejecutivos en la carpeta seleccionada');return;}
  
  ALL_DATA = [];
  
  // Mostrar overlay de progreso (funciona tanto si upload-zone está visible como oculto)
  let overlay = document.getElementById('load-overlay');
  if(!overlay){
    overlay = document.createElement('div');
    overlay.id = 'load-overlay';
    overlay.style.cssText = 'position:fixed;inset:0;z-index:9999;background:rgba(3,5,14,.88);backdrop-filter:blur(8px);display:flex;flex-direction:column;align-items:center;justify-content:center;gap:20px';
    overlay.innerHTML = `
      <div style="font-family:var(--font-display);font-size:13px;font-weight:700;letter-spacing:1px;text-transform:uppercase;color:var(--corp-cyan)">Cargando datos...</div>
      <div id="load-status" style="font-size:12px;color:var(--text3);font-family:var(--font-body);min-width:320px;text-align:center"></div>
      <div style="background:var(--border);border-radius:6px;height:6px;overflow:hidden;width:320px">
        <div id="load-bar" style="height:100%;background:linear-gradient(90deg,var(--corp-blue),var(--corp-cyan));width:0%;transition:width .25s;border-radius:6px"></div>
      </div>
      <div id="folder-structure" style="max-width:480px;text-align:left"></div>
    `;
    document.body.appendChild(overlay);
  } else {
    overlay.style.display = 'flex';
  }
  const progBar  = document.getElementById('load-bar');
  const progTxt  = document.getElementById('load-status');
  const fstruc   = document.getElementById('folder-structure');
  fstruc.innerHTML = '';
  
  // Agrupar por director según path
  const byDir = {};
  xlsFiles.forEach(f=>{
    const path = f.webkitRelativePath || f.name;
    const dir  = directorFromPath(path) || 'Sin Director';
    if(!byDir[dir]) byDir[dir]=[];
    byDir[dir].push(f);
  });
  LOADED_FILES_BY_DIR = byDir; // save for persona display
  
  // Procesar uno a uno con progreso
  let done=0;
  const total=xlsFiles.length;
  const dirColors={'0':'#2D7BF7','1':'#00C2D4','2':'#0DBF82','3':'#F0A020','4':'#E040A0'};
  
  const dirNames = Object.keys(byDir);
  for(let di=0;di<dirNames.length;di++){
    const dirName = dirNames[di];
    for(const f of byDir[dirName]){
      progTxt.textContent = `Leyendo: ${f.name} (${done+1}/${total})`;
      progBar.style.width = Math.round(((done)/total)*100)+'%';
      const recs = await parseXlsx(f, dirName);
      ALL_DATA.push(...recs);
      done++;
    }
  }
  progBar.style.width = '100%';
  progTxt.textContent = `✅ ${total} archivo${total!==1?'s':''} cargados — ${ALL_DATA.length} negocios`;
  
  // Mostrar estructura detectada
  let structHTML = '<div style="font-size:10px;font-family:var(--font-display);font-weight:700;letter-spacing:1px;text-transform:uppercase;color:var(--text2);margin-bottom:10px">Estructura detectada:</div>';
  dirNames.forEach((d,di)=>{
    const c = Object.values(dirColors)[di%5];
    const ejecs = [...new Set(byDir[d].map(f=>f.name.replace(/\.(xlsx|xls)$/i,'')))];
    structHTML += `<div style="margin-bottom:8px">
      <div style="display:flex;align-items:center;gap:6px;font-size:11px;color:${c};font-family:var(--font-display);font-weight:700">
        <span>📁</span> ${d} <span style="color:var(--text2);font-weight:400">(${ejecs.length} ejecutivo${ejecs.length!==1?'s':''})</span>
      </div>
      <div style="margin-left:20px;margin-top:4px;display:flex;flex-wrap:wrap;gap:6px">
        ${ejecs.map(e=>`<span style="font-size:10px;background:${c}15;border:1px solid ${c}30;border-radius:4px;padding:2px 8px;color:var(--text3)">📄 ${e}</span>`).join('')}
      </div>
    </div>`;
  });
  fstruc.innerHTML = structHTML;
  fstruc.style.display = 'block';
  
  // Ocultar overlay y mostrar dashboard
  setTimeout(()=>{
    const ov = document.getElementById('load-overlay');
    if(ov) ov.style.display='none';
    // Asegurarse que el dashboard esté visible
    document.getElementById('upload-zone-g').style.display='none';
    document.getElementById('gerencia-content').style.display='block';
    document.getElementById('hoy-strip').style.display='block';
    finalizeLoad();
  }, 1200);
}

// Evento: carpeta seleccionada (fallback manual)
document.getElementById('folder-input').addEventListener('change', function(e){
  const files = Array.from(e.target.files);
  if(!files.length) return;
  processFiles(files);
});

// Evento: archivos sueltos
document.getElementById('file-input').addEventListener('change', function(e){
  const files = Array.from(e.target.files);
  if(!files.length) return;
  processFiles(files);
});

// Drag & drop — acepta tanto carpeta como archivos
const uzDragDrop = document.getElementById('upload-zone-g');
if(uzDragDrop){
  uzDragDrop.addEventListener('dragover', e=>{e.preventDefault();uzDragDrop.classList.add('drag');});
  uzDragDrop.addEventListener('dragleave', ()=>uzDragDrop.classList.remove('drag'));
  uzDragDrop.addEventListener('drop', async e=>{
    e.preventDefault();
    uzDragDrop.classList.remove('drag');
    const items = Array.from(e.dataTransfer.items||[]);
    
    // Intentar leer como carpeta via FileSystemAPI
    if(items.length && items[0].webkitGetAsEntry){
      const allFiles = [];
      let pending = items.length;
      const readEntry = (entry, path='')=>{
        return new Promise(res=>{
          if(entry.isFile){
            entry.file(f=>{
              // Agregar path relativo manualmente
              Object.defineProperty(f,'webkitRelativePath',{value: path+'/'+f.name, writable:false});
              if(/\.(xlsx|xls)$/i.test(f.name)) allFiles.push(f);
              res();
            });
          } else if(entry.isDirectory){
            const reader = entry.createReader();
            reader.readEntries(async entries=>{
              for(const en of entries) await readEntry(en, path+'/'+entry.name);
              res();
            });
          } else res();
        });
      };
      for(const item of items){
        const entry = item.webkitGetAsEntry();
        if(entry) await readEntry(entry, entry.name);
      }
      if(allFiles.length) processFiles(allFiles);
      else {
        // Fallback a files
        const files = Array.from(e.dataTransfer.files);
        processFiles(files);
      }
    } else {
      const files = Array.from(e.dataTransfer.files);
      processFiles(files);
    }
  });
}

// TRM change
document.getElementById('trm-input').addEventListener('input',function(){
  TRM=getTRM();
  if(ALL_DATA.length) renderAll();
});

function finalizeLoad(){
  // Ocultar zona de carga, mostrar solo botón reload
  const uz = document.getElementById('upload-zone-g');
  if(uz) uz.style.display = 'none';
  const reloadBar = document.getElementById('reload-bar');
  if(reloadBar) reloadBar.style.display = 'flex';
  const reloadInfo = document.getElementById('reload-info');
  if(reloadInfo) {
    const nFiles = Object.values(LOADED_FILES_BY_DIR).reduce((s,a)=>s+a.length,0);
    const nDirs = Object.keys(LOADED_FILES_BY_DIR).length;
    reloadInfo.textContent = nFiles + ' archivos cargados · ' + nDirs + ' equipos · Última carga: ' + new Date().toLocaleTimeString('es-CO',{hour:'2-digit',minute:'2-digit'});
  }
  // Si los Excel traían TRM, aplicarla al input
  if(window._lastTRM && window._lastTRM > 100){
    document.getElementById('trm-input').value = window._lastTRM;
  }
  TRM=getTRM();
  // Asegurar visibilidad correcta sin importar el flujo que llamó
  const gc = document.getElementById('gerencia-content');
  if(gc) gc.style.display='block';
  const hs = document.getElementById('hoy-strip');
  if(hs) hs.style.display='block';

  // Status header (según visibilidad del rol)
  const visibleData = getVisibleData();
  var directorList=[...new Set(visibleData.map(r=>(r['DIRECTOR']||'').trim()).filter(Boolean))].sort();
  var executiveList=[...new Set(visibleData.map(r=>r['COMERCIAL']||'').filter(Boolean))].sort();
  var execsWithData = [...new Set(visibleData.map(r=>(r['COMERCIAL']||'').trim()).filter(Boolean))];
  document.getElementById('file-count-hd').textContent=
    visibleData.length+' negocios · '+directorList.length+' dir · '+execsWithData.length+' ejecutivos con datos';
  const directorList=[...new Set(visibleData.map(r=>(r['DIRECTOR']||'').trim()).filter(Boolean))].sort();
  const executiveList=[...new Set(visibleData.map(r=>r['COMERCIAL']||'').filter(Boolean))].sort();
  const execsWithData = [...new Set(visibleData.map(r=>(r['COMERCIAL']||'').trim()).filter(Boolean))];
  document.getElementById('file-count-hd').textContent=
    visibleData.length+' negocios · '+directorList.length+' dir · '+execsWithData.length+' ejecutivos con datos';
  const dirs=[...new Set(visibleData.map(r=>(r['DIRECTOR']||'').trim()).filter(Boolean))].sort();
  const execs=[...new Set(visibleData.map(r=>r['COMERCIAL']||'').filter(Boolean))].sort();
  const execsWithData = [...new Set(visibleData.map(r=>(r['COMERCIAL']||'').trim()).filter(Boolean))];
  document.getElementById('file-count-hd').textContent=
    visibleData.length+' negocios · '+dirs.length+' dir · '+execsWithData.length+' ejecutivos con datos';
  // Status header
  const dirs=[...new Set(ALL_DATA.map(r=>(r['DIRECTOR']||'').trim()).filter(Boolean))].sort();
  const execs=[...new Set(ALL_DATA.map(r=>r['COMERCIAL']||'').filter(Boolean))].sort();
  const execsWithData = [...new Set(ALL_DATA.map(r=>(r['COMERCIAL']||'').trim()).filter(Boolean))];
  document.getElementById('file-count-hd').textContent=
    ALL_DATA.length+' negocios · '+dirs.length+' dir · '+execsWithData.length+' ejecutivos con datos';
  const now = new Date();
  document.getElementById('last-update-hd').textContent=
    'Actualizado: '+now.toLocaleTimeString('es-CO',{hour:'2-digit',minute:'2-digit'})+' · '+
    now.toLocaleDateString('es-CO',{day:'2-digit',month:'short'});
  const dot = document.getElementById('status-dot-hd');
  if(dot){dot.style.background='var(--corp-green)';dot.style.boxShadow='0 0 8px var(--corp-green)';}
  const reloadBtn = document.getElementById('btn-reload');
  if(reloadBtn){
    reloadBtn.style.background='rgba(13,191,130,.15)';
    reloadBtn.style.border='1px solid rgba(13,191,130,.35)';
    reloadBtn.style.color='var(--corp-green)';
    const lbl=document.getElementById('btn-reload-txt');
    if(lbl) lbl.textContent='🔄 Recargar';
  }
  // Aplicar pestañas y badge según rol
  applyRoleTabs();
  showUserBadge();

  // Populate selects
  const selDir=document.getElementById('sel-director');
  const dirsForSel=[...new Set([
    ...visibleData.map(r=>(r['DIRECTOR']||'').trim()),
    ...ALL_DATA.map(r=>(r['DIRECTOR']||'').trim()),
    ...Object.keys(LOADED_FILES_BY_DIR||{}).map(d=>d.trim())
  ].filter(Boolean))].sort();
  selDir.innerHTML=dirsForSel.map(d=>`<option value="${d}">${d}</option>`).join('');
  
  const execsFromFiles2=Object.values(LOADED_FILES_BY_DIR||{}).flat()
    .map(f=>f.name.replace(/\.(xlsx|xls)$/i,'').trim()).filter(Boolean);
  const allExecsForSel=[...new Set([...executiveList,...execsFromFiles2])].sort();
  const allExecsForSel=[...new Set([...execs,...execsFromFiles2])].sort();
  const selEj=document.getElementById('sel-ejecutivo');
  selEj.innerHTML=allExecsForSel.map(e=>`<option value="${e}">${e}</option>`).join('');

  renderAll();
}

function getVisibleData() {
  if(!CURRENT_USER) return ALL_DATA;
  const { role, directorGroup, name } = CURRENT_USER;
  if(role === 'director') {
    return ALL_DATA.filter(r => (r['DIRECTOR']||'').trim() === (directorGroup||'').trim());
  }
  if(role === 'ejecutivo') {
    return ALL_DATA.filter(r => namesMatch(r['COMERCIAL'], name));
    return ALL_DATA.filter(r => (r['COMERCIAL']||'').trim().toLowerCase() === (name||'').trim().toLowerCase());
  }
  return ALL_DATA; // gerencia ve todo
}

function renderAll(){
  renderGerencia();
  renderDirector();
  renderEjecutivo();
  renderDivisas();
  renderMarcas();
  renderResumen();
}

/* ══════════════════════════════════════
   NAV
══════════════════════════════════════ */
function showPage(id,btn){
  document.querySelectorAll('.page').forEach(p=>p.classList.remove('active'));
  document.querySelectorAll('.nav-btn').forEach(b=>b.classList.remove('active'));
  document.getElementById('page-'+id).classList.add('active');
  if(btn) btn.classList.add('active');
}

/* ══════════════════════════════════════
   CHART HELPERS
══════════════════════════════════════ */
function renderBars(containerId, items, color, fmtFn){
  const el=document.getElementById(containerId);
  if(!el) return;
  const max=Math.max(...items.map(i=>i.val),1);
  el.innerHTML=items.map((it,idx)=>{
    const pct=Math.round((it.val/max)*100);
    const c=Array.isArray(color)?color[idx%color.length]:color;
    return `<div class="bar-row">
      <div class="bar-name w100" title="${it.name}">${it.name}</div>
      <div class="bar-track">
        <div class="bar-fill" style="width:${pct}%;background:${c}20;border:1px solid ${c}40">
          <span class="bar-val" style="color:${c}">${fmtFn?fmtFn(it.val):abr(it.val)}</span>
        </div>
      </div>
    </div>`;
  }).join('');
}

function renderDonut(svgId, legId, items){
  const svg=document.getElementById(svgId);
  const leg=document.getElementById(legId);
  if(!svg||!leg) return;
  const total=items.reduce((s,i)=>s+i.val,0)||1;
  let angle=-90;
  const r=38, cx=50, cy=50;
  let paths='';
  items.forEach((it,i)=>{
    const slice=(it.val/total)*360;
    const a1=angle*Math.PI/180;
    const a2=(angle+slice)*Math.PI/180;
    const x1=cx+r*Math.cos(a1), y1=cy+r*Math.sin(a1);
    const x2=cx+r*Math.cos(a2), y2=cy+r*Math.sin(a2);
    const large=slice>180?1:0;
    const c=COLORS[i%COLORS.length];
    paths+=`<path d="M${cx},${cy} L${x1},${y1} A${r},${r} 0 ${large},1 ${x2},${y2} Z" fill="${c}" opacity=".85"/>`;
    angle+=slice;
  });
  svg.innerHTML=paths+`<circle cx="50" cy="50" r="22" fill="#0B0F1E"/>`;
  
  leg.innerHTML=items.slice(0,6).map((it,i)=>{
    const pct=((it.val/total)*100).toFixed(1);
    return `<div class="leg-item" style="display:flex;align-items:center;gap:8px;font-size:12px"><div class="leg-dot" style="background:${COLORS[i%COLORS.length]};width:10px;height:10px;border-radius:50%;flex-shrink:0"></div><span title="${it.name}" style="flex:1;color:var(--text2)">${it.name.substring(0,18)}</span><span class="leg-pct" style="color:var(--text);font-family:var(--font-mono);font-weight:600">${pct}%</span></div>`;
  }).join('');
}

function renderEvoChart(containerId, dataByDir){
  const el=document.getElementById(containerId);
  if(!el) return;
  const months=['2026-01','2026-02','2026-03'];
  const dirs=[...new Set(Object.keys(dataByDir))];
  const nDirs=dirs.length;
  const allVals=dirs.flatMap(d=>months.map(m=>dataByDir[d][m]||0));
  const maxVal=Math.max(...allVals,1);
  const W=520,H=200,padL=56,padR=12,padT=20,padB=36;
  const gW=W-padL-padR, gH=H-padT-padB;
  const grpW=gW/months.length;
  const barW=Math.min(18, (grpW-8)/Math.max(nDirs,1));
  const gap=2;

  let svg=`<svg viewBox="0 0 ${W} ${H}" style="width:100%;overflow:visible">`;
  svg+=`<defs>${dirs.map((d,di)=>`<linearGradient id="bg${di}" x1="0" y1="0" x2="0" y2="1"><stop offset="0%" stop-color="${COLORS[di%COLORS.length]}" stop-opacity=".95"/><stop offset="100%" stop-color="${COLORS[di%COLORS.length]}" stop-opacity=".4"/></linearGradient>`).join('')}</defs>`;

  // Grid lines
  [0,.25,.5,.75,1].forEach(t=>{
    const y=padT+gH*(1-t);
    svg+=`<line x1="${padL}" y1="${y}" x2="${W-padR}" y2="${y}" stroke="#1A2240" stroke-width="${t===0?1.5:.7}"/>`;
    if(t>0) svg+=`<text x="${padL-5}" y="${y+3.5}" text-anchor="end" font-size="7.5" fill="#4A5A7A" font-family="IBM Plex Mono,monospace">${abr(maxVal*t)}</text>`;
  });

  // Bars
  months.forEach((m,mi)=>{
    const grpCenter=padL+(mi+0.5)*grpW;
    const totalBarW=(barW+gap)*nDirs-gap;
    dirs.forEach((d,di)=>{
      const v=dataByDir[d][m]||0;
      const bh=Math.max(v/maxVal*gH, v>0?2:0);
      const x=grpCenter - totalBarW/2 + di*(barW+gap);
      const y=padT+gH-bh;
      svg+=`<rect x="${x.toFixed(1)}" y="${y.toFixed(1)}" width="${barW}" height="${bh.toFixed(1)}" rx="2" fill="url(#bg${di})"><title>${dirs[di]}: ${abr(v)}</title></rect>`;
    });
    // Month label
    svg+=`<text x="${grpCenter.toFixed(1)}" y="${H-8}" text-anchor="middle" font-size="9" fill="#7A8AAA" font-family="IBM Plex Sans,sans-serif" font-weight="600">${MES_LABELS[m]||m}</text>`;
  });

  // Legend top
  let lx=padL;
  dirs.forEach((d,di)=>{
    const c=COLORS[di%COLORS.length];
    const short=d.split(' ').slice(0,2).join(' ');
    svg+=`<rect x="${lx}" y="4" width="9" height="9" rx="2" fill="${c}"/>`;
    svg+=`<text x="${lx+12}" y="12" font-size="8" fill="#8A9ACC" font-family="IBM Plex Sans,sans-serif">${short}</text>`;
    lx+=short.length*4.8+20;
  });

  svg+='</svg>';
  el.innerHTML=svg;
}

/* ══════════════════════════════════════
   GERENCIA GENERAL
══════════════════════════════════════ */
function renderGerencia(){
  const ALL_DATA = getVisibleData();
  if(!ALL_DATA.length) return;
  const trm=getTRM();
  const today=todayStr();
  
  // Hoy strip
  const hoy=ALL_DATA.filter(r=>{
    const f=r['FECHA DIA/MES/AÑO'];
    if(!f) return false;
    const d=f instanceof Date?f.toISOString().substring(0,10):(typeof f==='number'?new Date(Math.round((f-25569)*86400*1000)).toISOString().substring(0,10):String(f)).substring(0,10);
    return d===today;
  });
  document.getElementById('hoy-fecha').textContent=new Date().toLocaleDateString('es-CO',{weekday:'long',year:'numeric',month:'long',day:'numeric'});
  const hoyCOP=hoy.filter(r=>(r['MONEDA 2']||'').trim()==='COP').reduce((s,r)=>s+(parseMonto(r['MONTO VENTA CLIENTE'])||0),0);
  const hoyUSD=hoy.filter(r=>(r['MONEDA 2']||'').trim()==='USD').reduce((s,r)=>s+(parseMonto(r['MONTO VENTA CLIENTE'])||0),0);
  document.getElementById('hoy-cop').textContent=fmtCOP(hoyCOP);
  document.getElementById('hoy-usd').textContent=fmtCOP(hoyUSD*trm);
  document.getElementById('hoy-count').textContent=hoy.length;
  document.getElementById('hoy-trm').textContent='$ '+trm.toLocaleString('es-CO');
  
  // Total COP (all)
  const totalCOP=ALL_DATA.reduce((s,r)=>s+toCOP(r),0);
  const totalUSD=ALL_DATA.filter(r=>(r['MONEDA 2']||'').trim()==='USD').reduce((s,r)=>s+(parseMonto(r['MONTO VENTA CLIENTE'])||0),0);
  const totalCOPonly=ALL_DATA.filter(r=>(r['MONEDA 2']||'').trim()==='COP').reduce((s,r)=>s+(parseMonto(r['MONTO VENTA CLIENTE'])||0),0);
  const ganadas=ALL_DATA.filter(r=>r['ESTADO']==='GANADA');
  const totalGanada=ganadas.reduce((s,r)=>s+toCOP(r),0);
  
  document.getElementById('kpi-gerencia').innerHTML=`
    <div class="kpi" style="--ac:var(--corp-blue2)"><div class="kpi-accent"></div>
      <div class="kpi-label">Total Forecast COP</div>
      <div class="kpi-val">${abr(totalCOP)}</div>
      <div class="kpi-sub">${fmtCOP(totalCOP)}</div>
      <span class="kpi-badge cop">COP</span>
    </div>
    <div class="kpi" style="--ac:var(--usd-color)"><div class="kpi-accent"></div>
      <div class="kpi-label">Total USD</div>
      <div class="kpi-val">${fmtUSD(totalUSD)}</div>
      <div class="kpi-sub">Liq: ${abr(totalUSD*trm)}</div>
      <span class="kpi-badge usd">USD</span>
    </div>
    <div class="kpi" style="--ac:var(--corp-green)"><div class="kpi-accent"></div>
      <div class="kpi-label">Ventas Ganadas</div>
      <div class="kpi-val">${abr(totalGanada)}</div>
      <div class="kpi-sub">${ganadas.length} negocios</div>
    </div>
    <div class="kpi" style="--ac:var(--corp-amber)"><div class="kpi-accent"></div>
      <div class="kpi-label">Negocios Totales</div>
      <div class="kpi-val">${ALL_DATA.length}</div>
      <div class="kpi-sub">${[...new Set(ALL_DATA.map(r=>r['COMERCIAL']))].length} ejecutivos</div>
    </div>
    <div class="kpi" style="--ac:var(--corp-cyan)"><div class="kpi-accent"></div>
      <div class="kpi-label">TRM Día</div>
      <div class="kpi-val">$ ${trm.toLocaleString('es-CO')}</div>
      <div class="kpi-sub">COP por USD</div>
    </div>
  `;
  
  // Bar directores — todos los de LOADED_FILES_BY_DIR aunque no tengan datos
  const dirsFromData=[...new Set(ALL_DATA.map(r=>(r['DIRECTOR']||'').trim()).filter(Boolean))];
  const dirsFromFiles=Object.keys(LOADED_FILES_BY_DIR||{}).map(d=>d.trim()).filter(Boolean);
  const dirsUniq=[...new Set([...dirsFromData,...dirsFromFiles])].sort();
  const dirData=dirsUniq.map((d)=>({
    name:d, val:ALL_DATA.filter(r=>(r['DIRECTOR']||'').trim()===d).reduce((s,r)=>s+toCOP(r),0)
  })).sort((a,b)=>b.val-a.val);
  renderBars('bar-directores',dirData,COLORS);
  
  // Evo por director mensual
  const dirsForEvo=[...new Set([
    ...ALL_DATA.map(r=>(r['DIRECTOR']||'').trim()),
    ...Object.keys(LOADED_FILES_BY_DIR||{}).map(d=>d.trim())
  ].filter(Boolean))].sort();
  const evoByDir={};
  dirsForEvo.forEach(d=>{
    evoByDir[d]={};
    ['2026-01','2026-02','2026-03'].forEach(m=>{
      evoByDir[d][m]=ALL_DATA.filter(r=>(r['DIRECTOR']||'').trim()===d&&getMonth(r['FECHA DIA/MES/AÑO'])===m).reduce((s,r)=>s+toCOP(r),0);
    });
  });
  renderEvoChart('evo-dir-chart',evoByDir);
  
  // Bar ejecutivos
  const execs=[...new Set(ALL_DATA.map(r=>r['COMERCIAL']||'').filter(Boolean))];
  const ejData=execs.map(e=>({name:e.split(' ')[0],val:ALL_DATA.filter(r=>r['COMERCIAL']===e).reduce((s,r)=>s+toCOP(r),0)})).sort((a,b)=>b.val-a.val);
  renderBars('bar-ejecutivos',ejData,COLORS);
  
  // Donuts
  const estados=['GANADA','PENDIENTE','PEDIDA','APLAZADO'];
  const estData=estados.map(e=>({name:e,val:ALL_DATA.filter(r=>r['ESTADO']===e).reduce((s,r)=>s+toCOP(r),0)}));
  renderDonut('donut-estado','leg-estado',estData);
  
  const lineas=[...new Set(ALL_DATA.map(r=>r['LINEA DE PRODUCTO']||'').filter(Boolean))];
  const linData=lineas.map(l=>({name:l,val:ALL_DATA.filter(r=>r['LINEA DE PRODUCTO']===l).reduce((s,r)=>s+toCOP(r),0)})).sort((a,b)=>b.val-a.val);
  renderDonut('donut-linea','leg-linea',linData);

  // Tablas de estados
  renderGerenciaEstadoTables(ALL_DATA);
}

function renderGerenciaEstadoTables(data) {
  const el = document.getElementById('gerencia-estado-tables');
  if(!el) return;
  const estados = ['GANADA','PENDIENTE','PEDIDA','APLAZADO'];
  const colores = {'GANADA':'#0DBF82','PENDIENTE':'#F0A020','PEDIDA':'#2D4FD6','APLAZADO':'#E84040'};

  el.innerHTML = estados.map(estado => {
    const rows = data
      .filter(r => (r['ESTADO']||'').toUpperCase() === estado)
      .sort((a,b) => toCOP(b) - toCOP(a)); // ordenar por valor desc
    const total = rows.reduce((s,r) => s+toCOP(r), 0);

    if(!rows.length) return `
      <div style="background:var(--card);border:1px solid var(--border);border-left:3px solid ${colores[estado]};border-radius:12px;padding:16px">
        <div style="font-family:var(--font-display);font-size:10px;font-weight:700;letter-spacing:1px;color:${colores[estado]};margin-bottom:6px">${estado}</div>
        <div style="font-size:11px;color:var(--text3)">Sin registros</div>
      </div>`;

    const rows_html = rows.slice(0, 30).map(r => {
      const cliente  = (r['CLIENTE'] || r['NOMBRE CLIENTE'] || r['EMPRESA'] || '—').toString().trim();
      const comercial = (r['COMERCIAL']||'—').split(' ')[0];
      const valor    = toCOP(r);
      const linea    = (r['LINEA DE PRODUCTO']||'').trim();
      return `
        <tr style="border-top:1px solid var(--border)">
          <td style="padding:5px 8px;font-size:10px;color:var(--text);max-width:140px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap" title="${cliente}">${cliente}</td>
          <td style="padding:5px 8px;font-size:10px;color:var(--text2);white-space:nowrap">${comercial}</td>
          <td style="padding:5px 8px;font-size:10px;color:var(--text3);white-space:nowrap;max-width:80px;overflow:hidden;text-overflow:ellipsis" title="${linea}">${linea||'—'}</td>
          <td style="padding:5px 8px;font-size:10px;color:${colores[estado]};text-align:right;font-family:var(--font-mono);font-weight:600;white-space:nowrap">${abr(valor)}</td>
        </tr>`;
    }).join('');

    const more = rows.length > 30 ? `<tr><td colspan="4" style="padding:5px 8px;font-size:9px;color:var(--text3);text-align:center">+${rows.length-30} más...</td></tr>` : '';

    return `
      <div style="background:var(--card);border:1px solid var(--border);border-left:3px solid ${colores[estado]};border-radius:12px;overflow:hidden">
        <div style="padding:10px 14px;display:flex;justify-content:space-between;align-items:center;border-bottom:1px solid var(--border)">
          <div>
            <span style="font-family:var(--font-display);font-size:10px;font-weight:700;letter-spacing:1px;color:${colores[estado]}">${estado}</span>
            <span style="font-size:9px;color:var(--text3);margin-left:8px;font-family:var(--font-body)">${rows.length} negocio${rows.length!==1?'s':''}</span>
          </div>
          <span style="font-family:var(--font-mono);font-size:13px;font-weight:700;color:var(--text)">${abr(total)}</span>
        </div>
        <div style="overflow-y:auto;max-height:280px">
          <table style="width:100%;border-collapse:collapse">
            <thead style="position:sticky;top:0;z-index:1;background:var(--bg2)">
              <tr>
                <th style="padding:5px 8px;font-size:8.5px;font-family:var(--font-display);letter-spacing:.8px;color:var(--text3);text-align:left">EMPRESA</th>
                <th style="padding:5px 8px;font-size:8.5px;font-family:var(--font-display);letter-spacing:.8px;color:var(--text3);text-align:left">EJECUTIVO</th>
                <th style="padding:5px 8px;font-size:8.5px;font-family:var(--font-display);letter-spacing:.8px;color:var(--text3);text-align:left">LÍNEA</th>
                <th style="padding:5px 8px;font-size:8.5px;font-family:var(--font-display);letter-spacing:.8px;color:var(--text3);text-align:right">VALOR</th>
              </tr>
            </thead>
            <tbody>${rows_html}${more}</tbody>
            <tfoot style="background:var(--bg2);border-top:1px solid var(--border)">
              <tr>
                <td colspan="3" style="padding:6px 8px;font-size:9px;font-family:var(--font-display);color:var(--text2);font-weight:600">TOTAL ${rows.length} NEGOCIOS</td>
                <td style="padding:6px 8px;font-size:11px;text-align:right;font-family:var(--font-mono);color:${colores[estado]};font-weight:700">${abr(total)}</td>
              </tr>
            </tfoot>
          </table>
        </div>
      </div>`;
  }).join('');
}

/* ══════════════════════════════════════
   DIRECTOR
══════════════════════════════════════ */
function renderDirector(){
  const ALL_DATA = getVisibleData();
  if(!ALL_DATA.length) return;
  const dir=document.getElementById('sel-director').value;
  const mes=document.getElementById('sel-dir-mes').value;
  const est=document.getElementById('sel-dir-estado').value;
  const trm=getTRM();
  
  let data=ALL_DATA.filter(r=>(r['DIRECTOR']||'').trim()===dir);
  if(mes) data=data.filter(r=>getMonth(r['FECHA DIA/MES/AÑO'])===mes);
  if(est) data=data.filter(r=>r['ESTADO']===est);
  
  const totalCOP=data.reduce((s,r)=>s+toCOP(r),0);
  const totalUSD=data.filter(r=>(r['MONEDA 2']||'').trim()==='USD').reduce((s,r)=>s+(parseMonto(r['MONTO VENTA CLIENTE'])||0),0);
  const ganadas=data.filter(r=>r['ESTADO']==='GANADA');
  // Todos los ejecutivos de este director (con o sin datos)
  const execsWithDataDir=[...new Set(data.map(r=>(r['COMERCIAL']||'').trim()).filter(Boolean))];
  const execsFromFilesDir=(LOADED_FILES_BY_DIR[dir]||[]).map(f=>f.name.replace(/\.(xlsx|xls)$/i,'').trim()).filter(Boolean);
  const execs=[...new Set([...execsWithDataDir,...execsFromFilesDir])].sort();
  
  // Evolution del director
  const evoMonths={'2026-01':0,'2026-02':0,'2026-03':0};
  data.forEach(r=>{const m=getMonth(r['FECHA DIA/MES/AÑO']);if(evoMonths[m]!==undefined)evoMonths[m]+=toCOP(r);});
  
  const evoSVG=()=>{
    const months=Object.keys(evoMonths);
    const vals=Object.values(evoMonths);
    const maxV=Math.max(...vals,1);
    const W=400,H=140,padL=44,padR=10,padT=16,padB=28;
    const gW=W-padL-padR,gH=H-padT-padB;
    const bW=Math.min(36,(gW/months.length)-12);
    let s=`<svg viewBox="0 0 ${W} ${H}" style="width:100%;overflow:visible">`;
    s+=`<defs><linearGradient id="bg-dir" x1="0" y1="0" x2="0" y2="1"><stop offset="0%" stop-color="#2D7BF7" stop-opacity=".9"/><stop offset="100%" stop-color="#2D7BF7" stop-opacity=".35"/></linearGradient></defs>`;
    // Grid
    [0,.5,1].forEach(t=>{
      const y=padT+gH*(1-t);
      s+=`<line x1="${padL}" y1="${y}" x2="${W-padR}" y2="${y}" stroke="#1A2240" stroke-width="${t===0?1.5:.6}"/>`;
      if(t>0) s+=`<text x="${padL-4}" y="${y+3}" text-anchor="end" font-size="7.5" fill="#4A5A7A" font-family="IBM Plex Mono,monospace">${abr(maxV*t)}</text>`;
    });
    months.forEach((m,i)=>{
      const v=vals[i]||0;
      const bh=Math.max(v/maxV*gH, v>0?2:0);
      const cx=padL+(i+0.5)*(gW/months.length);
      const x=cx-bW/2;
      const y=padT+gH-bh;
      s+=`<rect x="${x.toFixed(1)}" y="${y.toFixed(1)}" width="${bW}" height="${bh.toFixed(1)}" rx="3" fill="url(#bg-dir)"><title>${MES_LABELS[m]||m}: ${abr(v)}</title></rect>`;
      s+=`<text x="${cx.toFixed(1)}" y="${H-6}" text-anchor="middle" font-size="9" fill="#7A8AAA" font-family="IBM Plex Sans,sans-serif" font-weight="600">${MES_LABELS[m]||m}</text>`;
    });
    s+='</svg>';
    return s;
  };
  
  // Equipo execs
  const ejCards=execs.map((e,i)=>{
    const ejData=data.filter(r=>(r['COMERCIAL']||'').trim()===e);
    const ejCOP=ejData.reduce((s,r)=>s+toCOP(r),0);
    const ejGan=ejData.filter(r=>r['ESTADO']==='GANADA').length;
    const hasD=ejData.length>0;
    const c=COLORS[i%COLORS.length];
    const ini=e.split(' ').slice(0,2).map(w=>w[0]).join('');
    return `<div class="kpi" style="--ac:${c};min-width:160px;opacity:${hasD?1:.55}">
      <div class="kpi-accent"></div>
      <div style="display:flex;align-items:center;gap:8px;margin-bottom:10px">
        <div style="width:30px;height:30px;border-radius:50%;background:${c}30;border:1px solid ${c}60;display:flex;align-items:center;justify-content:center;font-family:var(--font-display);font-size:11px;font-weight:700;color:${c}">${ini}</div>
        <div>
          <div style="font-size:11px;font-family:var(--font-display);font-weight:700;color:var(--text)">${e.split(' ')[0]}</div>
          <div style="font-size:9px;color:var(--text2)">${hasD?ejData.length+' negocios':'Sin datos aún'}</div>
        </div>
      </div>
      <div class="kpi-val">${hasD?abr(ejCOP):'—'}</div>
      <div class="kpi-sub">${hasD?ejGan+' ganadas':'Pendiente'}</div>
    </div>`;
  }).join('');
  
  document.getElementById('director-content').innerHTML=`
    <div class="section-hd"><h2>${dir}</h2><span class="section-tag">DIRECTOR</span></div>
    
    <div class="kpi-grid kpi-grid-4" style="margin-bottom:16px">
      <div class="kpi" style="--ac:var(--corp-blue2)"><div class="kpi-accent"></div>
        <div class="kpi-label">Total COP</div>
        <div class="kpi-val">${abr(totalCOP)}</div>
        <div class="kpi-sub">${fmtCOP(totalCOP)}</div>
      </div>
      <div class="kpi" style="--ac:var(--usd-color)"><div class="kpi-accent"></div>
        <div class="kpi-label">Total USD</div>
        <div class="kpi-val">${fmtUSD(totalUSD)}</div>
        <div class="kpi-sub">Liq: ${abr(totalUSD*trm)}</div>
      </div>
      <div class="kpi" style="--ac:var(--corp-green)"><div class="kpi-accent"></div>
        <div class="kpi-label">Ganadas</div>
        <div class="kpi-val">${ganadas.length}</div>
        <div class="kpi-sub">${abr(ganadas.reduce((s,r)=>s+toCOP(r),0))}</div>
      </div>
      <div class="kpi" style="--ac:var(--corp-amber)"><div class="kpi-accent"></div>
        <div class="kpi-label">Ejecutivos</div>
        <div class="kpi-val">${execs.length}</div>
        <div class="kpi-sub">${data.length} negocios · ${execsWithDataDir.length} activos</div>
      </div>
    </div>

    <div class="g2b">
      <div class="chart-card">
        <div class="chart-hd">Evolución Mensual — ${dir.split(' ')[0]}</div>
        ${evoSVG()}
      </div>
      <div class="chart-card">
        <div class="chart-hd">Estado Pipeline</div>
        <div class="donut-wrap">
          <svg id="donut-dir-est" viewBox="0 0 100 100" style="width:130px;height:130px;flex-shrink:0"></svg>
          <div class="donut-leg" id="leg-dir-est"></div>
        </div>
      </div>
    </div>

    <div class="section-hd"><h2>Equipo</h2><span class="section-tag">${execs.length} EJECUTIVOS</span></div>
    <div style="display:grid;grid-template-columns:repeat(auto-fill,minmax(180px,1fr));gap:12px;margin-bottom:18px">
      ${ejCards}
    </div>
    
    <div class="chart-card g1">
      <div class="chart-hd">Detalle de Negocios</div>
      ${buildTable(data.slice(0,30))}
    </div>
  `;
  
  // Donut estado director — usar datos sin filtro de estado para mostrar distribución real
  const dataForDonut=ALL_DATA.filter(r=>(r['DIRECTOR']||'').trim()===dir && (!mes||getMonth(r['FECHA DIA/MES/AÑO'])===mes));
  const estD=['GANADA','PENDIENTE','PEDIDA','APLAZADO'].map(e=>({name:e,val:dataForDonut.filter(r=>r['ESTADO']===e).reduce((s,r)=>s+toCOP(r),0)}));
  renderDonut('donut-dir-est','leg-dir-est',estD);
}

/* ══════════════════════════════════════
   EJECUTIVO
══════════════════════════════════════ */
function initials(name){return name.split(' ').slice(0,2).map(w=>w[0]).join('');}

function renderEjecutivo(){
  const ALL_DATA = getVisibleData();
  if(!ALL_DATA.length) return;
  const ej=document.getElementById('sel-ejecutivo').value;
  const mes=document.getElementById('sel-ej-mes').value;
  const est=document.getElementById('sel-ej-estado').value;
  const trm=getTRM();
  
  // Persona grid — todos los ejecutivos cargados, tengan o no datos
  const execsFromData = [...new Set(ALL_DATA.map(r=>(r['COMERCIAL']||'').trim()).filter(Boolean))];
  // Agregar ejecutivos de archivos cargados aunque estén vacíos
  const execsFromFiles = Object.values(LOADED_FILES_BY_DIR||{}).flat()
    .map(f=>f.name.replace(/\.(xlsx|xls)$/i,'').trim());
  const allExecs = [...new Set([...execsFromData, ...execsFromFiles])].sort();
  const execs = allExecs;
  document.getElementById('persona-grid').innerHTML=execs.map((e,i)=>{
    const ed=ALL_DATA.filter(r=>(r['COMERCIAL']||'').trim()===e);
    const cop=ed.reduce((s,r)=>s+toCOP(r),0);
    const gan=ed.filter(r=>r['ESTADO']==='GANADA').length;
    const pen=ed.filter(r=>r['ESTADO']==='PENDIENTE').length;
    const c=COLORS[i%COLORS.length];
    const dirFromData=ed[0]?ed[0]['DIRECTOR']||'':'';
    const dirFromFile=Object.entries(LOADED_FILES_BY_DIR||{}).find(([d,fs])=>fs.some(f=>f.name.replace(/\.(xlsx|xls)$/i,'').trim()===e));
    const dir=dirFromData||(dirFromFile?dirFromFile[0]:'—');
    const hasData=ed.length>0;
    const selected=e===ej?'selected':'';
    return `<div class="persona-card ${selected} ${hasData?'':'no-data'}" onclick="selectEjecutivo('${e}')">
      <div class="persona-avatar" style="background:${c}${hasData?'25':'10'};border:2px solid ${c}${hasData?'50':'20'};color:${hasData?c:'var(--text2)'}">${initials(e)}</div>
      <div class="persona-name" style="color:${hasData?'var(--text)':'var(--text2)'}">${e}</div>
      <div class="persona-role">${dir}</div>
      ${hasData
        ?`<div class="persona-stats">
        <div class="p-stat"><div class="p-stat-label">Total</div><div class="p-stat-val" style="color:${c};font-size:11px">${abr(cop)}</div></div>
        <div class="p-stat"><div class="p-stat-label">Negoc.</div><div class="p-stat-val">${ed.length}</div></div>
        <div class="p-stat"><div class="p-stat-label">Ganadas</div><div class="p-stat-val" style="color:var(--corp-green)">${gan}</div></div>
        <div class="p-stat"><div class="p-stat-label">Pend.</div><div class="p-stat-val" style="color:var(--corp-amber)">${pen}</div></div>
      </div>`
        :`<div style="font-size:9px;color:var(--text2);font-family:var(--font-display);margin-top:8px;padding:5px 8px;background:rgba(255,255,255,.03);border-radius:6px;letter-spacing:.5px">📋 Sin registros aún</div>`}
    </div>`;
  }).join('');
  
  if(!ej) return;
  
  let data=ALL_DATA.filter(r=>r['COMERCIAL']===ej);
  if(mes) data=data.filter(r=>getMonth(r['FECHA DIA/MES/AÑO'])===mes);
  if(est) data=data.filter(r=>r['ESTADO']===est);
  
  const totalCOP=data.reduce((s,r)=>s+toCOP(r),0);
  const totalUSD=data.filter(r=>(r['MONEDA 2']||'').trim()==='USD').reduce((s,r)=>s+(parseMonto(r['MONTO VENTA CLIENTE'])||0),0);
  const ganadas=data.filter(r=>r['ESTADO']==='GANADA');
  const ejColor=COLORS[execs.indexOf(ej)%COLORS.length];
  
  const lineas=[...new Set(data.map(r=>r['LINEA DE PRODUCTO']||'').filter(Boolean))];
  const linData=lineas.map(l=>({name:l,val:data.filter(r=>r['LINEA DE PRODUCTO']===l).reduce((s,r)=>s+toCOP(r),0)})).sort((a,b)=>b.val-a.val);
  
  document.getElementById('ejecutivo-content').innerHTML=`
    <div class="section-hd" style="margin-top:16px"><h2>${ej}</h2><span class="section-tag" style="background:${ejColor}20;color:${ejColor};border-color:${ejColor}40">EJECUTIVO</span></div>
    
    <div class="kpi-grid kpi-grid-4" style="margin-bottom:16px">
      <div class="kpi" style="--ac:${ejColor}"><div class="kpi-accent"></div>
        <div class="kpi-label">Total COP</div>
        <div class="kpi-val">${abr(totalCOP)}</div>
        <div class="kpi-sub">${fmtCOP(totalCOP)}</div>
      </div>
      <div class="kpi" style="--ac:var(--usd-color)"><div class="kpi-accent"></div>
        <div class="kpi-label">Total USD</div>
        <div class="kpi-val">${fmtUSD(totalUSD)}</div>
        <div class="kpi-sub">Liq: ${abr(totalUSD*trm)}</div>
      </div>
      <div class="kpi" style="--ac:var(--corp-green)"><div class="kpi-accent"></div>
        <div class="kpi-label">Ganadas</div>
        <div class="kpi-val">${ganadas.length}</div>
        <div class="kpi-sub">${abr(ganadas.reduce((s,r)=>s+toCOP(r),0))}</div>
      </div>
      <div class="kpi" style="--ac:var(--corp-amber)"><div class="kpi-accent"></div>
        <div class="kpi-label">Negocios</div>
        <div class="kpi-val">${data.length}</div>
        <div class="kpi-sub">Director: ${(data[0]||{})['DIRECTOR']||'—'}</div>
      </div>
    </div>
    
    <div class="g2">
      <div class="chart-card">
        <div class="chart-hd">Top Líneas — ${ej.split(' ')[0]}</div>
        <div class="bar-list" id="bar-ej-lineas"></div>
      </div>
      <div class="chart-card">
        <div class="chart-hd">Estado Negocios</div>
        <div class="donut-wrap">
          <svg id="donut-ej-est" viewBox="0 0 100 100" style="width:130px;height:130px;flex-shrink:0"></svg>
          <div class="donut-leg" id="leg-ej-est"></div>
        </div>
      </div>
    </div>
    
    <div class="chart-card g1">
      <div class="chart-hd">Detalle de Negocios — ${ej}</div>
      ${buildTable(data)}
    </div>
  `;
  
  renderBars('bar-ej-lineas',linData,COLORS);
  const estD=['GANADA','PENDIENTE','PEDIDA','APLAZADO'].map(e=>({name:e,val:data.filter(r=>r['ESTADO']===e).length}));
  renderDonut('donut-ej-est','leg-ej-est',estD);
}

function selectEjecutivo(name){
  document.getElementById('sel-ejecutivo').value=name;
  renderEjecutivo();
}

/* ══════════════════════════════════════
   DIVISAS
══════════════════════════════════════ */
function renderDivisas(){
  const ALL_DATA = getVisibleData();
  if(!ALL_DATA.length) return;
  const trm=getTRM();
  
  const usdData=ALL_DATA.filter(r=>(r['MONEDA 2']||'').trim()==='USD');
  const copData=ALL_DATA.filter(r=>(r['MONEDA 2']||'').trim()==='COP');
  
  const totalUSD=usdData.reduce((s,r)=>s+(parseMonto(r['MONTO VENTA CLIENTE'])||0),0);
  const totalCOP=copData.reduce((s,r)=>s+(parseMonto(r['MONTO VENTA CLIENTE'])||0),0);
  const usdLiqCOP=totalUSD*trm;
  const granTotal=totalCOP+usdLiqCOP;
  
  document.getElementById('divisas-cards').innerHTML=`
    <div class="divisa-card cop">
      <div class="divisa-label" style="color:var(--cop-color)"><span class="flag">🇨🇴</span> COP — Peso Colombiano</div>
      <div class="divisa-main" style="color:var(--cop-color)">${abr(totalCOP)}</div>
      <div class="divisa-sub">${fmtCOP(totalCOP)}</div>
      <div class="divisa-stats">
        <div><div class="d-stat-label">Negocios</div><div class="d-stat-val" style="color:var(--cop-color)">${copData.length}</div></div>
        <div><div class="d-stat-label">Ganadas</div><div class="d-stat-val" style="color:var(--corp-green)">${copData.filter(r=>r['ESTADO']==='GANADA').length}</div></div>
        <div><div class="d-stat-label">Prom. Negocio</div><div class="d-stat-val" style="color:var(--cop-color)">${abr(totalCOP/(copData.length||1))}</div></div>
      </div>
    </div>
    <div class="divisa-card usd">
      <div class="divisa-label" style="color:var(--usd-color)"><span class="flag">🇺🇸</span> USD — Dólar Americano</div>
      <div class="divisa-main" style="color:var(--usd-color)">${fmtUSD(totalUSD)}</div>
      <div class="divisa-sub">TRM ${trm.toLocaleString('es-CO')} → Liquidado ${abr(usdLiqCOP)}</div>
      <div class="divisa-stats">
        <div><div class="d-stat-label">Negocios</div><div class="d-stat-val" style="color:var(--usd-color)">${usdData.length}</div></div>
        <div><div class="d-stat-label">Ganadas</div><div class="d-stat-val" style="color:var(--corp-green)">${usdData.filter(r=>r['ESTADO']==='GANADA').length}</div></div>
        <div><div class="d-stat-label">En COP</div><div class="d-stat-val" style="color:var(--usd-color)">${abr(usdLiqCOP)}</div></div>
      </div>
    </div>
  `;
  
  // Table USD detail
  document.getElementById('tbl-usd').innerHTML=`<table>
    <thead><tr><th>Ejecutivo</th><th>Cliente</th><th>Producto</th><th>USD</th><th>TRM</th><th>COP Liquidado</th><th>Estado</th></tr></thead>
    <tbody>${usdData.map(r=>{
      const usd=parseMonto(r['MONTO VENTA CLIENTE'])||0;
      const trmR=parseFloat(r['TRM REFERENCIA'])||trm;
      const liq=usd*trmR;
      return `<tr>
        <td>${(r['COMERCIAL']||'').split(' ')[0]}</td>
        <td>${r['CLIENTE']||'—'}</td>
        <td style="max-width:120px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis">${r['PRODUCTO']||'—'}</td>
        <td class="td-mono td-usd">${fmtUSD(usd)}</td>
        <td class="td-mono" style="color:var(--corp-cyan)">${trmR.toLocaleString('es-CO')}</td>
        <td class="td-mono td-cop">${fmtCOP(liq)}</td>
        <td><span class="badge badge-${r['ESTADO']}">${r['ESTADO']||'—'}</span></td>
      </tr>`;
    }).join('')}</tbody>
  </table>`;
  
  // Table COP detail
  document.getElementById('tbl-cop').innerHTML=`<table>
    <thead><tr><th>Ejecutivo</th><th>Cliente</th><th>Producto</th><th>COP</th><th>Estado</th></tr></thead>
    <tbody>${copData.map(r=>{
      const cop=parseMonto(r['MONTO VENTA CLIENTE'])||0;
      return `<tr>
        <td>${(r['COMERCIAL']||'').split(' ')[0]}</td>
        <td>${r['CLIENTE']||'—'}</td>
        <td style="max-width:120px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis">${r['PRODUCTO']||'—'}</td>
        <td class="td-mono td-cop">${fmtCOP(cop)}</td>
        <td><span class="badge badge-${r['ESTADO']}">${r['ESTADO']||'—'}</span></td>
      </tr>`;
    }).join('')}</tbody>
  </table>`;
  
  // Tabla resumen consolidado
  const dirs=[...new Set(ALL_DATA.map(r=>(r['DIRECTOR']||'').trim()).filter(Boolean))];
  document.getElementById('tbl-resumen-divisas').innerHTML=`<table>
    <thead><tr><th>Director</th><th>Negos COP</th><th>Valor COP</th><th>Negos USD</th><th>Valor USD</th><th>Liq. USD→COP</th><th>TOTAL COP</th></tr></thead>
    <tbody>${dirs.map(d=>{
      const dd=ALL_DATA.filter(r=>(r['DIRECTOR']||'').trim()===d);
      const dc=dd.filter(r=>(r['MONEDA 2']||'').trim()==='COP');
      const du=dd.filter(r=>(r['MONEDA 2']||'').trim()==='USD');
      const vCOP=dc.reduce((s,r)=>s+(parseMonto(r['MONTO VENTA CLIENTE'])||0),0);
      const vUSD=du.reduce((s,r)=>s+(parseMonto(r['MONTO VENTA CLIENTE'])||0),0);
      const liq=vUSD*trm;
      return `<tr>
        <td style="font-family:var(--font-display);font-weight:700;color:var(--text)">${d}</td>
        <td class="td-mono">${dc.length}</td>
        <td class="td-mono td-cop">${fmtCOP(vCOP)}</td>
        <td class="td-mono">${du.length}</td>
        <td class="td-mono td-usd">${fmtUSD(vUSD)}</td>
        <td class="td-mono td-usd">${fmtCOP(liq)}</td>
        <td class="td-mono" style="color:#fff;font-weight:600">${fmtCOP(vCOP+liq)}</td>
      </tr>`;
    }).join('')}
    <tr style="border-top:1px solid var(--border2)">
      <td style="font-family:var(--font-display);font-weight:800;color:#fff">TOTAL GENERAL</td>
      <td class="td-mono">${copData.length}</td>
      <td class="td-mono td-cop" style="font-weight:700">${fmtCOP(totalCOP)}</td>
      <td class="td-mono">${usdData.length}</td>
      <td class="td-mono td-usd" style="font-weight:700">${fmtUSD(totalUSD)}</td>
      <td class="td-mono td-usd" style="font-weight:700">${fmtCOP(usdLiqCOP)}</td>
      <td class="td-mono" style="color:var(--corp-cyan);font-weight:800;font-size:13px">${fmtCOP(granTotal)}</td>
    </tr></tbody>
  </table>`;
}

/* ══════════════════════════════════════
   MARCAS
══════════════════════════════════════ */
function renderMarcas(){
  const ALL_DATA = getVisibleData();
  if(!ALL_DATA.length) return;
  const trm=getTRM();
  
  const marcas=[...new Set(ALL_DATA.map(r=>r['MARCA']||'').filter(Boolean))];
  const marcaData=marcas.map(m=>({name:m,val:ALL_DATA.filter(r=>r['MARCA']===m).reduce((s,r)=>s+toCOP(r),0)})).sort((a,b)=>b.val-a.val);
  renderBars('bar-marcas',marcaData,COLORS);
  renderDonut('donut-marca','leg-marca',marcaData);
  
  const lineas=[...new Set(ALL_DATA.map(r=>r['LINEA DE PRODUCTO']||'').filter(Boolean))];
  const linData=lineas.map(l=>({name:l,val:ALL_DATA.filter(r=>r['LINEA DE PRODUCTO']===l).reduce((s,r)=>s+toCOP(r),0)})).sort((a,b)=>b.val-a.val);
  renderBars('bar-lineas',linData,COLORS);
  renderDonut('donut-linea2','leg-linea2',linData);
  
  // Marca por ejecutivo
  const execs=[...new Set(ALL_DATA.map(r=>r['COMERCIAL']||'').filter(Boolean))];
  document.getElementById('tbl-marca-ej').innerHTML=`<table>
    <thead><tr><th>Ejecutivo</th>${marcas.map(m=>`<th>${m}</th>`).join('')}<th>Top Marca</th></tr></thead>
    <tbody>${execs.map(e=>{
      const ed=ALL_DATA.filter(r=>r['COMERCIAL']===e);
      const marcaCounts=marcas.map(m=>ed.filter(r=>r['MARCA']===m).length);
      const topIdx=marcaCounts.indexOf(Math.max(...marcaCounts));
      return `<tr>
        <td style="font-family:var(--font-display);font-weight:600;color:var(--text)">${e.split(' ')[0]}</td>
        ${marcaCounts.map((c,i)=>`<td class="td-mono" style="color:${c>0?COLORS[i%COLORS.length]:'var(--text2)'}">${c||'—'}</td>`).join('')}
        <td style="font-family:var(--font-display);font-weight:700;color:${COLORS[topIdx%COLORS.length]}">${marcas[topIdx]||'—'}</td>
      </tr>`;
    }).join('')}</tbody>
  </table>`;
  
  // Productos únicos
  const prods=[...new Set(ALL_DATA.map(r=>r['PRODUCTO']||'').filter(Boolean))];
  document.getElementById('tbl-productos').innerHTML=`<table>
    <thead><tr><th>Producto / Tipo Máquina</th><th>Marca</th><th>Línea</th><th>Moneda</th><th>Negocios</th><th>Valor Total COP</th></tr></thead>
    <tbody>${prods.map(p=>{
      const pd=ALL_DATA.filter(r=>r['PRODUCTO']===p);
      const tot=pd.reduce((s,r)=>s+toCOP(r),0);
      const marca=pd[0]?pd[0]['MARCA']||'—':'—';
      const linea=pd[0]?pd[0]['LINEA DE PRODUCTO']||'—':'—';
      const moneda=(r=>r['MONEDA 2']||'—')(pd[0]||{});
      return `<tr>
        <td style="color:var(--text)">${p}</td>
        <td class="td-mono" style="color:var(--corp-cyan)">${marca}</td>
        <td>${linea}</td>
        <td><span class="badge ${moneda==='USD'?'badge-PEDIDA':'badge-PENDIENTE'}">${moneda}</span></td>
        <td class="td-mono">${pd.length}</td>
        <td class="td-mono td-cop">${fmtCOP(tot)}</td>
      </tr>`;
    }).join('')}</tbody>
  </table>`;
}

/* ══════════════════════════════════════
   RESUMEN TOTAL
══════════════════════════════════════ */
function renderResumen(){
  const ALL_DATA = getVisibleData();
  if(!ALL_DATA.length) return;
  const trm=getTRM();
  
  const usdData=ALL_DATA.filter(r=>(r['MONEDA 2']||'').trim()==='USD');
  const copData=ALL_DATA.filter(r=>(r['MONEDA 2']||'').trim()==='COP');
  const totalUSD=usdData.reduce((s,r)=>s+(parseMonto(r['MONTO VENTA CLIENTE'])||0),0);
  const totalCOP=copData.reduce((s,r)=>s+(parseMonto(r['MONTO VENTA CLIENTE'])||0),0);
  const usdLiq=totalUSD*trm;
  
  document.getElementById('kpi-resumen-top').innerHTML=`
    <div class="kpi" style="--ac:var(--cop-color)"><div class="kpi-accent"></div>
      <div class="kpi-label">Total Puro COP</div>
      <div class="kpi-val">${abr(totalCOP)}</div>
      <div class="kpi-sub">${fmtCOP(totalCOP)}</div>
      <span class="kpi-badge cop">COP</span>
    </div>
    <div class="kpi" style="--ac:var(--usd-color)"><div class="kpi-accent"></div>
      <div class="kpi-label">Total USD Liquidado COP</div>
      <div class="kpi-val">${abr(usdLiq)}</div>
      <div class="kpi-sub">${fmtUSD(totalUSD)} × ${trm.toLocaleString('es-CO')}</div>
      <span class="kpi-badge usd">USD→COP</span>
    </div>
    <div class="kpi" style="--ac:var(--corp-cyan)"><div class="kpi-accent"></div>
      <div class="kpi-label">GRAN TOTAL COP</div>
      <div class="kpi-val">${abr(totalCOP+usdLiq)}</div>
      <div class="kpi-sub">${fmtCOP(totalCOP+usdLiq)}</div>
    </div>
  `;
  
  // Tabla solo COP por director/ejecutivo
  const dirs=[...new Set(ALL_DATA.map(r=>(r['DIRECTOR']||'').trim()).filter(Boolean))];
  const execs=[...new Set(ALL_DATA.map(r=>r['COMERCIAL']||'').filter(Boolean))];
  
  document.getElementById('tbl-total-cop').innerHTML=`<table>
    <thead><tr><th>Ejecutivo</th><th>Director</th><th>Negocios COP</th><th>Ene</th><th>Feb</th><th>Mar</th><th>Total COP</th></tr></thead>
    <tbody>${execs.map(e=>{
      const ed=copData.filter(r=>r['COMERCIAL']===e);
      const dir=ALL_DATA.find(r=>r['COMERCIAL']===e);
      const jan=ed.filter(r=>getMonth(r['FECHA DIA/MES/AÑO'])==='2026-01').reduce((s,r)=>s+(parseMonto(r['MONTO VENTA CLIENTE'])||0),0);
      const feb=ed.filter(r=>getMonth(r['FECHA DIA/MES/AÑO'])==='2026-02').reduce((s,r)=>s+(parseMonto(r['MONTO VENTA CLIENTE'])||0),0);
      const mar=ed.filter(r=>getMonth(r['FECHA DIA/MES/AÑO'])==='2026-03').reduce((s,r)=>s+(parseMonto(r['MONTO VENTA CLIENTE'])||0),0);
      const tot=jan+feb+mar;
      return `<tr>
        <td style="font-family:var(--font-display);font-weight:600;color:var(--text)">${e}</td>
        <td style="color:var(--text2)">${dir?dir['DIRECTOR']||'—':'—'}</td>
        <td class="td-mono">${ed.length}</td>
        <td class="td-mono td-cop">${jan>0?abr(jan):'—'}</td>
        <td class="td-mono td-cop">${feb>0?abr(feb):'—'}</td>
        <td class="td-mono td-cop">${mar>0?abr(mar):'—'}</td>
        <td class="td-mono td-cop" style="font-weight:700">${fmtCOP(tot)}</td>
      </tr>`;
    }).join('')}
    <tr style="border-top:1px solid var(--border2)">
      <td colspan="2" style="font-family:var(--font-display);font-weight:800;color:var(--cop-color)">SUBTOTAL COP</td>
      <td class="td-mono">${copData.length}</td>
      <td class="td-mono td-cop">${abr(copData.filter(r=>getMonth(r['FECHA DIA/MES/AÑO'])==='2026-01').reduce((s,r)=>s+(parseMonto(r['MONTO VENTA CLIENTE'])||0),0))}</td>
      <td class="td-mono td-cop">${abr(copData.filter(r=>getMonth(r['FECHA DIA/MES/AÑO'])==='2026-02').reduce((s,r)=>s+(parseMonto(r['MONTO VENTA CLIENTE'])||0),0))}</td>
      <td class="td-mono td-cop">${abr(copData.filter(r=>getMonth(r['FECHA DIA/MES/AÑO'])==='2026-03').reduce((s,r)=>s+(parseMonto(r['MONTO VENTA CLIENTE'])||0),0))}</td>
      <td class="td-mono td-cop" style="font-weight:800;font-size:13px">${fmtCOP(totalCOP)}</td>
    </tr></tbody>
  </table>`;
  
  // Tabla USD detail + liquidado
  document.getElementById('tbl-total-usd').innerHTML=`<table>
    <thead><tr><th>Ejecutivo</th><th>Director</th><th>Negocios USD</th><th>Ene USD</th><th>Feb USD</th><th>Mar USD</th><th>Total USD</th><th>TRM</th><th>Total Liquidado COP</th></tr></thead>
    <tbody>${execs.map(e=>{
      const ed=usdData.filter(r=>r['COMERCIAL']===e);
      if(!ed.length) return '';
      const dir=ALL_DATA.find(r=>r['COMERCIAL']===e);
      const jan=ed.filter(r=>getMonth(r['FECHA DIA/MES/AÑO'])==='2026-01').reduce((s,r)=>s+(parseMonto(r['MONTO VENTA CLIENTE'])||0),0);
      const feb=ed.filter(r=>getMonth(r['FECHA DIA/MES/AÑO'])==='2026-02').reduce((s,r)=>s+(parseMonto(r['MONTO VENTA CLIENTE'])||0),0);
      const mar=ed.filter(r=>getMonth(r['FECHA DIA/MES/AÑO'])==='2026-03').reduce((s,r)=>s+(parseMonto(r['MONTO VENTA CLIENTE'])||0),0);
      const tot=jan+feb+mar;
      return `<tr>
        <td style="font-family:var(--font-display);font-weight:600;color:var(--text)">${e}</td>
        <td style="color:var(--text2)">${dir?dir['DIRECTOR']||'—':'—'}</td>
        <td class="td-mono">${ed.length}</td>
        <td class="td-mono td-usd">${jan>0?fmtUSD(jan):'—'}</td>
        <td class="td-mono td-usd">${feb>0?fmtUSD(feb):'—'}</td>
        <td class="td-mono td-usd">${mar>0?fmtUSD(mar):'—'}</td>
        <td class="td-mono td-usd" style="font-weight:700">${fmtUSD(tot)}</td>
        <td class="td-mono" style="color:var(--corp-cyan)">${trm.toLocaleString('es-CO')}</td>
        <td class="td-mono td-cop" style="font-weight:700">${fmtCOP(tot*trm)}</td>
      </tr>`;
    }).join('')}
    <tr style="border-top:1px solid var(--border2)">
      <td colspan="2" style="font-family:var(--font-display);font-weight:800;color:var(--usd-color)">SUBTOTAL USD</td>
      <td class="td-mono">${usdData.length}</td>
      <td colspan="3"></td>
      <td class="td-mono td-usd" style="font-weight:800">${fmtUSD(totalUSD)}</td>
      <td class="td-mono" style="color:var(--corp-cyan)">${trm.toLocaleString('es-CO')}</td>
      <td class="td-mono td-usd" style="font-weight:800;font-size:13px">${fmtCOP(usdLiq)}</td>
    </tr></tbody>
  </table>`;
  
  // Consolidado final
  document.getElementById('tbl-consolidado').innerHTML=`<table>
    <thead><tr><th>Director</th><th>Ejecutivo</th><th>Total COP</th><th>Total USD</th><th>USD Liq. COP</th><th>TOTAL CONSOLIDADO</th></tr></thead>
    <tbody>${dirs.flatMap(d=>{
      const dejecs=[...new Set(ALL_DATA.filter(r=>(r['DIRECTOR']||'').trim()===d).map(r=>r['COMERCIAL']))];
      return dejecs.map((e,ei)=>{
        const ed=ALL_DATA.filter(r=>(r['DIRECTOR']||'').trim()===d&&r['COMERCIAL']===e);
        const eCOP=ed.filter(r=>(r['MONEDA 2']||'').trim()==='COP').reduce((s,r)=>s+(parseMonto(r['MONTO VENTA CLIENTE'])||0),0);
        const eUSD=ed.filter(r=>(r['MONEDA 2']||'').trim()==='USD').reduce((s,r)=>s+(parseMonto(r['MONTO VENTA CLIENTE'])||0),0);
        const total=eCOP+eUSD*trm;
        return `<tr>
          <td style="font-family:var(--font-display);font-weight:700;color:var(--text2)">${ei===0?d:''}</td>
          <td style="color:var(--text)">${e}</td>
          <td class="td-mono td-cop">${fmtCOP(eCOP)}</td>
          <td class="td-mono td-usd">${eUSD>0?fmtUSD(eUSD):'—'}</td>
          <td class="td-mono td-usd">${eUSD>0?fmtCOP(eUSD*trm):'—'}</td>
          <td class="td-mono" style="color:#fff;font-weight:700">${fmtCOP(total)}</td>
        </tr>`;
      });
    }).join('')}
    <tr style="border-top:2px solid var(--corp-blue2)">
      <td colspan="2" style="font-family:var(--font-display);font-weight:800;font-size:13px;color:var(--corp-cyan)">GRAN TOTAL</td>
      <td class="td-mono td-cop" style="font-weight:800">${fmtCOP(totalCOP)}</td>
      <td class="td-mono td-usd" style="font-weight:800">${fmtUSD(totalUSD)}</td>
      <td class="td-mono td-usd" style="font-weight:800">${fmtCOP(usdLiq)}</td>
      <td class="td-mono" style="color:var(--corp-cyan);font-weight:800;font-size:14px">${fmtCOP(totalCOP+usdLiq)}</td>
    </tr></tbody>
  </table>`;
}

/* ══════════════════════════════════════
   GENERIC TABLE
══════════════════════════════════════ */
function buildTable(data){
  const trm=getTRM();
  return `<table>
    <thead><tr><th>Fecha</th><th>Cliente</th><th>Producto</th><th>Marca</th><th>Línea</th><th>Moneda</th><th>Valor</th><th>COP Total</th><th>Margen</th><th>Estado</th></tr></thead>
    <tbody>${data.map(r=>{
      const mon=(r['MONEDA 2']||'COP').trim();
      const val=parseMonto(r['MONTO VENTA CLIENTE'])||0;
      const cop=toCOP(r);
      const fecha=r['FECHA DIA/MES/AÑO']?(()=>{const fd=r['FECHA DIA/MES/AÑO'];if(fd instanceof Date)return fd.toISOString().substring(0,10);if(typeof fd==='number')return new Date(Math.round((fd-25569)*86400*1000)).toISOString().substring(0,10);const s=String(fd);const meses={ene:'01',feb:'02',mar:'03',abr:'04',may:'05',jun:'06',jul:'07',ago:'08',sep:'09',oct:'10',nov:'11',dic:'12'};const m1=s.match(/(\d{2})[-/](\w{3})\.?[-/](\d{2,4})/i);if(m1){const mes=meses[m1[2].toLowerCase()]||'01';const anio=m1[3].length===2?'20'+m1[3]:m1[3];return anio+'-'+mes+'-'+m1[1].padStart(2,'0');}return s.substring(0,10);})():'—';
      return `<tr>
        <td class="td-mono" style="font-size:10px">${fecha}</td>
        <td style="color:var(--text)">${r['CLIENTE']||'—'}</td>
        <td style="max-width:130px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis">${r['PRODUCTO']||'—'}</td>
        <td style="color:var(--corp-cyan)">${r['MARCA']||'—'}</td>
        <td style="font-size:10px">${r['LINEA DE PRODUCTO']||'—'}</td>
        <td><span class="badge ${mon==='USD'?'badge-PEDIDA':'badge-PENDIENTE'}">${mon}</span></td>
        <td class="td-mono ${mon==='USD'?'td-usd':'td-cop'}">${mon==='USD'?fmtUSD(val):fmtCOP(val)}</td>
        <td class="td-mono td-cop">${fmtCOP(cop)}</td>
        <td class="td-mono" style="color:var(--corp-amber)">${fmtPct(parseMonto(r['MARGEN'])||0)}</td>
        <td><span class="badge badge-${r['ESTADO']||''}">${r['ESTADO']||'—'}</span></td>
      </tr>`;
    }).join('')}</tbody>
  </table>`;
}

/* ══════════════════════════════════════
   DEMO DATA (pre-load for testing)
══════════════════════════════════════ */



async function loadFolderFromSharePoint() {
  // Cargar MSAL
  await new Promise(r => loadMSAL(r));
  if(!initMsalApp()) {
    alert('No se pudo cargar autenticación'); return;
  }
  showLoadingOverlay('Conectando con Microsoft 365...');
  try {
    await spLogin();
  } catch(e) {
    hideLoadingOverlay();
    alert('Error de autenticación: ' + e.message);
    return;
  }
  updateLoadingStatus('Cargando archivos de SharePoint...');
  try {
    const siteId = await getSiteId();
    const { role, directorGroup } = CURRENT_USER;
    ALL_DATA = [];
    LOADED_FILES_BY_DIR = {};

    if(role === 'ejecutivo') {
      await loadEjecutivoFile(siteId);
    } else if(role === 'director') {
      const folder = directorGroup.includes('Miller') ? 'Gupo Miller Romero' : 'Grupo ' + directorGroup;
      await loadDirectorFolder(siteId, folder);
    } else {
      // Gerencia — todas las carpetas
      const folders = ['Grupo Juan David Novoa','Grupo Maria Angelica Caballero','Grupo Oscar Beltran','Gupo Miller Romero'];
      for(const f of folders) await loadDirectorFolder(siteId, f);
    }
    finalizeLoad();
    hideLoadingOverlay();
  } catch(e) {
    hideLoadingOverlay();
    console.error(e);
    alert('Error cargando datos: ' + e.message);
  }
}

let _siteId = null;
let _driveId = null;
async function getSiteId() {
  if(_siteId) return _siteId;
  const token = await getToken(['Files.Read.All']);
  const r = await fetch('https://graph.microsoft.com/v1.0/sites/provexpress.sharepoint.com:/sites/ProvexpressIntranet/comercial', {
    headers: { Authorization: 'Bearer ' + token }
  });
  const d = await r.json();
  if(!d.id) throw new Error('Site no encontrado: ' + JSON.stringify(d.error||d));
  _siteId = d.id;
  console.log('[SITE OK]', _siteId);
  // Get drive ID for "Documentos compartidos"
  const r2 = await fetch('https://graph.microsoft.com/v1.0/sites/' + _siteId + '/drives', {
    headers: { Authorization: 'Bearer ' + token }
  });
  const d2 = await r2.json();
  console.log('[DRIVES]', JSON.stringify(d2).slice(0,500));
  const drive = (d2.value||[]).find(dv =>
    dv.name === 'Documentos compartidos' || dv.name === 'Documents' || dv.name === 'Documentos'
  );
  if(drive) { _driveId = drive.id; console.log('[DRIVE OK]', drive.name, _driveId); }
  else { console.warn('[DRIVE] not found, using default'); }
  return _siteId;
}

async function loadDirectorFolder(siteId, folderName) {
  const token = await getToken(['Files.Read.All']);
  updateLoadingStatus('Cargando: ' + folderName + '...');
  const folderPath = 'COMERCIAL/FORECAST 2026/' + folderName;
  // Use specific drive if found, otherwise default drive
  const driveBase = _driveId
    ? 'https://graph.microsoft.com/v1.0/drives/' + _driveId + '/root:/' + encodeURIComponent(folderPath) + ':/children?$top=50'
    : 'https://graph.microsoft.com/v1.0/sites/' + siteId + '/drive/root:/' + encodeURIComponent(folderPath) + ':/children?$top=50';
  console.log('[FOLDER]', driveBase);
  const r = await fetch(driveBase, { headers: { Authorization: 'Bearer ' + token } });
  const d = await r.json();
  if(!d.value) { console.warn('Sin archivos en', folderName, d); return; }
  const dirName = folderName.replace(/^(Grupo|Gupo)\s+/i,'').trim();
  if(!LOADED_FILES_BY_DIR[dirName]) LOADED_FILES_BY_DIR[dirName] = [];
  for(const item of d.value) {
    if(!item.name.match(/\.xlsx?$/i)) continue;
    if(item.name.startsWith('~$')) continue;
    if(item.name.toLowerCase().includes('base de datos')) continue;
    updateLoadingStatus('Leyendo: ' + item.name);
    const recs = await loadSpFile(item, dirName);
    ALL_DATA.push(...recs);
    LOADED_FILES_BY_DIR[dirName].push({ name: item.name });
  }
}

async function loadEjecutivoFile(siteId) {
  const folders = ['Grupo Juan David Novoa','Grupo Maria Angelica Caballero','Grupo Oscar Beltran','Gupo Miller Romero'];
  for(const folder of folders) {
    const path = AZURE_CONFIG.driveBase + '/' + folder;
    const token = await getToken(['Files.Read.All']);
    try {
      const r = await fetch(
        _driveId ? 'https://graph.microsoft.com/v1.0/drives/' + _driveId + '/root:/' + encodeURIComponent('COMERCIAL/FORECAST 2026/' + folder) + ':/children?$top=50' : 'https://graph.microsoft.com/v1.0/sites/' + siteId + '/drive/root:/' + encodeURIComponent(path) + ':/children?$top=50',
        { headers: { Authorization: 'Bearer ' + token } }
      );
      const d = await r.json();
      if(!d.value) continue;
      const file = d.value.find(f => namesMatch(f.name, CURRENT_USER.name));
      const file = d.value.find(f => f.name.toLowerCase().replace(/\.xlsx?$/i,'').trim() === CURRENT_USER.name.toLowerCase().trim());
      if(file) {
        const dirName = folder.replace(/^(Grupo|Gupo)\s+/i,'').trim();
        if(!LOADED_FILES_BY_DIR[dirName]) LOADED_FILES_BY_DIR[dirName] = [];
        const recs = await loadSpFile(file, dirName);
        ALL_DATA.push(...recs);
        LOADED_FILES_BY_DIR[dirName].push({ name: file.name });
        return;
      }
    } catch(e) { continue; }
  }
}

async function loadSpFile(item, dirName) {
  const url = item['@microsoft.graph.downloadUrl'];
  if(!url) return [];
  try {
    const buf = await (await fetch(url)).arrayBuffer();
    const wb  = XLSX.read(buf, { type:'array', cellDates:true });
    const wsName = wb.SheetNames.find(s=>s.includes('Gerencia')||s.includes('Comercial'))||wb.SheetNames[0];
    const ws  = wb.Sheets[wsName];
    const raw = XLSX.utils.sheet_to_json(ws, { header:1, defval:null });
    let hdrIdx = -1;
    for(let i=0;i<raw.length;i++) { if(raw[i]&&raw[i].some(c=>c&&String(c).includes('CLIENTE'))){ hdrIdx=i; break; } }
    if(hdrIdx<0) return [];
    const KEY_MAP = {'VENTA CLIENTE':'MONTO VENTA CLIENTE','MONEDA':'MONEDA 2','FECHA':'FECHA DIA/MES/AÑO','TRM':'TRM REFERENCIA','LINEA':'LINEA DE PRODUCTO'};
    const hdrs = raw[hdrIdx].map(h=>h?KEY_MAP[String(h).trim()]||String(h).trim():'');
    const toTitle = s => s?String(s).replace(/^[' ]+/,'').trim().replace(/\w\S*/g,w=>w.charAt(0).toUpperCase()+w.slice(1).toLowerCase()):'';
    const fileExec = toTitle(item.name.replace(/\.xlsx?$/i,'').trim());
    const recs = [];
    for(let i=hdrIdx+1;i<raw.length;i++) {
      const row=raw[i]; if(!row[1]) continue;
      const rec={};
      hdrs.forEach((h,j)=>{ if(h) rec[h]=row[j]!==undefined?row[j]:null; });
      rec['DIRECTOR']  = toTitle(dirName);
      rec['COMERCIAL'] = fileExec;
      recs.push(rec);
    }
    const consulta = wb.Sheets['Consulta1'];
    if(consulta) {
      const cData = XLSX.utils.sheet_to_json(consulta,{header:1,defval:null});
      if(cData[1]&&cData[1][0]){ const t=parseFloat(cData[1][0]); if(t>100) window._lastTRM=t; }
    }
    return recs;
  } catch(e) { console.warn('Error leyendo', item.name, e); return []; }
}

// ── Tabs por rol ─────────────────────────────
function applyRoleTabs() {
  if(!CURRENT_USER) return;
  const { role } = CURRENT_USER;
  const tabs = {
    gerencia:  document.getElementById('tab-gerencia'),
    director:  document.getElementById('tab-director'),
    ejecutivo: document.getElementById('tab-ejecutivo'),
    divisas:   document.getElementById('tab-divisas'),
    marcas:    document.getElementById('tab-marcas'),
    resumen:   document.getElementById('tab-resumen'),
  };
  // Reset — mostrar todas
  Object.values(tabs).forEach(t => { if(t) t.style.display = ''; });

  if(role === 'ejecutivo') {
    tabs.gerencia && (tabs.gerencia.style.display = 'none');
    tabs.director && (tabs.director.style.display = 'none');
    tabs.divisas  && (tabs.divisas.style.display  = 'none');
    tabs.marcas   && (tabs.marcas.style.display   = 'none');
    tabs.resumen  && (tabs.resumen.style.display  = 'none');
    showPage('ejecutivo', tabs.ejecutivo);
  } else if(role === 'director') {
    tabs.gerencia && (tabs.gerencia.style.display = 'none');
    tabs.ejecutivo&& (tabs.ejecutivo.style.display= 'none');
    tabs.resumen  && (tabs.resumen.style.display  = 'none');
    showPage('director', tabs.director);
  } else {
    // gerencia / gerencia_director — ven todo
    showPage('gerencia', tabs.gerencia);
  }
}

// ── Auto-cargar al abrir desde SharePoint ────
window.addEventListener('DOMContentLoaded', () => {
  // Cargar TRM automáticamente
  fetchTRM();
  // Pre-cargar MSAL y auto-login siempre
  loadMSAL(() => {
    initMsalApp();
    // Auto-cargar siempre al abrir
    loadFolderFromSharePoint();
  });
});

// Exponer handlers usados por atributos inline (onclick/onchange) en index.html
window.loadFolderFromSharePoint = loadFolderFromSharePoint;
window.showPage = showPage;
window.renderDirector = renderDirector;
window.renderEjecutivo = renderEjecutivo;
window.selectEjecutivo = selectEjecutivo;
