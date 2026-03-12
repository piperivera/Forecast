// ══════════════════════════════════════
//   AUTENTICACIÓN MICROSOFT 365 + MSAL
// ══════════════════════════════════════

var CURRENT_USER = null;
var SP_TOKEN     = null;
var msalApp      = null;

var AZURE_CONFIG = {
  clientId:  '4a2b9726-2736-4f72-9e7e-c64cfdc80253',
  tenantId:  'e6805558-f5bb-444c-8af2-5f3a4d6dd3fc',
  redirectUri: 'https://piperivera.github.io/ForeCast/',
  siteUrl:   'https://provexpress.sharepoint.com/sites/ProvexpressIntranet',
  driveBase: 'Documentos compartidos/COMERCIAL/FORECAST 2026',
};

function initMsalApp() {
  if(typeof msal === 'undefined') return false;
  if(msalApp) return true;
  msalApp = new msal.PublicClientApplication({
    auth: {
      clientId:    AZURE_CONFIG.clientId,
      authority:   'https://login.microsoftonline.com/' + AZURE_CONFIG.tenantId,
      redirectUri: AZURE_CONFIG.redirectUri,
    },
    cache: { cacheLocation: 'sessionStorage' }
  });
  return true;
}

function loadMSAL(callback) {
  if(typeof msal !== 'undefined') { callback(); return; }
  const s = document.createElement('script');
  s.src = 'https://alcdn.msauth.net/browser/2.38.3/js/msal-browser.min.js';
  s.onload = () => callback();
  s.onerror = () => {
    const s2 = document.createElement('script');
    s2.src = 'https://cdn.jsdelivr.net/npm/@azure/msal-browser@2.38.3/lib/msal-browser.min.js';
    s2.onload = () => callback();
    s2.onerror = () => { console.error('[MSAL] Failed'); callback(); };
    document.head.appendChild(s2);
  };
  document.head.appendChild(s);
}

async function getToken(scopes) {
  const accounts = msalApp.getAllAccounts();
  if(!accounts.length) throw new Error('No session');
  try {
    const r = await msalApp.acquireTokenSilent({ scopes, account: accounts[0] });
    return r.accessToken;
  } catch {
    const r = await msalApp.acquireTokenPopup({ scopes });
    return r.accessToken;
  }
}

async function spLogin() {
  if(!initMsalApp()) throw new Error('MSAL no disponible');
  await msalApp.handleRedirectPromise();
  let account = msalApp.getAllAccounts()[0];
  if(!account) {
    await msalApp.loginPopup({ scopes: ['User.Read'] });
    account = msalApp.getAllAccounts()[0];
  }
  const token = await getToken(['User.Read']);
  const res   = await fetch('https://graph.microsoft.com/v1.0/me', {
    headers: { Authorization: 'Bearer ' + token }
  });
  const profile = await res.json();
  const email   = (profile.mail || profile.userPrincipalName || '').toLowerCase();
  const { role, directorGroup } = getUserRole(email);
  CURRENT_USER = { email, name: profile.displayName, role, directorGroup };
  SP_TOKEN = await getToken(['Files.Read.All']);
  sessionStorage.setItem('forecast_user', JSON.stringify(CURRENT_USER));
  console.log('[AUTH]', email, role);
  return true;
}

const ROLES = {
  gerencia: [
    'juannovoa@provexpress.com.co',
    'c.estrategica@provexpress.com.co',
    'especialista.preventa@provexpress.com.co',
    'preventa.software@provexpress.com.co',
  ],
  directores: {
    'juannovoa@provexpress.com.co':          'Juan David Novoa',
    'angelica.caballero@provexpress.com.co': 'Maria Angelica Caballero',
    'oscar.beltran@provexpress.com.co':      'Oscar Beltran',
    'miller.romero@provexpress.com.co':      'Miller Romero',
  }
};

// Mapea correo -> nombre de archivo Excel del ejecutivo
const EXECUTIVO_BY_EMAIL = {
  'dafne.ruiz@provexpress.com.co': 'Dafne Lizeth Ruiz',
  'diana.castro@provexpress.com.co': 'Diana Catalina Castro',
  'jessica.valencia@provexpress.com.co': 'Jessica Lorena Valencia',
  'jhonatan.acevedo@provexpress.com.co': 'Jhonatan Acevedo',
  'camilo.hernandez@provexpress.com.co': 'Jhonatan Camilo Hernández',
  'juan.velasquez@provexpress.com.co': 'Juan Camilo Velásquez',
  'astrid.jimenez@provexpress.com.co': 'Leidy Astrid Jiménez',
  'maria.briceno@provexpress.com.co': 'María Paola Briceño',
  'yeison.urrego@provexpress.com.co': 'Yeison Urrego',
  'alejandra.velasquez@provexpress.com.co': 'Alejandra Velásquez',
  'angela.torres@provexpress.com.co': 'Angela Torres',
  'cesar.cespedes@provexpress.com.co': 'César Cespedes',
  'fernando.quinonez@provexpress.com.co': 'Fernando Alberto Quiñonez',
  'jenny.gonzalez@provexpress.com.co': 'Jenny González',
  'johanna.jaime@provexpress.com.co': 'Johanna Jaime Murcia',
  'juan.martinez@provexpress.com.co': 'Juan David Martínez',
  'mariela.ramirez@provexpress.com.co': 'Mariela Ramírez',
  'rosa.mendoza@provexpress.com.co': 'Rosa María Mendoza',
  'wilson.sanchez@provexpress.com.co': 'Wilson Fernando Sánchez',
  'tatiana.parra@provexpress.com.co': 'Angie Tatiana Parra',
  'claudia.triana@provexpress.com.co': 'Claudia Patricia Triana',
  'dilma.cuesta@provexpress.com.co': 'Dilma Cuesta',
  'andres.pena@provexpress.com.co': 'Freddy Andrés Peña',
  'paola.garcia@provexpress.com.co': 'Gina Paola García',
  'javier.cortes@provexpress.com.co': 'Javier Antonio Cortés',
  'julieth.galindo@provexpress.com.co': 'Juliet Milena Galindo Fino',
  'karen.carrillo@provexpress.com.co': 'Karent Carrillo',
  'lington.linares@provexpress.com.co': 'Lington Linares',
  'maria.cruz@provexpress.com.co': 'María Eugenia Cruz',
  'mario.reyes@provexpress.com.co': 'Mario Reyes',
  'daniel.galindo@provexpress.com.co': 'Daniel Galindo Girón',
  'dayana.chala@provexpress.com.co': 'Dayana Chala',
  'angelica.alvarez@provexpress.com.co': 'María Angélica Alvarez',
  'rosmira.rojas@provexpress.com.co': 'Rosmira Rojas',
  'yovanny.herrera@provexpress.com.co': 'Yovanny Herrera',
  'andrea.vargas@provexpress.com.co': 'Yurany Andrea Vargas',
};

window.EXECUTIVO_BY_EMAIL = EXECUTIVO_BY_EMAIL;

function getUserRole(email) {
  const e = (email||'').toLowerCase().trim();
  const isGerencia = ROLES.gerencia.includes(e);
  const isDirector = e in ROLES.directores;
  const dirGroup   = ROLES.directores[e] || null;
  if(isGerencia && isDirector) return { role:'gerencia_director', directorGroup: dirGroup };
  if(isGerencia)  return { role:'gerencia',  directorGroup: null };
  if(isDirector)  return { role:'director',  directorGroup: dirGroup };
  return { role:'ejecutivo', directorGroup: null };
}

function showLoginScreen() {
  // Crear pantalla de login si no existe
  let loginDiv = document.getElementById('login-screen');
  if(!loginDiv) {
    loginDiv = document.createElement('div');
    loginDiv.id = 'login-screen';
    loginDiv.style.cssText = 'position:fixed;inset:0;z-index:99999;background:var(--bg);display:flex;align-items:center;justify-content:center;flex-direction:column;gap:0';
    loginDiv.innerHTML = `
      <div style="background:var(--card);border:1px solid var(--border);border-radius:16px;padding:40px 48px;max-width:420px;width:90%;box-shadow:0 24px 80px rgba(0,0,0,.5)">
        <div style="display:flex;align-items:center;gap:12px;margin-bottom:32px">
          <img src="data:image/webp;base64,${window._logoB64||''}" style="height:32px" onerror="this.style.display='none'">
          <div>
            <div style="font-family:var(--font-display);font-size:14px;font-weight:800;letter-spacing:1px;color:var(--text)">FORECAST 2026</div>
            <div style="font-size:10px;color:var(--text2);font-family:var(--font-display);letter-spacing:.5px">ÁREA COMERCIAL</div>
          </div>
        </div>
        <div style="font-size:13px;color:var(--text2);margin-bottom:8px;font-family:var(--font-display);letter-spacing:.5px">CORREO CORPORATIVO</div>
        <input id="login-email" type="email" placeholder="tu.nombre@provexpress.com.co"
          style="width:100%;box-sizing:border-box;padding:12px 16px;border-radius:8px;border:1px solid var(--border);background:var(--bg2);color:var(--text);font-size:13px;font-family:var(--font-body);outline:none;margin-bottom:8px"
          onkeydown="if(event.key==='Enter') doLogin()">
        <div id="login-error" style="font-size:11px;color:#ff6b6b;min-height:18px;margin-bottom:12px;font-family:var(--font-body)"></div>
        <button onclick="doLogin()" style="width:100%;padding:13px;border-radius:8px;border:none;background:var(--corp-blue2);color:#fff;font-family:var(--font-display);font-size:12px;font-weight:700;letter-spacing:1px;cursor:pointer;transition:opacity .2s"
          onmouseover="this.style.opacity='.85'" onmouseout="this.style.opacity='1'">
          INGRESAR →
        </button>
        <div style="margin-top:20px;padding-top:16px;border-top:1px solid var(--border);font-size:10px;color:var(--text3);font-family:var(--font-body);text-align:center">
          Acceso restringido · Solo personal Provexpress
        </div>
      </div>`;
    document.body.appendChild(loginDiv);
  }
  loginDiv.style.display = 'flex';
  setTimeout(()=>{ const inp=document.getElementById('login-email'); if(inp) inp.focus(); }, 100);
}

function doLogin() {
  const inp   = document.getElementById('login-email');
  const errEl = document.getElementById('login-error');
  const email = (inp ? inp.value : '').toLowerCase().trim();

  if(!email || !email.includes('@')) {
    if(errEl) errEl.textContent = 'Ingresa un correo válido';
    return;
  }
  if(!email.endsWith('@provexpress.com.co')) {
    if(errEl) errEl.textContent = 'Solo se permite correo @provexpress.com.co';
    return;
  }

  const { role, directorGroup } = getUserRole(email);
  // Inferir nombre del email: juan.novoa@ → "Juan Novoa"
  const namePart = email.split('@')[0].replace(/\./g,' ');
  const name = namePart.replace(/\w\S*/g,w=>w.charAt(0).toUpperCase()+w.slice(1));

  CURRENT_USER = { email, name, role, directorGroup };

  // Guardar sesión en sessionStorage
  sessionStorage.setItem('forecast_user', JSON.stringify(CURRENT_USER));

  // Ocultar login
  const loginDiv = document.getElementById('login-screen');
  if(loginDiv) loginDiv.style.display = 'none';

  // Aplicar vistas por rol y mostrar badge
  applyRoleTabs();
  showUserBadge();

  // Cambiar label del botón y mensaje de bienvenida
  const lbl = document.getElementById('btn-reload-txt');
  if(lbl) lbl.textContent = '📂 Cargar Carpeta';

  // Mostrar mensaje de bienvenida en upload zone
  const uzD = document.getElementById('upload-zone-g');
  if(uzD) {
    const roleMsg = {
      gerencia: 'Carga la carpeta ÁREA COMERCIAL - FORECAST 2026 para ver todos los datos',
      gerencia_director: 'Carga la carpeta ÁREA COMERCIAL - FORECAST 2026 para ver todos los datos',
      director: `Carga tu carpeta Grupo ${CURRENT_USER.directorGroup} para ver tu equipo`,
      ejecutivo: 'Carga tu archivo Excel para ver tu forecast'
    };
    const msgEl = uzD.querySelector('p');
    if(msgEl) msgEl.textContent = roleMsg[CURRENT_USER.role] || 'Selecciona la carpeta para continuar';
  }
}

function showUserBadge() {
  if(!CURRENT_USER) return;
  const badge = document.getElementById('user-badge');
  if(badge) badge.style.display = 'flex';
  const av = document.getElementById('user-avatar');
  if(av) av.textContent = CURRENT_USER.name.split(' ').slice(0,2).map(w=>w[0]).join('');
  const nm = document.getElementById('user-name');
  if(nm) nm.textContent = CURRENT_USER.name.split(' ')[0];
  const rb = document.getElementById('user-role-badge');
  const roleLabels = {gerencia:'Gerencia',gerencia_director:'Gerencia · Director',director:'Director',ejecutivo:'Ejecutivo'};
  if(rb) rb.textContent = roleLabels[CURRENT_USER.role]||CURRENT_USER.role;
  // Mostrar botón de cambio de vista solo para especialista.preventa
  const gearBtn = document.getElementById('view-switcher-btn');
  if(gearBtn && CURRENT_USER.email === 'especialista.preventa@provexpress.com.co') {
    gearBtn.style.display = 'block';
  }
}

function toggleViewPanel() {
  const panel = document.getElementById('view-panel');
  if(!panel) return;
  panel.style.display = panel.style.display === 'none' ? 'block' : 'none';
}

// Close view panel clicking outside
document.addEventListener('click', e => {
  const panel = document.getElementById('view-panel');
  const btn   = document.getElementById('view-switcher-btn');
  if(panel && btn && !panel.contains(e.target) && !btn.contains(e.target)) {
    panel.style.display = 'none';
  }
});

function switchView(role, directorGroup) {
  // Override CURRENT_USER view without changing real identity
  const prev = CURRENT_USER;
  CURRENT_USER = { ...prev, role, directorGroup: directorGroup||null };
  applyRoleTabs();
  renderAll();
  // Update role badge
  const rb = document.getElementById('user-role-badge');
  const roleLabels = {gerencia:'Gerencia',director:'Director',ejecutivo:'Ejecutivo'};
  if(rb) rb.textContent = (roleLabels[role]||role) + (directorGroup ? ' · '+directorGroup.split(' ')[0] : '') + ' ⚙';
  // Highlight active button
  document.querySelectorAll('.view-opt-btn').forEach(b => b.classList.remove('active'));
  event.target.classList.add('active');
  // Close panel
  const panel = document.getElementById('view-panel');
  if(panel) panel.style.display = 'none';
}

// Restaurar sesión si ya inició sesión antes
function restoreSession() {
  try {
    const saved = sessionStorage.getItem('forecast_user');
    if(saved) {
      CURRENT_USER = JSON.parse(saved);
      showUserBadge();
      applyRoleTabs();
      return true;
    }
  } catch(e) {}
  return false;
}

// ── overlay helpers ──────────────────────────
function showLoadingOverlay(msg) {
  let ov = document.getElementById('load-overlay');
  if(!ov){
    ov = document.createElement('div');
    ov.id = 'load-overlay';
    ov.style.cssText='position:fixed;inset:0;z-index:9999;background:rgba(3,5,14,.92);backdrop-filter:blur(8px);display:flex;flex-direction:column;align-items:center;justify-content:center;gap:16px';
    ov.innerHTML=`
      <div style="font-family:var(--font-display);font-size:13px;font-weight:700;letter-spacing:1px;text-transform:uppercase;color:var(--corp-cyan)">Conectando...</div>
      <div id="load-status" style="font-size:12px;color:var(--text3);font-family:var(--font-body);min-width:320px;text-align:center"></div>
      <div style="background:var(--border);border-radius:6px;height:4px;overflow:hidden;width:320px">
        <div id="load-bar" style="height:100%;background:linear-gradient(90deg,var(--corp-blue),var(--corp-cyan));width:100%;animation:pulse 1.5s ease-in-out infinite;border-radius:6px"></div>
      </div>`;
    document.body.appendChild(ov);
  }
  ov.style.display='flex';
  updateLoadingStatus(msg);
}
function updateLoadingStatus(msg){ const el=document.getElementById('load-status'); if(el) el.textContent=msg; }
function hideLoadingOverlay(){ const ov=document.getElementById('load-overlay'); if(ov) ov.style.display='none'; }
