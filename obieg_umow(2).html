<!DOCTYPE html>
<html lang="pl">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>ObiegUmów — System obiegu umów</title>
<script src="https://alcdn.msauth.net/browser/2.38.3/js/msal-browser.min.js"></script>
<script src="https://cdnjs.cloudflare.com/ajax/libs/mammoth/1.6.0/mammoth.browser.min.js"></script>
<script src="https://cdnjs.cloudflare.com/ajax/libs/jspdf/2.5.1/jspdf.umd.min.js"></script>
<script src="https://cdnjs.cloudflare.com/ajax/libs/html2canvas/1.4.1/html2canvas.min.js"></script>
<link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/@tabler/icons-webfont@latest/tabler-icons.min.css">
<style>
*{box-sizing:border-box;margin:0;padding:0}
body{font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',sans-serif;background:#f5f5f3;color:#111;font-size:14px}
:root{
  --bg-primary:#fff;--bg-secondary:#f5f5f3;--bg-tertiary:#eeece8;
  --bg-info:#e6f1fb;--bg-success:#eaf3de;--bg-warning:#faeeda;--bg-danger:#fcebeb;
  --text-primary:#111;--text-secondary:#666;--text-tertiary:#999;
  --text-info:#0c447c;--text-success:#27500a;--text-warning:#633806;--text-danger:#a32d2d;
  --border-light:rgba(0,0,0,.1);--border-mid:rgba(0,0,0,.18);--border-strong:rgba(0,0,0,.28);
  --radius-md:8px;--radius-lg:12px;
}
@media(prefers-color-scheme:dark){:root{
  --bg-primary:#1e1e1e;--bg-secondary:#252525;--bg-tertiary:#1a1a1a;
  --bg-info:#0c2a44;--bg-success:#1a2e0a;--bg-warning:#3a2408;--bg-danger:#3a1010;
  --text-primary:#e8e8e8;--text-secondary:#aaa;--text-tertiary:#666;
  --text-info:#85b7eb;--text-success:#97c459;--text-warning:#fac775;--text-danger:#f09595;
  --border-light:rgba(255,255,255,.08);--border-mid:rgba(255,255,255,.14);--border-strong:rgba(255,255,255,.22);
}}

/* ── LOGIN SCREEN ── */
#loginScreen{position:fixed;inset:0;background:var(--bg-tertiary);display:flex;align-items:center;justify-content:center;z-index:200}
.login-card{background:var(--bg-primary);border:0.5px solid var(--border-light);border-radius:var(--radius-lg);padding:40px 48px;width:380px;text-align:center;box-shadow:0 8px 32px rgba(0,0,0,.08)}
.login-logo{font-size:20px;font-weight:700;color:var(--text-primary);display:flex;align-items:center;justify-content:center;gap:8px;margin-bottom:8px}
.login-logo i{font-size:28px;color:#185FA5}
.login-subtitle{font-size:13px;color:var(--text-secondary);margin-bottom:32px}
.login-ms-btn{display:flex;align-items:center;justify-content:center;gap:10px;width:100%;padding:11px 20px;border-radius:var(--radius-md);border:0.5px solid var(--border-mid);background:var(--bg-primary);color:var(--text-primary);font-size:14px;font-weight:500;cursor:pointer;transition:background .15s}
.login-ms-btn:hover{background:var(--bg-secondary)}
.login-ms-btn svg{flex-shrink:0}
.login-note{font-size:11px;color:var(--text-tertiary);margin-top:16px;line-height:1.5}
.login-config-warn{background:var(--bg-warning);border-radius:var(--radius-md);padding:10px 14px;font-size:12px;color:var(--text-warning);margin-bottom:20px;text-align:left;line-height:1.5}
.login-config-warn code{background:rgba(0,0,0,.08);padding:1px 4px;border-radius:3px;font-size:11px}

/* ── APP SHELL ── */
.app{min-height:100vh;background:var(--bg-tertiary);display:none}
.topbar{background:var(--bg-primary);border-bottom:0.5px solid var(--border-light);padding:0 24px;display:flex;align-items:center;justify-content:space-between;height:52px;position:sticky;top:0;z-index:50}
.logo{font-size:15px;font-weight:700;color:var(--text-primary);display:flex;align-items:center;gap:8px}
.logo i{font-size:20px;color:#185FA5}
.topbar-right{display:flex;align-items:center;gap:10px}
.role-badge{font-size:11px;padding:3px 8px;border-radius:10px;font-weight:600;background:var(--bg-info);color:var(--text-info)}
.role-badge.prawny{background:#eeedfe;color:#3c3489}
.role-badge.podpisujacy{background:var(--bg-success);color:var(--text-success)}
.user-chip{display:flex;align-items:center;gap:6px;font-size:13px;color:var(--text-secondary);cursor:pointer;padding:4px 8px;border-radius:var(--radius-md);transition:background .15s}
.user-chip:hover{background:var(--bg-secondary)}
.avatar{width:28px;height:28px;border-radius:50%;background:var(--bg-info);display:flex;align-items:center;justify-content:center;font-size:11px;font-weight:700;color:var(--text-info);overflow:hidden;flex-shrink:0}
.avatar img{width:100%;height:100%;object-fit:cover}
.layout{display:flex;height:calc(100vh - 52px)}
.sidebar{width:210px;background:var(--bg-primary);border-right:0.5px solid var(--border-light);padding:16px 0;flex-shrink:0;overflow-y:auto}
.sidebar-label{font-size:11px;font-weight:600;color:var(--text-tertiary);text-transform:uppercase;letter-spacing:.6px;padding:0 20px;margin-bottom:4px}
.nav-item{display:flex;align-items:center;gap:8px;padding:7px 12px;margin:0 8px;border-radius:var(--radius-md);cursor:pointer;font-size:13px;color:var(--text-secondary);transition:background .15s}
.nav-item:hover{background:var(--bg-secondary)}
.nav-item.active{background:var(--bg-info);color:var(--text-info)}
.nav-item i{font-size:16px}
.badge-count{margin-left:auto;font-size:11px;background:var(--bg-danger);color:var(--text-danger);padding:1px 6px;border-radius:10px;font-weight:600}
.sidebar-section{margin-bottom:12px}
.main{flex:1;overflow-y:auto;padding:24px}
.page{display:none}.page.active{display:block}
.ph{margin-bottom:20px}
.ph h1{font-size:18px;font-weight:600;color:var(--text-primary)}
.ph p{font-size:13px;color:var(--text-secondary);margin-top:2px}
.btn{display:inline-flex;align-items:center;gap:6px;padding:7px 14px;border-radius:var(--radius-md);font-size:13px;font-weight:500;cursor:pointer;border:0.5px solid var(--border-mid);background:var(--bg-primary);color:var(--text-primary);transition:background .15s}
.btn:hover{background:var(--bg-secondary)}
.btn:disabled{opacity:.45;cursor:not-allowed}
.btn-primary{background:#185FA5;color:#fff;border-color:#185FA5}
.btn-primary:hover{background:#0C447C}
.btn-success{background:#3B6D11;color:#fff;border-color:#3B6D11}
.btn-success:hover{background:#27500A}
.btn-sm{padding:4px 10px;font-size:12px}
.card{background:var(--bg-primary);border:0.5px solid var(--border-light);border-radius:var(--radius-lg);padding:16px 20px;margin-bottom:12px}
.crow{background:var(--bg-primary);border:0.5px solid var(--border-light);border-radius:var(--radius-lg);padding:14px 18px;margin-bottom:8px;display:flex;align-items:center;gap:12px;cursor:pointer;transition:border-color .15s}
.crow:hover{border-color:var(--border-strong)}
.cicon{width:36px;height:36px;border-radius:var(--radius-md);background:var(--bg-info);display:flex;align-items:center;justify-content:center;flex-shrink:0}
.cicon i{font-size:18px;color:#185FA5}
.cinfo{flex:1;min-width:0}
.ctitle{font-size:13px;font-weight:500;color:var(--text-primary);white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
.cmeta{font-size:12px;color:var(--text-secondary);margin-top:1px}
.sp{font-size:11px;padding:3px 8px;border-radius:10px;font-weight:600;white-space:nowrap;flex-shrink:0}
.s-draft{background:var(--bg-info);color:var(--text-info)}
.s-legal{background:var(--bg-warning);color:var(--text-warning)}
.s-revision{background:#fbeaf0;color:#72243e}
.s-approval{background:var(--bg-success);color:var(--text-success)}
.s-signing{background:#eeedfe;color:#3c3489}
.s-signed{background:var(--bg-success);color:#173404}
.s-archived{background:var(--bg-secondary);color:var(--text-tertiary)}
.tl{display:flex;align-items:flex-start;margin-bottom:20px;overflow-x:auto;padding-bottom:4px}
.tls{display:flex;flex-direction:column;align-items:center;flex:1;min-width:72px;position:relative}
.tls:not(:last-child)::after{content:'';position:absolute;top:14px;left:50%;width:100%;height:1px;background:var(--border-light);z-index:0}
.tld{width:28px;height:28px;border-radius:50%;border:0.5px solid var(--border-mid);background:var(--bg-primary);display:flex;align-items:center;justify-content:center;font-size:14px;z-index:1;position:relative}
.tld.done{background:#185FA5;border-color:#185FA5;color:#fff}
.tld.active{background:var(--bg-info);border-color:#185FA5;color:#185FA5}
.tld i{font-size:13px}
.tll{font-size:10px;color:var(--text-secondary);margin-top:4px;text-align:center;line-height:1.3}
.tll.active{color:#185FA5;font-weight:600}
.dl{font-size:11px;font-weight:600;color:var(--text-tertiary);text-transform:uppercase;letter-spacing:.5px;margin-bottom:6px}
.fc{display:inline-flex;align-items:center;gap:6px;padding:6px 10px;border:0.5px solid var(--border-mid);border-radius:var(--radius-md);font-size:12px;color:var(--text-secondary);background:var(--bg-secondary);cursor:pointer;transition:background .15s;margin-right:4px;margin-bottom:4px}
.fc:hover{background:var(--bg-primary)}
.fc i{font-size:14px}
.fc.pdf-ready{border-color:#185FA5;color:#185FA5;background:var(--bg-info)}
.cb{background:var(--bg-secondary);border-radius:var(--radius-lg);padding:10px 14px;margin-bottom:8px}
.ca{font-size:12px;font-weight:600;color:var(--text-primary)}
.ct{font-size:11px;color:var(--text-tertiary);margin-left:6px}
.cx{font-size:13px;color:var(--text-secondary);margin-top:3px}
.uz{border:1px dashed var(--border-mid);border-radius:var(--radius-lg);padding:20px;text-align:center;cursor:pointer;transition:background .15s;position:relative}
.uz:hover{background:var(--bg-secondary)}
.uz i{font-size:26px;color:var(--text-tertiary);display:block;margin-bottom:6px}
.uz p{font-size:13px;color:var(--text-secondary)}
.uz span{font-size:11px;color:var(--text-tertiary)}
.uz input[type=file]{position:absolute;inset:0;opacity:0;cursor:pointer}
textarea{width:100%;border:0.5px solid var(--border-mid);border-radius:var(--radius-md);padding:10px 12px;font-size:13px;font-family:inherit;color:var(--text-primary);background:var(--bg-primary);resize:vertical;min-height:80px}
textarea:focus{outline:none;border-color:#185FA5}
input[type=text],select{width:100%;border:0.5px solid var(--border-mid);border-radius:var(--radius-md);padding:8px 12px;font-size:13px;font-family:inherit;color:var(--text-primary);background:var(--bg-primary)}
input[type=text]:focus,select:focus{outline:none;border-color:#185FA5}
.fr{margin-bottom:12px}
.fr label{display:block;font-size:12px;font-weight:600;color:var(--text-secondary);margin-bottom:4px}
.ab{display:flex;gap:8px;flex-wrap:wrap;padding-top:12px;border-top:0.5px solid var(--border-light);margin-top:16px}
.divider{height:1px;background:var(--border-light);margin:12px 0}
.ib{background:var(--bg-info);border-radius:var(--radius-md);padding:10px 14px;font-size:13px;color:var(--text-info);display:flex;gap:8px;align-items:flex-start;margin-bottom:12px}
.ib i{font-size:16px;flex-shrink:0;margin-top:1px}
.sb{background:var(--bg-success);border-radius:var(--radius-md);padding:10px 14px;font-size:13px;color:var(--text-success);display:flex;gap:8px;align-items:flex-start;margin-bottom:12px}
.sb i{font-size:16px;flex-shrink:0;margin-top:1px}
.wb{background:var(--bg-warning);border-radius:var(--radius-md);padding:10px 14px;font-size:13px;color:var(--text-warning);display:flex;gap:8px;align-items:flex-start;margin-bottom:12px}
.wb i{font-size:16px;flex-shrink:0;margin-top:1px}
.stat-grid{display:grid;grid-template-columns:repeat(3,1fr);gap:10px;margin-bottom:20px}
.sc{background:var(--bg-secondary);border-radius:var(--radius-md);padding:12px}
.sn{font-size:24px;font-weight:600;color:var(--text-primary)}
.sl{font-size:12px;color:var(--text-secondary);margin-top:2px}
.pdf-preview{border:0.5px solid var(--border-mid);border-radius:var(--radius-md);padding:16px;background:#fff;max-height:260px;overflow-y:auto;font-size:13px;line-height:1.6;color:#222}
.pdf-preview h1,.pdf-preview h2,.pdf-preview h3{color:#111;margin:8px 0 4px}
.pdf-preview p{margin-bottom:6px}
.pdf-preview table{border-collapse:collapse;width:100%;margin:8px 0}
.pdf-preview td,.pdf-preview th{border:1px solid #ccc;padding:4px 8px;font-size:12px}
.converting{display:flex;align-items:center;gap:8px;font-size:13px;color:var(--text-secondary);padding:8px 0}
.spin{width:16px;height:16px;border:2px solid var(--border-mid);border-top-color:#185FA5;border-radius:50%;animation:spin .7s linear infinite;flex-shrink:0}
@keyframes spin{to{transform:rotate(360deg)}}
/* ── SIGNER PICKER ── */
.signer-pick{display:flex;gap:10px;margin-top:6px}
.signer-opt{flex:1;border:0.5px solid var(--border-mid);border-radius:var(--radius-md);padding:10px 14px;cursor:pointer;transition:all .15s;text-align:center}
.signer-opt:hover{border-color:#185FA5;background:var(--bg-info)}
.signer-opt.selected{border-color:#185FA5;background:var(--bg-info);color:var(--text-info)}
.signer-opt .so-name{font-size:13px;font-weight:600;color:var(--text-primary);margin-bottom:2px}
.signer-opt .so-email{font-size:11px;color:var(--text-tertiary)}
.signer-opt.selected .so-name,.signer-opt.selected .so-email{color:var(--text-info)}
/* ── TOAST ── */
#toast-container{position:fixed;bottom:24px;right:24px;z-index:9999;display:flex;flex-direction:column;gap:8px;pointer-events:none}
.toast{background:var(--bg-primary);border:0.5px solid var(--border-mid);border-radius:var(--radius-lg);padding:12px 16px;box-shadow:0 4px 20px rgba(0,0,0,.12);display:flex;align-items:flex-start;gap:10px;min-width:280px;max-width:380px;pointer-events:all;animation:toastIn .25s ease;position:relative;overflow:hidden}
.toast::before{content:'';position:absolute;left:0;top:0;bottom:0;width:3px;background:#185FA5}
.toast-icon{font-size:18px;color:#185FA5;flex-shrink:0;margin-top:1px}
.toast-body{flex:1;min-width:0}
.toast-title{font-size:12px;font-weight:600;color:var(--text-primary);margin-bottom:2px}
.toast-to{font-size:11px;color:var(--text-secondary);white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
.toast-subject{font-size:11px;color:var(--text-tertiary);margin-top:2px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
.toast-close{font-size:14px;color:var(--text-tertiary);cursor:pointer;flex-shrink:0;padding:2px}
.toast-close:hover{color:var(--text-primary)}
.toast-progress{position:absolute;bottom:0;left:0;height:2px;background:#185FA5;animation:toastProgress 4.5s linear forwards}
@keyframes toastIn{from{opacity:0;transform:translateX(20px)}to{opacity:1;transform:translateX(0)}}
@keyframes toastProgress{from{width:100%}to{width:0%}}
.email-log-row{display:flex;align-items:flex-start;gap:10px;padding:10px 0;border-bottom:0.5px solid var(--border-light)}
.email-log-row:last-child{border-bottom:none}
.email-dot{width:8px;height:8px;border-radius:50%;background:#185FA5;flex-shrink:0;margin-top:5px}
.email-log-subject{font-size:13px;color:var(--text-primary);font-weight:500}
.email-log-to{font-size:12px;color:var(--text-secondary)}
.email-log-meta{font-size:11px;color:var(--text-tertiary);margin-top:2px}
/* ── LOGOUT MENU ── */
.user-menu{position:absolute;top:48px;right:16px;background:var(--bg-primary);border:0.5px solid var(--border-mid);border-radius:var(--radius-lg);padding:8px;box-shadow:0 4px 20px rgba(0,0,0,.12);z-index:100;min-width:200px;display:none}
.user-menu.open{display:block}
.user-menu-header{padding:8px 10px 10px;border-bottom:0.5px solid var(--border-light);margin-bottom:6px}
.user-menu-name{font-size:13px;font-weight:600;color:var(--text-primary)}
.user-menu-email{font-size:11px;color:var(--text-tertiary);margin-top:1px}
.user-menu-role{font-size:11px;margin-top:4px}
.user-menu-item{display:flex;align-items:center;gap:8px;padding:7px 10px;border-radius:var(--radius-md);cursor:pointer;font-size:13px;color:var(--text-secondary);transition:background .15s}
.user-menu-item:hover{background:var(--bg-secondary);color:var(--text-primary)}
.ti-brand-sharepoint{color:#038387}
</style>
</head>
<body>

<!-- LOGIN -->
<div id="loginScreen">
  <div class="login-card">
    <div class="login-logo"><i class="ti ti-file-certificate"></i>ObiegUmów</div>
    <p class="login-subtitle">System obiegu umów firmowych</p>

    <!-- Widok produkcyjny (Microsoft 365) -->
    <div id="loginProd">
      <button class="login-ms-btn" onclick="handleLogin()">
        <svg width="20" height="20" viewBox="0 0 21 21" xmlns="http://www.w3.org/2000/svg">
          <rect x="1" y="1" width="9" height="9" fill="#f25022"/>
          <rect x="11" y="1" width="9" height="9" fill="#7fba00"/>
          <rect x="1" y="11" width="9" height="9" fill="#00a4ef"/>
          <rect x="11" y="11" width="9" height="9" fill="#ffb900"/>
        </svg>
        Zaloguj się kontem Microsoft
      </button>
      <p class="login-note">Logowanie przez konto firmowe Microsoft 365.<br>Rola przypisywana automatycznie na podstawie adresu email.</p>
    </div>

    <!-- Widok demo (gdy brak konfiguracji Azure) -->
    <div id="loginDemo" style="display:none">
      <div class="login-config-warn">
        <strong>Tryb DEMO</strong> — brak konfiguracji Azure AD.<br>
        Wybierz użytkownika aby zobaczyć widok danej roli.
      </div>
      <div style="display:flex;flex-direction:column;gap:8px;margin-bottom:16px" id="demoUserList"></div>
      <p class="login-note" style="margin-top:0">IT: uzupełnij <code style="background:rgba(0,0,0,.08);padding:1px 4px;border-radius:3px">MSAL_CLIENT_ID</code> i <code style="background:rgba(0,0,0,.08);padding:1px 4px;border-radius:3px">MSAL_TENANT_ID</code> w kodzie aby włączyć prawdziwe logowanie.</p>
    </div>
  </div>
</div>

<!-- APP -->
<div class="app" id="appShell">
<div class="topbar" style="position:relative">
  <div class="logo"><i class="ti ti-file-certificate"></i>ObiegUmów</div>
  <div class="topbar-right">
    <span class="role-badge" id="roleBadgeEl"></span>
    <div class="user-chip" onclick="toggleUserMenu()">
      <div class="avatar" id="avatarEl">?</div>
      <span id="userNameEl">Ładowanie...</span>
      <i class="ti ti-chevron-down" style="font-size:12px"></i>
    </div>
  </div>
  <div class="user-menu" id="userMenu">
    <div class="user-menu-header">
      <div class="user-menu-name" id="menuName"></div>
      <div class="user-menu-email" id="menuEmail"></div>
      <div class="user-menu-role" id="menuRole"></div>
    </div>
    <div class="user-menu-item" onclick="handleLogout()"><i class="ti ti-logout"></i>Wyloguj się</div>
  </div>
</div>

<div class="layout">
<div class="sidebar">
  <div class="sidebar-section" style="margin-bottom:16px;padding-top:4px">
    <div class="sidebar-label" style="margin-bottom:6px">Menu</div>
    <div class="nav-item active" id="nav-dashboard" onclick="showPage('dashboard')"><i class="ti ti-layout-dashboard"></i>Pulpit</div>
    <div class="nav-item" id="nav-contracts" onclick="showPage('contracts')"><i class="ti ti-files"></i>Umowy<span class="badge-count" id="bc">0</span></div>
    <div class="nav-item" id="nav-new" onclick="showPage('new')"><i class="ti ti-plus"></i>Nowa umowa</div>
  </div>
  <div class="sidebar-section">
    <div class="sidebar-label" style="margin-bottom:6px">Akcje</div>
    <div class="nav-item" id="nav-review" onclick="showPage('review')"><i class="ti ti-message-circle"></i>Do przeglądu<span class="badge-count" id="br" style="display:none">0</span></div>
    <div class="nav-item" id="nav-sign" onclick="showPage('sign')"><i class="ti ti-pen"></i>Do podpisu<span class="badge-count" id="bs" style="display:none">0</span></div>
    <div class="nav-item" id="nav-archive" onclick="showPage('archive')"><i class="ti ti-archive"></i>Archiwum</div>
    <div class="nav-item" id="nav-emails" onclick="showPage('emails')"><i class="ti ti-mail"></i>Log maili<span class="badge-count" id="be" style="display:none">0</span></div>
  </div>
</div>

<div class="main">

  <div class="page active" id="page-dashboard">
    <div class="ph"><h1 id="dg">Dzień dobry!</h1><p id="ds">Podsumowanie obiegu umów.</p></div>
    <div class="stat-grid">
      <div class="sc"><div class="sn" id="st">0</div><div class="sl">Wszystkich umów</div></div>
      <div class="sc"><div class="sn" id="sa">0</div><div class="sl">W toku</div></div>
      <div class="sc"><div class="sn" id="ss">0</div><div class="sl">Podpisanych</div></div>
    </div>
    <div id="dri"></div>
    <div class="ph" style="margin-bottom:12px;margin-top:4px"><h1 style="font-size:15px">Ostatnia aktywność</h1></div>
    <div id="rc"></div>
  </div>

  <div class="page" id="page-contracts">
    <div class="ph" style="display:flex;justify-content:space-between;align-items:flex-start">
      <div><h1>Umowy</h1><p>Wszystkie umowy w systemie</p></div>
      <button class="btn btn-primary" id="newContractBtn" onclick="showPage('new')" style="display:none"><i class="ti ti-plus"></i>Nowa umowa</button>
    </div>
    <div id="cl"></div>
  </div>

  <div class="page" id="page-new">
    <div class="ph"><h1>Nowa umowa</h1><p>Wyślij szkic lub wytyczne do działu prawnego</p></div>
    <div class="card" id="newFormCard">
      <div class="fr"><label>Tytuł / opis umowy</label><input type="text" id="nt" placeholder="np. Umowa z dostawcą IT — firma XYZ"/></div>
      <div class="fr"><label>Typ umowy</label>
        <select id="nty"><option>Umowa o współpracy</option><option>Umowa NDA</option><option>Umowa licencyjna</option><option>Umowa zakupu</option><option>Umowa o świadczenie usług</option><option>Inny</option></select>
      </div>
      <div class="fr"><label>Wytyczne dla prawników</label><textarea id="nd" placeholder="Opisz zakres, kluczowe warunki, termin, wymagania specjalne..."></textarea></div>
      <div class="fr"><label>Plik DOCX — szkic umowy</label>
        <div class="uz" id="docxZone"><i class="ti ti-file-type-doc"></i><p>Kliknij lub przeciągnij plik DOCX</p><span id="docxLbl">Brak pliku</span><input type="file" accept=".docx" onchange="handleDocx(this)"></div>
      </div>
      <div class="fr"><label>Plik XLSX — dane handlowe (opcjonalnie)</label>
        <div class="uz" id="xlsxZone"><i class="ti ti-file-spreadsheet"></i><p>Kliknij lub przeciągnij plik XLSX</p><span id="xlsxLbl">Brak pliku</span><input type="file" accept=".xlsx,.xls" onchange="handleXlsx(this)"></div>
      </div>
      <div class="ab">
        <button class="btn btn-primary" onclick="submitContract()"><i class="ti ti-send"></i>Wyślij do działu prawnego</button>
        <button class="btn" onclick="showPage('dashboard')">Anuluj</button>
      </div>
    </div>
  </div>

  <div class="page" id="page-review">
    <div class="ph"><h1 id="rt">Do przeglądu</h1><p id="rsub"></p></div>
    <div id="rl"></div>
  </div>

  <div class="page" id="page-sign">
    <div class="ph"><h1>Do podpisu</h1><p>Umowy przypisane do Ciebie</p></div>
    <div id="sl2"></div>
  </div>

  <div class="page" id="page-archive">
    <div class="ph"><h1>Archiwum</h1><p>Podpisane i zarchiwizowane umowy</p></div>
    <div id="al"></div>
  </div>

  <div class="page" id="page-emails">
    <div class="ph"><h1>Log powiadomień email</h1><p>Historia wszystkich wysłanych powiadomień z systemu</p></div>
    <div class="card">
      <div class="wb" style="margin-bottom:16px">
        <i class="ti ti-tool"></i>
        <div><strong>TODO dla IT:</strong> W funkcji <code style="background:rgba(0,0,0,.08);padding:1px 5px;border-radius:4px;font-size:12px">sendEmailNotification()</code> zastąp blok <code style="background:rgba(0,0,0,.08);padding:1px 5px;border-radius:4px;font-size:12px">// TODO: BACKEND</code> wywołaniem <code style="background:rgba(0,0,0,.08);padding:1px 5px;border-radius:4px;font-size:12px">POST /api/send-email</code> — analogicznie do istniejącego systemu potwierdzeń terminów. Payload jest już gotowy.</div>
      </div>
      <div id="emailLogList"><p style="font-size:13px;color:var(--text-tertiary)">Brak wysłanych powiadomień w tej sesji.</p></div>
    </div>
  </div>

  <div class="page" id="page-detail">
    <div style="display:flex;align-items:center;gap:8px;margin-bottom:16px">
      <button class="btn btn-sm" onclick="goBack()"><i class="ti ti-arrow-left"></i>Wróć</button>
      <span id="dsp" class="sp"></span>
    </div>
    <div id="dc"></div>
  </div>

</div>
</div>
</div>

<div id="pdfRenderArea" style="position:fixed;left:-9999px;top:0;width:794px;padding:48px;background:#fff;font-family:Arial,sans-serif;font-size:13px;line-height:1.7;color:#111"></div>
<div id="toast-container"></div>

<script>
/* ═══════════════════════════════════════════════════════════════
   KONFIGURACJA — IT UZUPEŁNIA TE DWIE WARTOŚCI
   ═══════════════════════════════════════════════════════════════
   1. Wejdź na portal.azure.com → Microsoft Entra ID
      → Rejestracje aplikacji → Nowa rejestracja
   2. Skopiuj Client ID i Tenant ID po rejestracji
   3. W "Redirect URI" wpisz adres tej aplikacji
   ═══════════════════════════════════════════════════════════════ */
const MSAL_CLIENT_ID = 'TUTAJ_WKLEJ_CLIENT_ID';   // ← IT: uzupełnić
const MSAL_TENANT_ID = 'TUTAJ_WKLEJ_TENANT_ID';   // ← IT: uzupełnić

/* ═══════════════════════════════════════════════════════════════
   ROLE — przypisanie na podstawie adresu email
   Biznes = wszyscy pozostali pracownicy (rola domyślna)
   ═══════════════════════════════════════════════════════════════ */
const ROLE_CONFIG = {
  prawny: [
    'aurelia.kowalska@firma.pl',   // ← IT: zmienić na prawdziwe adresy
    'ania.nowak@firma.pl'
  ],
  podpisujacy: [
    { name: 'Jan Kowalski (Prezes)',  email: 'prezes@firma.pl' },  // ← IT: zmienić
    { name: 'Magda Wiśniewska',       email: 'magda.wisniewska@firma.pl' }
  ]
};

/* ─── MSAL setup ─── */
const msalConfigured = MSAL_CLIENT_ID !== 'TUTAJ_WKLEJ_CLIENT_ID';
let msalInstance = null;
if (msalConfigured) {
  msalInstance = new msal.PublicClientApplication({
    auth: {
      clientId: MSAL_CLIENT_ID,
      authority: `https://login.microsoftonline.com/${MSAL_TENANT_ID}`,
      redirectUri: window.location.origin + window.location.pathname
    },
    cache: { cacheLocation: 'sessionStorage' }
  });
}

/* ─── Bieżący użytkownik ─── */
let currentUser = null; // { name, email, initials, role, photo }

function detectRole(email) {
  const e = (email || '').toLowerCase();
  if (ROLE_CONFIG.prawny.map(x=>x.toLowerCase()).includes(e)) return 'prawny';
  if (ROLE_CONFIG.podpisujacy.map(x=>x.email.toLowerCase()).includes(e)) return 'podpisujacy';
  return 'biznes';
}

function initials(name) {
  return name.split(' ').slice(0,2).map(w=>w[0]).join('').toUpperCase();
}

function applyUser(name, email, photoUrl) {
  const role = detectRole(email);
  currentUser = { name, email, initials: initials(name), role, photo: photoUrl || null };
  document.getElementById('userNameEl').textContent = name.split(' ')[0];
  document.getElementById('avatarEl').innerHTML = photoUrl
    ? `<img src="${photoUrl}" alt="${name}">`
    : currentUser.initials;
  const rb = document.getElementById('roleBadgeEl');
  const roleLabels = { biznes: 'Biznes', prawny: 'Dział prawny', podpisujacy: 'Podpisujący' };
  rb.textContent = roleLabels[role];
  rb.className = 'role-badge ' + (role !== 'biznes' ? role : '');
  document.getElementById('menuName').textContent = name;
  document.getElementById('menuEmail').textContent = email;
  document.getElementById('menuRole').innerHTML = `<span class="role-badge ${role !== 'biznes' ? role : ''}" style="font-size:10px">${roleLabels[role]}</span>`;
  // Pokaż/ukryj "Nowa umowa" — tylko biznes
  const nb = document.getElementById('newContractBtn');
  if (nb) nb.style.display = role === 'biznes' ? '' : 'none';
  const navNew = document.getElementById('nav-new');
  if (navNew) navNew.style.display = role === 'biznes' ? '' : 'none';
  showApp();
  renderDash();
  updateBadges();
}

function showApp() {
  document.getElementById('loginScreen').style.display = 'none';
  document.getElementById('appShell').style.display = 'block';
}

/* ─── Demo users — widoczne tylko gdy brak konfiguracji Azure ─── */
const DEMO_USERS = [
  { name: 'Anna Kowalska',      email: 'anna.kowalska@firma.pl',        role: 'biznes',       desc: 'Biznes — inicjuje umowy' },
  { name: 'Tomasz Maj',         email: 'tomasz.maj@firma.pl',           role: 'biznes',       desc: 'Biznes — pracownik' },
  { name: 'Aurelia Kowalska',   email: 'aurelia.kowalska@firma.pl',     role: 'prawny',       desc: 'Dział prawny' },
  { name: 'Ania Nowak',         email: 'ania.nowak@firma.pl',           role: 'prawny',       desc: 'Dział prawny' },
  { name: 'Jan Kowalski',       email: 'prezes@firma.pl',               role: 'podpisujacy',  desc: 'Prezes — podpisuje umowy' },
  { name: 'Magda Wiśniewska',   email: 'magda.wisniewska@firma.pl',     role: 'podpisujacy',  desc: 'Podpisująca' },
];

function buildDemoUserList() {
  const roleColors = { biznes:'var(--text-info)', prawny:'#3c3489', podpisujacy:'var(--text-success)' };
  const roleLabels = { biznes:'Biznes', prawny:'Dział prawny', podpisujacy:'Podpisujący' };
  document.getElementById('demoUserList').innerHTML = DEMO_USERS.map(u => `
    <button class="login-ms-btn" style="justify-content:flex-start;gap:12px;text-align:left" onclick="loginAsDemo('${u.email}')">
      <div style="width:32px;height:32px;border-radius:50%;background:var(--bg-info);display:flex;align-items:center;justify-content:center;font-size:12px;font-weight:700;color:var(--text-info);flex-shrink:0">${u.name.split(' ').map(w=>w[0]).join('').substring(0,2)}</div>
      <div>
        <div style="font-size:13px;font-weight:600;color:var(--text-primary)">${u.name}</div>
        <div style="font-size:11px;color:${roleColors[u.role]};margin-top:1px">${roleLabels[u.role]} · ${u.desc}</div>
      </div>
    </button>`).join('');
}

function loginAsDemo(email) {
  const u = DEMO_USERS.find(x => x.email === email);
  if (u) applyUser(u.name + ' (DEMO)', u.email, null);
}

/* ─── Login / Logout ─── */
async function handleLogin() {
  if (!msalConfigured) {
    document.getElementById('loginProd').style.display = 'none';
    document.getElementById('loginDemo').style.display = 'block';
    buildDemoUserList();
    return;
  }
  try {
    await msalInstance.handleRedirectPromise();
    const accounts = msalInstance.getAllAccounts();
    if (accounts.length > 0) {
      await loadUserFromMsal(accounts[0]);
    } else {
      await msalInstance.loginRedirect({ scopes: ['User.Read'] });
    }
  } catch(e) {
    console.error('MSAL login error:', e);
    alert('Błąd logowania: ' + e.message);
  }
}

async function loadUserFromMsal(account) {
  try {
    const tokenRes = await msalInstance.acquireTokenSilent({
      scopes: ['User.Read'], account
    });
    // Pobierz zdjęcie z Graph API
    let photoUrl = null;
    try {
      const photoRes = await fetch('https://graph.microsoft.com/v1.0/me/photo/$value', {
        headers: { Authorization: `Bearer ${tokenRes.accessToken}` }
      });
      if (photoRes.ok) {
        const blob = await photoRes.blob();
        photoUrl = URL.createObjectURL(blob);
      }
    } catch(_) {}
    applyUser(account.name, account.username, photoUrl);
  } catch(e) {
    console.error('Token error:', e);
    await msalInstance.loginRedirect({ scopes: ['User.Read'] });
  }
}

async function handleLogout() {
  toggleUserMenu();
  if (msalConfigured && msalInstance) {
    await msalInstance.logoutRedirect();
  } else {
    document.getElementById('appShell').style.display = 'none';
    document.getElementById('loginScreen').style.display = 'flex';
    // Wróć do widoku wyboru użytkownika demo
    document.getElementById('loginProd').style.display = 'none';
    document.getElementById('loginDemo').style.display = 'block';
    buildDemoUserList();
    currentUser = null;
  }
}

function toggleUserMenu() {
  document.getElementById('userMenu').classList.toggle('open');
}
document.addEventListener('click', e => {
  const chip = document.querySelector('.user-chip');
  const menu = document.getElementById('userMenu');
  if (!chip?.contains(e.target) && !menu?.contains(e.target)) menu?.classList.remove('open');
});

/* ─── Sprawdź czy już zalogowany (po redirect) ─── */
window.addEventListener('load', async () => {
  if (!msalConfigured) {
    // Pokaż od razu picker użytkowników demo
    document.getElementById('loginProd').style.display = 'none';
    document.getElementById('loginDemo').style.display = 'block';
    buildDemoUserList();
    return;
  }
  try {
    await msalInstance.handleRedirectPromise();
    const accounts = msalInstance.getAllAccounts();
    if (accounts.length > 0) await loadUserFromMsal(accounts[0]);
  } catch(e) { console.error(e); }
});

/* ══════════════════════════════════════════════
   DANE APLIKACJI
   ══════════════════════════════════════════════ */
const sLabels = {draft:'Szkic',legal:'U prawników',revision:'Korekta — biznes',approval:'Akceptacja prawna',signing:'Do podpisu',signed:'Podpisana',archived:'Zarchiwizowana'};
const sCls    = {draft:'s-draft',legal:'s-legal',revision:'s-revision',approval:'s-approval',signing:'s-signing',signed:'s-signed',archived:'s-archived'};

let prevPage = 'dashboard';
let newDocxFile = null, newXlsxFile = null;
let selectedSigner = null; // email wybranego podpisującego

const contracts = [
  {id:1,title:'Umowa z dostawcą IT — Firma XYZ',type:'Umowa o współpracy',status:'revision',owner:'anna.kowalska@firma.pl',ownerName:'Anna Kowalska',created:'2024-05-10',docxName:'umowa_xyz_v2.docx',docxHtml:'<h2>Umowa o współpracy</h2><p>Zawarta dnia 10 maja 2024 r. pomiędzy <strong>Firma XYZ Sp. z o.o.</strong> a <strong>Nasza Firma S.A.</strong></p><h3>§1 Przedmiot umowy</h3><p>Strony postanawiają nawiązać współpracę w zakresie dostarczania usług IT.</p><h3>§4 Odpowiedzialność</h3><p>Dostawca odpowiada za prawidłowe wykonanie usług zgodnie z SLA.</p>',xlsxName:'warunki_xyz.xlsx',pdfBlob:null,pdfName:null,signerEmail:null,signerName:null,comments:[{author:'Ania Nowak',time:'12.05 10:22',text:'Należy doprecyzować §4 — odpowiedzialność za opóźnienia. Proszę o wskazanie kar umownych.'},{author:'Ania Nowak',time:'12.05 10:25',text:'Punkt 7.2 wymaga zmiany — obecny zapis jest niezgodny z KSH.'}]},
  {id:2,title:'NDA — Partner Zagraniczny GmbH',type:'Umowa NDA',status:'legal',owner:'anna.kowalska@firma.pl',ownerName:'Anna Kowalska',created:'2024-05-14',docxName:'nda_gmbh_v1.docx',docxHtml:'<h2>Umowa o zachowaniu poufności (NDA)</h2><p>Zawarta pomiędzy <strong>Partner Zagraniczny GmbH</strong> a <strong>Nasza Firma S.A.</strong></p><h3>§1 Informacje poufne</h3><p>Strony zobowiązują się do zachowania w tajemnicy wszelkich informacji handlowych.</p>',xlsxName:null,pdfBlob:null,pdfName:null,signerEmail:null,signerName:null,comments:[]},
  {id:3,title:'Umowa licencyjna — Oprogramowanie CRM',type:'Umowa licencyjna',status:'signing',owner:'tomasz.maj@firma.pl',ownerName:'Tomasz Maj',created:'2024-05-08',docxName:'licencja_crm_final.docx',docxHtml:'<h2>Umowa licencyjna oprogramowania CRM</h2><p>Licencjodawca: <strong>CRM Solutions Sp. z o.o.</strong><br>Licencjobiorca: <strong>Nasza Firma S.A.</strong></p><h3>§2 Opłaty licencyjne</h3><p>Roczna opłata wynosi 48 000 PLN netto.</p>',xlsxName:'warunki_crm.xlsx',pdfBlob:null,pdfName:'licencja_crm_DO_PODPISU.pdf',signerEmail:'prezes@firma.pl',signerName:'Jan Kowalski (Prezes)',comments:[{author:'Aurelia Kowalska',time:'11.05 14:00',text:'Umowa przeszła pełną weryfikację. Wysłano do Prezesa do podpisu.'}]},
  {id:4,title:'Umowa zakupu — Sprzęt serwerowy',type:'Umowa zakupu',status:'signed',owner:'katarzyna.zajac@firma.pl',ownerName:'Katarzyna Zając',created:'2024-04-28',docxName:'zakup_serwery_final.docx',docxHtml:'<h2>Umowa zakupu</h2><p>Sprzedający: <strong>IT Hardware Polska S.A.</strong></p><h3>§2 Cena</h3><p>Łączna cena zakupu wynosi 240 000 PLN netto.</p>',xlsxName:'spec_serwery.xlsx',pdfBlob:null,pdfName:'zakup_serwery_PODPISANA.pdf',signerEmail:'magda.wisniewska@firma.pl',signerName:'Magda Wiśniewska',comments:[]},
  {id:5,title:'Umowa o świadczenie usług — Agencja PR',type:'Umowa o świadczenie usług',status:'approval',owner:'anna.kowalska@firma.pl',ownerName:'Anna Kowalska',created:'2024-05-16',docxName:'uslugi_pr_v3.docx',docxHtml:'<h2>Umowa o świadczenie usług PR</h2><p>Usługodawca: <strong>PR Masters Sp. z o.o.</strong></p><h3>§4 Wynagrodzenie</h3><p>Miesięczne wynagrodzenie ryczałtowe: 8 000 PLN netto.</p>',xlsxName:null,pdfBlob:null,pdfName:null,signerEmail:null,signerName:null,comments:[{author:'Anna Kowalska',time:'16.05 09:10',text:'Zaktualizowałam §4 i §7 zgodnie z uwagami. Proszę o ponowny przegląd.'}]}
];

/* ══════════════════════════════════════════════
   EMAIL
   ══════════════════════════════════════════════ */
const emailLog = [];

function sendEmailNotification({ to, subject, body, contractTitle }) {
  emailLog.unshift({ time: new Date(), to, subject, body, contractTitle });
  renderEmailLog();
  showEmailToast({ to, subject });
  const be = document.getElementById('be');
  be.textContent = emailLog.length; be.style.display = '';

  /* ── TODO: BACKEND ────────────────────────────────────────────
     IT: zastąpić wywołaniem Waszego endpointu, np.:

     fetch('/api/send-email', {
       method: 'POST',
       headers: { 'Content-Type': 'application/json',
                  'Authorization': `Bearer ${await getAccessToken()}` },
       body: JSON.stringify({ to, subject, body, contractTitle,
                              from: currentUser.email })
     }).catch(err => console.error('Email error:', err));

  ─────────────────────────────────────────────────────────────── */
}

function getEmailPayload(event, contract) {
  const ownerEmail = contract.owner;
  const link = `[LINK: ${window.location.href}#${contract.id}]`;
  const legalEmails = ROLE_CONFIG.prawny.join(', ');
  const map = {
    toLegal:      { to: legalEmails,           subject: `[ObiegUmów] Nowa umowa do przeglądu: ${contract.title}`,            body: `Dzień dobry,\n\n${contract.ownerName} przesłał(a) nową umowę do weryfikacji.\n\nUmowa: ${contract.title}\nTyp: ${contract.type}\n\n${link}` },
    sendRevision: { to: ownerEmail,             subject: `[ObiegUmów] Uwagi do umowy: ${contract.title}`,                     body: `Dzień dobry,\n\nDział prawny naniósł uwagi i odsyła umowę do korekty.\n\nUmowa: ${contract.title}\n\n${link}` },
    respondBiz:   { to: legalEmails,            subject: `[ObiegUmów] Odpowiedź na uwagi: ${contract.title}`,                 body: `Dzień dobry,\n\n${contract.ownerName} odniósł(a) się do uwag i odsyła do ponownego przeglądu.\n\nUmowa: ${contract.title}\n\n${link}` },
    toApproval:   { to: ownerEmail,             subject: `[ObiegUmów] Umowa wstępnie zatwierdzona: ${contract.title}`,        body: `Dzień dobry,\n\nDział prawny wstępnie zatwierdził umowę. Trwa finalna weryfikacja.\n\nUmowa: ${contract.title}\n\n${link}` },
    pdfSentToSigner: { to: contract.signerEmail, subject: `[ObiegUmów] Umowa do podpisu: ${contract.title}`,                  body: `Dzień dobry,\n\nDział prawny przesyła umowę do podpisu.\n\nUmowa: ${contract.title}\nTyp: ${contract.type}\n\nProszę pobrać PDF, podpisać i wgrać przez system.\n\n${link}` },
    signed:       { to: legalEmails,            subject: `[ObiegUmów] Umowa podpisana: ${contract.title}`,                   body: `Dzień dobry,\n\n${contract.signerName} wgrał(a) podpisany PDF.\n\nUmowa: ${contract.title}\n\nProszę wysłać pliki na SharePoint i zarchiwizować.\n\n${link}` },
    archived:     { to: ownerEmail,             subject: `[ObiegUmów] Umowa zarchiwizowana: ${contract.title}`,              body: `Dzień dobry,\n\nUmowa została podpisana i zarchiwizowana w SharePoint.\n\n${link}` }
  };
  return map[event] || null;
}

function triggerEmail(event, contract) {
  const p = getEmailPayload(event, contract);
  if (p) sendEmailNotification({ ...p, contractTitle: contract.title });
}

let toastCtr = 0;
function showEmailToast({ to, subject }) {
  const id = 'toast-' + (++toastCtr);
  const el = document.createElement('div');
  el.className = 'toast'; el.id = id;
  el.innerHTML = `<i class="ti ti-mail toast-icon"></i><div class="toast-body"><div class="toast-title">✉ Powiadomienie wysłane</div><div class="toast-to">Do: ${to}</div><div class="toast-subject">${subject}</div></div><span class="toast-close" onclick="closeToast('${id}')"><i class="ti ti-x"></i></span><div class="toast-progress"></div>`;
  document.getElementById('toast-container').appendChild(el);
  setTimeout(() => closeToast(id), 4500);
}
function closeToast(id) {
  const el = document.getElementById(id);
  if(el){el.style.opacity='0';el.style.transform='translateX(20px)';el.style.transition='all .2s';setTimeout(()=>el.remove(),200);}
}
function renderEmailLog() {
  const el = document.getElementById('emailLogList'); if(!el) return;
  if(!emailLog.length){el.innerHTML='<p style="font-size:13px;color:var(--text-tertiary)">Brak wysłanych powiadomień w tej sesji.</p>';return;}
  el.innerHTML = emailLog.map(e=>{
    const t=e.time, ts=t.getDate()+'.'+(t.getMonth()+1)+' '+t.getHours()+':'+String(t.getMinutes()).padStart(2,'0');
    return `<div class="email-log-row"><div class="email-dot"></div><div><div class="email-log-subject">${e.subject}</div><div class="email-log-to">Do: ${e.to}</div><div class="email-log-meta">${ts} · ${e.contractTitle}</div></div></div>`;
  }).join('');
}

/* ══════════════════════════════════════════════
   ROUTING / PAGES
   ══════════════════════════════════════════════ */
function showPage(p) {
  if(p!=='detail') prevPage = document.querySelector('.page.active')?.id?.replace('page-','')||'dashboard';
  document.querySelectorAll('.page').forEach(x=>x.classList.remove('active'));
  document.querySelectorAll('.nav-item').forEach(x=>x.classList.remove('active'));
  document.getElementById('page-'+p).classList.add('active');
  const n = document.getElementById('nav-'+p); if(n) n.classList.add('active');
  if(p==='dashboard') renderDash();
  if(p==='contracts') {
    const visible = currentUser?.role==='biznes'
      ? contracts.filter(c=>c.owner===currentUser.email)
      : currentUser?.role==='podpisujacy'
      ? contracts.filter(c=>c.signerEmail===currentUser.email)
      : contracts;
    renderList('cl', visible);
  }
  if(p==='review') renderReview();
  if(p==='sign') renderSign();
  if(p==='archive') {
    const visible = currentUser?.role==='biznes'
      ? contracts.filter(c=>['signed','archived'].includes(c.status) && c.owner===currentUser.email)
      : currentUser?.role==='podpisujacy'
      ? contracts.filter(c=>['signed','archived'].includes(c.status) && c.signerEmail===currentUser.email)
      : contracts.filter(c=>['signed','archived'].includes(c.status));
    renderList('al', visible);
  }
  if(p==='emails') renderEmailLog();
  if(p==='new') { resetNew(); document.getElementById('nav-new').classList.add('active'); }
}
function goBack() { showPage(prevPage||'dashboard'); }

/* ══════════════════════════════════════════════
   NOWA UMOWA
   ══════════════════════════════════════════════ */
function resetNew() {
  newDocxFile=null; newXlsxFile=null;
  const lbl1=document.getElementById('docxLbl'); if(lbl1) lbl1.textContent='Brak pliku';
  const lbl2=document.getElementById('xlsxLbl'); if(lbl2) lbl2.textContent='Brak pliku';
  const t=document.getElementById('nt'); if(t) t.value='';
  const d=document.getElementById('nd'); if(d) d.value='';
  if(!currentUser || currentUser.role!=='biznes') {
    document.getElementById('newFormCard').innerHTML='<div class="ib"><i class="ti ti-lock"></i><span>Nowe umowy mogą inicjować tylko pracownicy działu biznesowego.</span></div>';
  }
}
function handleDocx(input) {
  if(input.files[0]){newDocxFile=input.files[0];document.getElementById('docxLbl').textContent=input.files[0].name+' ✓';document.getElementById('docxZone').style.borderColor='#185FA5';}
}
function handleXlsx(input) {
  if(input.files[0]){newXlsxFile=input.files[0];document.getElementById('xlsxLbl').textContent=input.files[0].name+' ✓';document.getElementById('xlsxZone').style.borderColor='#3B6D11';}
}
function submitContract() {
  const title = document.getElementById('nt').value.trim();
  if(!title){alert('Proszę podać tytuł umowy.');return;}
  const nc = {
    id:Date.now(), title, type:document.getElementById('nty').value,
    status:'legal', owner:currentUser.email, ownerName:currentUser.name,
    created:new Date().toISOString().split('T')[0],
    docxFile:newDocxFile, docxName:newDocxFile?newDocxFile.name:'szkic.docx',
    docxHtml:null, xlsxName:newXlsxFile?newXlsxFile.name:null,
    pdfBlob:null, pdfName:null, signerEmail:null, signerName:null, comments:[]
  };
  if(newDocxFile){
    newDocxFile.arrayBuffer().then(buf=>mammoth.convertToHtml({arrayBuffer:buf})).then(r=>{nc.docxHtml=r.value;});
  } else {
    nc.docxHtml='<p>(Brak wgranego pliku DOCX)</p>';
  }
  contracts.unshift(nc);
  updateBadges();
  triggerEmail('toLegal', nc);
  showPage('contracts');
}

/* ══════════════════════════════════════════════
   SZCZEGÓŁY UMOWY
   ══════════════════════════════════════════════ */
function openContract(id) {
  const c = contracts.find(x=>x.id===id);
  // Biznes widzi tylko swoje umowy
  if(currentUser?.role==='biznes' && c.owner !== currentUser.email) return;
  // Podpisujący widzi tylko umowy przypisane do nich
  if(currentUser?.role==='podpisujacy' && c.signerEmail !== currentUser.email) return;
  prevPage = document.querySelector('.page.active')?.id?.replace('page-','')||'dashboard';
  const c = contracts.find(x=>x.id===id);
  document.querySelectorAll('.page').forEach(x=>x.classList.remove('active'));
  document.getElementById('page-detail').classList.add('active');
  document.getElementById('dsp').textContent = sLabels[c.status]||c.status;
  document.getElementById('dsp').className = 'sp '+(sCls[c.status]||'');
  renderDetail(c);
}

function tlHtml(status) {
  const steps=[{k:'draft',ic:'ti-building',l:'Biznes'},{k:'legal',ic:'ti-gavel',l:'Prawnicy'},{k:'revision',ic:'ti-pencil',l:'Korekta'},{k:'approval',ic:'ti-check',l:'Akceptacja'},{k:'signing',ic:'ti-pen',l:'Do podpisu'},{k:'signed',ic:'ti-award',l:'Podpisana'}];
  const order=['draft','legal','revision','approval','signing','signed'];
  const ci=order.indexOf(status);
  return '<div class="tl">'+steps.map((s,i)=>{const done=i<ci,active=i===ci;return`<div class="tls"><div class="tld ${done?'done':active?'active':''}"><i class="ti ${s.ic}"></i></div><div class="tll ${active?'active':''}">${s.l}</div></div>`;}).join('')+'</div>';
}

function fChips(c) {
  let h='';
  if(c.docxName) h+=`<span class="fc"><i class="ti ti-file-type-doc"></i>${c.docxName}</span>`;
  if(c.xlsxName) h+=`<span class="fc"><i class="ti ti-file-spreadsheet"></i>${c.xlsxName}</span>`;
  if(c.pdfName)  h+=`<span class="fc pdf-ready" onclick="downloadPdfFile(${c.id})"><i class="ti ti-file-type-pdf"></i>${c.pdfName} <i class="ti ti-download" style="font-size:11px"></i></span>`;
  return h||'<span style="font-size:13px;color:var(--text-tertiary)">Brak załączników</span>';
}

function signerPickerHtml(selectedEmail) {
  return `<div class="fr" style="margin-top:12px">
    <label>Wybierz osobę do podpisu</label>
    <div class="signer-pick" id="signerPick">
      ${ROLE_CONFIG.podpisujacy.map(s=>`
        <div class="signer-opt ${selectedEmail===s.email?'selected':''}" onclick="selectSigner('${s.email}','${s.name}',this)">
          <div class="so-name">${s.name}</div>
          <div class="so-email">${s.email}</div>
        </div>`).join('')}
    </div>
  </div>`;
}

function selectSigner(email, name, el) {
  selectedSigner = { email, name };
  el.closest('.signer-pick').querySelectorAll('.signer-opt').forEach(o=>o.classList.remove('selected'));
  el.classList.add('selected');
}

function renderDetail(c) {
  const role = currentUser?.role;
  const isOwner = currentUser?.email === c.owner;
  let actions = '';

  if(role==='biznes' && isOwner && c.status==='draft')
    actions+=`<button class="btn btn-primary" onclick="act(${c.id},'toLegal')"><i class="ti ti-send"></i>Wyślij do prawników</button>`;
  if(role==='biznes' && isOwner && c.status==='revision')
    actions+=`<button class="btn btn-primary" onclick="act(${c.id},'respondBiz')"><i class="ti ti-send"></i>Wyślij odpowiedź do prawników</button>`;
  if(role==='prawny' && c.status==='legal'){
    actions+=`<button class="btn" onclick="act(${c.id},'sendRevision')"><i class="ti ti-arrow-back-up"></i>Odeślij z uwagami</button>`;
    actions+=`<button class="btn btn-success" onclick="act(${c.id},'toApproval')"><i class="ti ti-check"></i>Zatwierdź wstępnie</button>`;
  }
  if(role==='prawny' && c.status==='approval'){
    actions+=`<button class="btn" onclick="act(${c.id},'sendRevision')"><i class="ti ti-arrow-back-up"></i>Odeślij do korekty</button>`;
    actions+=`<button class="btn btn-primary" id="genPdfBtn" onclick="generateAndSendPdf(${c.id})"><i class="ti ti-file-type-pdf"></i>Generuj PDF i wyślij do podpisu</button>`;
  }
  if(role==='podpisujacy' && c.status==='signing' && c.signerEmail===currentUser?.email){
    actions+=`<button class="btn btn-primary" onclick="downloadPdfFile(${c.id})"><i class="ti ti-download"></i>Pobierz PDF do podpisu</button>`;
    actions+=`<button class="btn btn-success" onclick="uploadSignedPdf(${c.id})"><i class="ti ti-upload"></i>Wgraj podpisany PDF</button>`;
  }
  if(role==='prawny' && c.status==='signed'){
    actions+=`<button class="btn btn-success" onclick="sendToSharepoint(${c.id})"><i class="ti ti-cloud-upload"></i>Wyślij na SharePoint i archiwizuj</button>`;
  }

  const cHtml = c.comments.map(cm=>`<div class="cb"><span class="ca">${cm.author}</span><span class="ct">${cm.time}</span><p class="cx">${cm.text}</p></div>`).join('');

  let addComment = '';
  if(['legal','revision','approval'].includes(c.status) && (role==='prawny'||isOwner)){
    addComment=`<div style="margin-top:8px"><textarea id="cmtBox" placeholder="Dodaj komentarz..."></textarea><div style="margin-top:6px"><button class="btn btn-sm" onclick="addCmt(${c.id})"><i class="ti ti-message-plus"></i>Dodaj</button></div></div>`;
  }

  let infoBox = '';
  if(c.status==='signing') infoBox=`<div class="ib"><i class="ti ti-info-circle"></i><span>PDF oczekuje na podpis: <strong>${c.signerName||'—'}</strong></span></div>`;
  if(c.status==='signed')  infoBox=`<div class="sb"><i class="ti ti-circle-check"></i><span>Umowa podpisana przez <strong>${c.signerName||'—'}</strong>. Dział prawny może wysłać na SharePoint.</span></div>`;
  if(c.status==='archived') infoBox=`<div class="sb"><i class="ti ti-archive"></i><span>Umowa zarchiwizowana i dostępna na SharePoint.</span></div>`;

  // Picker osoby podpisującej przy approval
  let signerSection = '';
  if(role==='prawny' && c.status==='approval'){
    selectedSigner = c.signerEmail ? {email:c.signerEmail,name:c.signerName} : null;
    signerSection = `<div class="divider"></div><div class="dl">Podgląd DOCX</div><div style="margin-bottom:8px">${c.docxHtml?'<div class="pdf-preview">'+c.docxHtml+'</div>':'<p style="font-size:13px;color:var(--text-tertiary)">Brak podglądu</p>'}</div>${signerPickerHtml(c.signerEmail)}<div id="convStatus${c.id}"></div>`;
  }

  document.getElementById('dc').innerHTML = `
    <div class="card" style="margin-bottom:12px">
      <h2 style="font-size:16px;font-weight:600;margin-bottom:4px">${c.title}</h2>
      <p style="font-size:12px;color:var(--text-secondary)">${c.type} &nbsp;·&nbsp; ${c.ownerName} &nbsp;·&nbsp; ${c.created}</p>
    </div>
    ${tlHtml(c.status)}
    ${infoBox}
    <div class="card">
      <div style="margin-bottom:12px"><div class="dl">Załączniki</div>${fChips(c)}</div>
      ${signerSection}
      <div class="divider"></div>
      <div><div class="dl">Komentarze i historia</div>${cHtml||'<p style="font-size:13px;color:var(--text-tertiary)">Brak komentarzy</p>'}${addComment}</div>
      ${actions?`<div class="ab">${actions}</div>`:''}
    </div>`;
}

/* ══════════════════════════════════════════════
   AKCJE NA UMOWIE
   ══════════════════════════════════════════════ */
function act(id, action) {
  const c = contracts.find(x=>x.id===id);
  const cmtEl = document.getElementById('cmtBox');
  const txt = cmtEl?cmtEl.value.trim():'';
  if(action==='toLegal')    { c.status='legal';    triggerEmail('toLegal',c); }
  if(action==='respondBiz') { if(txt) pushCmt(c,txt); c.status='legal'; triggerEmail('respondBiz',c); }
  if(action==='sendRevision'){ pushCmt(c,txt||'[Odesłano do korekty]'); c.status='revision'; triggerEmail('sendRevision',c); }
  if(action==='toApproval') { c.status='approval'; triggerEmail('toApproval',c); }
  if(action==='archive')    { c.status='archived'; triggerEmail('archived',c); }
  document.getElementById('dsp').textContent = sLabels[c.status];
  document.getElementById('dsp').className = 'sp '+(sCls[c.status]||'');
  renderDetail(c); updateBadges();
}

function addCmt(id) {
  const c = contracts.find(x=>x.id===id);
  const txt = document.getElementById('cmtBox').value.trim();
  if(!txt) return;
  pushCmt(c,txt); renderDetail(c);
}
function pushCmt(c,txt) {
  c.comments.push({author:currentUser?.name||'?', time:now(), text:txt});
}

/* ══════════════════════════════════════════════
   PDF
   ══════════════════════════════════════════════ */
async function generateAndSendPdf(id) {
  const c = contracts.find(x=>x.id===id);
  if(!selectedSigner){alert('Proszę wybrać osobę do podpisu przed wygenerowaniem PDF.');return;}
  c.signerEmail = selectedSigner.email;
  c.signerName  = selectedSigner.name;
  const btn = document.getElementById('genPdfBtn');
  const statusEl = document.getElementById('convStatus'+id);
  if(btn){btn.disabled=true;btn.innerHTML='<div class="spin"></div>Generowanie PDF...';}
  if(statusEl) statusEl.innerHTML='<div class="converting"><div class="spin"></div>Konwertowanie DOCX → PDF...</div>';
  try {
    const renderArea = document.getElementById('pdfRenderArea');
    const stamp = `<div style="border-bottom:2px solid #185FA5;padding-bottom:12px;margin-bottom:20px;display:flex;justify-content:space-between"><div><div style="font-size:18px;font-weight:700;color:#185FA5">ObiegUmów</div><div style="font-size:11px;color:#666">System obiegu umów</div></div><div style="text-align:right;font-size:11px;color:#666"><div>Wygenerowano: ${new Date().toLocaleString('pl-PL')}</div><div>Do podpisu: ${c.signerName}</div></div></div>`;
    const footer = `<div style="border-top:1px solid #ccc;margin-top:32px;padding-top:12px;font-size:11px;color:#999;display:flex;justify-content:space-between"><span>${c.title}</span><span>Strona 1</span></div>`;
    renderArea.innerHTML = stamp + (c.docxHtml||'<h2>'+c.title+'</h2>') + footer;
    await new Promise(r=>setTimeout(r,300));
    const canvas = await html2canvas(renderArea,{scale:1.5,useCORS:true,backgroundColor:'#ffffff'});
    const {jsPDF} = window.jspdf;
    const pdf = new jsPDF('p','pt','a4');
    const W=595.28; const ratio=canvas.height/canvas.width;
    pdf.addImage(canvas.toDataURL('image/jpeg',.92),'JPEG',0,0,W,Math.min(W*ratio,841.89));
    c.pdfBlob = pdf.output('blob');
    c.pdfName = c.title.substring(0,25).replace(/[^a-zA-Z0-9]/g,'_')+'_DO_PODPISU.pdf';
    c.status = 'signing';
    c.comments.push({author:currentUser?.name||'Dział prawny', time:now(), text:`PDF wygenerowany i wysłany do podpisu: ${c.signerName} (${c.signerEmail})`});
    renderArea.innerHTML = '';
    updateBadges();
    triggerEmail('pdfSentToSigner', c);
    if(statusEl) statusEl.innerHTML='<div class="sb" style="margin-top:8px"><i class="ti ti-check"></i><span>PDF wygenerowany i wysłany do: <strong>${c.signerName}</strong></span></div>';
    setTimeout(()=>{document.getElementById('dsp').textContent=sLabels[c.status];document.getElementById('dsp').className='sp '+(sCls[c.status]||'');renderDetail(c);},1200);
  } catch(e) {
    if(statusEl) statusEl.innerHTML=`<div class="wb"><i class="ti ti-alert-circle"></i><span>Błąd: ${e.message}</span></div>`;
    if(btn){btn.disabled=false;btn.innerHTML='<i class="ti ti-file-type-pdf"></i>Generuj PDF i wyślij do podpisu';}
  }
}

function downloadPdfFile(id) {
  const c = contracts.find(x=>x.id===id);
  if(c.pdfBlob){ const url=URL.createObjectURL(c.pdfBlob);const a=document.createElement('a');a.href=url;a.download=c.pdfName||'umowa.pdf';a.click();setTimeout(()=>URL.revokeObjectURL(url),2000); }
  else alert('Plik PDF: '+c.pdfName+'\n(Wygeneruj PDF przez dział prawny)');
}

function uploadSignedPdf(id) {
  const input = document.createElement('input');
  input.type='file';input.accept='.pdf';
  input.onchange = e => {
    if(e.target.files[0]){
      const c = contracts.find(x=>x.id===id);
      c.pdfBlob=e.target.files[0];c.pdfName=e.target.files[0].name;
      c.status='signed';
      c.comments.push({author:currentUser?.name||'Podpisujący',time:now(),text:'Podpisany PDF wgrany: '+e.target.files[0].name});
      document.getElementById('dsp').textContent=sLabels[c.status];
      document.getElementById('dsp').className='sp '+(sCls[c.status]||'');
      triggerEmail('signed',c);
      renderDetail(c);updateBadges();
    }
  };
  input.click();
}

/* ═══════════════════════════════════════════════════════════════
   SHAREPOINT — wysyłka przez Microsoft Graph API
   ═══════════════════════════════════════════════════════════════
   IT: uzupełnij SHAREPOINT_CONFIG poniżej.

   Jak znaleźć siteId i driveId:
   1. Otwórz: https://graph.microsoft.com/v1.0/sites?search=NazwaWaszejFirmy
   2. Skopiuj "id" z wyniku → to jest SITE_ID
   3. Otwórz: https://graph.microsoft.com/v1.0/sites/{siteId}/drives
   4. Znajdź "Dokumenty" lub docelową bibliotekę → skopiuj "id" → DRIVE_ID
   5. FOLDER_PATH = ścieżka folderu w bibliotece, np. "/Prawny/Umowy"
   ═══════════════════════════════════════════════════════════════ */
const SHAREPOINT_CONFIG = {
  siteId:   'TUTAJ_WKLEJ_SITE_ID',    // ← IT: uzupełnić
  driveId:  'TUTAJ_WKLEJ_DRIVE_ID',   // ← IT: uzupełnić
  folderPath: '/Prawny/Umowy'          // ← IT: dostosować ścieżkę
};

const spConfigured = SHAREPOINT_CONFIG.siteId !== 'TUTAJ_WKLEJ_SITE_ID';

/* Pobierz aktualny access token z MSAL (do Graph API) */
async function getGraphToken() {
  if (!msalConfigured || !msalInstance) return null;
  const accounts = msalInstance.getAllAccounts();
  if (!accounts.length) return null;
  try {
    const res = await msalInstance.acquireTokenSilent({
      scopes: ['Files.ReadWrite', 'Sites.ReadWrite.All'],
      account: accounts[0]
    });
    return res.accessToken;
  } catch(e) {
    console.error('Token error:', e);
    return null;
  }
}

/* Upload jednego pliku na SharePoint przez Graph API */
async function uploadToSharePoint(token, fileName, fileBlob, yearFolder, contractId) {
  const folder = `${SHAREPOINT_CONFIG.folderPath}/${yearFolder}`;
  const url = `https://graph.microsoft.com/v1.0/sites/${SHAREPOINT_CONFIG.siteId}/drives/${SHAREPOINT_CONFIG.driveId}/root:${folder}/${fileName}:/content`;
  const res = await fetch(url, {
    method: 'PUT',
    headers: {
      'Authorization': `Bearer ${token}`,
      'Content-Type': 'application/octet-stream'
    },
    body: fileBlob
  });
  if (!res.ok) {
    const err = await res.json().catch(()=>({error:{message:res.statusText}}));
    throw new Error(err.error?.message || res.statusText);
  }
  return await res.json(); // zwraca metadane pliku, w tym webUrl
}

/* Główna funkcja — wysyłka na SharePoint */
async function sendToSharepoint(id) {
  const c = contracts.find(x=>x.id===id);
  showSpModal(c);
}

/* ── MODAL SharePoint ── */
function showSpModal(c) {
  const year = new Date().getFullYear();
  const folderDisplay = `${SHAREPOINT_CONFIG.folderPath}/${year}/`;
  const existing = document.getElementById('spModal');
  if (existing) existing.remove();

  const modal = document.createElement('div');
  modal.id = 'spModal';
  modal.style.cssText = 'position:fixed;inset:0;background:rgba(0,0,0,.45);z-index:500;display:flex;align-items:center;justify-content:center;padding:16px';
  modal.innerHTML = `
    <div style="background:var(--bg-primary);border-radius:var(--radius-lg);padding:28px;width:460px;max-width:100%;box-shadow:0 8px 40px rgba(0,0,0,.18)">
      <div style="display:flex;align-items:center;gap:10px;margin-bottom:4px">
        <i class="ti ti-brand-sharepoint" style="font-size:22px;color:#038387"></i>
        <h3 style="font-size:15px;font-weight:600;color:var(--text-primary)">Wyślij na SharePoint</h3>
      </div>
      <p style="font-size:12px;color:var(--text-tertiary);margin-bottom:20px">Folder docelowy: <code style="background:rgba(0,0,0,.07);padding:1px 5px;border-radius:3px;font-size:11px">${folderDisplay}</code></p>

      <div style="margin-bottom:16px">
        <div class="dl">Pliki do wysłania</div>
        <div id="spFileList" style="display:flex;flex-direction:column;gap:6px;margin-top:6px">
          ${c.docxName ? `<div class="fc" style="cursor:default"><i class="ti ti-file-type-doc"></i>${c.docxName}</div>` : ''}
          ${c.pdfName  ? `<div class="fc pdf-ready" style="cursor:default"><i class="ti ti-file-type-pdf"></i>${c.pdfName}</div>` : ''}
        </div>
      </div>

      <div id="spStatus" style="margin-bottom:16px"></div>

      ${!spConfigured ? `<div class="wb" style="margin-bottom:16px"><i class="ti ti-tool"></i><div><strong>Tryb DEMO</strong> — brak konfiguracji SharePoint.<br>IT uzupełnia <code style="background:rgba(0,0,0,.08);padding:1px 4px;border-radius:3px;font-size:11px">SHAREPOINT_CONFIG</code> w kodzie. Symulacja uploadu.</div></div>` : ''}

      <div style="display:flex;gap:8px;justify-content:flex-end">
        <button class="btn" onclick="closeSpModal()">Anuluj</button>
        <button class="btn btn-success" id="spSendBtn" onclick="doSendToSharepoint(${c.id})">
          <i class="ti ti-cloud-upload"></i>${spConfigured ? 'Wyślij na SharePoint' : 'Symuluj wysyłkę (DEMO)'}
        </button>
      </div>
    </div>`;
  document.body.appendChild(modal);
}

function closeSpModal() {
  const m = document.getElementById('spModal');
  if (m) m.remove();
}

async function doSendToSharepoint(id) {
  const c = contracts.find(x=>x.id===id);
  const btn = document.getElementById('spSendBtn');
  const statusEl = document.getElementById('spStatus');
  btn.disabled = true;
  btn.innerHTML = '<div class="spin"></div>Wysyłanie...';

  const year = new Date().getFullYear();
  const results = [];

  if (!spConfigured) {
    /* ── TRYB DEMO — symulacja ── */
    statusEl.innerHTML = '<div class="converting"><div class="spin"></div>Symulowanie uploadu DOCX...</div>';
    await new Promise(r => setTimeout(r, 800));
    if (c.docxName) results.push({ name: c.docxName, ok: true, url: '#demo' });
    statusEl.innerHTML = '<div class="converting"><div class="spin"></div>Symulowanie uploadu PDF...</div>';
    await new Promise(r => setTimeout(r, 800));
    if (c.pdfName) results.push({ name: c.pdfName, ok: true, url: '#demo' });
  } else {
    /* ── PRODUKCJA — prawdziwy upload przez Graph API ── */
    const token = await getGraphToken();
    if (!token) {
      statusEl.innerHTML = '<div class="wb"><i class="ti ti-alert-circle"></i><span>Brak tokenu — zaloguj się przez Microsoft 365.</span></div>';
      btn.disabled = false; btn.innerHTML = '<i class="ti ti-cloud-upload"></i>Wyślij na SharePoint';
      return;
    }
    // Upload DOCX
    if (c.docxFile || c.docxName) {
      statusEl.innerHTML = `<div class="converting"><div class="spin"></div>Wysyłanie ${c.docxName}...</div>`;
      try {
        const blob = c.docxFile || new Blob([''], {type:'application/vnd.openxmlformats-officedocument.wordprocessingml.document'});
        const meta = await uploadToSharePoint(token, c.docxName, blob, year, c.id);
        results.push({ name: c.docxName, ok: true, url: meta.webUrl });
      } catch(e) {
        results.push({ name: c.docxName, ok: false, err: e.message });
      }
    }
    // Upload PDF
    if (c.pdfBlob || c.pdfName) {
      statusEl.innerHTML = `<div class="converting"><div class="spin"></div>Wysyłanie ${c.pdfName}...</div>`;
      try {
        const meta = await uploadToSharePoint(token, c.pdfName, c.pdfBlob, year, c.id);
        results.push({ name: c.pdfName, ok: true, url: meta.webUrl });
      } catch(e) {
        results.push({ name: c.pdfName, ok: false, err: e.message });
      }
    }
  }

  /* Podsumowanie */
  const allOk = results.every(r => r.ok);
  const resultHtml = results.map(r => r.ok
    ? `<div class="fc pdf-ready" style="cursor:default"><i class="ti ti-circle-check"></i>${r.name} ${r.url !== '#demo' ? `<a href="${r.url}" target="_blank" style="font-size:10px;margin-left:4px;color:var(--text-info)">otwórz ↗</a>` : '(DEMO)'}</div>`
    : `<div class="fc" style="cursor:default;border-color:var(--text-danger);color:var(--text-danger)"><i class="ti ti-alert-circle"></i>${r.name}: ${r.err}</div>`
  ).join('');

  statusEl.innerHTML = allOk
    ? `<div class="sb"><i class="ti ti-circle-check"></i><span>Pliki wysłane pomyślnie na SharePoint!</span></div>${resultHtml}`
    : `<div class="wb"><i class="ti ti-alert-circle"></i><span>Część plików nie została wysłana.</span></div>${resultHtml}`;

  if (allOk) {
    /* Archiwizuj i zamknij po chwili */
    c.status = 'archived';
    c.comments.push({
      author: currentUser?.name || 'Dział prawny', time: now(),
      text: `Pliki wysłane na SharePoint${spConfigured ? '' : ' (DEMO)'}. Umowa zarchiwizowana.`
    });
    triggerEmail('archived', c);
    updateBadges();
    btn.innerHTML = '<i class="ti ti-check"></i>Gotowe';
    setTimeout(() => {
      closeSpModal();
      document.getElementById('dsp').textContent = sLabels[c.status];
      document.getElementById('dsp').className = 'sp ' + (sCls[c.status] || '');
      renderDetail(c);
    }, 2000);
  } else {
    btn.disabled = false;
    btn.innerHTML = '<i class="ti ti-refresh"></i>Spróbuj ponownie';
  }
}

/* ══════════════════════════════════════════════
   RENDER HELPERS
   ══════════════════════════════════════════════ */
function crow(c) {
  const isMine = currentUser?.email===c.owner;
  const isMySign = currentUser?.email===c.signerEmail;
  const roleTag = (isMine && c.status!=='archived') ? '<span style="font-size:10px;color:var(--text-info);margin-left:6px">moja</span>' : (isMySign?'<span style="font-size:10px;color:#3c3489;margin-left:6px">do podpisu</span>':'');
  return `<div class="crow" onclick="openContract(${c.id})"><div class="cicon"><i class="ti ti-file-description"></i></div><div class="cinfo"><div class="ctitle">${c.title}${roleTag}</div><div class="cmeta">${c.type} · ${c.ownerName} · ${c.created}</div></div><span class="sp ${sCls[c.status]}">${sLabels[c.status]}</span></div>`;
}

function renderList(elId, list) {
  document.getElementById(elId).innerHTML = list.length ? list.map(crow).join('') : '<p style="font-size:13px;color:var(--text-tertiary);padding:16px 0">Brak umów.</p>';
}

function renderDash() {
  if(!currentUser) return;
  document.getElementById('dg').textContent = 'Dzień dobry, '+currentUser.name.replace(' (DEMO)','').split(' ')[0]+'!';
  const mine = currentUser.role==='biznes'
    ? contracts.filter(c=>c.owner===currentUser.email)
    : currentUser.role==='podpisujacy'
    ? contracts.filter(c=>c.signerEmail===currentUser.email)
    : contracts;
  document.getElementById('st').textContent = mine.length;
  document.getElementById('sa').textContent = mine.filter(c=>!['signed','archived'].includes(c.status)).length;
  document.getElementById('ss').textContent = mine.filter(c=>c.status==='signed').length;
  let info='';
  if(currentUser.role==='biznes'){const r=contracts.filter(c=>c.status==='revision'&&c.owner===currentUser.email);if(r.length) info=`<div class="ib"><i class="ti ti-alert-circle"></i><span>${r.length} Twoja umowa czeka na odpowiedź po uwagach prawników.</span></div>`;}
  if(currentUser.role==='prawny'){const r=contracts.filter(c=>c.status==='legal'||c.status==='approval');if(r.length) info=`<div class="ib"><i class="ti ti-gavel"></i><span>${r.length} umów czeka na przegląd lub zatwierdzenie.</span></div>`;}
  if(currentUser.role==='podpisujacy'){const r=contracts.filter(c=>c.status==='signing'&&c.signerEmail===currentUser.email);if(r.length) info=`<div class="ib"><i class="ti ti-pen"></i><span>${r.length} umowa czeka na Twój podpis.</span></div>`;}
  document.getElementById('dri').innerHTML=info;
  const recent = currentUser.role==='biznes'
    ? contracts.filter(c=>c.owner===currentUser.email).slice(0,5)
    : currentUser.role==='podpisujacy'
    ? contracts.filter(c=>c.signerEmail===currentUser.email).slice(0,5)
    : contracts.slice(0,5);
  document.getElementById('rc').innerHTML=recent.map(crow).join('');
}

function renderReview() {
  if(!currentUser) return;
  let visible=[];
  if(currentUser.role==='biznes'){
    visible=contracts.filter(c=>c.status==='revision'&&c.owner===currentUser.email);
    document.getElementById('rt').textContent='Odpowiedź na uwagi prawników';
    document.getElementById('rsub').textContent='Twoje umowy odesłane do korekty';
  } else if(currentUser.role==='prawny'){
    visible=contracts.filter(c=>c.status==='legal'||c.status==='approval');
    document.getElementById('rt').textContent='Umowy do przeglądu';
    document.getElementById('rsub').textContent='Przeglądaj, komentuj, zatwierdź lub odeślij do korekty';
  } else {
    document.getElementById('rt').textContent='Brak umów do przeglądu';
    document.getElementById('rsub').textContent='';
  }
  renderList('rl',visible);
}

function renderSign() {
  if(!currentUser) return;
  const visible = currentUser.role==='podpisujacy'
    ? contracts.filter(c=>c.status==='signing'&&c.signerEmail===currentUser.email)
    : [];
  document.getElementById('sl2').innerHTML = visible.length ? visible.map(crow).join('') : '<p style="font-size:13px;color:var(--text-tertiary);padding:16px 0">Brak umów oczekujących na Twój podpis.</p>';
}

function updateBadges() {
  if(!currentUser) return;
  const rev = currentUser.role==='biznes'
    ? contracts.filter(c=>c.status==='revision'&&c.owner===currentUser.email).length
    : currentUser.role==='prawny'
    ? contracts.filter(c=>c.status==='legal'||c.status==='approval').length : 0;
  const sign = currentUser.role==='podpisujacy'
    ? contracts.filter(c=>c.status==='signing'&&c.signerEmail===currentUser.email).length : 0;
  const br=document.getElementById('br'); br.textContent=rev; br.style.display=rev?'':'none';
  const bsEl=document.getElementById('bs'); bsEl.textContent=sign; bsEl.style.display=sign?'':'none';
  document.getElementById('bc').textContent=contracts.length;
}

function now(){const d=new Date();return d.getDate()+'.'+(d.getMonth()+1)+' '+d.getHours()+':'+String(d.getMinutes()).padStart(2,'0');}
</script>
</body>
</html>
