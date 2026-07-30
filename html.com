<!DOCTYPE html>
<html lang="pt-BR">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no">
<meta name="apple-mobile-web-app-capable" content="yes">
<meta name="apple-mobile-web-app-status-bar-style" content="black-translucent">
<title>Visitas — Astra &amp; Japi</title>
<script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js"></script>
<style>
  :root {
    --bg: #F4F6FA;
    --surface: #FFFFFF;
    --border: rgba(15,30,60,0.09);
    --text: #16202E;
    --muted: #7A8494;
    --accent: #2454B0;
    --accent-light: #E8EFFC;
    --accent-dark: #173B82;
    --red: #D14343;
    --red-light: #FDECEC;
    --amber: #C97A00;
    --amber-light: #FFF6E5;
    --green: #1E8E5A;
    --green-light: #E9FAF1;
    --gray-light: #EEF0F3;
    --radius: 14px;
    --radius-sm: 8px;
    --shadow: 0 2px 12px rgba(36,84,176,0.10);
  }
  * { box-sizing: border-box; margin: 0; padding: 0; -webkit-tap-highlight-color: transparent; }
  body { font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', sans-serif; background: var(--bg); color: var(--text); min-height: 100vh; padding-bottom: 80px; }

  /* TOP BAR */
  .topbar { background: var(--accent); color: #fff; position: sticky; top: 0; z-index: 100; border-bottom: 3px solid var(--accent-dark); }
  .topbar-row { display: flex; align-items: center; justify-content: space-between; padding: 14px 16px 10px; }
  .brand { display: flex; align-items: center; gap: 10px; }
  .brand-icon { font-size: 22px; }
  .brand-title { font-size: 16px; font-weight: 700; letter-spacing: .2px; }
  .brand-sub { font-size: 11.5px; opacity: .85; margin-top: 1px; }
  .sync-btn { background: rgba(255,255,255,0.15); border: 1px solid rgba(255,255,255,0.35); color: #fff; padding: 6px 12px; border-radius: 20px; font-size: 12.5px; cursor: pointer; display: flex; align-items: center; gap: 5px; white-space: nowrap; }
  .sync-btn:active { background: rgba(255,255,255,0.25); }
  .offline-bar { background: var(--amber); color: #fff; text-align: center; font-size: 12px; padding: 5px; display: none; }
  .offline-bar.show { display: block; }

  /* SEARCH */
  .search-wrap { padding: 14px 16px 8px; position: relative; }
  .search-input { width: 100%; padding: 12px 16px 12px 42px; border: 1.5px solid var(--border); border-radius: 12px; font-size: 16px; background: var(--surface); color: var(--text); outline: none; }
  .search-input:focus { border-color: var(--accent); }
  .search-icon { position: absolute; left: 28px; top: 50%; transform: translateY(-50%); color: var(--muted); font-size: 18px; pointer-events: none; }

  /* FILTER CHIPS */
  .chip-row { display: flex; gap: 6px; padding: 0 16px 10px; overflow-x: auto; }
  .chip { flex-shrink: 0; padding: 7px 13px; border-radius: 20px; border: 1.5px solid var(--border); background: var(--surface); font-size: 12.5px; font-weight: 600; color: var(--muted); cursor: pointer; white-space: nowrap; }
  .chip.active { background: var(--accent); border-color: var(--accent); color: #fff; }
  .select-row { display: flex; gap: 8px; padding: 0 16px 10px; }
  .select-row select { flex: 1; padding: 9px 10px; border-radius: 10px; border: 1.5px solid var(--border); background: var(--surface); font-size: 13px; color: var(--text); }

  /* TABS */
  .tabs { display: flex; padding: 0 16px; gap: 6px; margin-bottom: 4px; }
  .tab { flex: 1; padding: 9px 4px; border: none; background: transparent; font-size: 13px; color: var(--muted); cursor: pointer; border-bottom: 2px solid transparent; font-weight: 500; }
  .tab.active { color: var(--accent); border-bottom-color: var(--accent); font-weight: 600; }

  /* CARDS */
  .list { padding: 8px 16px; display: flex; flex-direction: column; gap: 10px; }
  .card { background: var(--surface); border-radius: var(--radius); padding: 14px 16px; box-shadow: var(--shadow); border: 1px solid var(--border); cursor: pointer; transition: transform .1s; }
  .card:active { transform: scale(0.98); }
  .card-top { display: flex; justify-content: space-between; align-items: flex-start; gap: 8px; }
  .card-name { font-size: 15px; font-weight: 600; line-height: 1.3; flex: 1; }
  .card-meta { display: flex; gap: 10px; margin-top: 6px; font-size: 12px; color: var(--muted); flex-wrap: wrap; align-items: center; }
  .card-obs { font-size: 12px; color: var(--muted); margin-top: 6px; line-height: 1.5; display: -webkit-box; -webkit-line-clamp: 1; -webkit-box-orient: vertical; overflow: hidden; }
  .badge { font-size: 11px; font-weight: 600; padding: 3px 9px; border-radius: 20px; white-space: nowrap; flex-shrink: 0; }
  .badge-blue { background: var(--accent-light); color: var(--accent); }
  .badge-green { background: var(--green-light); color: var(--green); }
  .badge-amber { background: var(--amber-light); color: var(--amber); }
  .badge-gray { background: var(--gray-light); color: #666; }
  .badge-red { background: var(--red-light); color: var(--red); }
  .alert-dot { width: 8px; height: 8px; border-radius: 50%; flex-shrink: 0; }
  .dot-red { background: var(--red); }
  .dot-amber { background: var(--amber); }
  .dot-green { background: var(--green); }
  .dot-gray { background: #ccc; }
  .empty-state { text-align: center; padding: 60px 20px; color: var(--muted); }
  .empty-state div:first-child { font-size: 34px; margin-bottom: 10px; }

  /* BOTTOM NAV */
  .bottom-nav { position: fixed; bottom: 0; left: 0; right: 0; background: var(--surface); border-top: 1px solid var(--border); display: flex; z-index: 100; padding-bottom: env(safe-area-inset-bottom); }
  .nav-btn { flex: 1; display: flex; flex-direction: column; align-items: center; padding: 10px 4px 8px; border: none; background: transparent; cursor: pointer; color: var(--muted); font-size: 11px; gap: 3px; }
  .nav-btn.active { color: var(--accent); font-weight: 600; }
  .nav-btn.active .nav-icon { color: var(--accent); }
  .nav-icon { font-size: 22px; }

  /* DRAWER / FORM */
  .overlay { position: fixed; inset: 0; background: rgba(10,20,40,0.45); z-index: 200; display: none; }
  .overlay.show { display: block; }
  .drawer { position: fixed; bottom: 0; left: 0; right: 0; background: var(--surface); border-radius: 20px 20px 0 0; z-index: 201; max-height: 92vh; overflow-y: auto; transform: translateY(100%); transition: transform .3s cubic-bezier(.32,.72,0,1); padding-bottom: env(safe-area-inset-bottom); }
  .drawer.open { transform: translateY(0); }
  .drawer-handle { width: 36px; height: 4px; background: #ddd; border-radius: 2px; margin: 12px auto 0; }
  .drawer-header { padding: 12px 20px 16px; border-bottom: 2px solid var(--accent); display: flex; align-items: center; justify-content: space-between; }
  .drawer-title { font-size: 17px; font-weight: 700; color: var(--accent); }
  .close-btn { width: 30px; height: 30px; border-radius: 50%; background: var(--gray-light); border: none; cursor: pointer; font-size: 16px; display: flex; align-items: center; justify-content: center; }
  .drawer-body { padding: 16px 20px; display: flex; flex-direction: column; gap: 14px; }
  .section-label { font-size: 13px; font-weight: 700; color: var(--muted); text-transform: uppercase; letter-spacing: .4px; margin-bottom: 2px; }

  /* FORM FIELDS */
  .field { display: flex; flex-direction: column; gap: 5px; }
  .field label { font-size: 13px; font-weight: 600; color: var(--muted); text-transform: uppercase; letter-spacing: 0.4px; }
  .field input, .field select, .field textarea { padding: 12px 14px; border: 1.5px solid var(--border); border-radius: var(--radius-sm); font-size: 16px; background: var(--bg); color: var(--text); outline: none; width: 100%; }
  .field input:focus, .field select:focus, .field textarea:focus { border-color: var(--accent); background: #fff; }
  .field textarea { min-height: 70px; resize: none; line-height: 1.5; }
  .field-row { display: grid; grid-template-columns: 1fr 1fr; gap: 12px; }

  /* VISIT HISTORY */
  .visit-list { display: flex; flex-direction: column; gap: 8px; }
  .visit-item { display: flex; align-items: center; justify-content: space-between; background: var(--bg); border-radius: var(--radius-sm); padding: 10px 12px; font-size: 13.5px; }
  .visit-item button { border: none; background: none; color: var(--red); font-size: 13px; cursor: pointer; padding: 4px 6px; }
  .add-visit-row { display: flex; gap: 8px; }
  .add-visit-row input { flex: 1; }
  .btn-ghost { background: var(--accent-light); color: var(--accent); border: none; padding: 12px 14px; border-radius: var(--radius-sm); font-size: 14px; font-weight: 600; cursor: pointer; white-space: nowrap; }
  .btn-today { background: var(--green-light); color: var(--green); border: none; padding: 13px 14px; border-radius: var(--radius-sm); font-size: 14px; font-weight: 700; cursor: pointer; width: 100%; }

  /* BUTTONS */
  .btn-primary { background: var(--accent); color: #fff; border: none; padding: 15px; border-radius: var(--radius-sm); font-size: 16px; font-weight: 600; cursor: pointer; width: 100%; letter-spacing: .3px; }
  .btn-primary:active { opacity: .85; }
  .btn-danger { background: none; color: var(--red); border: 1.5px solid var(--red-light); padding: 13px; border-radius: var(--radius-sm); font-size: 14px; font-weight: 600; cursor: pointer; width: 100%; }
  .origin-grid { display: grid; grid-template-columns: 1fr 1fr; gap: 8px; }
  .origin-pill { padding: 10px 6px; border-radius: var(--radius-sm); border: 2px solid var(--border); text-align: center; font-size: 13px; font-weight: 600; cursor: pointer; }
  .origin-pill.sel-Astra { background: var(--accent-light); border-color: var(--accent); color: var(--accent); }
  .origin-pill.sel-Japi { background: var(--green-light); border-color: var(--green); color: var(--green); }

  /* DASHBOARD */
  .dash-wrap { padding: 14px 16px 24px; display: flex; flex-direction: column; gap: 14px; }
  .kpi-grid { display: grid; grid-template-columns: 1fr 1fr; gap: 10px; }
  .kpi-card { background: var(--surface); border-radius: var(--radius); padding: 14px; box-shadow: var(--shadow); border: 1px solid var(--border); }
  .kpi-num { font-size: 24px; font-weight: 800; color: var(--accent); }
  .kpi-label { font-size: 11.5px; color: var(--muted); margin-top: 2px; font-weight: 600; text-transform: uppercase; letter-spacing: .3px; }
  .dash-card { background: var(--surface); border-radius: var(--radius); padding: 16px; box-shadow: var(--shadow); border: 1px solid var(--border); }
  .dash-card-title { font-size: 13px; font-weight: 700; color: var(--muted); text-transform: uppercase; letter-spacing: .4px; margin-bottom: 12px; }
  .bar-row { display: flex; align-items: center; gap: 10px; margin-bottom: 10px; }
  .bar-label { font-size: 12px; color: var(--muted); width: 46px; flex-shrink: 0; }
  .bar-track { flex: 1; background: var(--bg); border-radius: 5px; height: 14px; overflow: hidden; }
  .bar-fill { background: var(--accent); height: 100%; border-radius: 5px; }
  .bar-val { font-size: 12px; font-weight: 700; width: 22px; text-align: right; flex-shrink: 0; }
  .attn-row { display: flex; justify-content: space-between; align-items: center; padding: 9px 0; border-bottom: 1px solid var(--border); font-size: 13px; cursor: pointer; }
  .attn-row:last-child { border-bottom: none; }
  .attn-name { flex: 1; padding-right: 8px; }
  .attn-days { font-weight: 700; font-size: 12px; flex-shrink: 0; }

  /* TOAST */
  .toast { position: fixed; bottom: 90px; left: 50%; transform: translateX(-50%) translateY(20px); background: #16202E; color: #fff; padding: 11px 20px; border-radius: 30px; font-size: 13.5px; z-index: 300; opacity: 0; transition: all .25s; pointer-events: none; text-align: center; max-width: 88%; }
  .toast.show { opacity: 1; transform: translateX(-50%) translateY(0); }

  /* CONFIG MODAL */
  .cfg-body p { font-size: 13.5px; line-height: 1.6; color: var(--text); margin-bottom: 10px; }
  .cfg-body code { background: var(--gray-light); padding: 2px 6px; border-radius: 4px; font-size: 12.5px; }
  .cfg-body a { color: var(--accent); }
  .cfg-body ol { padding-left: 18px; font-size: 13.5px; line-height: 1.7; }
</style>
</head>
<body>

  <div class="topbar">
    <div class="topbar-row">
      <div class="brand">
        <span class="brand-icon">🚗</span>
        <div>
          <div class="brand-title">Controle de Visitas</div>
          <div class="brand-sub">Astra &amp; Japi</div>
        </div>
      </div>
      <div style="display:flex;gap:6px">
        <button class="sync-btn" onclick="exportExcel()" title="Exportar Excel">📥</button>
        <button class="sync-btn" onclick="handleSyncClick()">⇅ <span id="syncLabel">Sincronizar</span></button>
      </div>
    </div>
    <div class="offline-bar" id="offlineBar">📡 Sem conexão — alterações salvas localmente</div>
  </div>

  <!-- LISTA -->
  <div id="view-lista">
    <div class="search-wrap">
      <span class="search-icon">🔍</span>
      <input class="search-input" id="searchInput" placeholder="Buscar por nome, CNPJ ou contato..." oninput="renderLista()">
    </div>
    <div class="chip-row" id="urgChips"></div>
    <div class="select-row">
      <select id="filterOrigem" onchange="renderLista()">
        <option value="Todos">Todas as origens</option>
        <option value="Astra">Astra</option>
        <option value="Japi">Japi</option>
      </select>
      <select id="sortBy" onchange="renderLista()">
        <option value="urgencia">Mais tempo sem visita</option>
        <option value="nome">Nome (A-Z)</option>
        <option value="compra">Última compra</option>
      </select>
    </div>
    <div class="list" id="clientList"></div>
  </div>

  <!-- DASHBOARD -->
  <div id="view-dash" style="display:none"></div>

  <div class="bottom-nav">
    <button class="nav-btn active" id="nav-lista" onclick="switchView('lista')"><span class="nav-icon">📋</span>Lista</button>
    <button class="nav-btn" id="nav-dash" onclick="switchView('dash')"><span class="nav-icon">📊</span>Dashboard</button>
    <button class="nav-btn" id="nav-novo" onclick="openNewDrawer()"><span class="nav-icon">➕</span>Novo</button>
  </div>

  <!-- OVERLAY -->
  <div class="overlay" id="overlay" onclick="closeAllDrawers()"></div>

  <!-- DRAWER: EDITAR/VER CLIENTE -->
  <div class="drawer" id="editDrawer">
    <div class="drawer-handle"></div>
    <div class="drawer-header">
      <div class="drawer-title">Cliente</div>
      <button class="close-btn" onclick="closeAllDrawers()">✕</button>
    </div>
    <div class="drawer-body" id="editDrawerBody"></div>
  </div>

  <!-- DRAWER: NOVO CLIENTE -->
  <div class="drawer" id="newDrawer">
    <div class="drawer-handle"></div>
    <div class="drawer-header">
      <div class="drawer-title">Novo cliente</div>
      <button class="close-btn" onclick="closeAllDrawers()">✕</button>
    </div>
    <div class="drawer-body">
      <div class="field"><label>Cliente *</label><input type="text" id="n-nome" placeholder="Razão social"></div>
      <div class="field"><label>CNPJ</label><input type="text" id="n-cnpj" placeholder="00.000.000/0000-00"></div>
      <div class="field"><label>Contato / referência</label><input type="text" id="n-contato" placeholder="Nome - bairro / referência"></div>
      <div class="field"><label>Última compra</label><input type="date" id="n-ultima"></div>
      <div class="field"><label>Origem</label>
        <div class="origin-grid" id="n-origem-grid">
          <div class="origin-pill" onclick="selectOrigem('Astra',this)">Astra</div>
          <div class="origin-pill" onclick="selectOrigem('Japi',this)">Japi</div>
        </div>
      </div>
      <button class="btn-primary" onclick="salvarNovo()">Salvar cliente</button>
    </div>
  </div>

  <!-- MODAL: CONFIGURAR SYNC -->
  <div class="drawer" id="cfgDrawer">
    <div class="drawer-handle"></div>
    <div class="drawer-header">
      <div class="drawer-title">Sincronização com Google Sheets</div>
      <button class="close-btn" onclick="closeAllDrawers()">✕</button>
    </div>
    <div class="drawer-body cfg-body">
      <p>Por padrão, os dados ficam salvos <strong>somente neste navegador</strong>. Para sincronizar entre aparelhos, usamos uma planilha do Google como "banco de dados" central, via Google Apps Script. Configuração única, sem necessidade de login em cada acesso:</p>
      <ol>
        <li>Crie uma planilha nova em <a href="https://sheets.new" target="_blank">sheets.new</a> (ex: "Visitas Astra Japi - Dados").</li>
        <li>No menu, vá em <strong>Extensões → Apps Script</strong>.</li>
        <li>Apague o conteúdo padrão e cole o código do arquivo <code>google-apps-script.gs</code> (fornecido junto com este app).</li>
        <li>Clique em <strong>Implantar → Nova implantação</strong>. Tipo: <strong>App da Web</strong>.</li>
        <li>Em "Quem pode acessar", escolha <strong>Qualquer pessoa</strong> e clique em Implantar (autorize as permissões pedidas).</li>
        <li>Copie a <strong>URL do app da Web</strong> gerada e cole na constante <code>APPS_SCRIPT_URL</code> no início do arquivo <code>index.html</code>.</li>
      </ol>
      <p>Depois disso, o botão "⇅ Sincronizar" salva e busca os dados diretamente nessa planilha — de qualquer aparelho, sem login.</p>
      <p style="color:var(--muted)">⚠️ Trate a URL do Web App como uma senha: quem tiver o link consegue ler e alterar os dados. Não a compartilhe publicamente.</p>
      <p style="color:var(--muted)">Enquanto isso não for configurado, o app continua funcionando normalmente, apenas offline/local.</p>
    </div>
  </div>

  <div class="toast" id="toast"></div>

<script>
/* ======================= CONFIGURAÇÃO ======================= */
// Cole aqui a URL do seu Google Apps Script (Web App) para ativar a sincronização
// via Google Sheets. Veja o passo a passo no botão "⇅ Sincronizar". Deixe vazio
// para usar apenas armazenamento local neste navegador.
const APPS_SCRIPT_URL = '';

/* ======================= DADOS SEED (importados da planilha) ======================= */
const SEED_CLIENTES = [{"cnpj": "78232865000170", "nome": "COM MAT ELETR HIDR CAJURU LTDA", "ultimaCompra": "2026-05-26", "contato": "Eduardo - Rodoviaria - Cristo Rei", "visitas": ["2025-11-17", "2026-02-04"], "origem": "Astra", "id": "seed-1"}, {"cnpj": "75056713000176", "nome": "ELETRO SAO MARCOS LTDA", "ultimaCompra": "2026-05-04", "contato": "Marcos - Bacacheri", "visitas": ["2025-09-09"], "origem": "Astra", "id": "seed-2"}, {"cnpj": "79468286000194", "nome": "H B MAT ELETR HIDR LTDA", "ultimaCompra": "2026-07-14", "contato": "Natasha - São Braz", "visitas": ["2025-12-15", "2026-04-06"], "origem": "Astra", "id": "seed-3"}, {"cnpj": "8247485000124", "nome": "MONTANA MAT CONSTR LTDA", "ultimaCompra": "2026-05-08", "contato": "São Braz", "visitas": ["2026-01-12"], "origem": "Astra", "id": "seed-4"}, {"cnpj": "11977828000110", "nome": "LOJA DO AVO COM PROD ESPEC T IDADE LT ME", "ultimaCompra": "2026-06-16", "contato": "Centro - Guadalupe", "visitas": ["2025-09-24", "2026-03-04"], "origem": "Astra", "id": "seed-5"}, {"cnpj": "48066586000113", "nome": "W S MANUT RES LTDA", "ultimaCompra": "2026-01-07", "contato": "Instalador- Marido de Aluguel", "visitas": [], "origem": "Astra", "id": "seed-6"}, {"cnpj": "73282253000179", "nome": "GULIN MAT CONSTR LTDA", "ultimaCompra": "2026-06-18", "contato": "São Braz", "visitas": ["2025-12-16"], "origem": "Astra", "id": "seed-7"}, {"cnpj": "11594300000162", "nome": "CRIATIVA COM MAT ELETR HIDR LTDA", "ultimaCompra": "2026-04-28", "contato": "Marcelo - Próximo Hipolito - São Lourenço", "visitas": ["2025-12-03"], "origem": "Astra", "id": "seed-8"}, {"cnpj": "75958439000120", "nome": "AQUECELAR COM MAT ELETR HIDR LTDA ME", "ultimaCompra": "2026-03-31", "contato": "Eduardo e Dalton", "visitas": ["2025-11-26"], "origem": "Astra", "id": "seed-9"}, {"cnpj": "78408218000177", "nome": "AQUECEBEM COM AQUECEDORES LTDA", "ultimaCompra": "2026-06-29", "contato": "Michelli - Centro", "visitas": ["2025-11-17", "2026-01-20", "2026-02-11", "2026-03-23", "2026-04-21"], "origem": "Astra", "id": "seed-10"}, {"cnpj": "180964000169", "nome": "GPAZA COM MAT ELETR HIDR LTDA", "ultimaCompra": "2026-07-17", "contato": "Gustavo - Portão", "visitas": ["2026-01-07", "2026-03-23"], "origem": "Astra", "id": "seed-11"}, {"cnpj": "29981882000103", "nome": "APARECIDA FIRMIANO COM MAT CONSTR", "ultimaCompra": "2025-12-05", "contato": "Dedéia - Disk Capanema - Jardim Botanico", "visitas": [], "origem": "Astra", "id": "seed-12"}, {"cnpj": "33227696000142", "nome": "CARLA JIONARA CORREA DO VALLE", "ultimaCompra": "2026-05-26", "contato": "Robinsin - Familia - Barreirinha", "visitas": ["2025-11-05", "2026-02-03"], "origem": "Astra", "id": "seed-13"}, {"cnpj": "51048371000184", "nome": "BARUC MMA ILUMINACAO FERR LTDA", "ultimaCompra": "2026-02-23", "contato": "Bivolt - São Lourenço", "visitas": ["2025-01-16"], "origem": "Astra", "id": "seed-14"}, {"cnpj": "1278006000198", "nome": "S 4 SCANDELARI MAT CONSTR LTDA EPP", "ultimaCompra": "2026-07-20", "contato": "Nilo -  Scandelari - Santa Candida", "visitas": ["2025-03-24", "2026-01-21"], "origem": "Astra", "id": "seed-15"}, {"cnpj": "2689584000180", "nome": "FRANCO CAVASSANI COM MAT ELETR HIDR FERR", "ultimaCompra": "2026-06-11", "contato": "Pedro - Casa Franco - Fanny", "visitas": ["2025-03-13", "2026-01-28"], "origem": "Astra", "id": "seed-16"}, {"cnpj": "30244394000194", "nome": "TOSTA SERV ELETR HIDR LTDA ME", "ultimaCompra": "2036-03-17", "contato": "SOS Reparos Tosta - Nilson - Centro", "visitas": ["2026-10-30"], "origem": "Astra", "id": "seed-17"}, {"cnpj": "40620129000117", "nome": "RECANTO CASA LAR MAT CONSTR LTDA", "ultimaCompra": "2026-07-08", "contato": "Viviane - Pedreira - Pilarzinho", "visitas": ["2025-12-12", "2026-03-20"], "origem": "Astra", "id": "seed-18"}, {"cnpj": "72229255000131", "nome": "HIDROVAR COM REPRES HIDRAULICA LTDA", "ultimaCompra": "2026-07-13", "contato": "Jose Carlos - Parolin", "visitas": ["2025-08-28", "2026-02-24", "2026-04-20"], "origem": "Astra", "id": "seed-19"}, {"cnpj": "77382190000183", "nome": "IRAI MAT CONSTR LTDA", "ultimaCompra": "2026-05-28", "contato": "Perto Atuba - Tingui", "visitas": ["2025-11-11", "2026-01-27"], "origem": "Astra", "id": "seed-20"}, {"cnpj": "81671661000140", "nome": "CASA ARDOZIA MAT CONSTR LTDA", "ultimaCompra": "2026-04-13", "contato": "Josiane - Santa Quiteria", "visitas": ["2025-12-05", "2026-03-10"], "origem": "Astra", "id": "seed-21"}, {"cnpj": "2073119000110", "nome": "GILBERTO GIL & CIA LTDA ME", "ultimaCompra": "2026-06-23", "contato": "Gilforce -  Fazendinha", "visitas": ["2025-12-09", "2026-02-26"], "origem": "Astra", "id": "seed-22"}, {"cnpj": "7923388000141", "nome": "BRENNY MAT CONSTR LTDA", "ultimaCompra": "2026-03-16", "contato": "Abranches", "visitas": ["2025-12-12"], "origem": "Astra", "id": "seed-23"}, {"cnpj": "8271498000139", "nome": "NIPO COM EQUIP ORTOPEDICOS LTDA", "ultimaCompra": "2025-04-15", "contato": "Mauro = Nipo | Esta comprando em outro cnpj", "visitas": [], "origem": "Astra", "id": "seed-24"}, {"cnpj": "24960069000125", "nome": "BIANO STORE C W B LTDA", "ultimaCompra": "2025-08-08", "contato": "Cli internet", "visitas": [], "origem": "Astra", "id": "seed-25"}, {"cnpj": "28724989000103", "nome": "ALFA COM PROD MEDICOS HOSPIT EIRELI ME", "ultimaCompra": "2025-03-27", "contato": "Fran / Cirurgica Alfa", "visitas": [], "origem": "Astra", "id": "seed-26"}, {"cnpj": "77375020000171", "nome": "LAURITA BINECK TEIXEIRA", "ultimaCompra": "2026-06-16", "contato": "Valdir - Fazendinha", "visitas": ["2025-08-27", "2026-02-18", "2026-04-15"], "origem": "Astra", "id": "seed-27"}, {"cnpj": "96663000151", "nome": "VILAS BOAS & PEDROSO LTDA", "ultimaCompra": "2026-03-01", "contato": "Mercês", "visitas": ["2025-04-29"], "origem": "Astra", "id": "seed-28"}, {"cnpj": "3133011000138", "nome": "HIDROVOLTS MAT ELETR HIDR FERR LTDA", "ultimaCompra": "2026-07-08", "contato": "Prox Guadalupe - centro", "visitas": ["2025-11-07", "2026-01-16"], "origem": "Astra", "id": "seed-29"}, {"cnpj": "5704147000186", "nome": "COML ARIGON FERR LTDA", "ultimaCompra": "2026-06-23", "contato": "Bairro Alto / Perto Godoy", "visitas": ["2025-11-17", "2026-05-04"], "origem": "Astra", "id": "seed-30"}, {"cnpj": "11088993000111", "nome": "TATA COM EQUIP SAUDE ODONTO MEDICO LTDA", "ultimaCompra": "2025-11-10", "contato": "Doutor Boca / compra somente lavatorios", "visitas": [], "origem": "Astra", "id": "seed-31"}, {"cnpj": "23258358000114", "nome": "23.258.358 CARLOS AGUINEL SILVA", "ultimaCompra": "2026-04-09", "contato": "Construsanta / Santa Candida", "visitas": ["2025-09-17", "2026-01-30"], "origem": "Astra", "id": "seed-32"}, {"cnpj": "27290941000163", "nome": "27.290.941 ISAQUE RODRIGUES DE SOUZA", "ultimaCompra": "2026-01-07", "contato": "Conexao  - Parolin", "visitas": [], "origem": "Astra", "id": "seed-33"}, {"cnpj": "42644048000182", "nome": "PILARZINHO ALMEIDA MAT CONSTR LTDA", "ultimaCompra": "2026-05-28", "contato": "Andre -  Pilarzinho", "visitas": ["2026-01-08", "2026-03-17"], "origem": "Astra", "id": "seed-34"}, {"cnpj": "53039407000199", "nome": "CASAS MARES COM VAREJISTA MAT CONSTR LTD", "ultimaCompra": "2026-06-15", "contato": "Japonesa / 2 lojas / Tuiuti", "visitas": ["2025-10-14"], "origem": "Astra", "id": "seed-35"}, {"cnpj": "68829662000102", "nome": "LUDELMA COM MAT CONSTR LTDA ME", "ultimaCompra": "2026-05-06", "contato": "Silvia - Taruma", "visitas": ["2025-11-17", "2025-01-16"], "origem": "Astra", "id": "seed-36"}, {"cnpj": "73662272000120", "nome": "JOSE DOMINGOS SIQUEIRA & FILHOS LTDA", "ultimaCompra": "2026-07-21", "contato": "Diogo - Guararema", "visitas": ["2025-08-12", "2026-02-10", "2026-03-11"], "origem": "Astra", "id": "seed-37"}, {"cnpj": "76599901000294", "nome": "EDMUNDO HAISI & CIA LTDA", "ultimaCompra": "2026-02-26", "contato": "Casa Haisi - Merces", "visitas": ["2025-11-04"], "origem": "Astra", "id": "seed-38"}, {"cnpj": "77976488000111", "nome": "STRAPASSON MAT CONSTR LTDA", "ultimaCompra": "2026-07-13", "contato": "Divonsir - Santa Quiteria", "visitas": ["2025-05-15", "2026-01-21"], "origem": "Astra", "id": "seed-39"}, {"cnpj": "1626232000112", "nome": "REI DO CHUVEIRO MAT ELETR HIDR LTDA", "ultimaCompra": "2026-07-07", "contato": "Joao Carlos - Prox Praça Rui Barbosa - Centro", "visitas": ["2026-01-06", "2026-01-28", "2026-03-19"], "origem": "Astra", "id": "seed-40"}, {"cnpj": "2024273000100", "nome": "CASA DOS REGISTROS MAT CONSTR LTDA ME", "ultimaCompra": "2026-07-09", "contato": "Domingos / Prox Pop House - Centro", "visitas": ["2025-11-12", "2026-03-01"], "origem": "Astra", "id": "seed-41"}, {"cnpj": "2969030000136", "nome": "ELETROWANN COM MAT ELETR HIDR FERR", "ultimaCompra": "2026-07-07", "contato": "Tingui / Perto Atenas", "visitas": ["2025-12-16", "2026-03-20"], "origem": "Astra", "id": "seed-42"}, {"cnpj": "4847058000126", "nome": "EVANGELISTA & EVANGELISTA LTDA", "ultimaCompra": "2026-04-07", "contato": "Piratini", "visitas": ["2025-09-29", "2026-01-28"], "origem": "Astra", "id": "seed-43"}, {"cnpj": "11927821000194", "nome": "NILSON C COUTINHO COML HIDROELETRICA", "ultimaCompra": "2026-06-23", "contato": "Solução - Parolin", "visitas": ["2025-12-02", "2026-01-27", "2026-03-16", "2026-05-05"], "origem": "Astra", "id": "seed-44"}, {"cnpj": "33619998000166", "nome": "MAT CONSTR FANNY LTDA", "ultimaCompra": "2026-07-07", "contato": "Andre - Fanny", "visitas": ["2025-12-03", "2026-01-28", "2026-04-07"], "origem": "Astra", "id": "seed-45"}, {"cnpj": "44772422000150", "nome": "JULIO TOSIN 62882961987", "ultimaCompra": "2026-01-08", "contato": "Andrea - Av Iguaçu", "visitas": [], "origem": "Astra", "id": "seed-46"}, {"cnpj": "48203522000117", "nome": "MAT ELETR COLOMBO LTDA", "ultimaCompra": "2026-07-15", "contato": "Boa Vista", "visitas": ["2025-09-03", "2026-01-26"], "origem": "Astra", "id": "seed-47"}, {"cnpj": "75118679000117", "nome": "CASA CONEXAO DE MAT HIDRAULICO LTDA", "ultimaCompra": "2026-07-10", "contato": "Sampaio - Prox Zat - Prado Velho", "visitas": ["2025-12-12", "2026-02-12", "2026-04-01", "2026-04-23", "2026-06-10"], "origem": "Astra", "id": "seed-48"}, {"cnpj": "76486810000161", "nome": "BUCHHOLTZ & CIA LTDA", "ultimaCompra": "2026-07-10", "contato": "Ricardo - Merces", "visitas": ["2025-09-05", "2026-02-03", "2026-04-10", "2026-04-27"], "origem": "Astra", "id": "seed-49"}, {"cnpj": "76631324000190", "nome": "SIQUEIRA & CIA MAT CONSTR LTDA", "ultimaCompra": "2026-03-12", "contato": "Salete - Campina do Siqueira", "visitas": ["2025-08-21"], "origem": "Astra", "id": "seed-50"}, {"cnpj": "82222696000165", "nome": "NEURI CARLOS GODOY ME", "ultimaCompra": "2026-06-23", "contato": "Godoy - Bairro Alto", "visitas": ["2025-11-17"], "origem": "Astra", "id": "seed-51"}, {"cnpj": "85042737000120", "nome": "J PERIN COM MAT CONSTR LTDA ME", "ultimaCompra": "2026-07-17", "contato": "Luana - Boa Vista", "visitas": ["2025-11-11", "2026-01-26", "2026-05-26"], "origem": "Astra", "id": "seed-52"}, {"cnpj": "95448882000169", "nome": "DEL MAZO COM MAT CONSTR SERV LTDA", "ultimaCompra": "2026-01-29", "contato": "Passeio Publico - Alto da Gloria", "visitas": ["2025-07-22"], "origem": "Astra", "id": "seed-53"}, {"cnpj": "1688758000127", "nome": "1001 MAT CONSTR LTDA ME", "ultimaCompra": "2026-07-03", "contato": "Paulo - Capao da Imbuia", "visitas": ["2025-10-09", "2026-02-06"], "origem": "Astra", "id": "seed-54"}, {"cnpj": "2053663000108", "nome": "WARELLA COM MAT CONSTR LTDA", "ultimaCompra": "2026-06-08", "contato": "Fabio - Jardim Social", "visitas": ["2026-01-13", "2026-03-16"], "origem": "Astra", "id": "seed-55"}, {"cnpj": "2377070000199", "nome": "MATSUI COM MAT ELETR HIDR LTDA", "ultimaCompra": "2026-07-16", "contato": "Cristo Rei", "visitas": ["2025-12-08", "2026-04-10", "2026-04-20"], "origem": "Astra", "id": "seed-56"}, {"cnpj": "2742848000111", "nome": "CASAS MARES MAT CONSTR LTDA", "ultimaCompra": "2026-01-12", "contato": "Santo Inacio", "visitas": [], "origem": "Astra", "id": "seed-57"}, {"cnpj": "4442353000100", "nome": "E H F ELETR HIDR LTDA", "ultimaCompra": "2026-03-30", "contato": "Lucas - Santa Felicidade", "visitas": ["2025-07-21"], "origem": "Astra", "id": "seed-58"}, {"cnpj": "4576403000134", "nome": "TELHAFER COM MAT CONSTR LTDA", "ultimaCompra": "2026-04-14", "contato": "Anderson - Abranches", "visitas": ["2025-09-24", "2026-03-20"], "origem": "Astra", "id": "seed-59"}, {"cnpj": "5294648000131", "nome": "GOBI MAT CONSTR LTDA", "ultimaCompra": "2026-07-14", "contato": "Sheila - Vila Guaira", "visitas": ["2025-12-03", "2026-03-13"], "origem": "Astra", "id": "seed-60"}, {"cnpj": "5459175000185", "nome": "CASA FORTE MAT CONSTR LTDA", "ultimaCompra": "2026-07-15", "contato": "Marcelo - Santa Candida", "visitas": ["2026-01-13", "2026-01-28", "2026-02-12", "2026-03-19", "2026-03-27", "2026-04-02", "2026-06-01"], "origem": "Astra", "id": "seed-61"}, {"cnpj": "58410673000108", "nome": "58.410.673/0001-08 - DANILO WISCHNIEVSKI", "ultimaCompra": "2026-04-10", "contato": "Diogo Mat Const", "visitas": ["2025-12-12"], "origem": "Astra", "id": "seed-62"}, {"cnpj": "7005521000180", "nome": "DELTA D AGUA APARELHOS A GAS LTDA ME", "ultimaCompra": "2025-07-14", "contato": "Oscar - Pilarzinho", "visitas": [], "origem": "Astra", "id": "seed-63"}, {"cnpj": "7107918000183", "nome": "COM ASSENTOS SANIT PEREIRA AMERICO LTDA", "ultimaCompra": "2026-06-10", "contato": "Ricardo - São Lourenço", "visitas": ["2026-01-19", "2025-12-03", "2026-02-03", "2026-03-12"], "origem": "Astra", "id": "seed-64"}, {"cnpj": "7296109000167", "nome": "LAPOLA & BECEGATO MAT CONSTR LTDA ME", "ultimaCompra": "2026-07-01", "contato": "Gabriel - Barreirinha", "visitas": ["2025-11-11", "2025-01-14", "2026-02-23", "2026-03-18"], "origem": "Astra", "id": "seed-65"}, {"cnpj": "12669695000188", "nome": "L D I COM PROD ORTOPEDICOS LTDA ME", "ultimaCompra": "2025-07-04", "contato": "Nipo", "visitas": [], "origem": "Astra", "id": "seed-66"}, {"cnpj": "15287459000195", "nome": "LONGINO CARLOS SOCZEK ME", "ultimaCompra": "2026-06-22", "contato": "Abdul - Armazem do Eletro - Centro", "visitas": ["2025-11-12"], "origem": "Astra", "id": "seed-67"}, {"cnpj": "23720971000101", "nome": "V T DE SOUZA MAT CONSTR ME", "ultimaCompra": "2026-03-20", "contato": "Walter - Batex - Bairro Alto", "visitas": ["2026-01-23"], "origem": "Astra", "id": "seed-68"}, {"cnpj": "28361828000194", "nome": "B D N AQUECIMENTO LTDA", "ultimaCompra": "2025-10-24", "contato": "Danilo - Centro = Volpato", "visitas": [], "origem": "Astra", "id": "seed-69"}, {"cnpj": "15630554000140", "nome": "VOLPATO ASSUNCAO E CIA LTDA ME", "ultimaCompra": "2026-06-30", "contato": "Danilo - Centro = B D N", "visitas": ["2025-12-09"], "origem": "Astra", "id": "seed-70"}, {"cnpj": "29178862000190", "nome": "MURATRON COM EQUIP MAT ELETRO ELETRN LTD", "ultimaCompra": "2026-07-09", "contato": "Rodrigo - Portão", "visitas": ["2025-10-14", "2026-01-29", "2026-06-02"], "origem": "Astra", "id": "seed-71"}, {"cnpj": "33852528000148", "nome": "T S IPOLITA MAT CONSTR EIRELI", "ultimaCompra": "2026-07-13", "contato": "Hipolito - Boa Vista", "visitas": ["2025-12-02", "2026-03-10", "2026-05-13"], "origem": "Astra", "id": "seed-72"}, {"cnpj": "38299332000191", "nome": "REI REPAROS MAT HIDR LTDA", "ultimaCompra": "2026-05-27", "contato": "Gerson - Bacacheri", "visitas": ["2025-11-05", "2026-01-27", "2026-03-18"], "origem": "Astra", "id": "seed-73"}, {"cnpj": "49183496000175", "nome": "ANTONIO JOSE CORREA LTDA", "ultimaCompra": "2026-03-10", "contato": "JKAJ Licitação", "visitas": ["2025-07-16"], "origem": "Astra", "id": "seed-74"}, {"cnpj": "75103739000128", "nome": "MANACA MAT CONSTR LTDA", "ultimaCompra": "2026-03-20", "contato": "Bairro Alto", "visitas": ["2025-07-04", "2026-01-23"], "origem": "Astra", "id": "seed-75"}, {"cnpj": "77373801000127", "nome": "LAPOLA & CORADASSI LTDA", "ultimaCompra": "2026-06-26", "contato": "Loja Bella Vista - Boa Vista", "visitas": ["2025-06-23", "2026-01-26"], "origem": "Astra", "id": "seed-76"}, {"cnpj": "78697455000103", "nome": "BIZ COM FERR REPRES LTDA", "ultimaCompra": "2026-06-15", "contato": "Bruno - Boa Vista - Av Parana", "visitas": ["2025-07-24", "2026-01-27"], "origem": "Astra", "id": "seed-77"}, {"cnpj": "79056669000155", "nome": "COML DISTR ARGENTA LTDA", "ultimaCompra": "2026-07-08", "contato": "São Francisco - Prox Mueller", "visitas": ["2026-01-09"], "origem": "Astra", "id": "seed-78"}, {"cnpj": "79109088000134", "nome": "ATENAS MAT CONSTR LTDA", "ultimaCompra": "2026-07-15", "contato": "Fortes - Bacacheri", "visitas": ["2025-12-10", "2036-03-20"], "origem": "Astra", "id": "seed-79"}, {"cnpj": "79166989000168", "nome": "CASA DOS ASSENTOS AMERICOS LTDA", "ultimaCompra": "2026-07-09", "contato": "Rafael - Rebouças", "visitas": ["2026-01-07", "2026-02-03", "2036-03-16", "2026-03-20", "2026-05-15", "2026-06-11", "2026-06-25"], "origem": "Astra", "id": "seed-80"}, {"cnpj": "80816465000154", "nome": "SIDMAR ELETRO LTDA ME", "ultimaCompra": "2026-07-13", "contato": "Leonardo - Alto da XV", "visitas": ["2025-12-15", "2026-01-26", "2036-03-19", "2026-05-12"], "origem": "Astra", "id": "seed-81"}, {"cnpj": "80851330000120", "nome": "COM MAT CONSTR W R DE ALMEIDA LTDA", "ultimaCompra": "2026-05-27", "contato": "WR - Pilarzinho", "visitas": ["2026-01-07", "2026-03-18", "2026-03-28"], "origem": "Astra", "id": "seed-82"}, {"cnpj": "81679508000160", "nome": "PURKOTT & MOLETTA LTDA ME", "ultimaCompra": "2026-07-10", "contato": "Murilo - Aliança - Santa Candida", "visitas": ["2025-11-10", "2026-03-19", "2026-05-11"], "origem": "Astra", "id": "seed-83"}, {"cnpj": "82285529000163", "nome": "P W COML HIDR LTDA", "ultimaCompra": "2026-05-28", "contato": "Regina - São Lourenço", "visitas": ["2025-11-05", "2026-02-09", "2026-03-27"], "origem": "Astra", "id": "seed-84"}, {"cnpj": "82288952000117", "nome": "O G COML ELETRO FERR LTDA ME", "ultimaCompra": "2026-05-07", "contato": "Casa do Chuveiro - Centro", "visitas": ["2025-10-21", "2026-02-24"], "origem": "Astra", "id": "seed-85"}, {"cnpj": "82436007000115", "nome": "E W G COM MAT ELETR HIDR LTDA", "ultimaCompra": "2026-06-15", "contato": "Elaine - Vila Izabel", "visitas": ["2026-01-14", "2026-01-30", "2026-03-09", "2026-04-23"], "origem": "Astra", "id": "seed-86"}, {"cnpj": "84966738000106", "nome": "FERR DONDA LTDA", "ultimaCompra": "2026-06-15", "contato": "Barreirinha", "visitas": ["2026-01-14"], "origem": "Astra", "id": "seed-87"}, {"cnpj": "84987999000102", "nome": "ZZAT MAT CONSTR LTDA", "ultimaCompra": "2026-04-30", "contato": "Lia - Rebouças", "visitas": ["2025-11-21", "2026-01-23", "2026-02-18", "2026-03-19", "2026-04-30"], "origem": "Astra", "id": "seed-88"}, {"cnpj": "80372493000120", "nome": "FERMATTI COM FERR LTDA", "ultimaCompra": "2026-05-28", "contato": "Pilarzinho", "visitas": ["2025-10-21", "2026-07-03"], "origem": "Astra", "id": "seed-89"}, {"cnpj": "61107693000100", "nome": "A3 FERRM UTIL LTDA", "ultimaCompra": "2026-04-13", "contato": "Fernando - G2 - Lindoia", "visitas": ["2025-12-02", "2026-03-09"], "origem": "Astra", "id": "seed-90"}, {"cnpj": "22072818000152", "nome": "CONTAINER ROTATIVO LTDA", "ultimaCompra": "2026-07-03", "contato": "Leonardo - Umbara", "visitas": ["2025-12-10", "2026-02-05", "2026-03-31", "2026-05-28"], "origem": "Astra", "id": "seed-91"}, {"cnpj": "84963685000161", "nome": "ZACARIAS PEREIRA DE ARAUJO ME", "ultimaCompra": "2026-06-22", "contato": "Zaca - Fazendinha", "visitas": ["2026-01-08"], "origem": "Astra", "id": "seed-92"}, {"cnpj": "3565616000106", "nome": "TERRA PROMETIDA MATERIAIS PARA CONST LTDA", "ultimaCompra": "2026-01-08", "contato": "Adalgisa - Fazendinha", "visitas": [], "origem": "Astra", "id": "seed-93"}, {"cnpj": "5078014000141", "nome": "LOJAO DO ALOISIO", "ultimaCompra": "2026-01-07", "contato": "Aloisio - Catedral - Centro", "visitas": [], "origem": "Astra", "id": "seed-94"}, {"cnpj": "12415821000178", "nome": "DEP CASA FORTE LTDA", "ultimaCompra": "2026-07-22", "contato": "Celso - Santa Candida", "visitas": ["2026-01-21", "2026-03-17"], "origem": "Astra", "id": "seed-95"}, {"cnpj": "38542019000132", "nome": "CAMPOS E FARIAS LTDA", "ultimaCompra": "2026-02-19", "contato": "Lana - Centro Tam | Devendo", "visitas": [], "origem": "Astra", "id": "seed-96"}, {"cnpj": "61504570000103", "nome": "J e C MATERIAIS DE CONSTRUCAO LTDA", "ultimaCompra": "2026-06-04", "contato": "Eduardo - Centro", "visitas": ["2026-03-10", "2026-05-02"], "origem": "Astra", "id": "seed-97"}, {"cnpj": "24971930000150", "nome": "ARTEMIO SANTOS DE SOUZA ME", "ultimaCompra": "2026-03-11", "contato": "Artemio - Figuras - Novo Mundo", "visitas": [], "origem": "Astra", "id": "seed-98"}, {"cnpj": "4384908000105", "nome": "ELETROZINA MAT ELETR HIDR LTDA", "ultimaCompra": "2026-03-16", "contato": "Valdinei - Alto da XV", "visitas": [], "origem": "Astra", "id": "seed-99"}, {"cnpj": "10861744000154", "nome": "ORTOMEDICA CAPITAL LTDA", "ultimaCompra": "2026-07-24", "contato": "Mauro = Nipo", "visitas": ["2026-03-19"], "origem": "Astra", "id": "seed-100"}, {"cnpj": "5650919000144", "nome": "J D M MAT CONSTR LTDA", "ultimaCompra": "2026-03-17", "contato": "Jose - Vila Fanny", "visitas": [], "origem": "Astra", "id": "seed-101"}, {"cnpj": "6576044000140", "nome": "HOSPINET COM E ASSISTENCIA TECNICA LTDA", "ultimaCompra": "2026-01-20", "contato": "Leticia - Dado = Hospitel", "visitas": [], "origem": "Astra", "id": "seed-102"}, {"cnpj": "85008274000180", "nome": "CELINO DE SOUZA E CIA LTDA ME", "ultimaCompra": "2026-05-07", "contato": "Murilo - Novo Mundo mat const", "visitas": [], "origem": "Astra", "id": "seed-103"}, {"cnpj": "63084824000107", "nome": "S D C MAT SERV LTDA", "ultimaCompra": "2026-04-24", "contato": "Suzan - Licitaçoes", "visitas": [], "origem": "Astra", "id": "seed-104"}, {"cnpj": "24297355000104", "nome": "LAERCIO DE OLIVEIRA", "ultimaCompra": "2026-05-28", "contato": "Laercio - Capanema", "visitas": [], "origem": "Astra", "id": "seed-105"}, {"cnpj": "30381186000137", "nome": "F L AQUECEDORES A GAS E SOLAR LTDA", "ultimaCompra": "2026-07-07", "contato": "Kelli - Fanny", "visitas": [], "origem": "Astra", "id": "seed-106"}, {"cnpj": "11419279000169", "nome": "KARESI MAT CONSTR DO VIZINHO LTDA", "ultimaCompra": "2026-07-03", "contato": "", "visitas": [], "origem": "Astra", "id": "seed-107"}, {"cnpj": "7881860000120", "nome": "A ZEM MAT ELETR", "ultimaCompra": "2026-07-17", "contato": "Antonio - Fazendinha", "visitas": [], "origem": "Astra", "id": "seed-108"}, {"cnpj": "80271950000190", "nome": "ALTIMAR MEDEIROS SANTOS", "ultimaCompra": null, "contato": "Adrianopolis", "visitas": [], "origem": "Astra", "id": "seed-109"}, {"cnpj": "79629457000110", "nome": "MARLI APARECIDA BONTORIN", "ultimaCompra": null, "contato": "Adrianopolis", "visitas": [], "origem": "Astra", "id": "seed-110"}, {"cnpj": "76152578000125", "nome": "B S COM MAT CONST LTDA ME", "ultimaCompra": null, "contato": "Bocaiuva do Sul", "visitas": [], "origem": "Astra", "id": "seed-111"}, {"cnpj": "9406759000143", "nome": "CORADIN & LOVATO MAT CONSTR", "ultimaCompra": null, "contato": "Bocaiuva do Sul", "visitas": [], "origem": "Astra", "id": "seed-112"}, {"cnpj": "", "nome": "Mercado Xambre Com e Exportação LTDA ME", "ultimaCompra": "2026-01-16", "contato": "Darci - Vista Alegre", "visitas": [], "origem": "Japi", "id": "seed-113"}, {"cnpj": "", "nome": "BUCHHOLTZ & CIA LTDA", "ultimaCompra": "2026-07-10", "contato": "Ricardo - Merces", "visitas": ["2026-04-28"], "origem": "Japi", "id": "seed-114"}, {"cnpj": "", "nome": "FERR DONDA LTDA", "ultimaCompra": "2026-04-28", "contato": "Barreirinha", "visitas": [], "origem": "Japi", "id": "seed-115"}, {"cnpj": "", "nome": "LAURITA BINECK TEIXEIRA", "ultimaCompra": "2026-04-15", "contato": "Fazendinha", "visitas": [], "origem": "Japi", "id": "seed-116"}, {"cnpj": "", "nome": "MAT CONSTR FANNY LTDA", "ultimaCompra": "2026-04-07", "contato": "Andre", "visitas": [], "origem": "Japi", "id": "seed-117"}, {"cnpj": "", "nome": "GULIN MAT CONSTR LTDA", "ultimaCompra": "2026-05-08", "contato": "Ação Brinde  - copo + garrafa", "visitas": [], "origem": "Japi", "id": "seed-118"}, {"cnpj": "", "nome": "MONTANA MAT CONST LTDA", "ultimaCompra": "2026-05-08", "contato": "Ação Brinde  - copo + garrafa", "visitas": [], "origem": "Japi", "id": "seed-119"}, {"cnpj": "", "nome": "TERRA PROMETIDA MATERIAIS PARA CONST LTDA", "ultimaCompra": "2026-05-11", "contato": "Ação Brinde - garrafa", "visitas": [], "origem": "Japi", "id": "seed-120"}, {"cnpj": "", "nome": "SIDMAR ELETRO LTDA ME", "ultimaCompra": "2026-07-13", "contato": "", "visitas": ["2026-05-12"], "origem": "Japi", "id": "seed-121"}, {"cnpj": "", "nome": "PURKOTT & MOLETTA LTDA ME", "ultimaCompra": "2026-05-12", "contato": "Ação Brinde - garrafa", "visitas": [], "origem": "Japi", "id": "seed-122"}, {"cnpj": "", "nome": "FERMATTI COM FERR LTDA", "ultimaCompra": "2026-05-28", "contato": "", "visitas": [], "origem": "Japi", "id": "seed-123"}, {"cnpj": "", "nome": "PILARZINHO ALMEIDA MAT CONSTR LTDA", "ultimaCompra": "2026-05-28", "contato": "", "visitas": [], "origem": "Japi", "id": "seed-124"}, {"cnpj": "", "nome": "COM MAT ELETR HIDR CAJURU LTDA", "ultimaCompra": "2026-05-26", "contato": "", "visitas": [], "origem": "Japi", "id": "seed-125"}, {"cnpj": "", "nome": "CARLA JIONARA CORREA DO VALLE", "ultimaCompra": "2026-05-26", "contato": "", "visitas": [], "origem": "Japi", "id": "seed-126"}, {"cnpj": "", "nome": "J PERIN COM MAT CONSTR LTDA ME", "ultimaCompra": "2026-05-26", "contato": "", "visitas": [], "origem": "Japi", "id": "seed-127"}, {"cnpj": "", "nome": "VOLPATO ASSUNCAO E CIA LTDA ME", "ultimaCompra": "2026-06-30", "contato": "", "visitas": [], "origem": "Japi", "id": "seed-128"}, {"cnpj": "", "nome": "FLORES DE JUNHO DECORAÇOES E ARRANJOS", "ultimaCompra": "2026-07-02", "contato": "", "visitas": [], "origem": "Japi", "id": "seed-129"}, {"cnpj": "", "nome": "T S IPOLITA MAT CONSTR EIRELI", "ultimaCompra": "2026-07-13", "contato": "", "visitas": [], "origem": "Japi", "id": "seed-130"}];

/* ======================= ESTADO ======================= */
let clientes = [];
let currentView = 'lista';
let editingId = null;
let newOrigem = 'Astra';
let syncing = false;

const URG_CHIPS = [
  { key: 'todos', label: 'Todos' },
  { key: 'nunca', label: '🔴 Nunca visitado' },
  { key: '60', label: '+60 dias' },
  { key: '30', label: '+30 dias' },
  { key: 'ok', label: '✅ Em dia' },
];
let activeUrgChip = 'todos';

/* ======================= PERSISTÊNCIA LOCAL ======================= */
function loadClientes() {
  try {
    const raw = localStorage.getItem('visitas_clientes');
    if (raw) { clientes = JSON.parse(raw); return; }
  } catch (e) {}
  clientes = JSON.parse(JSON.stringify(SEED_CLIENTES));
  saveClientes(false);
}

function saveClientes(sync) {
  localStorage.setItem('visitas_clientes', JSON.stringify(clientes));
  if (sync !== false) maybeAutoSync();
}

/* ======================= HELPERS DE DATA ======================= */
function todayISO() { return new Date().toISOString().split('T')[0]; }

function fmtDate(iso) {
  if (!iso) return '—';
  const p = iso.split('-');
  if (p.length !== 3) return iso;
  return p[2] + '/' + p[1] + '/' + p[0];
}

function diasDesde(iso) {
  if (!iso) return null;
  const d1 = new Date(iso + 'T00:00:00');
  const d2 = new Date(todayISO() + 'T00:00:00');
  return Math.round((d2 - d1) / 86400000);
}

function ultimaVisita(c) {
  if (!c.visitas || c.visitas.length === 0) return null;
  return c.visitas.slice().sort().reverse()[0];
}

function urgenciaInfo(c) {
  const uv = ultimaVisita(c);
  if (!uv) return { dias: null, cor: 'red', label: 'Nunca visitado' };
  const dias = diasDesde(uv);
  if (dias > 60) return { dias, cor: 'red', label: dias + ' dias sem visita' };
  if (dias > 30) return { dias, cor: 'amber', label: dias + ' dias sem visita' };
  return { dias, cor: 'green', label: dias + ' dias sem visita' };
}

/* ======================= NAVEGAÇÃO ======================= */
function switchView(v) {
  currentView = v;
  document.getElementById('view-lista').style.display = v === 'lista' ? '' : 'none';
  document.getElementById('view-dash').style.display = v === 'dash' ? '' : 'none';
  document.getElementById('nav-lista').classList.toggle('active', v === 'lista');
  document.getElementById('nav-dash').classList.toggle('active', v === 'dash');
  if (v === 'lista') renderLista();
  if (v === 'dash') renderDashboard();
}

/* ======================= LISTA ======================= */
function renderUrgChips() {
  const wrap = document.getElementById('urgChips');
  wrap.innerHTML = URG_CHIPS.map(ch =>
    `<div class="chip ${ch.key === activeUrgChip ? 'active' : ''}" onclick="setUrgChip('${ch.key}')">${ch.label}</div>`
  ).join('');
}

function setUrgChip(key) {
  activeUrgChip = key;
  renderUrgChips();
  renderLista();
}

function passaUrgencia(c, key) {
  const info = urgenciaInfo(c);
  if (key === 'todos') return true;
  if (key === 'nunca') return info.dias === null;
  if (key === '60') return info.dias !== null && info.dias > 60;
  if (key === '30') return info.dias !== null && info.dias > 30;
  if (key === 'ok') return info.dias !== null && info.dias <= 30;
  return true;
}

function renderLista() {
  renderUrgChips();
  const q = document.getElementById('searchInput').value.trim().toLowerCase();
  const origem = document.getElementById('filterOrigem').value;
  const sortBy = document.getElementById('sortBy').value;

  let lista = clientes.filter(c => {
    if (origem !== 'Todos' && c.origem !== origem) return false;
    if (!passaUrgencia(c, activeUrgChip)) return false;
    if (q) {
      const hay = (c.nome + ' ' + (c.cnpj||'') + ' ' + (c.contato||'')).toLowerCase();
      if (!hay.includes(q)) return false;
    }
    return true;
  });

  if (sortBy === 'nome') {
    lista.sort((a,b) => a.nome.localeCompare(b.nome));
  } else if (sortBy === 'compra') {
    lista.sort((a,b) => (b.ultimaCompra||'').localeCompare(a.ultimaCompra||''));
  } else {
    lista.sort((a,b) => {
      const da = urgenciaInfo(a).dias; const db = urgenciaInfo(b).dias;
      if (da === null && db === null) return 0;
      if (da === null) return -1;
      if (db === null) return 1;
      return db - da;
    });
  }

  const wrap = document.getElementById('clientList');
  if (lista.length === 0) {
    wrap.innerHTML = `<div class="empty-state"><div>🔍</div>Nenhum cliente encontrado</div>`;
    return;
  }

  wrap.innerHTML = lista.map(c => {
    const info = urgenciaInfo(c);
    const badgeClass = c.origem === 'Astra' ? 'badge-blue' : 'badge-green';
    return `
    <div class="card" onclick="openEditDrawer('${c.id}')">
      <div class="card-top">
        <div class="alert-dot dot-${info.cor}" style="margin-top:5px"></div>
        <div class="card-name">${c.nome}</div>
        <span class="badge ${badgeClass}">${c.origem}</span>
      </div>
      <div class="card-meta">
        <span>${info.label}</span>
        <span>·</span>
        <span>Última compra: ${fmtDate(c.ultimaCompra)}</span>
      </div>
      ${c.contato ? `<div class="card-obs">${c.contato}</div>` : ''}
    </div>`;
  }).join('');
}

/* ======================= DRAWER: EDITAR CLIENTE ======================= */
function openEditDrawer(id) {
  editingId = id;
  const c = clientes.find(x => x.id === id);
  if (!c) return;
  const info = urgenciaInfo(c);
  const visitasOrdenadas = (c.visitas || []).slice().sort().reverse();

  document.getElementById('editDrawerBody').innerHTML = `
    <div class="field"><label>Cliente</label><input type="text" id="e-nome" value="${(c.nome||'').replace(/"/g,'&quot;')}"></div>
    <div class="field-row">
      <div class="field"><label>CNPJ</label><input type="text" id="e-cnpj" value="${c.cnpj||''}"></div>
      <div class="field"><label>Origem</label>
        <select id="e-origem"><option value="Astra" ${c.origem==='Astra'?'selected':''}>Astra</option><option value="Japi" ${c.origem==='Japi'?'selected':''}>Japi</option></select>
      </div>
    </div>
    <div class="field"><label>Contato / referência</label><input type="text" id="e-contato" value="${(c.contato||'').replace(/"/g,'&quot;')}"></div>
    <div class="field"><label>Última compra</label><input type="date" id="e-ultima" value="${c.ultimaCompra||''}"></div>

    <div>
      <div class="section-label" style="margin-bottom:8px">${info.dias === null ? '🔴 Nunca visitado' : info.label}</div>
      <button class="btn-today" onclick="registrarVisitaHoje()">📍 Registrar visita hoje</button>
    </div>

    <div>
      <div class="section-label" style="margin-bottom:8px">Histórico de visitas (${visitasOrdenadas.length})</div>
      <div class="add-visit-row" style="margin-bottom:10px">
        <input type="date" id="e-nova-visita">
        <button class="btn-ghost" onclick="adicionarVisitaData()">Adicionar</button>
      </div>
      <div class="visit-list" id="visitList">
        ${visitasOrdenadas.length ? visitasOrdenadas.map(v => `
          <div class="visit-item"><span>${fmtDate(v)}</span><button onclick="removerVisita('${v}')">Remover</button></div>
        `).join('') : '<div style="color:var(--muted);font-size:13px">Nenhuma visita registrada ainda.</div>'}
      </div>
    </div>

    <button class="btn-primary" onclick="salvarEdicao()">Salvar alterações</button>
    <button class="btn-danger" onclick="excluirCliente()">Excluir cliente</button>
  `;

  document.getElementById('overlay').classList.add('show');
  document.getElementById('editDrawer').classList.add('open');
}

function registrarVisitaHoje() {
  const c = clientes.find(x => x.id === editingId);
  if (!c) return;
  const hoje = todayISO();
  c.visitas = c.visitas || [];
  if (!c.visitas.includes(hoje)) c.visitas.push(hoje);
  saveClientes();
  openEditDrawer(editingId);
  showToast('Visita registrada hoje ✅');
}

function adicionarVisitaData() {
  const v = document.getElementById('e-nova-visita').value;
  if (!v) { showToast('Escolha uma data'); return; }
  const c = clientes.find(x => x.id === editingId);
  if (!c) return;
  c.visitas = c.visitas || [];
  if (!c.visitas.includes(v)) c.visitas.push(v);
  saveClientes();
  openEditDrawer(editingId);
}

function removerVisita(v) {
  const c = clientes.find(x => x.id === editingId);
  if (!c) return;
  c.visitas = (c.visitas || []).filter(x => x !== v);
  saveClientes();
  openEditDrawer(editingId);
}

function salvarEdicao() {
  const c = clientes.find(x => x.id === editingId);
  if (!c) return;
  c.nome = document.getElementById('e-nome').value.trim() || c.nome;
  c.cnpj = document.getElementById('e-cnpj').value.trim();
  c.origem = document.getElementById('e-origem').value;
  c.contato = document.getElementById('e-contato').value.trim();
  c.ultimaCompra = document.getElementById('e-ultima').value || null;
  saveClientes();
  closeAllDrawers();
  renderLista();
  showToast('Cliente atualizado');
}

function excluirCliente() {
  if (!confirm('Excluir este cliente? Essa ação não pode ser desfeita.')) return;
  clientes = clientes.filter(x => x.id !== editingId);
  saveClientes();
  closeAllDrawers();
  renderLista();
  showToast('Cliente excluído');
}

/* ======================= DRAWER: NOVO CLIENTE ======================= */
function openNewDrawer() {
  document.getElementById('n-nome').value = '';
  document.getElementById('n-cnpj').value = '';
  document.getElementById('n-contato').value = '';
  document.getElementById('n-ultima').value = '';
  newOrigem = 'Astra';
  document.querySelectorAll('#n-origem-grid .origin-pill').forEach(p => p.className = 'origin-pill');
  document.querySelector('#n-origem-grid .origin-pill').className = 'origin-pill sel-Astra';
  document.getElementById('overlay').classList.add('show');
  document.getElementById('newDrawer').classList.add('open');
}

function selectOrigem(o, el) {
  newOrigem = o;
  document.querySelectorAll('#n-origem-grid .origin-pill').forEach(p => p.className = 'origin-pill');
  el.className = 'origin-pill sel-' + o;
}

function salvarNovo() {
  const nome = document.getElementById('n-nome').value.trim();
  if (!nome) { showToast('Informe o nome do cliente'); return; }
  const novo = {
    id: 'c-' + Date.now() + '-' + Math.floor(Math.random()*10000),
    nome,
    cnpj: document.getElementById('n-cnpj').value.trim(),
    contato: document.getElementById('n-contato').value.trim(),
    ultimaCompra: document.getElementById('n-ultima').value || null,
    visitas: [],
    origem: newOrigem
  };
  clientes.push(novo);
  saveClientes();
  closeAllDrawers();
  switchView('lista');
  showToast('Cliente cadastrado');
}

/* ======================= DRAWERS GERAIS ======================= */
function closeAllDrawers() {
  document.getElementById('overlay').classList.remove('show');
  ['editDrawer','newDrawer','cfgDrawer'].forEach(id => document.getElementById(id).classList.remove('open'));
  editingId = null;
}

function showToast(msg) {
  const t = document.getElementById('toast');
  t.textContent = msg;
  t.classList.add('show');
  setTimeout(() => t.classList.remove('show'), 2200);
}

/* ======================= DASHBOARD ======================= */
function renderDashboard() {
  const total = clientes.length;
  const semVisita60 = clientes.filter(c => { const i = urgenciaInfo(c); return i.dias !== null && i.dias > 60; }).length;
  const nuncaVisitado = clientes.filter(c => urgenciaInfo(c).dias === null).length;
  const inicio30 = new Date(); inicio30.setDate(inicio30.getDate() - 30);
  const visitas30 = clientes.reduce((acc, c) => acc + (c.visitas||[]).filter(v => new Date(v+'T00:00:00') >= inicio30).length, 0);

  // --- Visitas por mês (últimos 6 meses) ---
  const meses = [];
  const now = new Date();
  for (let i = 5; i >= 0; i--) {
    const d = new Date(now.getFullYear(), now.getMonth() - i, 1);
    meses.push({ key: d.getFullYear() + '-' + String(d.getMonth()+1).padStart(2,'0'), label: d.toLocaleDateString('pt-BR',{month:'short'}).replace('.','') });
  }
  const contagem = {};
  meses.forEach(m => contagem[m.key] = 0);
  clientes.forEach(c => (c.visitas||[]).forEach(v => { const k = v.slice(0,7); if (contagem[k] !== undefined) contagem[k]++; }));
  const maxVisitas = Math.max(1, ...meses.map(m => contagem[m.key]));

  // --- Atenção: sem visita há mais tempo (top 8) ---
  const semVisitaLista = clientes
    .map(c => ({ c, info: urgenciaInfo(c) }))
    .filter(x => x.info.dias === null || x.info.dias > 30)
    .sort((a,b) => { if (a.info.dias === null) return -1; if (b.info.dias === null) return 1; return b.info.dias - a.info.dias; })
    .slice(0, 8);

  // --- Comprou recente mas não é visitado há tempo (top 8) ---
  const compraVsVisita = clientes
    .filter(c => c.ultimaCompra)
    .map(c => ({ c, diasCompra: diasDesde(c.ultimaCompra), diasVisita: urgenciaInfo(c).dias }))
    .filter(x => x.diasCompra <= 60 && (x.diasVisita === null || x.diasVisita > 45))
    .sort((a,b) => a.diasCompra - b.diasCompra)
    .slice(0, 8);

  const wrap = document.getElementById('view-dash');
  wrap.innerHTML = `
    <div class="dash-wrap">
      <div class="kpi-grid">
        <div class="kpi-card"><div class="kpi-num">${total}</div><div class="kpi-label">Clientes ativos</div></div>
        <div class="kpi-card"><div class="kpi-num" style="color:var(--red)">${semVisita60}</div><div class="kpi-label">Sem visita +60 dias</div></div>
        <div class="kpi-card"><div class="kpi-num" style="color:var(--amber)">${nuncaVisitado}</div><div class="kpi-label">Nunca visitados</div></div>
        <div class="kpi-card"><div class="kpi-num" style="color:var(--green)">${visitas30}</div><div class="kpi-label">Visitas últimos 30d</div></div>
      </div>

      <div class="dash-card">
        <div class="dash-card-title">Visitas por mês</div>
        ${meses.map(m => `
          <div class="bar-row">
            <div class="bar-label">${m.label}</div>
            <div class="bar-track"><div class="bar-fill" style="width:${(contagem[m.key]/maxVisitas*100)}%"></div></div>
            <div class="bar-val">${contagem[m.key]}</div>
          </div>
        `).join('')}
      </div>

      <div class="dash-card">
        <div class="dash-card-title">Clientes há mais tempo sem visita</div>
        ${semVisitaLista.length ? semVisitaLista.map(x => `
          <div class="attn-row" onclick="switchView('lista'); document.getElementById('nav-lista').click(); setTimeout(()=>openEditDrawer('${x.c.id}'),50)">
            <div class="attn-name">${x.c.nome}</div>
            <div class="attn-days" style="color:${x.info.dias===null?'var(--red)':(x.info.dias>60?'var(--red)':'var(--amber)')}">${x.info.dias===null?'Nunca':x.info.dias+'d'}</div>
          </div>
        `).join('') : '<div style="color:var(--muted);font-size:13px">Tudo em dia 🎉</div>'}
      </div>

      <div class="dash-card">
        <div class="dash-card-title">Comprou recente, mas sem visita há tempo</div>
        ${compraVsVisita.length ? compraVsVisita.map(x => `
          <div class="attn-row" onclick="openEditDrawer('${x.c.id}')">
            <div class="attn-name">${x.c.nome}<br><span style="color:var(--muted);font-size:11.5px">Comprou há ${x.diasCompra}d</span></div>
            <div class="attn-days" style="color:var(--amber)">${x.diasVisita===null?'Nunca visitado':x.diasVisita+'d s/ visita'}</div>
          </div>
        `).join('') : '<div style="color:var(--muted);font-size:13px">Nenhum caso de atenção agora.</div>'}
      </div>
    </div>
  `;
}

/* ======================= SINCRONIZAÇÃO COM GOOGLE SHEETS (opcional) ======================= */
function updateOfflineBar() {
  document.getElementById('offlineBar').classList.toggle('show', !navigator.onLine);
}
window.addEventListener('online', updateOfflineBar);
window.addEventListener('offline', updateOfflineBar);

function handleSyncClick() {
  if (!APPS_SCRIPT_URL) {
    document.getElementById('overlay').classList.add('show');
    document.getElementById('cfgDrawer').classList.add('open');
    return;
  }
  if (!navigator.onLine) { showToast('Sem conexão — não é possível sincronizar agora'); return; }
  syncWithSheet();
}

async function syncWithSheet() {
  if (syncing) return;
  syncing = true;
  document.getElementById('syncLabel').textContent = 'Sincronizando...';
  try {
    // 1) busca dados remotos da planilha
    const res = await fetch(APPS_SCRIPT_URL);
    if (!res.ok) throw new Error('GET falhou');
    const remoto = await res.json();
    if (Array.isArray(remoto) && remoto.length) clientes = remoto;

    // 2) envia os dados atuais (após o merge) de volta pra planilha
    await fetch(APPS_SCRIPT_URL, {
      method: 'POST',
      headers: { 'Content-Type': 'text/plain;charset=utf-8' }, // evita preflight CORS no Apps Script
      body: JSON.stringify(clientes)
    });

    saveClientes(false);
    if (currentView === 'lista') renderLista(); else renderDashboard();
    showToast('Sincronizado com a planilha ✅');
  } catch (e) {
    showToast('Falha ao sincronizar. Confira a URL configurada.');
  }
  document.getElementById('syncLabel').textContent = 'Sincronizar';
  syncing = false;
}

function maybeAutoSync() {
  // Sincronização automática silenciosa em segundo plano, se já configurado
  if (APPS_SCRIPT_URL && navigator.onLine) syncWithSheet();
}

/* ======================= EXPORTAR EXCEL ======================= */
function exportExcel() {
  const linhas = clientes.map(c => {
    const uv = ultimaVisita(c);
    const info = urgenciaInfo(c);
    return {
      'CNPJ': c.cnpj || '',
      'Nome Cliente': c.nome,
      'Origem': c.origem,
      'Contato': c.contato || '',
      'Última Compra': fmtDate(c.ultimaCompra),
      'Última Visita': uv ? fmtDate(uv) : 'Nunca visitado',
      'Dias sem Visita': info.dias === null ? 'Nunca visitado' : info.dias,
      'Histórico de Visitas': (c.visitas || []).slice().sort().map(fmtDate).join(', ')
    };
  });
  const ws = XLSX.utils.json_to_sheet(linhas);
  ws['!cols'] = [{wch:18},{wch:38},{wch:9},{wch:34},{wch:14},{wch:14},{wch:14},{wch:50}];
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, 'Clientes');
  XLSX.writeFile(wb, 'clientes_visitas_' + todayISO() + '.xlsx');
  showToast('Planilha exportada 📥');
}

/* ======================= INIT ======================= */
loadClientes();
renderUrgChips();
renderLista();
updateOfflineBar();
</script>
</body>
</html>
