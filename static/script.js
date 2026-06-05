/* ─── CHAMADOS PORTAL — script.js ───────────────────────────────── */

const DATA_PATH = 'data/base.xlsx';

/* ─── STATE ──────────────────────────────────────────────────────── */
let allData          = [];
let currentDept      = null;
let openClientRow    = null;
let activeCategories = new Set();
let activeTab        = 'pendentes'; // 'pendentes' | 'retornados'

/* ─── TEMA ───────────────────────────────────────────────────────── */
const THEME_KEY = 'portal-theme';
function initTheme() {
  const saved = localStorage.getItem(THEME_KEY) || 'dark';
  document.documentElement.setAttribute('data-theme', saved);
  updateThemeBtn(saved);
}
function toggleTheme() {
  const cur  = document.documentElement.getAttribute('data-theme') || 'dark';
  const next = cur === 'dark' ? 'light' : 'dark';
  document.documentElement.setAttribute('data-theme', next);
  localStorage.setItem(THEME_KEY, next);
  updateThemeBtn(next);
}
function updateThemeBtn(theme) {
  const btn = document.getElementById('theme-toggle');
  if (btn) btn.textContent = theme === 'dark' ? '☀ Claro' : '🌙 Escuro';
}

/* ─── DOM REFS ───────────────────────────────────────────────────── */
const $loading      = document.getElementById('loading');
const $errorMsg     = document.getElementById('error-msg');
const $viewHome     = document.getElementById('view-home');
const $viewDept     = document.getElementById('view-dept');
const $deptGrid     = document.getElementById('dept-grid');
const $totalRecs    = document.getElementById('total-records');
const $totalDeps    = document.getElementById('total-depts');
const $clientBody   = document.getElementById('client-body');
const $clientSearch = document.getElementById('client-search');
const $deptTitle    = document.getElementById('dept-title');
const $deptCount    = document.getElementById('dept-count');
const $headerDate   = document.getElementById('header-date');

/* ─── DEPT COM RETORNADOS: qualquer dept que tenha linhas Retornado=SIM ─ */
function deptTemRetornados(rows) {
  return rows.some(r => String(col(r,'Retornado')).toUpperCase() === 'SIM');
}

/* ─── INIT ───────────────────────────────────────────────────────── */
document.addEventListener('DOMContentLoaded', () => {
  initTheme();
  document.getElementById('theme-toggle')?.addEventListener('click', toggleTheme);

  $headerDate.textContent = new Date().toLocaleDateString('pt-BR', {
    day: '2-digit', month: 'short', year: 'numeric'
  }).toUpperCase();

  loadData();
  buildCategoryDropdown();
});

/* ─── LOAD XLSX ──────────────────────────────────────────────────── */
async function loadData() {
  try {
    const res = await fetch(DATA_PATH);
    if (!res.ok) throw new Error(`Arquivo não encontrado: ${DATA_PATH} (HTTP ${res.status})`);
    const buf = await res.arrayBuffer();
    const wb  = XLSX.read(new Uint8Array(buf), { type: 'array', cellDates: true });
    const ws  = wb.Sheets[wb.SheetNames[0]];
    allData   = XLSX.utils.sheet_to_json(ws, { defval: '' });
    renderHome();
  } catch (e) {
    showError(e.message);
  }
}

/* ─── SHOW ERROR ─────────────────────────────────────────────────── */
function showError(msg) {
  $loading.style.display = 'none';
  $errorMsg.style.display = 'block';
  $errorMsg.textContent = '⚠  ' + msg;
}

/* ─── COLUMN RESOLVER ────────────────────────────────────────────── */
function col(row, ...candidates) {
  const keys = Object.keys(row);
  for (const c of candidates) {
    const found = keys.find(k => k.trim().toLowerCase() === c.trim().toLowerCase());
    if (found !== undefined) return row[found] ?? '';
  }
  return '';
}

/* ─── DATE FORMATTER ─────────────────────────────────────────────── */
function fmt(val) {
  if (!val && val !== 0) return '—';
  if (val instanceof Date) return val.toLocaleDateString('pt-BR');
  if (typeof val === 'number') {
    const d = new Date(Math.round((val - 25569) * 86400 * 1000));
    if (!isNaN(d)) return d.toLocaleDateString('pt-BR');
  }
  return String(val);
}

/* ─── PRAZO BADGE ────────────────────────────────────────────────── */
function fmtPrazo(val) {
  if (!val && val !== 0) return '<span class="date-cell">—</span>';
  let date;
  if (val instanceof Date) date = val;
  else if (typeof val === 'number') date = new Date(Math.round((val - 25569) * 86400 * 1000));
  else return `<span class="date-cell">${escHtml(String(val))}</span>`;
  if (isNaN(date)) return '<span class="date-cell">—</span>';
  const today = new Date(); today.setHours(0,0,0,0);
  const diffDays = Math.ceil((date - today) / (1000 * 60 * 60 * 24));
  const label = date.toLocaleDateString('pt-BR');
  if (diffDays < 0)  return `<span class="badge badge-open" title="Vencido">⚠ ${label}</span>`;
  if (diffDays <= 3) return `<span class="badge badge-pending" title="Vence em breve">${label}</span>`;
  return `<span class="date-cell">${label}</span>`;
}

/* ─── STATUS BADGE ───────────────────────────────────────────────── */
function badge(status) {
  const s = String(status).toLowerCase().trim();
  let cls = 'badge-default';
  if (/aberto|open|pendente/i.test(s))             cls = 'badge-open';
  else if (/fechado|closed|resolvido/i.test(s))    cls = 'badge-closed';
  else if (/em atend|andamento|progress/i.test(s)) cls = 'badge-pending';
  return `<span class="badge ${cls}">${status || '—'}</span>`;
}

/* ─── HOME VIEW ──────────────────────────────────────────────────── */
function renderHome() {
  $loading.style.display = 'none';

  const deptMap = {};
  for (const row of allData) {
    const retornado = String(col(row, 'Retornado')).toUpperCase();
    if (retornado === 'SIM') continue; // não conta retornados no card da home
    const dept = String(col(row, 'Departamento Responsavel', 'Departamento Responsável', 'Departamento') || 'Sem Departamento').trim();
    deptMap[dept] = (deptMap[dept] || 0) + 1;
  }

  const depts = Object.entries(deptMap).sort((a, b) => b[1] - a[1]);
  $totalRecs.textContent = allData.filter(r => String(col(r,'Retornado')).toUpperCase() !== 'SIM').length.toLocaleString('pt-BR');
  $totalDeps.textContent = depts.length;

  $deptGrid.innerHTML = '';
  for (const [name, count] of depts) {
    // Conta retornados deste dept (qualquer dept pode ter, mas na prática só GC)
    const retCount = allData.filter(r => {
      const dept = String(col(r,'Departamento Responsavel','Departamento Responsável','Departamento')||'').trim();
      const solicitante = String(col(r,'Departamento Solicitante')||'').trim();
      const retornado = String(col(r,'Retornado')).toUpperCase() === 'SIM';
      if (name === 'GERENCIA DE CONTAS' && retornado && solicitante === 'GC - ADMINISTRATIVO') return true;
      return dept === name && retornado;
    }).length;

    const card = document.createElement('div');
    card.className = 'dept-card';
    card.innerHTML = `
      <div class="card-label">Departamento</div>
      <div class="card-name">${escHtml(name)}</div>
      <div class="card-count-wrap">
        <div>
          <div class="card-count">${count.toLocaleString('pt-BR')}</div>
          <div class="card-unit">registros</div>
        </div>
        <div class="card-arrow">→</div>
      </div>
      ${retCount > 0 ? `<div class="card-returned-badge">🔄 ${retCount} retornado${retCount !== 1 ? 's' : ''}</div>` : ''}`;
    card.addEventListener('click', () => openDept(name));
    $deptGrid.appendChild(card);
  }

  showView('home');
}

/* ─── DEPT VIEW ──────────────────────────────────────────────────── */
function openDept(deptName) {
  currentDept      = deptName;
  openClientRow    = null;
  activeCategories = new Set();
  activeTab        = 'pendentes';

  $deptTitle.textContent = deptName;

  const GC_ADM = 'GC - ADMINISTRATIVO';
  const allDeptRows = allData.filter(r => {
    const dept = String(col(r,'Departamento Responsavel','Departamento Responsável','Departamento')||'').trim();
    const solicitante = String(col(r,'Departamento Solicitante')||'').trim();
    const retornado = String(col(r,'Retornado')).toUpperCase() === 'SIM';
    // Inclui rows do próprio dept + retornados do GC-ADM se estivermos em GERENCIA DE CONTAS
    if (deptName === 'GERENCIA DE CONTAS' && retornado && solicitante === GC_ADM) return true;
    return dept === deptName;
  });

  // Monta abas se for dept GC
  buildTabs(deptName, allDeptRows);
  populateCategoryDropdown(allDeptRows.filter(r => String(col(r,'Retornado')).toUpperCase() !== 'SIM'));
  renderDeptRows(allDeptRows);
  showView('dept');
}

/* ─── ABAS ───────────────────────────────────────────────────────── */
function buildTabs(deptName, allDeptRows) {
  const tabsWrap = document.getElementById('dept-tabs');
  if (!tabsWrap) return;

  const hasRetornados = deptTemRetornados(allDeptRows);

  if (!hasRetornados) {
    tabsWrap.style.display = 'none';
    return;
  }

  const pendCount = allDeptRows.filter(r => String(col(r,'Retornado')).toUpperCase() !== 'SIM').length;
  const retCount  = allDeptRows.filter(r => String(col(r,'Retornado')).toUpperCase() === 'SIM').length;

  tabsWrap.style.display = 'flex';
  tabsWrap.innerHTML = `
    <button class="tab-btn tab-btn--active" data-tab="pendentes">
      Pendentes <span class="tab-count">${pendCount}</span>
    </button>
    <button class="tab-btn" data-tab="retornados">
      Retornados <span class="tab-count tab-count--ret">${retCount}</span>
    </button>`;

  tabsWrap.querySelectorAll('.tab-btn').forEach(btn => {
    btn.addEventListener('click', () => {
      activeTab = btn.dataset.tab;
      tabsWrap.querySelectorAll('.tab-btn').forEach(b => b.classList.remove('tab-btn--active'));
      btn.classList.add('tab-btn--active');
      openClientRow = null;
      activeCategories = new Set();
      $clientSearch.value = '';

      const rows = allData.filter(r =>
        String(col(r,'Departamento Responsavel','Departamento Responsável','Departamento')||'').trim() === currentDept
      );
      const tabRows = activeTab === 'retornados'
        ? rows.filter(r => String(col(r,'Retornado')).toUpperCase() === 'SIM')
        : rows.filter(r => String(col(r,'Retornado')).toUpperCase() !== 'SIM');

      populateCategoryDropdown(tabRows);
      renderDeptRows(rows);
      updateCategoryLabel();
    });
  });
}

/* ─── CATEGORY DROPDOWN — BUILD SHELL ───────────────────────────── */
function buildCategoryDropdown() {
  const wrapper = document.createElement('div');
  wrapper.id        = 'cat-filter-wrap';
  wrapper.className = 'cat-filter-wrap';
  wrapper.innerHTML = `
    <button id="cat-btn" class="cat-btn" type="button" aria-haspopup="true" aria-expanded="false">
      <span id="cat-btn-label">Categoria</span>
      <span class="cat-chevron">▾</span>
    </button>
    <div id="cat-dropdown" class="cat-dropdown" hidden>
      <div class="cat-dropdown-actions">
        <button type="button" id="cat-select-all">Todas</button>
        <button type="button" id="cat-clear-all">Limpar</button>
      </div>
      <ul id="cat-list" class="cat-list"></ul>
    </div>`;

  // Cria wrapper .table-header-controls ao redor do input + dropdown
  const searchInput = $clientSearch;
  const searchWrap  = searchInput.parentElement;
  const controls    = document.createElement('div');
  controls.className = 'table-header-controls';
  searchWrap.insertBefore(controls, searchInput);
  controls.appendChild(searchInput);
  controls.appendChild(wrapper);

  document.getElementById('cat-btn').addEventListener('click', (e) => {
    e.stopPropagation();
    const dd  = document.getElementById('cat-dropdown');
    const btn = document.getElementById('cat-btn');
    const isHidden = dd.hidden;
    dd.hidden = !isHidden;
    btn.setAttribute('aria-expanded', String(isHidden));
  });

  document.addEventListener('click', (e) => {
    const dd = document.getElementById('cat-dropdown');
    if (!wrapper.contains(e.target)) {
      dd.hidden = true;
      document.getElementById('cat-btn').setAttribute('aria-expanded', 'false');
    }
  });

  document.getElementById('cat-select-all').addEventListener('click', () => {
    document.querySelectorAll('#cat-list input[type=checkbox]').forEach(cb => {
      cb.checked = true;
      activeCategories.add(cb.value);
    });
    updateCategoryLabel();
    applyFilters();
  });

  document.getElementById('cat-clear-all').addEventListener('click', () => {
    document.querySelectorAll('#cat-list input[type=checkbox]').forEach(cb => cb.checked = false);
    activeCategories.clear();
    updateCategoryLabel();
    applyFilters();
  });
}

/* ─── CATEGORY DROPDOWN — POPULAR ───────────────────────────────── */
function populateCategoryDropdown(rows) {
  const cats = [...new Set(
    rows.map(r => String(col(r,'Categoria','categoria')||'').trim()).filter(Boolean)
  )].sort();

  const $list = document.getElementById('cat-list');
  $list.innerHTML = '';
  activeCategories.clear();

  for (const cat of cats) {
    const li = document.createElement('li');
    li.innerHTML = `
      <label class="cat-option">
        <input type="checkbox" value="${escHtml(cat)}">
        <span>${escHtml(cat)}</span>
      </label>`;
    li.querySelector('input').addEventListener('change', (e) => {
      if (e.target.checked) activeCategories.add(cat);
      else activeCategories.delete(cat);
      updateCategoryLabel();
      applyFilters();
    });
    $list.appendChild(li);
  }
  updateCategoryLabel();
}

/* ─── CATEGORY LABEL ─────────────────────────────────────────────── */
function updateCategoryLabel() {
  const $label = document.getElementById('cat-btn-label');
  const total  = document.querySelectorAll('#cat-list input[type=checkbox]').length;
  if (activeCategories.size === 0 || activeCategories.size === total) {
    $label.textContent = 'Categoria';
    document.getElementById('cat-btn').classList.remove('cat-btn--active');
  } else {
    $label.textContent = `Categoria (${activeCategories.size})`;
    document.getElementById('cat-btn').classList.add('cat-btn--active');
  }
}

/* ─── APLICAR FILTROS ────────────────────────────────────────────── */
function applyFilters() {
  const q = $clientSearch.value.toLowerCase();

  $clientBody.querySelectorAll('tr:not(.detail-row)').forEach(tr => {
    const clientId  = tr.dataset.clientId;
    const matchText = !q || tr.textContent.toLowerCase().includes(q);

    let matchCat = true;
    if (activeCategories.size > 0) {
      const clientRows = allData.filter(r => {
        const dept = String(col(r,'Departamento Responsavel','Departamento Responsável','Departamento')||'').trim();
        const id   = String(col(r,'IdCliente','Id Cliente','ID Cliente','id_cliente')||'').trim();
        const ret  = String(col(r,'Retornado')).toUpperCase() === 'SIM';
        return dept === currentDept && id === clientId &&
               (activeTab === 'retornados' ? ret : !ret);
      });
      matchCat = clientRows.some(r =>
        activeCategories.has(String(col(r,'Categoria','categoria')||'').trim())
      );
    }

    const visible = matchText && matchCat;
    tr.style.display = visible ? '' : 'none';
    const next = tr.nextElementSibling;
    if (next && next.classList.contains('detail-row')) next.style.display = visible ? '' : 'none';
  });

  const visible = $clientBody.querySelectorAll('tr:not(.detail-row):not([style*="display: none"])').length;
  $deptCount.textContent = `${visible} registro${visible !== 1 ? 's' : ''}`;
}

/* ─── RENDER DEPT ROWS ───────────────────────────────────────────── */
function renderDeptRows(allDeptRows) {
  const isRet = activeTab === 'retornados';
  const rows  = allDeptRows.filter(r =>
    isRet
      ? String(col(r,'Retornado')).toUpperCase() === 'SIM'
      : String(col(r,'Retornado')).toUpperCase() !== 'SIM'
  );

  const clientMap = {};
  for (const r of rows) {
    const id = String(col(r,'IdCliente','Id Cliente','ID Cliente','id_cliente')||'').trim();
    if (!clientMap[id]) {
      clientMap[id] = {
        id,
        name:      col(r,'Cliente','Nome Cliente','NomeCliente') || id,
        dataCad:   col(r,'Data Cadastro','DataCadastro','Data_Cadastro'),
        prazoVenc: col(r,'Prazo Vencimento','Prazo de Vencimento','PrazoVencimento'),
        rows: []
      };
    }
    clientMap[id].rows.push(r);
  }

  const clients = Object.values(clientMap);
  $deptCount.textContent = `${rows.length} registro${rows.length !== 1 ? 's' : ''}`;
  renderClientTable(clients);
}

/* ─── RENDER CLIENT TABLE ────────────────────────────────────────── */
function renderClientTable(clients) {
  $clientBody.innerHTML = '';

  if (!clients.length) {
    $clientBody.innerHTML = `<tr><td colspan="4" class="empty">Nenhum registro encontrado.</td></tr>`;
    return;
  }

  for (const c of clients) {
    const tr = document.createElement('tr');
    tr.dataset.clientId = c.id;
    tr.innerHTML = `
      <td class="id-cell">${escHtml(c.id)||'—'}</td>
      <td class="name-cell">${escHtml(String(c.name))}</td>
      <td class="date-cell">${fmt(c.dataCad)}</td>
      <td>${fmtPrazo(c.prazoVenc)}</td>`;
    tr.addEventListener('click', () => toggleClientDetail(tr, c));
    $clientBody.appendChild(tr);
  }
}

/* ─── TOGGLE CLIENT DETAIL ───────────────────────────────────────── */
function toggleClientDetail(tr, client) {
  const existingDetail = tr.nextElementSibling;
  const isOpen = existingDetail && existingDetail.classList.contains('detail-row');

  if (openClientRow) {
    const prev = $clientBody.querySelector(`tr[data-client-id="${openClientRow}"]`);
    if (prev) {
      const prevDetail = prev.nextElementSibling;
      if (prevDetail && prevDetail.classList.contains('detail-row')) prevDetail.remove();
    }
  }

  if (isOpen) { openClientRow = null; return; }
  openClientRow = client.id;

  const visibleRows = activeCategories.size > 0
    ? client.rows.filter(r => activeCategories.has(String(col(r,'Categoria','categoria')||'').trim()))
    : client.rows;

  const detailTr = document.createElement('tr');
  detailTr.className = 'detail-row';
  const td = document.createElement('td');
  td.colSpan = 3;

  td.innerHTML = `
    <div class="detail-inner">
      <h4>Pendências do cliente — ${escHtml(String(client.name))}</h4>
      <div class="detail-table-wrap">
        <table class="detail-table">
          <thead>
            <tr>
              <th>ID</th><th>Categoria</th><th>Solicitação</th><th>Responsável</th>
              <th>Depto. Solicitante</th><th>Data Cadastro</th><th>Prazo Vencimento</th><th>Status</th>
              <th>Solicitante</th><th>Início Atend.</th>
            </tr>
          </thead>
          <tbody>
            ${visibleRows.map(r => `
              <tr>
                <td class="id-cell">${escHtml(String(col(r,'Id','ID','id')||'—'))}</td>
                <td>${escHtml(String(col(r,'Categoria','categoria')||'—'))}</td>
                <td>${escHtml(String(col(r,'Solicitacao','Solicitação','solicitacao','solicitação')||'—'))}</td>
                <td>${escHtml(String(col(r,'Responsavel','Responsável','responsavel')||'—'))}</td>
                <td>${escHtml(String(col(r,'Departamento Solicitante','Departamento_Solicitante')||'—'))}</td>
                <td class="date-cell">${fmt(col(r,'Data Cadastro','DataCadastro','Data_Cadastro'))}</td>
                <td>${fmtPrazo(col(r,'Prazo Vencimento','Prazo de Vencimento','PrazoVencimento','prazo_vencimento'))}</td>
                <td>${badge(col(r,'Status','status'))}</td>
                <td>${escHtml(String(col(r,'Solicitante','solicitante')||'—'))}</td>
                <td class="date-cell">${fmt(col(r,'Inicio Atend.','Início Atend.','Inicio Atendimento','InicioAtend','Inicio_Atend'))}</td>
              </tr>`).join('')}
          </tbody>
        </table>
      </div>
    </div>`;

  detailTr.appendChild(td);
  tr.insertAdjacentElement('afterend', detailTr);
}

/* ─── SEARCH ─────────────────────────────────────────────────────── */
$clientSearch.addEventListener('input', applyFilters);

/* ─── BACK ───────────────────────────────────────────────────────── */
document.getElementById('btn-back').addEventListener('click', () => {
  openClientRow = null; activeCategories = new Set(); $clientSearch.value = '';
  showView('home');
});

/* ─── BREADCRUMB ─────────────────────────────────────────────────── */
document.getElementById('bc-home').addEventListener('click', () => {
  if ($viewDept.classList.contains('active')) {
    openClientRow = null; activeCategories = new Set(); $clientSearch.value = '';
    showView('home');
  }
});

/* ─── SHOW VIEW ──────────────────────────────────────────────────── */
function showView(name) {
  $viewHome.classList.toggle('active', name === 'home');
  $viewDept.classList.toggle('active', name === 'dept');
  const bcDept = document.getElementById('bc-dept');
  const bcSep  = document.getElementById('bc-sep');
  if (name === 'dept') {
    bcDept.textContent = currentDept; bcDept.style.display = ''; bcSep.style.display = '';
  } else {
    bcDept.style.display = 'none'; bcSep.style.display = 'none';
  }
}

/* ─── HTML ESCAPE ────────────────────────────────────────────────── */
function escHtml(str) {
  return String(str)
    .replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;');
}