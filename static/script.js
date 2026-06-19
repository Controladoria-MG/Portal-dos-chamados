/* ─── CHAMADOS PORTAL — script.js ───────────────────────────────── */

const DATA_PATH = 'data/base.xlsx';

/* ─── STATE ──────────────────────────────────────────────────────── */
let allData          = [];
let currentDept      = null;
let openClientRow    = null;
let activeTab        = 'pendentes'; // 'pendentes' | 'retornados'
let currentDeptTemDeptoAnterior = false; // dept atual tem chamados "Devolvido para Solicitante"?

/* ─── FILTROS (CATEGORIA / RESPONSÁVEL / STATUS) ─────────────────── */
const FILTER_DEFS = [
  { key: 'categoria',   label: 'Categoria',   cols: ['Categoria','categoria'] },
  { key: 'responsavel', label: 'Responsável', cols: ['Responsavel','Responsável','responsavel'] },
  { key: 'status',      label: 'Status',      cols: ['Status','status'] },
];
const activeFilters = {
  categoria:   new Set(),
  responsavel: new Set(),
  status:      new Set(),
};
function resetFilters() {
  FILTER_DEFS.forEach(def => activeFilters[def.key].clear());
}
function filterValue(row, def) {
  return String(col(row, ...def.cols) || '').trim();
}
function rowMatchesFilters(row) {
  return FILTER_DEFS.every(def => {
    const active = activeFilters[def.key];
    return active.size === 0 || active.has(filterValue(row, def));
  });
}

/* ─── TEMA ───────────────────────────────────────────────────────── */
function initTheme() {
  document.documentElement.setAttribute('data-theme', 'light');
  updateThemeBtn('light');
}
function toggleTheme() {
  const cur  = document.documentElement.getAttribute('data-theme') || 'light';
  const next = cur === 'dark' ? 'light' : 'dark';
  document.documentElement.setAttribute('data-theme', next);
  updateThemeBtn(next);
}
function updateThemeBtn(theme) {
  const btn = document.getElementById('theme-toggle');
  if (btn) btn.textContent = theme === 'light' ? '☀️ Alterar Tema: Claro' : '🌙 Alterar Tema: Escuro';
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

/* ─── DEPT COM DEVOLVIDOS: qualquer linha com Depto. Responsavel Original ─ */
function deptTemDeptoAnterior(rows) {
  return rows.some(r => String(col(r,'Departamento Responsavel Original','Departamento Responsável Original')||'').trim() !== '');
}

/* ─── INIT ───────────────────────────────────────────────────────── */
document.addEventListener('DOMContentLoaded', () => {
  initTheme();
  document.getElementById('theme-toggle')?.addEventListener('click', toggleTheme);

  $headerDate.textContent = new Date().toLocaleDateString('pt-BR', {
    day: '2-digit', month: 'short', year: 'numeric'
  }).toUpperCase();

  loadData();
  buildFilterDropdowns();
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

/* ─── PRAZO PARSER ───────────────────────────────────────────────── */
function parsePrazoDate(val) {
  if (!val && val !== 0) return null;
  let date;
  if (val instanceof Date) date = val;
  else if (typeof val === 'number') date = new Date(Math.round((val - 25569) * 86400 * 1000));
  else return null;
  return isNaN(date) ? null : date;
}

/* ─── PRAZO BADGE ────────────────────────────────────────────────── */
function fmtPrazo(val) {
  const date = parsePrazoDate(val);
  if (!date) {
    if (!val && val !== 0) return '<span class="date-cell">—</span>';
    return `<span class="date-cell">${escHtml(String(val))}</span>`;
  }
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
    if (retornado === 'SIM') continue; // retornados não contam nos cards normais
    const dept = String(col(row, 'Departamento Responsavel', 'Departamento Responsável', 'Departamento') || 'Sem Departamento').trim();
    deptMap[dept] = (deptMap[dept] || 0) + 1;
  }
  // Garante que GERENCIA DE CONTAS aparece mesmo que não tenha pendentes
  if (!deptMap['GERENCIA DE CONTAS']) deptMap['GERENCIA DE CONTAS'] = 0;

  const depts = Object.entries(deptMap).sort((a, b) => b[1] - a[1]);
  $totalRecs.textContent = allData.filter(r => String(col(r,'Retornado')).toUpperCase() !== 'SIM').length.toLocaleString('pt-BR');
  const totalRetornados = allData.filter(r => String(col(r,'Retornado')).toUpperCase() === 'SIM').length;
  $totalDeps.textContent = depts.length;

  $deptGrid.innerHTML = '';
  for (const [name, count] of depts) {
    // Conta retornados deste dept (qualquer dept pode ter, mas na prática só GC)
    // Todos os retornados ficam sob GERENCIA DE CONTAS
    const retCount = name === 'GERENCIA DE CONTAS'
      ? allData.filter(r => String(col(r,'Retornado')).toUpperCase() === 'SIM').length
      : 0;

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
  resetFilters();
  activeTab        = 'pendentes';

  $deptTitle.textContent = deptName;

  const allDeptRows = allData.filter(r => {
    const dept = String(col(r,'Departamento Responsavel','Departamento Responsável','Departamento')||'').trim();
    const retornado = String(col(r,'Retornado')).toUpperCase() === 'SIM';
    // Retornados sempre ficam em GERENCIA DE CONTAS, independente do dept original
    if (retornado) return deptName === 'GERENCIA DE CONTAS';
    return dept === deptName;
  });

  currentDeptTemDeptoAnterior = deptTemDeptoAnterior(allDeptRows);

  // Monta abas se for dept GC
  buildTabs(deptName, allDeptRows);
  populateFilterDropdowns(allDeptRows.filter(r => String(col(r,'Retornado')).toUpperCase() !== 'SIM'));
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
      resetFilters();
      $clientSearch.value = '';

      const rows = allData.filter(r => {
        const dept = String(col(r,'Departamento Responsavel','Departamento Responsável','Departamento')||'').trim();
        const retornado = String(col(r,'Retornado')).toUpperCase() === 'SIM';
        if (retornado) return currentDept === 'GERENCIA DE CONTAS';
        return dept === currentDept;
      });

      populateFilterDropdowns(rows.filter(r => String(col(r,'Retornado')).toUpperCase() !== 'SIM'));
      renderDeptRows(rows);
    });
  });
}

/* ─── FILTROS — MONTAR ESTRUTURA (CATEGORIA / RESPONSÁVEL / STATUS) ─ */
function buildFilterDropdowns() {
  // Cria wrapper .table-header-controls ao redor do input + filtros
  const searchInput = $clientSearch;
  const searchWrap  = searchInput.parentElement;
  const controls    = document.createElement('div');
  controls.className = 'table-header-controls';
  searchWrap.insertBefore(controls, searchInput);
  controls.appendChild(searchInput);

  const filtersRow = document.createElement('div');
  filtersRow.className = 'filters-row';
  controls.appendChild(filtersRow);

  FILTER_DEFS.forEach(def => {
    const wrapper = document.createElement('div');
    wrapper.id        = `${def.key}-filter-wrap`;
    wrapper.className = 'cat-filter-wrap';
    wrapper.innerHTML = `
      <button id="${def.key}-btn" class="cat-btn" type="button" aria-haspopup="true" aria-expanded="false">
        <span id="${def.key}-btn-label">${def.label}</span>
        <span class="cat-chevron">▾</span>
      </button>`;
    filtersRow.appendChild(wrapper);

    // Dropdown é anexado direto no <body> ("portal"), para nunca ser
    // cortado por um ancestral com overflow:hidden/auto (ex: .table-wrap).
    const dd = document.createElement('div');
    dd.id        = `${def.key}-dropdown`;
    dd.className = 'cat-dropdown';
    dd.hidden    = true;
    dd.innerHTML = `
      <div class="cat-dropdown-actions">
        <button type="button" id="${def.key}-select-all">Todos</button>
        <button type="button" id="${def.key}-clear-all">Limpar</button>
      </div>
      <ul id="${def.key}-list" class="cat-list"></ul>`;
    document.body.appendChild(dd);

    const btn = wrapper.querySelector('.cat-btn');

    const openDropdown = () => {
      document.querySelectorAll('.cat-dropdown').forEach(other => { if (other !== dd) other.hidden = true; });
      document.querySelectorAll('.cat-btn').forEach(other => { if (other !== btn) other.setAttribute('aria-expanded', 'false'); });
      dd.hidden = false;
      btn.setAttribute('aria-expanded', 'true');
      positionDropdown(btn, dd);
    };
    const closeDropdown = () => {
      dd.hidden = true;
      btn.setAttribute('aria-expanded', 'false');
    };

    btn.addEventListener('click', (e) => {
      e.stopPropagation();
      if (dd.hidden) openDropdown(); else closeDropdown();
    });

    document.addEventListener('click', (e) => {
      if (!wrapper.contains(e.target) && !dd.contains(e.target)) closeDropdown();
    });

    dd.querySelector(`#${def.key}-select-all`).addEventListener('click', () => {
      dd.querySelectorAll('input[type=checkbox]').forEach(cb => {
        cb.checked = true;
        activeFilters[def.key].add(cb.value);
      });
      updateFilterLabel(def);
      applyFilters();
    });

    dd.querySelector(`#${def.key}-clear-all`).addEventListener('click', () => {
      dd.querySelectorAll('input[type=checkbox]').forEach(cb => cb.checked = false);
      activeFilters[def.key].clear();
      updateFilterLabel(def);
      applyFilters();
    });
  });

  // Reposiciona o dropdown aberto ao rolar (inclusive scroll interno de
  // containers, via capture) ou redimensionar a janela.
  const repositionOpenDropdown = () => {
    const dd = document.querySelector('.cat-dropdown:not([hidden])');
    if (!dd) return;
    const btn = document.getElementById(dd.id.replace('-dropdown', '-btn'));
    if (btn) positionDropdown(btn, dd);
  };
  window.addEventListener('scroll', repositionOpenDropdown, true);
  window.addEventListener('resize', repositionOpenDropdown);
}

/* ─── FILTROS — POSICIONAR DROPDOWN (fixed, ancorado no botão) ──── */
function positionDropdown(btn, dd) {
  const margin = 6;
  const rect   = btn.getBoundingClientRect();
  const ddW    = dd.offsetWidth;
  const ddH    = dd.offsetHeight;

  let left = rect.right - ddW;
  left = Math.max(8, Math.min(left, window.innerWidth - ddW - 8));

  let top = rect.bottom + margin;
  if (top + ddH > window.innerHeight - 8 && rect.top - ddH - margin > 0) {
    top = rect.top - ddH - margin; // sem espaço abaixo: abre para cima
  }

  dd.style.left = `${left}px`;
  dd.style.top  = `${top}px`;
}

/* ─── FILTROS — POPULAR OPÇÕES ───────────────────────────────────── */
function populateFilterDropdowns(rows) {
  FILTER_DEFS.forEach(def => {
    const values = [...new Set(
      rows.map(r => filterValue(r, def)).filter(Boolean)
    )].sort();

    const $list = document.getElementById(`${def.key}-list`);
    $list.innerHTML = '';
    activeFilters[def.key].clear();

    for (const val of values) {
      const li = document.createElement('li');
      li.innerHTML = `
        <label class="cat-option">
          <input type="checkbox" value="${escHtml(val)}">
          <span>${escHtml(val)}</span>
        </label>`;
      li.querySelector('input').addEventListener('change', (e) => {
        if (e.target.checked) activeFilters[def.key].add(val);
        else activeFilters[def.key].delete(val);
        updateFilterLabel(def);
        applyFilters();
      });
      $list.appendChild(li);
    }
    updateFilterLabel(def);
  });
}

/* ─── FILTROS — ATUALIZAR LABEL DO BOTÃO ─────────────────────────── */
function updateFilterLabel(def) {
  const $label = document.getElementById(`${def.key}-btn-label`);
  const total  = document.querySelectorAll(`#${def.key}-list input[type=checkbox]`).length;
  const active = activeFilters[def.key];
  const $btn   = document.getElementById(`${def.key}-btn`);
  if (active.size === 0 || active.size === total) {
    $label.textContent = def.label;
    $btn.classList.remove('cat-btn--active');
  } else {
    $label.textContent = `${def.label} (${active.size})`;
    $btn.classList.add('cat-btn--active');
  }
}

/* ─── APLICAR FILTROS ────────────────────────────────────────────── */
function applyFilters() {
  const q = $clientSearch.value.toLowerCase();
  const algumFiltroAtivo = FILTER_DEFS.some(def => activeFilters[def.key].size > 0);

  $clientBody.querySelectorAll('tr:not(.detail-row)').forEach(tr => {
    const clientId  = tr.dataset.clientId;
    const matchText = !q || tr.textContent.toLowerCase().includes(q);

    let matchFiltros = true;
    if (algumFiltroAtivo) {
      const clientRows = allData.filter(r => {
        const dept = String(col(r,'Departamento Responsavel','Departamento Responsável','Departamento')||'').trim();
        const id   = String(col(r,'IdCliente','Id Cliente','ID Cliente','id_cliente')||'').trim();
        const ret  = String(col(r,'Retornado')).toUpperCase() === 'SIM';
        return dept === currentDept && id === clientId &&
               (activeTab === 'retornados' ? ret : !ret);
      });
      matchFiltros = clientRows.some(rowMatchesFilters);
    }

    const visible = matchText && matchFiltros;
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

  $deptCount.textContent = `${rows.length} registro${rows.length !== 1 ? 's' : ''}`;

  // Ambas as abas: agrupado por cliente com detalhamento expansível
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

  // Restaura cabeçalho padrão
  const thead = $clientBody.closest('table').querySelector('thead tr');
  thead.innerHTML = '<th>ID Cliente</th><th>Cliente</th><th>Data Cadastro</th><th class="prazo-cell">Prazo Vencimento</th>';

  const clients = Object.values(clientMap).sort((a, b) => {
    const da = parsePrazoDate(a.prazoVenc);
    const db = parsePrazoDate(b.prazoVenc);
    if (!da && !db) return 0;
    if (!da) return 1;
    if (!db) return -1;
    return da - db;
  });

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
      <td class="prazo-cell">${fmtPrazo(c.prazoVenc)}</td>`;
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

  const visibleRows = client.rows.filter(rowMatchesFilters).sort((a, b) => {
    const da = parsePrazoDate(col(a,'Prazo Vencimento','Prazo de Vencimento','PrazoVencimento','prazo_vencimento'));
    const db = parsePrazoDate(col(b,'Prazo Vencimento','Prazo de Vencimento','PrazoVencimento','prazo_vencimento'));
    if (!da && !db) return 0;
    if (!da) return 1;
    if (!db) return -1;
    return da - db;
  });

  const detailTr = document.createElement('tr');
  detailTr.className = 'detail-row';
  const td = document.createElement('td');
  td.colSpan = 3;

  const showDeptoAnterior = currentDeptTemDeptoAnterior;

  td.innerHTML = `
    <div class="detail-inner">
      <h4>Pendências do cliente — ${escHtml(String(client.name))}</h4>
      <div class="ticket-list">
        ${visibleRows.map(r => `
          <div class="ticket-card">
            <div class="ticket-top">
              <div class="ticket-field"><span class="f-label">ID</span><span class="f-value f-id">${escHtml(String(col(r,'Id','ID','id')||'—'))}</span></div>
              <div class="ticket-field"><span class="f-label">Categoria</span><span class="f-value">${escHtml(String(col(r,'Categoria','categoria')||'—'))}</span></div>
              <div class="ticket-field"><span class="f-label">Responsável</span><span class="f-value">${escHtml(String(col(r,'Responsavel','Responsável','responsavel')||'—'))}</span></div>
              <div class="ticket-field"><span class="f-label">Depto. Solicitante</span><span class="f-value">${escHtml(String(col(r,'Departamento Solicitante','Departamento_Solicitante')||'—'))}</span></div>
              <div class="ticket-field"><span class="f-label">Data Cadastro</span><span class="f-value">${fmt(col(r,'Data Cadastro','DataCadastro','Data_Cadastro'))}</span></div>
              <div class="ticket-field"><span class="f-label">Prazo Vencimento</span><span class="f-value">${fmtPrazo(col(r,'Prazo Vencimento','Prazo de Vencimento','PrazoVencimento','prazo_vencimento'))}</span></div>
              <div class="ticket-field"><span class="f-label">Status</span><span class="f-value">${badge(col(r,'Status','status'))}</span></div>
              ${showDeptoAnterior ? `<div class="ticket-field"><span class="f-label">Depto. Anterior</span><span class="f-value">${escHtml(String(col(r,'Departamento Responsavel Original','Departamento Responsável Original')||'—'))}</span></div>` : ''}
              <div class="ticket-field"><span class="f-label">Solicitante</span><span class="f-value">${escHtml(String(col(r,'Solicitante','solicitante')||'—'))}</span></div>
              <div class="ticket-field"><span class="f-label">Início Atend.</span><span class="f-value">${fmt(col(r,'Inicio Atend.','Início Atend.','Inicio Atendimento','InicioAtend','Inicio_Atend'))}</span></div>
            </div>
            <div class="ticket-bottom">
              <span class="f-label">Solicitação</span>
              <p class="f-solicitacao">${escHtml(String(col(r,'Solicitacao','Solicitação','solicitacao','solicitação')||'—'))}</p>
            </div>
          </div>`).join('')}
      </div>
    </div>`;

  detailTr.appendChild(td);
  tr.insertAdjacentElement('afterend', detailTr);
}

/* ─── SEARCH ─────────────────────────────────────────────────────── */
$clientSearch.addEventListener('input', applyFilters);

/* ─── BACK ───────────────────────────────────────────────────────── */
document.getElementById('btn-back').addEventListener('click', () => {
  openClientRow = null; resetFilters(); $clientSearch.value = '';
  showView('home');
});

/* ─── BREADCRUMB ─────────────────────────────────────────────────── */
document.getElementById('bc-home').addEventListener('click', () => {
  if ($viewDept.classList.contains('active')) {
    openClientRow = null; resetFilters(); $clientSearch.value = '';
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