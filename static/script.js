/* ─── CHAMADOS PORTAL — script.js ───────────────────────────────── */

const DATA_PATH = 'data/base.xlsx';
const INFO_PATH = 'data/base_info.json';

/* ─── STATE ──────────────────────────────────────────────────────── */
let allData          = [];
let currentDept      = null;
let openClientRow    = null;
let activeTab        = 'pendentes'; // 'pendentes' | 'retornados'
let currentDeptTemDeptoAnterior = false; // dept atual tem chamados "Devolvido para Solicitante"?

/* ─── FILTROS (CATEGORIA / RESPONSÁVEL / STATUS / VENCIMENTO) ────── */
const FILTER_DEFS = [
  { key: 'categoria',       label: 'Categoria',           cols: ['Categoria','categoria'] },
  { key: 'deptoSolicitante', label: 'Depto. Solicitante', cols: ['Departamento Solicitante','Departamento_Solicitante','departamento solicitante'] },
  { key: 'responsavel',     label: 'Responsável',         cols: ['Responsavel','Responsável','responsavel'] },
  { key: 'coordenador',     label: 'Coordenador',         cols: ['Coordenador','coordenador'] },
  { key: 'status',          label: 'Status',              cols: ['Status','status'] },
  { key: 'vencimento',      label: 'Vencimento',          value: vencimentoStatus, sort: ['Vencido','Não Vencido'] },
];
const activeFilters = {
  categoria:        new Set(),
  deptoSolicitante: new Set(),
  responsavel:      new Set(),
  coordenador:      new Set(),
  status:           new Set(),
  vencimento:       new Set(),
};
function resetFilters() {
  FILTER_DEFS.forEach(def => activeFilters[def.key].clear());
}
function filterValue(row, def) {
  if (def.value) return def.value(row);
  return String(col(row, ...def.cols) || '').trim();
}
/* Classifica o chamado em 'Vencido' / 'Não Vencido' a partir do Prazo Vencimento */
function vencimentoStatus(row) {
  const date = parsePrazoDate(prazoVal(row));
  if (!date) return '';
  const today = new Date(); today.setHours(0,0,0,0);
  return date < today ? 'Vencido' : 'Não Vencido';
}
function prazoVal(row) {
  return col(row,'Prazo Vencimento','Prazo de Vencimento','PrazoVencimento','prazo_vencimento');
}
/* Prazo Vencimento mais urgente (data mais antiga) entre um conjunto de chamados */
function earliestPrazo(rows) {
  let best = null;
  for (const r of rows) {
    const raw = prazoVal(r);
    const date = parsePrazoDate(raw);
    if (date && (!best || date < best.date)) best = { raw, date };
  }
  return best ? best.raw : '';
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
  if (btn) btn.textContent = theme === 'light' ? 'Alterar Tema: Claro' : 'Alterar Tema: Escuro';
}

/* ─── DOM REFS ───────────────────────────────────────────────────── */
const $loading      = document.getElementById('loading');
const $errorMsg     = document.getElementById('error-msg');
const $viewHome     = document.getElementById('view-home');
const $viewDept     = document.getElementById('view-dept');
const $viewRelatorios = document.getElementById('view-relatorios');
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

/* ─── DATA DE ATUALIZAÇÃO DA BASE ───────────────────────────────────
   Lê data/base_info.json (gravado pelo pipeline toda vez que base.xlsx
   é regerado) para mostrar quando os dados foram atualizados de fato -
   a data de hoje sozinha não diz nada sobre isso. */
async function carregarDataAtualizacao() {
  try {
    const res = await fetch(INFO_PATH);
    if (!res.ok) throw new Error('base_info.json não encontrado');
    const info = await res.json();
    $headerDate.textContent = `Base atualizada em ${info.atualizado_em}`;
  } catch (e) {
    $headerDate.textContent = '';
  }
}

/* ─── INIT ───────────────────────────────────────────────────────── */
document.addEventListener('DOMContentLoaded', () => {
  initTheme();
  document.getElementById('theme-toggle')?.addEventListener('click', toggleTheme);

  carregarDataAtualizacao();
  loadData();
  buildFilterDropdowns();
  carregarHistoricoMensal();
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

/* ─── AGRUPAMENTO DE DEPARTAMENTO ────────────────────────────────────
   GC - Administrativo é tratado como parte de GERENCIA DE CONTAS —
   um único card/departamento no portal, nunca separados. ─────────── */
function deptGroupName(row) {
  const raw = String(col(row, 'Departamento Responsavel', 'Departamento Responsável', 'Departamento') || '').trim();
  return raw.toUpperCase() === 'GC - ADMINISTRATIVO' ? 'GERENCIA DE CONTAS' : (raw || 'Sem Departamento');
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
  const date = parsePrazoDate(val);
  if (date) return date.toLocaleDateString('pt-BR');
  if (!val && val !== 0) return '—';
  return String(val);
}

/* ─── PRAZO PARSER ───────────────────────────────────────────────── */
function parsePrazoDate(val) {
  if (!val && val !== 0) return null;
  let date;
  if (val instanceof Date) date = val;
  else if (typeof val === 'number') date = new Date(Math.round((val - 25569) * 86400 * 1000));
  else return null;
  if (isNaN(date)) return null;
  // Datas do Excel/SheetJS chegam como meia-noite UTC. Em fuso negativo
  // (Brasil, UTC-3) isso corresponde a 21h do dia anterior no horário local,
  // o que adianta a data em 1 dia em comparações/exibições locais. Normaliza
  // para meia-noite LOCAL do mesmo dia-mês-ano UTC, eliminando esse desvio.
  return new Date(date.getUTCFullYear(), date.getUTCMonth(), date.getUTCDate());
}

/* ─── PRAZO BADGE ────────────────────────────────────────────────── */
function fmtPrazo(val) {
  const date = parsePrazoDate(val);
  if (!date) {
    if (!val && val !== 0) return '<span class="date-cell">—</span>';
    return `<span class="date-cell">${escHtml(String(val))}</span>`;
  }
  const today = new Date(); today.setHours(0,0,0,0);
  const label = date.toLocaleDateString('pt-BR');
  if (date < today) return `<span class="prazo-vencido">${label}</span>`;
  return `<span class="date-cell">${label}</span>`;
}

/* ─── STATUS BADGE ───────────────────────────────────────────────── */
function badge(status) {
  const s = String(status).toLowerCase().trim();
  let cls = 'badge-default';
  if (/fechado|closed|resolvido/i.test(s))              cls = 'badge-closed';
  else if (/em atend|andamento|progress/i.test(s))      cls = 'badge-yellow';
  else if (/espera|aguard|pendente|aberto|open/i.test(s)) cls = 'badge-red';
  return `<span class="badge ${cls}">${status || '—'}</span>`;
}

/* ─── HOME VIEW ──────────────────────────────────────────────────── */
function renderHome() {
  $loading.style.display = 'none';

  const deptMap = {};
  for (const row of allData) {
    const retornado = String(col(row, 'Retornado')).toUpperCase();
    if (retornado === 'SIM') continue; // retornados não contam nos cards normais
    const dept = deptGroupName(row);
    deptMap[dept] = (deptMap[dept] || 0) + 1;
  }
  // Garante que GERENCIA DE CONTAS aparece mesmo que não tenha pendentes
  if (!deptMap['GERENCIA DE CONTAS']) deptMap['GERENCIA DE CONTAS'] = 0;

  const depts = Object.entries(deptMap).sort((a, b) => b[1] - a[1]);
  if ($totalRecs) $totalRecs.textContent = allData.filter(r => String(col(r,'Retornado')).toUpperCase() !== 'SIM').length.toLocaleString('pt-BR');
  const totalRetornados = allData.filter(r => String(col(r,'Retornado')).toUpperCase() === 'SIM').length;
  if ($totalDeps) $totalDeps.textContent = depts.length;

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
      <div class="card-name">${escHtml(name)}</div>
      <div class="card-count">${count.toLocaleString('pt-BR')}</div>
      <div class="card-hint">Clique para ver os detalhes</div>`;
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

  if ($deptTitle) $deptTitle.textContent = deptName;

  const allDeptRows = allData.filter(r => {
    const dept = deptGroupName(r);
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
        const dept = deptGroupName(r);
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
    const found = [...new Set(rows.map(r => filterValue(r, def)).filter(Boolean))];
    const values = def.sort ? def.sort.filter(v => found.includes(v)) : found.sort();

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

    const clientRows = allData.filter(r => {
      const dept = deptGroupName(r);
      const id   = String(col(r,'IdCliente','Id Cliente','ID Cliente','id_cliente')||'').trim();
      const ret  = String(col(r,'Retornado')).toUpperCase() === 'SIM';
      return dept === currentDept && id === clientId &&
             (activeTab === 'retornados' ? ret : !ret);
    });
    // Apenas os chamados que de fato passam pelos filtros ativos — é a partir
    // deles que a célula de Prazo Vencimento da linha é recalculada abaixo,
    // para nunca exibir o prazo de um chamado que não está no resultado filtrado.
    const relevantRows = algumFiltroAtivo ? clientRows.filter(rowMatchesFilters) : clientRows;
    const matchFiltros  = relevantRows.length > 0;

    const prazoCell = tr.querySelector('.prazo-cell');
    if (prazoCell) prazoCell.innerHTML = fmtPrazo(earliestPrazo(relevantRows));

    const visible = matchText && matchFiltros;
    tr.style.display = visible ? '' : 'none';
    const next = tr.nextElementSibling;
    if (next && next.classList.contains('detail-row')) next.style.display = visible ? '' : 'none';
  });

  const visible = $clientBody.querySelectorAll('tr:not(.detail-row):not([style*="display: none"])').length;
  if ($deptCount) $deptCount.textContent = `${visible} registro${visible !== 1 ? 's' : ''}`;
}

/* ─── RENDER DEPT ROWS ───────────────────────────────────────────── */
function renderDeptRows(allDeptRows) {
  const isRet = activeTab === 'retornados';
  const rows  = allDeptRows.filter(r =>
    isRet
      ? String(col(r,'Retornado')).toUpperCase() === 'SIM'
      : String(col(r,'Retornado')).toUpperCase() !== 'SIM'
  );

  if ($deptCount) $deptCount.textContent = `${rows.length} registro${rows.length !== 1 ? 's' : ''}`;

  // Ambas as abas: agrupado por cliente com detalhamento expansível
  const clientMap = {};
  for (const r of rows) {
    const id = String(col(r,'IdCliente','Id Cliente','ID Cliente','id_cliente')||'').trim();
    if (!clientMap[id]) {
      clientMap[id] = {
        id,
        name:      col(r,'Cliente','Nome Cliente','NomeCliente') || id,
        dataCad:   col(r,'Data Cadastro','DataCadastro','Data_Cadastro'),
        rows: []
      };
    }
    clientMap[id].rows.push(r);
  }
  // Prazo Vencimento exibido na linha do cliente = o mais urgente entre
  // todos os chamados dele (não apenas o primeiro encontrado nos dados).
  Object.values(clientMap).forEach(c => { c.prazoVenc = earliestPrazo(c.rows); });

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

  clients.forEach((c, i) => {
    const tr = document.createElement('tr');
    tr.dataset.clientId = c.id;
    if (i % 2 === 1) tr.classList.add('row-even');
    tr.innerHTML = `
      <td class="id-cell">${escHtml(c.id)||'—'}</td>
      <td class="name-cell">${escHtml(String(c.name))}</td>
      <td class="date-cell">${fmt(c.dataCad)}</td>
      <td class="prazo-cell">${fmtPrazo(c.prazoVenc)}</td>`;
    tr.addEventListener('click', () => toggleClientDetail(tr, c));
    $clientBody.appendChild(tr);
  });
}

/* ─── TOGGLE CLIENT DETAIL ───────────────────────────────────────── */
function toggleClientDetail(tr, client) {
  const existingDetail = tr.nextElementSibling;
  const isOpen = existingDetail && existingDetail.classList.contains('detail-row');

  if (openClientRow !== null) {
    const prev = $clientBody.querySelector(`tr[data-client-id="${openClientRow}"]`);
    if (prev) {
      const prevDetail = prev.nextElementSibling;
      if (prevDetail && prevDetail.classList.contains('detail-row')) prevDetail.remove();
    }
  }

  if (isOpen) { openClientRow = null; return; }
  openClientRow = client.id;

  const visibleRows = client.rows.filter(rowMatchesFilters).sort((a, b) => {
    const da = parsePrazoDate(prazoVal(a));
    const db = parsePrazoDate(prazoVal(b));
    if (!da && !db) return 0;
    if (!da) return 1;
    if (!db) return -1;
    return da - db;
  });

  const detailTr = document.createElement('tr');
  detailTr.className = 'detail-row';
  const td = document.createElement('td');
  td.colSpan = 4;

  const showDeptoAnterior = currentDeptTemDeptoAnterior;

  td.innerHTML = `
    <div class="detail-inner">
      <div class="ticket-list">
        ${visibleRows.map(r => `
          <div class="ticket-card">
            <div class="ticket-top">
              <div class="ticket-field"><span class="f-label">ID</span><span class="f-value f-id">${escHtml(String(col(r,'Id','ID','id')||'—'))}</span></div>
              <div class="ticket-field"><span class="f-label">Categoria</span><span class="f-value">${escHtml(String(col(r,'Categoria','categoria')||'—'))}</span></div>
              <div class="ticket-field"><span class="f-label">Responsável</span><span class="f-value">${escHtml(String(col(r,'Responsavel','Responsável','responsavel')||'—'))}</span></div>
              <div class="ticket-field"><span class="f-label">Coordenador</span><span class="f-value">${escHtml(String(col(r,'Coordenador','coordenador')||'—'))}</span></div>
              <div class="ticket-field"><span class="f-label">Depto. Solicitante</span><span class="f-value">${escHtml(String(col(r,'Departamento Solicitante','Departamento_Solicitante')||'—'))}</span></div>
              <div class="ticket-field"><span class="f-label">Data Cadastro</span><span class="f-value">${fmt(col(r,'Data Cadastro','DataCadastro','Data_Cadastro'))}</span></div>
              <div class="ticket-field"><span class="f-label">Prazo Vencimento</span><span class="f-value">${fmtPrazo(prazoVal(r))}</span></div>
              <div class="ticket-field"><span class="f-label">Status</span><span class="f-value">${badge(col(r,'Status','status'))}</span></div>
              ${showDeptoAnterior ? `<div class="ticket-field"><span class="f-label">Depto. Anterior</span><span class="f-value">${escHtml(String(col(r,'Departamento Responsavel Original','Departamento Responsável Original')||'—'))}</span></div>` : ''}
              <div class="ticket-field"><span class="f-label">Solicitante</span><span class="f-value">${escHtml(String(col(r,'Solicitante','solicitante')||'—'))}</span></div>
              <div class="ticket-field"><span class="f-label">Início Atend.</span><span class="f-value">${fmt(col(r,'Inicio Atend.','Início Atend.','Inicio Atendimento','InicioAtend','Inicio_Atend'))}</span></div>
              <div class="ticket-field"><span class="f-label">Previsão Atend.</span><span class="f-value">${fmt(col(r,'DataPrevisaoAtendimento','Data Previsao Atendimento','Data Previsão Atendimento','Data Previsão de Atendimento','Data_Previsao_Atendimento'))}</span></div>
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
  $viewRelatorios.classList.toggle('active', name === 'relatorios');
  const bar    = document.getElementById('unidade-bar');
  const bcDept = document.getElementById('bc-dept');
  if (name === 'dept') {
    bcDept.textContent = currentDept;
    bar.style.display = 'flex';
  } else {
    bar.style.display = 'none';
  }
}

/* ─── RELATÓRIOS — NAVEGAÇÃO ─────────────────────────────────────── */
document.getElementById('btn-relatorios').addEventListener('click', openRelatorios);
document.getElementById('btn-back-relatorios').addEventListener('click', () => showView('home'));

/* ─── HTML ESCAPE ────────────────────────────────────────────────── */
function escHtml(str) {
  return String(str)
    .replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;');
}

/* ═══════════════════════════════════════════════════════════════════
   RELATÓRIOS — COMPARATIVO MENSAL DE ABERTURAS x BAIXAS
   Só trabalha com contagens agregadas (data/relatorios/historico_mensal.json)
   — nenhum detalhe de chamado individual é exibido aqui.
═══════════════════════════════════════════════════════════════════ */
const HIST_PATH = 'data/relatorios/historico_mensal.json';
let historicoMensal = null;

async function carregarHistoricoMensal() {
  try {
    const res = await fetch(HIST_PATH);
    if (!res.ok) throw new Error('histórico não encontrado');
    historicoMensal = await res.json();
  } catch (e) {
    historicoMensal = null;
  }
}

const MESES_ABREV = ['jan','fev','mar','abr','mai','jun','jul','ago','set','out','nov','dez'];
function fmtMes(chave) {
  const [ano, mes] = chave.split('-').map(Number);
  return `${MESES_ABREV[mes - 1]}/${String(ano).slice(2)}`;
}

/* Arredonda o teto do eixo Y para um número "redondo" (1/2/5 x 10^n) */
function niceCeil(v) {
  if (v <= 0) return 10;
  const mag  = Math.pow(10, Math.floor(Math.log10(v)));
  const norm = v / mag;
  let niceNorm;
  if (norm <= 1) niceNorm = 1;
  else if (norm <= 2) niceNorm = 2;
  else if (norm <= 5) niceNorm = 5;
  else niceNorm = 10;
  return niceNorm * mag;
}

async function openRelatorios() {
  if (!historicoMensal) await carregarHistoricoMensal();
  renderRelatorios();
  showView('relatorios');
}

function renderRelatorios() {
  const $empty   = document.getElementById('rel-empty');
  const $content = document.getElementById('rel-content');
  const $atualiz = document.getElementById('rel-atualizado');

  const meses = (historicoMensal && historicoMensal.meses) || [];
  if (!meses.length) {
    $empty.style.display   = 'block';
    $content.style.display = 'none';
    $atualiz.textContent   = '';
    return;
  }
  $empty.style.display   = 'none';
  $content.style.display = 'block';
  $atualiz.textContent   = historicoMensal.atualizado_em
    ? `Base atualizada em ${historicoMensal.atualizado_em}`
    : '';

  buildRelDeptFilter();
  populateRelDeptFilter(historicoMensal.departamentos || []);

  renderKpis(meses);
  renderLegend();
  buildChart(meses);
  renderResumoTable(meses);
}

/* ─── RELATÓRIOS — FILTRO DE DEPARTAMENTOS ────────────────────────── */
const activeRelDeptos    = new Set();
let   relDeptFilterBuilt = false;

/* Contagens do mês, restritas aos departamentos selecionados (soma).
   Nenhum selecionado = sem filtro = totais de todos os departamentos. */
function relMesStats(mes) {
  if (!activeRelDeptos.size) return { abertos: mes.abertos, baixados: mes.baixados };
  let abertos = 0, baixados = 0;
  for (const d of activeRelDeptos) {
    const s = mes.porDepto && mes.porDepto[d];
    if (s) { abertos += s.abertos; baixados += s.baixados; }
  }
  return { abertos, baixados };
}

function buildRelDeptFilter() {
  if (relDeptFilterBuilt) return;
  relDeptFilterBuilt = true;

  const wrap = document.getElementById('rel-filters-row');
  wrap.innerHTML = `
    <div class="cat-filter-wrap" id="rel-depto-filter-wrap">
      <button id="rel-depto-btn" class="cat-btn" type="button" aria-haspopup="true" aria-expanded="false">
        <span id="rel-depto-btn-label">Departamento</span>
        <span class="cat-chevron">▾</span>
      </button>
    </div>`;

  const dd = document.createElement('div');
  dd.id        = 'rel-depto-dropdown';
  dd.className = 'cat-dropdown';
  dd.hidden    = true;
  dd.innerHTML = `
    <div class="cat-dropdown-actions">
      <button type="button" id="rel-depto-select-all">Todos</button>
      <button type="button" id="rel-depto-clear-all">Limpar</button>
    </div>
    <ul id="rel-depto-list" class="cat-list"></ul>`;
  document.body.appendChild(dd);

  const wrapper = document.getElementById('rel-depto-filter-wrap');
  const btn     = document.getElementById('rel-depto-btn');

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

  dd.querySelector('#rel-depto-select-all').addEventListener('click', () => {
    dd.querySelectorAll('input[type=checkbox]').forEach(cb => {
      cb.checked = true;
      activeRelDeptos.add(cb.value);
    });
    updateRelDeptLabel();
    renderRelatorios();
  });
  dd.querySelector('#rel-depto-clear-all').addEventListener('click', () => {
    dd.querySelectorAll('input[type=checkbox]').forEach(cb => cb.checked = false);
    activeRelDeptos.clear();
    updateRelDeptLabel();
    renderRelatorios();
  });

  window.addEventListener('scroll', () => { if (!dd.hidden) positionDropdown(btn, dd); }, true);
  window.addEventListener('resize', () => { if (!dd.hidden) positionDropdown(btn, dd); });
}

function populateRelDeptFilter(departamentos) {
  const $list = document.getElementById('rel-depto-list');
  $list.innerHTML = '';
  for (const dep of departamentos) {
    const li = document.createElement('li');
    li.innerHTML = `
      <label class="cat-option">
        <input type="checkbox" value="${escHtml(dep)}" ${activeRelDeptos.has(dep) ? 'checked' : ''}>
        <span>${escHtml(dep)}</span>
      </label>`;
    li.querySelector('input').addEventListener('change', (e) => {
      if (e.target.checked) activeRelDeptos.add(dep);
      else activeRelDeptos.delete(dep);
      updateRelDeptLabel();
      renderRelatorios();
    });
    $list.appendChild(li);
  }
  updateRelDeptLabel();
}

function updateRelDeptLabel() {
  const $label = document.getElementById('rel-depto-btn-label');
  const $btn   = document.getElementById('rel-depto-btn');
  const total  = document.querySelectorAll('#rel-depto-list input[type=checkbox]').length;
  if (activeRelDeptos.size === 0 || activeRelDeptos.size === total) {
    $label.textContent = 'Departamento';
    $btn.classList.remove('cat-btn--active');
  } else {
    $label.textContent = `Departamento (${activeRelDeptos.size})`;
    $btn.classList.add('cat-btn--active');
  }
}

/* ─── KPIs DO MÊS ATUAL ───────────────────────────────────────────── */
function renderKpis(meses) {
  const $kpi = document.getElementById('kpi-row');
  const chaveAtual = new Date().toISOString().slice(0, 7);
  const mesAtual = meses.find(m => m.mes === chaveAtual);
  const atual = mesAtual ? relMesStats(mesAtual) : { abertos: 0, baixados: 0 };
  const saldo = atual.abertos - atual.baixados;

  const tiles = [
    { label: `Abertos em ${fmtMes(chaveAtual)}`,  value: atual.abertos },
    { label: `Baixados em ${fmtMes(chaveAtual)}`, value: atual.baixados },
    { label: `Saldo em ${fmtMes(chaveAtual)}`,     value: saldo, signed: true },
  ];

  $kpi.innerHTML = '';
  tiles.forEach(t => {
    const div = document.createElement('div');
    div.className = 'kpi-tile';

    const label = document.createElement('div');
    label.className = 'kpi-label';
    label.textContent = t.label;

    const value = document.createElement('div');
    value.className = 'kpi-value';
    value.textContent = (t.signed && t.value > 0 ? '+' : '') + t.value.toLocaleString('pt-BR');

    div.appendChild(label);
    div.appendChild(value);
    $kpi.appendChild(div);
  });
}

/* ─── LEGENDA ─────────────────────────────────────────────────────── */
function renderLegend() {
  const $legend = document.getElementById('chart-legend');
  $legend.innerHTML = '';
  [['sw-abertos', 'Abertos'], ['sw-baixados', 'Baixados']].forEach(([cls, label]) => {
    const wrap = document.createElement('div');
    wrap.className = 'legend-item';
    const sw = document.createElement('span');
    sw.className = 'legend-swatch ' + cls;
    const txt = document.createElement('span');
    txt.textContent = label;
    wrap.appendChild(sw);
    wrap.appendChild(txt);
    $legend.appendChild(wrap);
  });
}

/* ─── TABELA RESUMO ───────────────────────────────────────────────── */
function renderResumoTable(meses) {
  const $body = document.getElementById('rel-table-body');
  $body.innerHTML = '';

  let totalAbertos = 0, totalBaixados = 0;

  meses.forEach((m, i) => {
    const stats = relMesStats(m);
    const saldo = stats.abertos - stats.baixados;
    totalAbertos  += stats.abertos;
    totalBaixados += stats.baixados;

    const tr = document.createElement('tr');
    if (i % 2 === 1) tr.classList.add('row-even');

    const tdMes = document.createElement('td');
    tdMes.textContent = fmtMes(m.mes);
    const tdAb = document.createElement('td');
    tdAb.textContent = stats.abertos.toLocaleString('pt-BR');
    const tdBx = document.createElement('td');
    tdBx.textContent = stats.baixados.toLocaleString('pt-BR');
    const tdSaldo = document.createElement('td');
    tdSaldo.textContent = (saldo > 0 ? '+' : '') + saldo.toLocaleString('pt-BR');

    tr.append(tdMes, tdAb, tdBx, tdSaldo);
    $body.appendChild(tr);
  });

  const totalSaldo = totalAbertos - totalBaixados;
  const trTotal = document.createElement('tr');
  trTotal.className = 'rel-total-row';

  const tdLabel = document.createElement('td');
  tdLabel.textContent = 'Total';
  const tdTotAb = document.createElement('td');
  tdTotAb.textContent = totalAbertos.toLocaleString('pt-BR');
  const tdTotBx = document.createElement('td');
  tdTotBx.textContent = totalBaixados.toLocaleString('pt-BR');
  const tdTotSaldo = document.createElement('td');
  tdTotSaldo.textContent = (totalSaldo > 0 ? '+' : '') + totalSaldo.toLocaleString('pt-BR');

  trTotal.append(tdLabel, tdTotAb, tdTotBx, tdTotSaldo);
  $body.appendChild(trTotal);
}

/* ─── TOOLTIP DO GRÁFICO ──────────────────────────────────────────── */
let $chartTooltip = null;
function getChartTooltip() {
  if (!$chartTooltip) {
    $chartTooltip = document.createElement('div');
    $chartTooltip.className = 'chart-tooltip';
    document.body.appendChild($chartTooltip);
  }
  return $chartTooltip;
}
function showChartTooltip(clientX, clientY, mesLabel, abertos, baixados) {
  const tt = getChartTooltip();
  tt.innerHTML = '';

  const title = document.createElement('div');
  title.className = 'tt-title';
  title.textContent = mesLabel;
  tt.appendChild(title);

  [['Abertos', abertos, 'sw-abertos'], ['Baixados', baixados, 'sw-baixados']].forEach(([label, val, cls]) => {
    const row  = document.createElement('div');
    row.className = 'tt-row';
    const key  = document.createElement('span');
    key.className = 'tt-key';
    const line = document.createElement('span');
    line.className = 'tt-line ' + cls;
    const lbl  = document.createElement('span');
    lbl.textContent = label;
    key.appendChild(line);
    key.appendChild(lbl);
    const value = document.createElement('span');
    value.className = 'tt-val';
    value.textContent = val.toLocaleString('pt-BR');
    row.appendChild(key);
    row.appendChild(value);
    tt.appendChild(row);
  });

  tt.classList.add('visible');
  positionChartTooltip(clientX, clientY, tt);
}
function positionChartTooltip(clientX, clientY, tt) {
  const margin = 12;
  const rect = tt.getBoundingClientRect();
  let x = clientX + margin;
  let y = clientY + margin;
  if (x + rect.width  > window.innerWidth  - 8) x = clientX - rect.width  - margin;
  if (y + rect.height > window.innerHeight - 8) y = clientY - rect.height - margin;
  tt.style.left = `${x}px`;
  tt.style.top  = `${y}px`;
}
function hideChartTooltip() {
  if ($chartTooltip) $chartTooltip.classList.remove('visible');
}

/* ─── GRÁFICO — BARRAS AGRUPADAS (SVG) ────────────────────────────── */
const SVG_NS = 'http://www.w3.org/2000/svg';
function svgEl(tag, attrs) {
  const el = document.createElementNS(SVG_NS, tag);
  for (const k in attrs) el.setAttribute(k, attrs[k]);
  return el;
}
/* Caminho de uma barra com topo arredondado (4px) e base quadrada, crescendo do baseline */
function roundedTopBarPath(x, y, w, h, r) {
  if (h <= 0) return '';
  r = Math.min(r, w / 2, h);
  return `M${x},${y + h} L${x},${y + r} Q${x},${y} ${x + r},${y} ` +
         `L${x + w - r},${y} Q${x + w},${y} ${x + w},${y + r} L${x + w},${y + h} Z`;
}

function buildChart(meses) {
  const svg = document.getElementById('rel-chart');
  svg.innerHTML = '';

  const W = 900, H = 320;
  const marginLeft = 44, marginRight = 12, marginTop = 12, marginBottom = 34;
  const plotW = W - marginLeft - marginRight;
  const plotH = H - marginTop - marginBottom;

  const maxVal  = Math.max(1, ...meses.map(m => { const s = relMesStats(m); return Math.max(s.abertos, s.baixados); }));
  const niceMax = niceCeil(maxVal);
  const steps   = 4;

  // Gridlines + rótulos do eixo Y
  for (let i = 0; i <= steps; i++) {
    const val = (niceMax / steps) * i;
    const y = marginTop + plotH - (val / niceMax) * plotH;
    svg.appendChild(svgEl('line', {
      x1: marginLeft, x2: W - marginRight, y1: y, y2: y,
      style: `stroke:var(--chart-grid);stroke-width:1;opacity:${i === 0 ? 0.6 : 0.35}`,
    }));
    const label = svgEl('text', {
      x: marginLeft - 8, y: y + 4, 'text-anchor': 'end',
      'font-size': 11, 'font-family': 'Segoe UI, Arial, sans-serif',
      style: 'fill:var(--muted)',
    });
    label.textContent = Math.round(val).toLocaleString('pt-BR');
    svg.appendChild(label);
  }

  // Barras agrupadas (Abertos / Baixados) por mês
  const groupW = plotW / meses.length;
  const barW   = Math.min(24, groupW * 0.28);
  const gap    = 2;

  meses.forEach(m => {
    const idx = meses.indexOf(m);
    const groupX = marginLeft + idx * groupW;
    const cx     = groupX + groupW / 2;
    const mesLabel = fmtMes(m.mes);
    const stats  = relMesStats(m);

    [
      { key: 'Abertos',  val: stats.abertos,  colorVar: '--chart-abertos',  x: cx - barW - gap / 2 },
      { key: 'Baixados', val: stats.baixados, colorVar: '--chart-baixados', x: cx + gap / 2 },
    ].forEach(b => {
      const h = niceMax > 0 ? (b.val / niceMax) * plotH : 0;
      const y = marginTop + plotH - h;

      const path = svgEl('path', {
        d: roundedTopBarPath(b.x, y, barW, h, 4),
        style: `fill:var(${b.colorVar})`,
        class: 'chart-bar',
        tabindex: '0',
        role: 'img',
        'aria-label': `${b.key} em ${mesLabel}: ${b.val}`,
      });

      const onEnter = (clientX, clientY) => showChartTooltip(clientX, clientY, mesLabel, stats.abertos, stats.baixados);
      path.addEventListener('pointerenter', e => onEnter(e.clientX, e.clientY));
      path.addEventListener('pointermove',  e => positionChartTooltip(e.clientX, e.clientY, getChartTooltip()));
      path.addEventListener('pointerleave', hideChartTooltip);
      path.addEventListener('focus', () => {
        const r = path.getBoundingClientRect();
        onEnter(r.left + r.width / 2, r.top);
      });
      path.addEventListener('blur', hideChartTooltip);

      svg.appendChild(path);
    });

    const xLabel = svgEl('text', {
      x: cx, y: H - marginBottom + 18, 'text-anchor': 'middle',
      'font-size': 11, 'font-family': 'Segoe UI, Arial, sans-serif',
      style: 'fill:var(--muted)',
    });
    xLabel.textContent = mesLabel;
    svg.appendChild(xLabel);
  });

  // Eixo base
  svg.appendChild(svgEl('line', {
    x1: marginLeft, x2: W - marginRight, y1: marginTop + plotH, y2: marginTop + plotH,
    style: 'stroke:var(--border-accent);stroke-width:1',
  }));
}