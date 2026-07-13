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
const $viewRelatoriosHome = document.getElementById('view-relatorios-home');
const $viewRelatorios     = document.getElementById('view-relatorios');
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
  carregarDetalheMensal();
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
  else if (typeof val === 'string') {
    // Datas do detalhe_mensal.json vêm como 'YYYY-MM-DD' (sem hora/fuso) —
    // trata como meia-noite UTC, igual às datas do Excel, para cair no
    // mesmo ajuste de fuso abaixo.
    const m = val.match(/^(\d{4})-(\d{2})-(\d{2})/);
    if (!m) return null;
    date = new Date(Date.UTC(Number(m[1]), Number(m[2]) - 1, Number(m[3])));
  }
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
  } else if ($viewRelatorios.classList.contains('active')) {
    showView('home');
  }
});

/* ─── SHOW VIEW ──────────────────────────────────────────────────── */
function showView(name) {
  $viewHome.classList.toggle('active', name === 'home');
  $viewDept.classList.toggle('active', name === 'dept');
  $viewRelatoriosHome.classList.toggle('active', name === 'relatorios-home');
  $viewRelatorios.classList.toggle('active', name === 'relatorios');
  const bar    = document.getElementById('unidade-bar');
  const bcDept = document.getElementById('bc-dept');
  if (name === 'dept') {
    bcDept.textContent = currentDept;
    bar.style.display = 'flex';
  } else if (name === 'relatorios') {
    bcDept.textContent = currentRelDept ? currentRelDept.nome : '';
    bar.style.display = 'flex';
  } else {
    bar.style.display = 'none';
  }
}

/* ─── RELATÓRIOS — NAVEGAÇÃO ─────────────────────────────────────── */
document.getElementById('btn-relatorios').addEventListener('click', openRelatorios);
document.getElementById('btn-back-relatorios-home').addEventListener('click', () => showView('home'));
document.getElementById('btn-back-relatorios').addEventListener('click', () => showView('relatorios-home'));

/* ─── HTML ESCAPE ────────────────────────────────────────────────── */
function escHtml(str) {
  return String(str)
    .replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;');
}

/* ═══════════════════════════════════════════════════════════════════
   RELATÓRIOS — POR DEPARTAMENTO, ATENDIDOS DENTRO x FORA DO PRAZO
   Trabalha com data/relatorios/detalhe_mensal.json — chamado a chamado,
   agrupado por depto/mês/categoria (mês = mês de abertura do chamado).
═══════════════════════════════════════════════════════════════════ */
const DETALHE_PATH = 'data/relatorios/detalhe_mensal.json';
let detalheMensal  = null;
let currentRelDept = null; // { nome, meses } do departamento aberto no relatório
let mesExpandido   = null; // chave 'YYYY-MM' do mês expandido na tabela, ou null

async function carregarDetalheMensal() {
  try {
    const res = await fetch(DETALHE_PATH);
    if (!res.ok) throw new Error('detalhe mensal não encontrado');
    detalheMensal = await res.json();
  } catch (e) {
    detalheMensal = null;
  }
}

const MESES_ABREV = ['jan','fev','mar','abr','mai','jun','jul','ago','set','out','nov','dez'];
function fmtMes(chave) {
  const [ano, mes] = chave.split('-').map(Number);
  return `${MESES_ABREV[mes - 1]}/${String(ano).slice(2)}`;
}

/* ─── RELATÓRIOS — ESCOLHA DE DEPARTAMENTO ────────────────────────── */
async function openRelatorios() {
  if (!detalheMensal) await carregarDetalheMensal();
  renderRelatoriosHome();
  showView('relatorios-home');
}

function totalAbertosDept(dept) {
  return dept.meses.reduce((soma, m) => soma + m.abertos, 0);
}

function renderRelatoriosHome() {
  const $empty   = document.getElementById('rel-home-empty');
  const $grid    = document.getElementById('rel-dept-grid');
  const $atualiz = document.getElementById('rel-home-atualizado');

  const deptos = (detalheMensal && detalheMensal.departamentos) || [];
  if (!deptos.length) {
    $empty.style.display = 'block';
    $grid.innerHTML = '';
    $atualiz.textContent = '';
    return;
  }
  $empty.style.display = 'none';
  $atualiz.textContent = detalheMensal.atualizado_em
    ? `Base atualizada em ${detalheMensal.atualizado_em}`
    : '';

  $grid.innerHTML = '';
  const ordenados = [...deptos].sort((a, b) => totalAbertosDept(b) - totalAbertosDept(a));
  for (const dept of ordenados) {
    const card = document.createElement('div');
    card.className = 'dept-card';
    card.innerHTML = `
      <div class="card-name">${escHtml(dept.nome)}</div>
      <div class="card-count">${totalAbertosDept(dept).toLocaleString('pt-BR')}</div>
      <div class="card-hint">Clique para ver o relatório</div>`;
    card.addEventListener('click', () => openRelatorioDept(dept.nome));
    $grid.appendChild(card);
  }
}

/* ─── RELATÓRIOS — TELA DO DEPARTAMENTO ───────────────────────────── */
function categoriasDoDept(dept) {
  const set = new Set();
  for (const m of dept.meses) for (const c of m.categorias) set.add(c.nome);
  return [...set].sort();
}

function openRelatorioDept(nome) {
  currentRelDept = (detalheMensal.departamentos || []).find(d => d.nome === nome) || { nome, meses: [] };
  mesExpandido   = null;
  activeRelCategorias.clear();

  document.getElementById('rel-dept-title').textContent = `Relatório de Chamados Entregues — ${nome}`;

  buildRelCategoriaFilter();
  populateRelCategoriaFilter(categoriasDoDept(currentRelDept));

  renderRelatorioDept();
  showView('relatorios');
}

function renderRelatorioDept() {
  const $empty   = document.getElementById('rel-empty');
  const $content = document.getElementById('rel-content');
  const $atualiz = document.getElementById('rel-atualizado');

  const meses = (currentRelDept && currentRelDept.meses) || [];
  if (!meses.length) {
    $empty.style.display   = 'block';
    $content.style.display = 'none';
    $atualiz.textContent   = '';
    return;
  }
  $empty.style.display   = 'none';
  $content.style.display = 'block';
  $atualiz.textContent   = detalheMensal.atualizado_em
    ? `Base atualizada em ${detalheMensal.atualizado_em}`
    : '';

  renderLegend();
  renderResumoTable(meses);
  buildPrazoChart();
}

/* ─── RELATÓRIOS — FILTRO DE CATEGORIA ────────────────────────────── */
const activeRelCategorias    = new Set();
let   relCategoriaFilterBuilt = false;

/* Restringe uma lista de categorias às selecionadas no filtro (soma).
   Nenhuma selecionada = sem filtro = todas as categorias. */
function filtrarCategorias(categorias) {
  if (!activeRelCategorias.size) return categorias;
  return categorias.filter(c => activeRelCategorias.has(c.nome));
}
function categoriasFiltradas(mesObj) {
  return filtrarCategorias(mesObj.categorias);
}
function mesStatsFiltradas(mesObj) {
  return categoriasFiltradas(mesObj).reduce((acc, c) => {
    acc.atendidosDentro += c.atendidosDentro;
    acc.atendidosFora   += c.atendidosFora;
    acc.pendentes       += c.pendentes;
    acc.atendidos       += c.atendidosDentro + c.atendidosFora;
    acc.abertos         += c.atendidosDentro + c.atendidosFora + c.pendentes;
    return acc;
  }, { abertos: 0, atendidos: 0, pendentes: 0, atendidosDentro: 0, atendidosFora: 0 });
}

/* Totais por categoria somando TODOS os meses do departamento — visão
   padrão do gráfico (sem nenhum mês selecionado na tabela). */
function categoriasAgregadas(dept) {
  const totais = new Map();
  for (const m of dept.meses) {
    for (const c of m.categorias) {
      const acc = totais.get(c.nome) || { nome: c.nome, atendidosDentro: 0, atendidosFora: 0 };
      acc.atendidosDentro += c.atendidosDentro;
      acc.atendidosFora   += c.atendidosFora;
      totais.set(c.nome, acc);
    }
  }
  return [...totais.values()].sort((a, b) => a.nome.localeCompare(b.nome));
}

function buildRelCategoriaFilter() {
  if (relCategoriaFilterBuilt) return;
  relCategoriaFilterBuilt = true;

  const wrap = document.getElementById('rel-filters-row');
  wrap.innerHTML = `
    <div class="cat-filter-wrap" id="rel-categoria-filter-wrap">
      <button id="rel-categoria-btn" class="cat-btn" type="button" aria-haspopup="true" aria-expanded="false">
        <span id="rel-categoria-btn-label">Categoria</span>
        <span class="cat-chevron">▾</span>
      </button>
    </div>`;

  const dd = document.createElement('div');
  dd.id        = 'rel-categoria-dropdown';
  dd.className = 'cat-dropdown';
  dd.hidden    = true;
  dd.innerHTML = `
    <div class="cat-dropdown-actions">
      <button type="button" id="rel-categoria-select-all">Todos</button>
      <button type="button" id="rel-categoria-clear-all">Limpar</button>
    </div>
    <ul id="rel-categoria-list" class="cat-list"></ul>`;
  document.body.appendChild(dd);

  const wrapper = document.getElementById('rel-categoria-filter-wrap');
  const btn     = document.getElementById('rel-categoria-btn');

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

  dd.querySelector('#rel-categoria-select-all').addEventListener('click', () => {
    dd.querySelectorAll('input[type=checkbox]').forEach(cb => {
      cb.checked = true;
      activeRelCategorias.add(cb.value);
    });
    updateRelCategoriaLabel();
    renderRelatorioDept();
  });
  dd.querySelector('#rel-categoria-clear-all').addEventListener('click', () => {
    dd.querySelectorAll('input[type=checkbox]').forEach(cb => cb.checked = false);
    activeRelCategorias.clear();
    updateRelCategoriaLabel();
    renderRelatorioDept();
  });

  window.addEventListener('scroll', () => { if (!dd.hidden) positionDropdown(btn, dd); }, true);
  window.addEventListener('resize', () => { if (!dd.hidden) positionDropdown(btn, dd); });
}

function populateRelCategoriaFilter(categorias) {
  const $list = document.getElementById('rel-categoria-list');
  $list.innerHTML = '';
  for (const cat of categorias) {
    const li = document.createElement('li');
    li.innerHTML = `
      <label class="cat-option">
        <input type="checkbox" value="${escHtml(cat)}">
        <span>${escHtml(cat)}</span>
      </label>`;
    li.querySelector('input').addEventListener('change', (e) => {
      if (e.target.checked) activeRelCategorias.add(cat);
      else activeRelCategorias.delete(cat);
      updateRelCategoriaLabel();
      renderRelatorioDept();
    });
    $list.appendChild(li);
  }
  updateRelCategoriaLabel();
}

function updateRelCategoriaLabel() {
  const $label = document.getElementById('rel-categoria-btn-label');
  const $btn   = document.getElementById('rel-categoria-btn');
  const total  = document.querySelectorAll('#rel-categoria-list input[type=checkbox]').length;
  if (activeRelCategorias.size === 0 || activeRelCategorias.size === total) {
    $label.textContent = 'Categoria';
    $btn.classList.remove('cat-btn--active');
  } else {
    $label.textContent = `Categoria (${activeRelCategorias.size})`;
    $btn.classList.add('cat-btn--active');
  }
}

/* ─── STATUS DO CHAMADO NO RELATÓRIO ──────────────────────────────── */
function badgeRelStatus(status) {
  if (status === 'atendido_dentro') return '<span class="badge badge-closed">Atendido dentro do prazo</span>';
  if (status === 'atendido_fora')   return '<span class="badge badge-red">Atendido fora do prazo</span>';
  return '<span class="badge badge-yellow">Pendente</span>';
}

/* ─── LEGENDA ─────────────────────────────────────────────────────── */
function renderLegend() {
  const $legend = document.getElementById('chart-legend');
  $legend.innerHTML = '';
  [
    ['sw-dentro', 'Dentro do prazo'],
    ['sw-fora',   'Fora do prazo'],
  ].forEach(([cls, label]) => {
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

  let totalAbertos = 0, totalAtendidos = 0, totalPendentes = 0;

  meses.forEach((m, i) => {
    const stats = mesStatsFiltradas(m);
    totalAbertos   += stats.abertos;
    totalAtendidos += stats.atendidos;
    totalPendentes += stats.pendentes;

    const tr = document.createElement('tr');
    tr.dataset.mes = m.mes;
    if (i % 2 === 1) tr.classList.add('row-even');

    const tdMes = document.createElement('td');
    tdMes.textContent = fmtMes(m.mes);
    const tdAb = document.createElement('td');
    tdAb.className = 'num-cell';
    tdAb.textContent = stats.abertos.toLocaleString('pt-BR');
    const tdAt = document.createElement('td');
    tdAt.className = 'num-cell';
    tdAt.textContent = stats.atendidos.toLocaleString('pt-BR');
    const tdPe = document.createElement('td');
    tdPe.className = 'num-cell';
    tdPe.textContent = stats.pendentes.toLocaleString('pt-BR');

    tr.append(tdMes, tdAb, tdAt, tdPe);
    tr.addEventListener('click', () => toggleMesDetail(tr, m));
    $body.appendChild(tr);

    if (mesExpandido === m.mes) abrirDetalheMes(tr, m);
  });

  const trTotal = document.createElement('tr');
  trTotal.className = 'rel-total-row';

  const tdLabel = document.createElement('td');
  tdLabel.textContent = 'Total';
  const tdTotAb = document.createElement('td');
  tdTotAb.className = 'num-cell';
  tdTotAb.textContent = totalAbertos.toLocaleString('pt-BR');
  const tdTotAt = document.createElement('td');
  tdTotAt.className = 'num-cell';
  tdTotAt.textContent = totalAtendidos.toLocaleString('pt-BR');
  const tdTotPe = document.createElement('td');
  tdTotPe.className = 'num-cell';
  tdTotPe.textContent = totalPendentes.toLocaleString('pt-BR');

  trTotal.append(tdLabel, tdTotAb, tdTotAt, tdTotPe);
  $body.appendChild(trTotal);
}

/* ─── EXPANDIR MÊS — CHAMADOS AGRUPADOS POR CATEGORIA ─────────────── */
function abrirDetalheMes(tr, mesObj) {
  const existing = tr.nextElementSibling;
  if (existing && existing.classList.contains('detail-row')) existing.remove();

  const categorias = categoriasFiltradas(mesObj);

  const detailTr = document.createElement('tr');
  detailTr.className = 'detail-row';
  const td = document.createElement('td');
  td.colSpan = 4;

  td.innerHTML = `
    <div class="detail-inner">
      ${categorias.length ? categorias.map(cat => `
        <div class="categoria-group">
          <div class="categoria-group-title">
            ${escHtml(cat.nome)}
            <span class="categoria-group-count">${cat.chamados.length} chamado${cat.chamados.length !== 1 ? 's' : ''}</span>
          </div>
          <div class="ticket-list">
            ${cat.chamados.map(c => `
              <div class="ticket-card">
                <div class="ticket-top">
                  <div class="ticket-field"><span class="f-label">ID</span><span class="f-value f-id">${escHtml(String(c.id ?? '—'))}</span></div>
                  <div class="ticket-field"><span class="f-label">Cliente</span><span class="f-value">${escHtml(c.cliente || '—')}</span></div>
                  <div class="ticket-field"><span class="f-label">Data Cadastro</span><span class="f-value">${fmt(c.dataCadastro)}</span></div>
                  <div class="ticket-field"><span class="f-label">Prazo Vencimento</span><span class="f-value">${c.status === 'pendente' ? fmtPrazo(c.prazo) : fmt(c.prazo)}</span></div>
                  <div class="ticket-field"><span class="f-label">Data Entrega</span><span class="f-value">${fmt(c.dataEntrega)}</span></div>
                  <div class="ticket-field"><span class="f-label">Status</span><span class="f-value">${badgeRelStatus(c.status)}</span></div>
                </div>
              </div>`).join('')}
          </div>
        </div>`).join('') : '<p class="empty">Nenhum chamado nessa combinação de filtros.</p>'}
    </div>`;

  detailTr.appendChild(td);
  tr.insertAdjacentElement('afterend', detailTr);
}

function toggleMesDetail(tr, mesObj) {
  const existingDetail = tr.nextElementSibling;
  const isOpen = existingDetail && existingDetail.classList.contains('detail-row');

  if (mesExpandido !== null) {
    const prevTr = document.querySelector(`#rel-table-body tr[data-mes="${mesExpandido}"]`);
    if (prevTr) {
      const prevDetail = prevTr.nextElementSibling;
      if (prevDetail && prevDetail.classList.contains('detail-row')) prevDetail.remove();
    }
  }

  if (isOpen) {
    mesExpandido = null;
  } else {
    mesExpandido = mesObj.mes;
    abrirDetalheMes(tr, mesObj);
  }
  buildPrazoChart();
}

/* Gráfico "Atendidos dentro x fora do prazo" — sempre por categoria (uma
   linha por categoria, crescendo pra baixo conforme necessário — cabe
   qualquer quantidade sem espremer nada). Sem mês selecionado, soma todas
   as categorias de todos os meses do departamento; com um mês selecionado
   na tabela, mostra só as categorias daquele mês. */
function buildPrazoChart() {
  const hbars      = document.getElementById('rel-chart-hbars');
  const $titulo    = document.getElementById('rel-chart-title');
  const $subtitulo = document.getElementById('rel-chart-subtitle');

  let categorias;
  if (mesExpandido !== null) {
    const mesObj = currentRelDept.meses.find(m => m.mes === mesExpandido);
    categorias = categoriasFiltradas(mesObj);
    $titulo.textContent    = `Atendidos dentro x fora do prazo — ${fmtMes(mesExpandido)}`;
    $subtitulo.textContent = 'Por categoria, só nesse mês — clique no mês de novo para ver todos os meses';
  } else {
    categorias = filtrarCategorias(categoriasAgregadas(currentRelDept));
    $titulo.textContent    = 'Atendidos dentro x fora do prazo';
    $subtitulo.textContent = 'Por categoria, somando todos os meses — clique num mês na tabela para ver só aquele mês';
  }

  const itens = categorias.map(c => ({
    label: c.nome, dentro: c.atendidosDentro, fora: c.atendidosFora,
  }));

  buildHBarChart(hbars, itens);
}

/* Gráfico de barras horizontais — uma linha por item (categoria), com duas
   barrinhas (dentro/fora do prazo) crescendo da esquerda pra direita. HTML
   puro (não SVG): o rótulo é texto normal (trunca com "..." + title nativo
   em vez de sobrepor), e a altura cresce com a quantidade de itens em vez
   de espremer a largura. */
function buildHBarChart(container, itens) {
  const maxVal = Math.max(1, ...itens.map(it => Math.max(it.dentro, it.fora)));

  container.innerHTML = itens.map(it => {
    const pctDentro = Math.round((it.dentro / maxVal) * 100);
    const pctFora   = Math.round((it.fora   / maxVal) * 100);
    return `
      <div class="hbar-row">
        <div class="hbar-label" title="${escHtml(it.label)}">${escHtml(it.label)}</div>
        <div class="hbar-series">
          <div class="hbar-serie-row" title="Dentro do prazo em ${escHtml(it.label)}: ${it.dentro}">
            <div class="hbar-track"><div class="hbar-fill hbar-fill--dentro" style="width:${pctDentro}%"></div></div>
            <span class="hbar-value">${it.dentro.toLocaleString('pt-BR')}</span>
          </div>
          <div class="hbar-serie-row" title="Fora do prazo em ${escHtml(it.label)}: ${it.fora}">
            <div class="hbar-track"><div class="hbar-fill hbar-fill--fora" style="width:${pctFora}%"></div></div>
            <span class="hbar-value">${it.fora.toLocaleString('pt-BR')}</span>
          </div>
        </div>
      </div>`;
  }).join('');
}