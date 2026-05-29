/* ─── CHAMADOS PORTAL — script.js ───────────────────────────────── */

const DATA_PATH = 'data/base.xlsx';

/* ─── STATE ──────────────────────────────────────────────────────── */
let allData       = [];      // raw rows from xlsx
let currentDept   = null;
let openClientRow = null;    // id of the currently expanded client row

/* ─── DOM REFS ───────────────────────────────────────────────────── */
const $loading   = document.getElementById('loading');
const $errorMsg  = document.getElementById('error-msg');
const $viewHome  = document.getElementById('view-home');
const $viewDept  = document.getElementById('view-dept');
const $deptGrid  = document.getElementById('dept-grid');
const $totalRecs = document.getElementById('total-records');
const $totalDeps = document.getElementById('total-depts');
const $clientBody   = document.getElementById('client-body');
const $clientSearch = document.getElementById('client-search');
const $deptTitle    = document.getElementById('dept-title');
const $deptCount    = document.getElementById('dept-count');
const $headerDate   = document.getElementById('header-date');

/* ─── INIT ───────────────────────────────────────────────────────── */
document.addEventListener('DOMContentLoaded', () => {
  $headerDate.textContent = new Date().toLocaleDateString('pt-BR', {
    day: '2-digit', month: 'short', year: 'numeric'
  }).toUpperCase();
  loadData();
});

/* ─── LOAD XLSX ──────────────────────────────────────────────────── */
async function loadData() {
  try {
    const res = await fetch(DATA_PATH);
    if (!res.ok) throw new Error(`Arquivo não encontrado: ${DATA_PATH} (HTTP ${res.status})`);
    const buf  = await res.arrayBuffer();
    const wb   = XLSX.read(new Uint8Array(buf), { type: 'array', cellDates: true });
    const ws   = wb.Sheets[wb.SheetNames[0]];
    allData    = XLSX.utils.sheet_to_json(ws, { defval: '' });
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
/* Finds the actual column key regardless of small spacing/case differences */
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
  if (val instanceof Date) {
    return val.toLocaleDateString('pt-BR');
  }
  if (typeof val === 'number') {
    // Excel serial date
    const d = new Date(Math.round((val - 25569) * 86400 * 1000));
    if (!isNaN(d)) return d.toLocaleDateString('pt-BR');
  }
  return String(val);
}

/* ─── STATUS BADGE ───────────────────────────────────────────────── */
function badge(status) {
  const s = String(status).toLowerCase().trim();
  let cls = 'badge-default';
  if (/aberto|open|pendente/i.test(s))    cls = 'badge-open';
  else if (/fechado|closed|resolvido/i.test(s)) cls = 'badge-closed';
  else if (/em atend|andamento|progress/i.test(s)) cls = 'badge-pending';
  return `<span class="badge ${cls}">${status || '—'}</span>`;
}

/* ─── HOME VIEW ──────────────────────────────────────────────────── */
function renderHome() {
  $loading.style.display = 'none';

  // Group by department
  const deptMap = {};
  for (const row of allData) {
    const dept = String(col(row, 'Departamento Responsavel', 'Departamento Responsável', 'Departamento') || 'Sem Departamento').trim();
    deptMap[dept] = (deptMap[dept] || 0) + 1;
  }

  const depts  = Object.entries(deptMap).sort((a, b) => b[1] - a[1]);

  $totalRecs.textContent = allData.length.toLocaleString('pt-BR');
  $totalDeps.textContent = depts.length;

  $deptGrid.innerHTML = '';
  for (const [name, count] of depts) {
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
      </div>`;
    card.addEventListener('click', () => openDept(name));
    $deptGrid.appendChild(card);
  }

  showView('home');
}

/* ─── DEPT VIEW ──────────────────────────────────────────────────── */
function openDept(deptName) {
  currentDept   = deptName;
  openClientRow = null;

  // Filter rows for this dept
  const rows = allData.filter(r =>
    String(col(r, 'Departamento Responsavel', 'Departamento Responsável', 'Departamento') || '').trim() === deptName
  );

  // Build unique client list (by IdCliente)
  const clientMap = {};
  for (const r of rows) {
    const id = String(col(r, 'IdCliente', 'Id Cliente', 'ID Cliente', 'id_cliente') || '').trim();
    if (!clientMap[id]) {
      clientMap[id] = {
        id,
        name: col(r, 'Cliente', 'Nome Cliente', 'NomeCliente') || id,
        dataCad: col(r, 'Data Cadastro', 'DataCadastro', 'Data_Cadastro'),
        rows: []
      };
    }
    clientMap[id].rows.push(r);
  }

  const clients = Object.values(clientMap);

  $deptTitle.textContent = deptName;
  $deptCount.textContent = `${rows.length} registro${rows.length !== 1 ? 's' : ''}`;

  renderClientTable(clients);
  showView('dept');
}

function renderClientTable(clients) {
  $clientBody.innerHTML = '';

  if (!clients.length) {
    $clientBody.innerHTML = `<tr><td colspan="3" class="empty">Nenhum cliente encontrado.</td></tr>`;
    return;
  }

  for (const c of clients) {
    const tr = document.createElement('tr');
    tr.dataset.clientId = c.id;
    tr.innerHTML = `
      <td class="id-cell">${escHtml(c.id) || '—'}</td>
      <td class="name-cell">${escHtml(String(c.name))}</td>
      <td class="date-cell">${fmt(c.dataCad)}</td>`;
    tr.addEventListener('click', () => toggleClientDetail(tr, c));
    $clientBody.appendChild(tr);
  }
}

/* ─── TOGGLE CLIENT DETAIL ───────────────────────────────────────── */
function toggleClientDetail(tr, client) {
  const existingDetail = tr.nextElementSibling;
  const isOpen = existingDetail && existingDetail.classList.contains('detail-row');

  // Close any open panel
  if (openClientRow) {
    const prev = $clientBody.querySelector(`tr[data-client-id="${openClientRow}"]`);
    if (prev) {
      const prevDetail = prev.nextElementSibling;
      if (prevDetail && prevDetail.classList.contains('detail-row')) {
        prevDetail.remove();
      }
    }
  }

  if (isOpen) {
    openClientRow = null;
    return;
  }

  openClientRow = client.id;

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
              <th>ID</th>
              <th>Categoria</th>
              <th>Solicitação</th>
              <th>Responsável</th>
              <th>Data Cadastro</th>
              <th>Status</th>
              <th>Solicitante</th>
              <th>Início Atend.</th>
            </tr>
          </thead>
          <tbody>
            ${client.rows.map(r => `
              <tr>
                <td class="id-cell">${escHtml(String(col(r,'Id','ID','id') || '—'))}</td>
                <td>${escHtml(String(col(r,'Categoria','categoria') || '—'))}</td>
                <td>${escHtml(String(col(r,'Solicitacao','Solicitação','solicitacao','solicitação') || '—'))}</td>
                <td>${escHtml(String(col(r,'Responsavel','Responsável','responsavel') || '—'))}</td>
                <td class="date-cell">${fmt(col(r,'Data Cadastro','DataCadastro','Data_Cadastro'))}</td>
                <td>${badge(col(r,'Status','status'))}</td>
                <td>${escHtml(String(col(r,'Solicitante','solicitante') || '—'))}</td>
                <td class="date-cell">${fmt(col(r,'Inicio Atend.','Início Atend.','Inicio Atendimento','InicioAtend','Inicio_Atend'))}</td>
              </tr>`).join('')}
          </tbody>
        </table>
      </div>
    </div>`;

  detailTr.appendChild(td);
  tr.insertAdjacentElement('afterend', detailTr);
}

/* ─── SEARCH CLIENTS ─────────────────────────────────────────────── */
$clientSearch.addEventListener('input', () => {
  const q = $clientSearch.value.toLowerCase();
  $clientBody.querySelectorAll('tr:not(.detail-row)').forEach(tr => {
    const visible = tr.textContent.toLowerCase().includes(q);
    tr.style.display = visible ? '' : 'none';
    // hide orphan detail panel when parent is hidden
    const next = tr.nextElementSibling;
    if (next && next.classList.contains('detail-row')) {
      next.style.display = visible ? '' : 'none';
    }
  });
});

/* ─── BACK BUTTON ────────────────────────────────────────────────── */
document.getElementById('btn-back').addEventListener('click', () => {
  openClientRow = null;
  $clientSearch.value = '';
  showView('home');
});

/* ─── BREADCRUMB ─────────────────────────────────────────────────── */
document.getElementById('bc-home').addEventListener('click', () => {
  if ($viewDept.classList.contains('active')) {
    openClientRow = null;
    $clientSearch.value = '';
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
    bcDept.textContent = currentDept;
    bcDept.style.display = '';
    bcSep.style.display  = '';
  } else {
    bcDept.style.display = 'none';
    bcSep.style.display  = 'none';
  }
}

/* ─── HTML ESCAPE ────────────────────────────────────────────────── */
function escHtml(str) {
  return String(str)
    .replace(/&/g,'&amp;')
    .replace(/</g,'&lt;')
    .replace(/>/g,'&gt;')
    .replace(/"/g,'&quot;');
}