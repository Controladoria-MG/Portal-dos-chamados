"""
atualizar_base.py
─────────────────
Lê os relatórios de chamados (principal, GC e CTR) da pasta Att Base,
substitui usuários AD das colunas I e Q pelo nome de exibição,
filtra status 'Entregue Ao Solicitante' — EXCETO para registros cujo
Departamento Responsavel seja 'GERENCIA DE CONTAS' ou 'GC - ADMINISTRATIVO',
que são salvos com a flag Retornado = TRUE na coluna R.

Estrutura esperada:
  Chamados/
  ├── data/
  │   ├── base.xlsx                          ← será sobrescrito
  │   └── Att Base/
  │       ├── relatorio_chamados_*.xlsx
  │       ├── *GC*.xlsx
  │       ├── *CTR*.xlsx
  │       └── Relatório_Colaboradores_*.xlsx
  └── atualizar_base.py
"""

import os
import sys
from pathlib import Path

try:
    import openpyxl
except ImportError:
    print("Instalando openpyxl...")
    os.system(f"{sys.executable} -m pip install openpyxl")
    import openpyxl

from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.styles import Font, Alignment
from openpyxl.utils import get_column_letter

# ─── CAMINHOS ────────────────────────────────────────────────────────
# BASE_DIR aponta para a pasta pai do script (funciona em backend/ ou raiz)
BASE_DIR = Path(__file__).parent
# Se o script está em backend/, sobe um nível para encontrar data/
DATA_DIR = BASE_DIR / "data" if (BASE_DIR / "data").exists() else BASE_DIR.parent / "data"
ATT_DIR  = DATA_DIR / "Att Base"
OUTPUT   = DATA_DIR / "base.xlsx"

# ─── DEPARTAMENTOS QUE SALVAM RETORNADOS ─────────────────────────────
# Verifica AMBAS as colunas: D (Depto Solicitante) E J (Depto Responsavel)
# Se qualquer uma delas for GC, o item é marcado como Retornado
DEPTS_RETORNADOS = {"gerencia de contas", "gc - administrativo"}

# ─── LOCALIZAR ARQUIVOS AUTOMATICAMENTE ──────────────────────────────
def find_file(folder, pattern, required=True):
    matches = list(folder.glob(pattern))
    if not matches:
        if required:
            raise FileNotFoundError(f"Nenhum arquivo encontrado com padrão '{pattern}' em {folder}")
        return None
    return max(matches, key=os.path.getmtime)

def find_all(folder, pattern):
    return list(folder.glob(pattern))

print("Localizando arquivos...")
colab_path = find_file(ATT_DIR, "*olaboradores*.xlsx")

# Arquivos específicos por sufixo (pegam o mais recente de cada tipo)
gc_path  = find_file(ATT_DIR, "*GC*.xlsx",  required=False)
ctr_path = find_file(ATT_DIR, "*CTR*.xlsx", required=False)

# Arquivos de chamados genéricos: todos com "chamados" no nome
# EXCLUINDO os que já foram identificados como GC ou CTR
gc_name  = gc_path.name  if gc_path  else None
ctr_name = ctr_path.name if ctr_path else None

chamados_paths = [
    p for p in find_all(ATT_DIR, "*chamados*.xlsx")
    if p.name != gc_name and p.name != ctr_name
]

if not chamados_paths:
    raise FileNotFoundError(f"Nenhum arquivo *chamados*.xlsx (não-GC/CTR) encontrado em {ATT_DIR}")

print(f"  Colaboradores: {colab_path.name}")
for p in chamados_paths:
    print(f"  Chamados:      {p.name}")
print(f"  GC:            {gc_name  if gc_name  else '⚠ não encontrado'}")
print(f"  CTR:           {ctr_name if ctr_name else '⚠ não encontrado'}")

# ─── CARREGAR COLABORADORES ───────────────────────────────────────────
print("\nCarregando colaboradores...")
wb_colab = openpyxl.load_workbook(colab_path, read_only=True, data_only=True)
ws_colab = wb_colab.active

colab_map = {}
for row in ws_colab.iter_rows(min_row=2, values_only=True):
    nome_exib = row[1]  # Coluna B - NOME EXIBICAO
    usuario   = row[2]  # Coluna C - USUARIO AD
    if usuario and nome_exib:
        colab_map[str(usuario).strip().lower()] = str(nome_exib).strip()

print(f"  {len(colab_map)} colaboradores carregados.")

def resolver_nome(valor):
    if not valor:
        return valor
    chave = str(valor).strip().lower()
    return colab_map.get(chave, str(valor).strip())

# ─── MAPEAMENTO DE COLUNAS (índice 0-based) ──────────────────────────
COL_MAP = [
    (0,  0),   # A → A  | Id
    (1,  1),   # B → B  | Categoria
    (2,  2),   # C → C  | Assunto
    (3,  3),   # D → D  | Departamento Solicitante
    (4,  4),   # E → E  | Solicitação
    (5,  5),   # F → F  | IdCliente
    (6,  6),   # G → G  | Cliente
    (8,  7),   # I → H  | Responsável (com lookup)
    (9,  8),   # J → I  | Departamento Responsavel
    (13, 9),   # N → J  | Data Cadastro
    (14, 10),  # O → K  | Prazo Vencimento
    (15, 11),  # P → L  | Status
    (16, 12),  # Q → M  | Solicitante (com lookup)
    (17, 13),  # R → N  | Inicio Atend.
    (18, 14),  # S → O  | Data Entrega
    (19, 15),  # T → P  | Responsável pela conclusão
    (20, 16),  # U → Q  | Ultimo Comentário
    # coluna R (índice 17 na base) = Retornado ← gerada pelo script
]

LOOKUP_COLS = {8, 16}  # I e Q do relatório de origem

# ─── CABEÇALHOS DA BASE ───────────────────────────────────────────────
HEADERS = [
    "Id", "Categoria", "Assunto", "Departamento Solicitante",
    "Solicitação", "IdCliente", "Cliente", "Responsável",
    "Departamento Responsavel", "Data Cadastro", "Prazo Vencimento",
    "Status", "Solicitante", "Inicio Atend.", "Data Entrega",
    "Responsável pela conclusão", "Ultimo Comentário",
    "Retornado",   # ← nova coluna R: TRUE para itens retornados ao solicitante
]

STATUS_ENTREGUE = "entregue ao solicitante"

# ─── FUNÇÃO GENÉRICA DE PROCESSAMENTO ────────────────────────────────
def processar_relatorio(ws, label, linhas_total, ignoradas_total,
                        retornados_total, nao_encontrados, ws_out):
    linhas    = 0
    ignoradas = 0
    retornados = 0

    for row in ws.iter_rows(min_row=2, values_only=True):
        if not any(row):
            continue

        # Coluna P (índice 15) = Status no relatório original
        status_val = str(row[15]).strip().lower() if row[15] else ''
        # Coluna D (índice 3) = Departamento Solicitante — único critério para GC
        dept_sol   = str(row[3]).strip().lower()  if row[3]  else ''

        eh_entregue = (status_val == STATUS_ENTREGUE)
        eh_gc       = (dept_sol in DEPTS_RETORNADOS)

        if eh_entregue and not eh_gc:
            # Ignora normalmente
            ignoradas += 1
            continue

        nova_linha = [None] * len(HEADERS)

        for col_src, col_dst in COL_MAP:
            valor = row[col_src] if col_src < len(row) else None
            if col_src in LOOKUP_COLS:
                nome_resolvido = resolver_nome(valor)
                if valor and nome_resolvido == str(valor).strip():
                    nao_encontrados.add(str(valor).strip())
                valor = nome_resolvido
            nova_linha[col_dst] = valor

        # Marca coluna Retornado (índice 17) como texto para leitura segura no browser
        if eh_entregue and eh_gc:
            nova_linha[17] = "SIM"
            retornados += 1
        else:
            nova_linha[17] = "NÃO"

        ws_out.append(nova_linha)
        linhas += 1

    print(f"  {label}: {linhas} incluídos  |  {ignoradas} ignorados  |  {retornados} retornados.")
    return linhas_total + linhas, ignoradas_total + ignoradas, retornados_total + retornados

# ─── PREPARAR WORKBOOK DE SAÍDA ───────────────────────────────────────
wb_out = openpyxl.Workbook()
ws_out = wb_out.active
ws_out.title = "Base"
ws_out.append(HEADERS)

linhas_total    = 0
ignoradas_total = 0
retornados_total = 0
nao_encontrados = set()

# ─── PROCESSAR OS TRÊS RELATÓRIOS ────────────────────────────────────
print("\nProcessando relatórios...")

# Todos os arquivos de chamados + GC + CTR
fontes = [(p.stem, p) for p in sorted(chamados_paths)]
if gc_path:  fontes.append(("GC",  gc_path))
if ctr_path: fontes.append(("CTR", ctr_path))

for label, path in fontes:
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    ws = wb.active
    linhas_total, ignoradas_total, retornados_total = processar_relatorio(
        ws, label, linhas_total, ignoradas_total, retornados_total, nao_encontrados, ws_out
    )
    wb.close()

# ─── FORMATAÇÃO VISUAL ────────────────────────────────────────────────
COL_WIDTHS = {
    'A': 10,    'B': 11.86, 'C': 10.43, 'D': 26.29,
    'E': 13.0,  'F': 11.57, 'G': 9.86,  'H': 14.71,
    'I': 28.14, 'J': 15.86, 'K': 19.29, 'L': 12.86,
    'M': 12.86, 'N': 14.43, 'O': 15.57, 'P': 28.57,
    'Q': 20.29, 'R': 12.0,
}
for col_letter, width in COL_WIDTHS.items():
    ws_out.column_dimensions[col_letter].width = width

font_padrao = Font(name='Aptos Narrow', size=11)
for row in ws_out.iter_rows():
    for cell in row:
        cell.font = font_padrao
        cell.alignment = Alignment(vertical='center', wrap_text=False)

DATE_COLS = ['J', 'K', 'N', 'O']
for col_letter in DATE_COLS:
    for cell in ws_out[col_letter][1:]:
        if cell.value:
            cell.number_format = 'DD/MM/YYYY'

last_row = ws_out.max_row
last_col = get_column_letter(ws_out.max_column)
table_ref = f"A1:{last_col}{last_row}"

table = Table(displayName="Tabela2", ref=table_ref)
table.tableStyleInfo = TableStyleInfo(
    name="TableStyleMedium2",
    showFirstColumn=False,
    showLastColumn=False,
    showRowStripes=True,
    showColumnStripes=False
)
ws_out.add_table(table)

# ─── SALVAR ───────────────────────────────────────────────────────────
wb_out.save(OUTPUT)
print(f"\n✅ base.xlsx atualizado com {linhas_total} registros → {OUTPUT}")
print(f"🗑  {ignoradas_total} registro(s) ignorado(s) (outros departamentos).")
print(f"🔄  {retornados_total} registro(s) marcado(s) como Retornado (GC/GC-ADM).")

if nao_encontrados:
    print(f"\n⚠  Usuários não encontrados no relatório de colaboradores ({len(nao_encontrados)}):")
    for u in sorted(nao_encontrados):
        print(f"   - {u}")
else:
    print("✅ Todos os usuários foram resolvidos com sucesso.")