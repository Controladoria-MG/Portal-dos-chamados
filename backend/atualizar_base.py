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
  │       ├── Chamados - Geral - *.xlsx
  │       ├── Chamados - GC - *.xlsx
  │       ├── Chamados - Controladoria - *.xlsx
  │       └── Relatório Colaboradores Ativos_*.xlsx
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
BASE_DIR = Path(__file__).parent
DATA_DIR = BASE_DIR / "data" if (BASE_DIR / "data").exists() else BASE_DIR.parent / "data"
ATT_DIR  = DATA_DIR / "Att Base"
OUTPUT   = DATA_DIR / "base.xlsx"

# ─── DEPARTAMENTOS QUE SALVAM RETORNADOS ─────────────────────────────
DEPTS_RETORNADOS = {"gerencia de contas", "gc - administrativo"}

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
]

LOOKUP_COLS = {8, 16}  # I e Q do relatório de origem

HEADERS = [
    "Id", "Categoria", "Assunto", "Departamento Solicitante",
    "Solicitação", "IdCliente", "Cliente", "Responsável",
    "Departamento Responsavel", "Data Cadastro", "Prazo Vencimento",
    "Status", "Solicitante", "Inicio Atend.", "Data Entrega",
    "Responsável pela conclusão", "Ultimo Comentário",
    "Retornado",
]

COL_WIDTHS = {
    'A': 10,    'B': 11.86, 'C': 10.43, 'D': 26.29,
    'E': 13.0,  'F': 11.57, 'G': 9.86,  'H': 14.71,
    'I': 28.14, 'J': 15.86, 'K': 19.29, 'L': 12.86,
    'M': 12.86, 'N': 14.43, 'O': 15.57, 'P': 28.57,
    'Q': 20.29, 'R': 12.0,
}

STATUS_ENTREGUE = "entregue ao solicitante"


# ─── HELPERS ─────────────────────────────────────────────────────────
def find_file(folder, pattern, required=True):
    matches = list(folder.glob(pattern))
    if not matches:
        if required:
            raise FileNotFoundError(f"Nenhum arquivo encontrado com padrão '{pattern}' em {folder}")
        return None
    return max(matches, key=os.path.getmtime)

def find_all(folder, pattern):
    return list(folder.glob(pattern))


def processar_relatorio(ws, label, linhas_total, ignoradas_total,
                        retornados_total, nao_encontrados, ws_out, colab_map):
    linhas     = 0
    ignoradas  = 0
    retornados = 0

    def resolver_nome(valor):
        if not valor:
            return valor
        chave = str(valor).strip().lower()
        return colab_map.get(chave, str(valor).strip())

    for row in ws.iter_rows(min_row=2, values_only=True):
        if not any(row):
            continue

        status_val = str(row[15]).strip().lower() if row[15] else ''
        dept_sol   = str(row[3]).strip().lower()  if row[3]  else ''

        eh_entregue = (status_val == STATUS_ENTREGUE)
        eh_gc       = (dept_sol in DEPTS_RETORNADOS)

        if eh_entregue and not eh_gc:
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

        nova_linha[17] = "SIM" if (eh_entregue and eh_gc) else "NÃO"
        if eh_entregue and eh_gc:
            retornados += 1

        ws_out.append(nova_linha)
        linhas += 1

    print(f"  {label}: {linhas} incluídos  |  {ignoradas} ignorados  |  retornados: {retornados}")
    return linhas_total + linhas, ignoradas_total + ignoradas, retornados_total + retornados


# ─── MAIN ────────────────────────────────────────────────────────────
def main():
    print("Localizando arquivos...")
    colab_path = find_file(ATT_DIR, "*olaboradores*.xlsx")

    # Arquivos específicos: GC e Controladoria (pega o mais recente de cada)
    gc_path  = find_file(ATT_DIR, "*Chamados - GC*.xlsx",             required=False)
    ctr_path = find_file(ATT_DIR, "*Chamados - Controladoria*.xlsx",  required=False)

    gc_name  = gc_path.name  if gc_path  else None
    ctr_name = ctr_path.name if ctr_path else None

    # Chamados gerais: todos com "Chamados" no nome, exceto GC e Controladoria
    chamados_paths = [
        p for p in find_all(ATT_DIR, "*Chamados*.xlsx")
        if p.name != gc_name and p.name != ctr_name
    ]

    if not chamados_paths:
        raise FileNotFoundError(f"Nenhum arquivo *Chamados*.xlsx (Geral) encontrado em {ATT_DIR}")

    print(f"  Colaboradores: {colab_path.name}")
    for p in chamados_paths:
        print(f"  Chamados:      {p.name}")
    print(f"  GC:            {gc_name  or 'nao encontrado'}")
    print(f"  CTR:           {ctr_name or 'nao encontrado'}")

    # Carregar mapa de colaboradores
    print("\nCarregando colaboradores...")
    wb_colab  = openpyxl.load_workbook(colab_path, read_only=True, data_only=True)
    ws_colab  = wb_colab.active
    colab_map = {}
    for row in ws_colab.iter_rows(min_row=2, values_only=True):
        nome_exib = row[1]
        usuario   = row[2]
        if usuario and nome_exib:
            colab_map[str(usuario).strip().lower()] = str(nome_exib).strip()
    wb_colab.close()
    print(f"  {len(colab_map)} colaboradores carregados.")

    # Preparar workbook de saída
    wb_out = openpyxl.Workbook()
    ws_out = wb_out.active
    ws_out.title = "Base"
    ws_out.append(HEADERS)

    linhas_total     = 0
    ignoradas_total  = 0
    retornados_total = 0
    nao_encontrados  = set()

    print("\nProcessando relatórios...")
    fontes = [(p.stem, p) for p in sorted(chamados_paths)]
    if gc_path:  fontes.append(("GC",           gc_path))
    if ctr_path: fontes.append(("Controladoria", ctr_path))

    for label, path in fontes:
        wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
        ws = wb.active
        linhas_total, ignoradas_total, retornados_total = processar_relatorio(
            ws, label, linhas_total, ignoradas_total, retornados_total,
            nao_encontrados, ws_out, colab_map
        )
        wb.close()

    # Formatação
    for col_letter, width in COL_WIDTHS.items():
        ws_out.column_dimensions[col_letter].width = width

    font_padrao = Font(name='Aptos Narrow', size=11)
    for row in ws_out.iter_rows():
        for cell in row:
            cell.font      = font_padrao
            cell.alignment = Alignment(vertical='center', wrap_text=False)

    for col_letter in ['J', 'K', 'N', 'O']:
        for cell in ws_out[col_letter][1:]:
            if cell.value:
                cell.number_format = 'DD/MM/YYYY'

    last_row  = ws_out.max_row
    last_col  = get_column_letter(ws_out.max_column)
    table_ref = f"A1:{last_col}{last_row}"
    table     = Table(displayName="Tabela2", ref=table_ref)
    table.tableStyleInfo = TableStyleInfo(
        name="TableStyleMedium2",
        showFirstColumn=False, showLastColumn=False,
        showRowStripes=True,  showColumnStripes=False,
    )
    ws_out.add_table(table)

    wb_out.save(OUTPUT)
    print(f"\nbase.xlsx atualizado com {linhas_total} registros -> {OUTPUT}")
    print(f"{ignoradas_total} registro(s) ignorado(s).")
    print(f"{retornados_total} registro(s) marcado(s) como Retornado (GC/GC-ADM).")

    if nao_encontrados:
        print(f"\nUsuarios nao encontrados no relatorio de colaboradores ({len(nao_encontrados)}):")
        for u in sorted(nao_encontrados):
            print(f"   - {u}")
    else:
        print("Todos os usuarios foram resolvidos com sucesso.")


if __name__ == "__main__":
    main()
