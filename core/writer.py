"""
core/writer.py
Construit l'Excel 5 feuilles.
"""
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter
from core.reader import parse_num

# ── Couleurs ──────────────────────────────────────────────────────────────────
C_DARK   = "1F4E79"
C_MED    = "2E75B6"
C_LIGHT  = "D6E4F0"
C_STRIPE = "F2F7FC"
C_WHITE  = "FFFFFF"
C_GOLD   = "FFF2CC"
C_BLACK  = "000000"
C_GRAY   = "888888"
NUM_FMT  = '#,##0.00'

# Mots-clés pour détecter les lignes de totaux/résultats → style gras
BOLD_KW = [
    'total', 'sous-total', 'résultat', 'resultat',
    'marge brute', 'valeur ajoutee', 'valeur ajoutée',
    'excedent brut', 'excédent brut', 'insuffisance brute',
    'capacite d', 'capacité d', 'autofinancement',
    'production de l', 'consommations de l',
]


def _is_bold_row(label: str) -> bool:
    ll = label.lower()
    return any(kw in ll for kw in BOLD_KW)


def _c(ws, row, col, value='', bg=C_WHITE, fg=C_BLACK, bold=False,
       align='left', sz=9, wrap=False, indent=0, num_fmt=None):
    cell = ws.cell(row, col)
    cell.value = value
    cell.font      = Font(name="Calibri", size=sz, bold=bold, color=fg)
    cell.fill      = PatternFill("solid", fgColor=bg)
    cell.alignment = Alignment(horizontal=align, vertical="center",
                               wrap_text=wrap, indent=indent)
    if num_fmt:
        cell.number_format = num_fmt
    return cell


def _title(ws, row, text, n_cols, sz=11):
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=n_cols)
    _c(ws, row, 1, text, bg=C_DARK, fg=C_WHITE, bold=True, align='center', sz=sz)
    ws.row_dimensions[row].height = 22


def _subinfo(ws, row, info, n_cols):
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=n_cols-2)
    _c(ws, row, 1, info.get('societe',''), bg=C_LIGHT, bold=True, sz=10)
    _c(ws, row, n_cols-1, f"IF: {info.get('identifiant_fiscal','')}", bg=C_LIGHT, sz=9, align='center')
    _c(ws, row, n_cols,   info.get('exercice',''), bg=C_LIGHT, sz=9, align='center', wrap=True)
    ws.row_dimensions[row].height = 18


def _headers(ws, row, hdrs, widths):
    for ci, (h, w) in enumerate(zip(hdrs, widths), 1):
        _c(ws, row, ci, h, bg=C_MED, fg=C_WHITE, bold=True,
           align='center', sz=9, wrap=True)
        ws.column_dimensions[get_column_letter(ci)].width = w
    ws.row_dimensions[row].height = 26
    ws.freeze_panes = f'A{row+1}'


def _data_rows(ws, start_row, rows, n_val_cols):
    """Écrit les lignes de données avec style."""
    r = start_row
    for i, (label, vals) in enumerate(rows):
        bold   = _is_bold_row(label)
        bg     = C_LIGHT if bold else (C_STRIPE if i % 2 == 0 else C_WHITE)
        indent = 0 if bold else 1

        _c(ws, r, 1, label, bg=bg, fg=C_BLACK, bold=bold, sz=9, indent=indent)

        for ci in range(n_val_cols):
            raw = vals[ci] if ci < len(vals) else None
            num = parse_num(raw) if raw is not None else None
            val = num if num is not None else (raw if raw else '')
            is_number = (num is not None)
            _c(ws, r, ci+2, val, bg=bg, bold=bold, align='right', sz=9,
               num_fmt=NUM_FMT if is_number else None)

        ws.row_dimensions[r].height = 15 if bold else 13
        r += 1
    return r


# ══════════════════════════════════════════════════════════════════════════════

def write_identification(wb, info):
    ws = wb.create_sheet("1 - Identification")
    ws.sheet_view.showGridLines = False
    ws.column_dimensions['A'].width = 35
    ws.column_dimensions['B'].width = 50

    _title(ws, 2, "IDENTIFICATION FISCALE", 2, sz=12)

    fields = [
        ("Société / Raison sociale",   info.get('societe','')),
        ("Identifiant Fiscal (IF)",    info.get('identifiant_fiscal','')),
        ("Article IS",                 info.get('article_is','')),
        ("ICE",                        info.get('ice','')),
        ("Exercice comptable",         info.get('exercice','')),
        ("Activité principale",        info.get('activite','')),
        ("Forme juridique",            info.get('forme_juridique','')),
    ]

    r = 4
    for i, (lbl, val) in enumerate(fields):
        bg = C_STRIPE if i % 2 == 0 else C_WHITE
        _c(ws, r, 1, lbl, bg=bg, bold=True, sz=10)
        _c(ws, r, 2, val, bg=bg, sz=10)
        ws.row_dimensions[r].height = 20
        r += 1


def write_actif(wb, info, rows):
    ws = wb.create_sheet("2 - Bilan Actif")
    ws.sheet_view.showGridLines = False
    _title(ws, 3, "BILAN ACTIF", 5)
    _subinfo(ws, 4, info, 5)
    _headers(ws, 5,
             ["DÉSIGNATION", "BRUT", "AMORT. & PROV.", "NET EXERCICE N", "NET EXERCICE N-1"],
             [48, 18, 18, 18, 18])
    _data_rows(ws, 6, rows, 4)


def write_passif(wb, info, rows):
    ws = wb.create_sheet("3 - Bilan Passif")
    ws.sheet_view.showGridLines = False
    _title(ws, 3, "BILAN PASSIF", 3)
    _subinfo(ws, 4, info, 3)
    _headers(ws, 5,
             ["DÉSIGNATION", "EXERCICE N", "EXERCICE N-1"],
             [52, 20, 20])
    _data_rows(ws, 6, rows, 2)


def write_cpc(wb, info, rows):
    ws = wb.create_sheet("4 - CPC")
    ws.sheet_view.showGridLines = False
    _title(ws, 3, "COMPTE DE PRODUITS ET CHARGES (Hors Taxes)", 5)
    _subinfo(ws, 4, info, 5)
    _headers(ws, 5,
             ["DÉSIGNATION", "PROPRES À\nL'EXERCICE", "EXERCICES\nPRÉCÉDENTS",
              "TOTAUX\nEXERCICE N", "TOTAUX\nEXERCICE N-1"],
             [48, 18, 18, 18, 18])
    _data_rows(ws, 6, rows, 4)


def write_esg(wb, info, rows):
    ws = wb.create_sheet("5 - ESG")
    ws.sheet_view.showGridLines = False
    _title(ws, 3, "ÉTAT DE SOLDES DE GESTION (E.S.G)", 3)
    _subinfo(ws, 4, info, 3)
    _headers(ws, 5,
             ["DÉSIGNATION", "EXERCICE N", "EXERCICE N-1"],
             [52, 20, 20])
    r = _data_rows(ws, 6, rows, 2)

    # Ligne distributions (saisie manuelle)
    if rows:
        r += 1
        _c(ws, r, 1, "Distributions de bénéfices", bg=C_GOLD, sz=9, indent=1)
        for ci in [2, 3]:
            c = ws.cell(r, ci)
            c.value = 0
            c.fill  = PatternFill("solid", fgColor=C_GOLD)
            c.font  = Font(name="Calibri", size=9, italic=True, color=C_GRAY)
            c.alignment = Alignment(horizontal="right", vertical="center")
            c.number_format = NUM_FMT
        ws.row_dimensions[r].height = 14


# ── Point d'entrée ────────────────────────────────────────────────────────────

def build_excel(data: dict, output_path: str):
    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    write_identification(wb, data['info'])
    write_actif(wb,  data['info'], data['actif'])
    write_passif(wb, data['info'], data['passif'])
    write_cpc(wb,    data['info'], data['cpc'])
    write_esg(wb,    data['info'], data['esg'])

    wb.save(output_path)
    return output_path
