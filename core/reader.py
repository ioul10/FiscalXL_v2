"""
core/reader.py
Lit les tableaux PDF et les nettoie.
Concept : lire → nettoyer → retourner.

Corrections appliquées :
1. Filtrer lignes parasites en haut de page (numéros tableau, dates)
2. Ignorer colonne de numérotation ESG (1,2,3...) 
3. CPC : inclure tableaux avec >= 1 ligne de données (pas >= 2)
4. SGTM : extraction X/Y pour PDFs avec cellules fusionnées multi-valeurs
5. Croix/Triangles = vide : filtrer les tokens non-numériques dans val_cols
"""
import re
import pdfplumber
from collections import defaultdict


# ── Parse nombre français ─────────────────────────────────────────────────────

def parse_num(s):
    """'1 234 567,89' → 1234567.89 ou None si pas un nombre."""
    if not s: return None
    s = str(s).strip()
    s = re.sub(r'[\xa0\u202f\s]', '', s)
    s = re.sub(r'[^\d,.\-\(\)]', '', s)
    if not s or s in ['-', '.', ',']: return None
    if s.startswith('(') and s.endswith(')'):
        s = '-' + s[1:-1]
    if ',' in s and '.' in s:
        s = s.replace('.', '').replace(',', '.')
    elif ',' in s:
        s = s.replace(',', '.')
    try:
        v = float(s)
        return v if abs(v) < 1e13 else None
    except:
        return None


def is_num(s):
    return parse_num(s) is not None


# ── Détection colonne numérotation (1,2,3...) ────────────────────────────────

def _is_numbering_col(table, col_idx: int) -> bool:
    """
    Vrai si une colonne contient uniquement des petits entiers séquentiels
    → c'est une colonne de numérotation (ESG : 1, 2, 3...) pas des montants.
    """
    vals = []
    for row in table[2:]:
        if not row or col_idx >= len(row): continue
        s = str(row[col_idx] or '').strip()
        if not s: continue
        v = parse_num(s)
        if v is None: return False          # contient du texte → pas numérotation
        if v != int(v) or v < 0: return False  # décimal ou négatif → pas numérotation
        vals.append(int(v))
    if len(vals) < 2: return False
    # Vérifier que c'est une séquence de petits entiers (1-30 max)
    return max(vals) <= 30 and min(vals) >= 1


# ── Mots à ignorer ────────────────────────────────────────────────────────────

HEADER_WORDS = {
    'brut', 'amortissements', 'amort', 'net', 'exercice', 'precedent',
    'désignation', 'designation', 'nature', 'propres', 'operations',
    'totaux', 'concernant', 'tableau', 'bilan', 'compte', 'hors taxes',
    'suite', 'modele', 'normal', 'immobilise', 'circulant',
    'exercice n', 'exercice n-1',
}

# Patterns de lignes parasites en tête de page
PARASITE_PATTERNS = [
    r'^tableau\s+n',           # Tableau n° 1(1/2)
    r'^\d+\s*\(\d+/\d+\)',     # 01 (1/2)
    r'^bilan\s*\(',            # Bilan (Actif)
    r'^compte\s+de\s+produits',# Compte de Produits
    r'^\(1\)',                 # notes de bas de page
    r'^\(2\)',
    r'^1\)\s',
    r'^2\)\s',
    r'^a\s+l.exclusion',
    r'^identifiant\s+fiscal',  # ligne IF en tête
]

def _is_parasite(label: str) -> bool:
    """Vrai si le label est une ligne parasite (numéro tableau, note...)."""
    if not label: return False
    ll = label.lower().strip()
    for pat in PARASITE_PATTERNS:
        if re.match(pat, ll):
            return True
    return False


def is_header(label: str) -> bool:
    if not label: return True
    ll = label.lower().strip()
    if len(ll) <= 1: return True
    if ll in HEADER_WORDS: return True
    return False


# ── Nettoyage ─────────────────────────────────────────────────────────────────

def clean_rows(rows: list) -> list:
    """
    Nettoie une liste de (label, [vals]).
    1. Supprimer lignes sans label ET sans valeurs
    2. Supprimer lignes parasites
    3. Supprimer doublons
    """
    cleaned = []
    seen = {}

    for label, vals in rows:
        label = str(label or '').strip()
        label = re.sub(r'\s+', ' ', label.replace('\n', ' '))
        label = re.sub(r'^[\*\.\s]+', '', label).strip()

        parsed = [parse_num(v) for v in vals]
        has_label  = bool(label and len(label) > 1)
        has_values = any(v is not None for v in parsed)

        # Règle 1 : rien du tout
        if not has_label and not has_values:
            continue

        # Règle 2 : parasite
        if _is_parasite(label):
            continue

        # Règle 3 : en-tête sans valeurs
        if is_header(label) and not has_values:
            continue

        # Règle 4 : doublons
        key = re.sub(r'\W', '', label.lower())[:25]
        if key and key in seen:
            idx = seen[key]
            ex_count  = sum(1 for v in cleaned[idx][1] if v is not None)
            new_count = sum(1 for v in parsed if v is not None)
            if new_count > ex_count:
                cleaned[idx] = (label, parsed)
            continue

        if key:
            seen[key] = len(cleaned)
        cleaned.append((label, parsed))

    return cleaned


# ── Détection intelligente label_col + val_cols ───────────────────────────────

def _detect_cols(table):
    """
    Retourne (label_col, val_cols) pour un tableau.
    - label_col : colonne avec le plus de textes non-numériques
    - val_cols  : colonnes avec >= 2 nombres, différentes du label,
                  et qui ne sont PAS des colonnes de numérotation
    """
    text_counts = defaultdict(int)
    num_counts  = defaultdict(int)
    for row in table[3:]:
        if not row: continue
        for ci, cell in enumerate(row):
            s = str(cell or '').strip()
            if not s: continue
            if parse_num(s) is not None:
                num_counts[ci] += 1
            elif len(s) > 2:
                text_counts[ci] += 1

    if not num_counts:
        return 0, []

    label_col = max(text_counts, key=text_counts.get) if text_counts else 0

    val_cols = sorted([
        ci for ci, cnt in num_counts.items()
        if cnt >= 2
        and ci != label_col
        and not _is_numbering_col(table, ci)   # FIX 2 : ignorer numérotation
    ])

    return label_col, val_cols


# ── Extraction X/Y pour PDFs fusionnés (SGTM) ────────────────────────────────

def _xy_extract(page) -> list:
    """
    FIX 4 : extraction par position X/Y pour PDFs avec cellules fusionnées (SGTM).
    Regroupe les mots par ligne Y, reconstruit les nombres fractionnés.
    """
    words = page.extract_words(x_tolerance=3, y_tolerance=3)
    if not words: return []

    page_width = page.width
    val_x_start = page_width * 0.38  # valeurs à droite de 38% de la largeur

    # Grouper par Y
    lines = defaultdict(list)
    for w in words:
        y = round(w['top'] / 4) * 4
        lines[y].append(w)
    for y in lines:
        lines[y].sort(key=lambda w: w['x0'])

    rows = []
    for y in sorted(lines.keys()):
        ws = lines[y]

        # Séparer label et zone valeurs
        label_words = []
        val_zone    = []
        for w in ws:
            if w['x0'] < val_x_start:
                # Zone label : ignorer si juste une lettre isolée
                if len(w['text']) > 1 or w['text'].lower() in 'abcdefghijklmnopqrstuvwxyz':
                    if not (len(w['text']) == 1 and w['text'].isupper()):
                        label_words.append(w['text'])
            else:
                val_zone.append(w)

        label = ' '.join(label_words).strip()
        label = re.sub(r'\s+', ' ', label).strip()

        # Reconstruire les nombres fractionnés dans la zone valeurs
        # Ex: ["2", "648", "773", "531,31"] → "2648773531,31" → un seul nombre
        vals = _reconstruct_numbers(val_zone)

        if label or any(v for v in vals if v):
            rows.append((label, vals))

    return rows


def _reconstruct_numbers(val_words: list) -> list:
    """
    Reconstruit les nombres fractionnés par les espaces.
    Groupe les tokens numériques adjacents (gap X < 30pt) en un seul nombre.
    Retourne une liste de valeurs reconstituées.
    """
    if not val_words: return []

    # Séparer en groupes par gap X
    groups = []
    current = [val_words[0]]
    for i in range(1, len(val_words)):
        prev = val_words[i-1]
        curr = val_words[i]
        gap  = curr['x0'] - (prev['x0'] + prev.get('width', len(prev['text'])*5))
        # Même groupe si gap < 25pt ET les deux sont "numériques"
        prev_num = re.match(r'^[\d\s]+$', prev['text'].replace(' ',''))
        curr_num = re.match(r'^[\d,\.]+$', curr['text'].replace(' ',''))
        if gap < 30 and prev_num and curr_num:
            current.append(curr)
        else:
            groups.append(current)
            current = [curr]
    groups.append(current)

    vals = []
    for grp in groups:
        text = ''.join(w['text'] for w in grp)
        # FIX 5 : croix/triangles = vide
        if re.match(r'^[x\+\*\/\|]{1,3}$', text.lower()):
            vals.append('')
            continue
        # Essayer de parser comme nombre FR
        v = parse_num(text)
        if v is not None:
            vals.append(text)  # garder raw pour parse_num dans clean_rows
        else:
            # Pas un nombre → texte dans la zone valeurs (ex: header)
            vals.append('')

    return vals


# ── Détection PDF fusionné ────────────────────────────────────────────────────

def _is_fused(table) -> bool:
    """Vrai si le tableau a beaucoup de cellules avec \n contenant des nombres."""
    count = 0
    for row in table:
        if not row: continue
        for cell in row:
            if cell and '\n' in str(cell):
                parts = str(cell).split('\n')
                if sum(1 for p in parts if parse_num(p.strip())) >= 2:
                    count += 1
    return count >= 3


# ── Lecture d'une page ────────────────────────────────────────────────────────

def read_page_tables(page) -> list:
    """
    Lit tous les tableaux d'une page.
    - Utilise X/Y si cellules fusionnées (SGTM)
    - Sinon lecture standard avec détection intelligente des colonnes
    """
    rows = []
    tables = page.extract_tables()

    # Filtrer les tableaux valides
    # FIX 3 : accepter tableaux avec >= 1 ligne de données (pas >= 2)
    good = [t for t in tables
            if t and len(t) >= 2
            and t[0] and len(t[0]) >= 2
            and sum(1 for r in t[1:] if any(c for c in r if c)) >= 1]

    if not good:
        return rows

    # FIX 4 : si tableau fusionné → X/Y
    if any(_is_fused(t) for t in good):
        return _xy_extract(page)

    # Lecture standard
    for table in good:
        label_col, val_cols = _detect_cols(table)
        if not val_cols:
            continue

        for row in table:
            if not row: continue
            cells = []
            for cell in row:
                s = str(cell or '').strip().replace('\n', ' ')
                s = re.sub(r'\s+', ' ', s).strip()
                # FIX 5 : croix/triangle = vide
                if re.match(r'^[x\+\*\/\\|]{1,3}$', s.lower()) and not is_num(s):
                    s = ''
                cells.append(s)

            label = cells[label_col] if label_col < len(cells) else ''
            vals  = [cells[ci] if ci < len(cells) else '' for ci in val_cols]
            rows.append((label, vals))

    return rows


# ── Lecture multi-pages ───────────────────────────────────────────────────────

def read_section(pdf, page_indices: list) -> list:
    """
    Lit une section sur plusieurs pages et retourne les lignes nettoyées.
    """
    all_rows = []
    for idx in page_indices:
        if idx >= len(pdf.pages): continue
        all_rows.extend(read_page_tables(pdf.pages[idx]))
    return clean_rows(all_rows)
