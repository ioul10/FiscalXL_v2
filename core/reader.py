"""
core/reader.py
Lit les tableaux PDF et les nettoie.
Concept : lire → nettoyer → retourner.
"""
import re
import pdfplumber
from collections import defaultdict


# ── Parse nombre français ─────────────────────────────────────────────────────

def parse_num(s):
    """'1 234 567,89' → 1234567.89 ou None si pas un nombre."""
    if not s: return None
    s = str(s).strip()
    s = re.sub(r'[\xa0\u202f\s]', '', s)  # espaces insécables
    s = re.sub(r'[^\d,.\-\(\)]', '', s)
    if not s or s in ['-', '.', ',']: return None
    # Parenthèses = négatif : (1234) → -1234
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


def detect_val_cols(table) -> list:
    """
    Détecte les colonnes avec des valeurs numériques (indices 0-based).
    Compte les colonnes avec au moins 2 nombres sur l'ensemble du tableau.
    """
    from collections import defaultdict
    counts = defaultdict(int)
    for row in table[2:]:  # skip header rows
        if not row: continue
        for ci, cell in enumerate(row):
            if parse_num(cell) is not None:
                counts[ci] += 1
    return sorted([ci for ci, cnt in counts.items() if cnt >= 2])


# ── Mots à ignorer (lignes d'en-tête) ────────────────────────────────────────

HEADER_WORDS = {
    'brut', 'amortissements', 'amort', 'net', 'exercice', 'precedent',
    'désignation', 'designation', 'nature', 'propres', 'operations',
    'totaux', 'concernant', 'tableau', 'bilan', 'compte', 'hors taxes',
    'suite', 'modele', 'normal', 'passif', 'actif', 'immobilise',
    'circulant', 'tresorerie', 'provisions', 'exercice n', 'exercice n-1',
}

def is_header(label: str) -> bool:
    """Vrai si le label est une ligne d'en-tête à ignorer."""
    if not label: return True
    ll = label.lower().strip()
    # Très court = probablement un code ou lettre de section
    if len(ll) <= 1: return True
    # Exactement un mot qui est dans la liste d'en-têtes
    if ll in HEADER_WORDS: return True
    # Commence par un de ces mots seuls
    for hw in HEADER_WORDS:
        if ll == hw: return True
    return False


# ── Nettoyage d'une liste de lignes ──────────────────────────────────────────

def clean_rows(rows: list) -> list:
    """
    Nettoie une liste de (label, [vals]).
    Règles :
    1. Supprimer lignes sans label ET sans valeurs
    2. Supprimer lignes d'en-tête
    3. Supprimer doublons — même label → garder une seule fois
    """
    cleaned = []
    seen_labels = {}  # label_key → index dans cleaned

    for label, vals in rows:
        label = str(label or '').strip()
        label = re.sub(r'\s+', ' ', label.replace('\n', ' '))
        label = re.sub(r'^[\*\.\s]+', '', label).strip()

        # Valeurs parsées
        parsed_vals = []
        for v in vals:
            n = parse_num(v)
            parsed_vals.append(n)

        has_label  = bool(label and len(label) > 1)
        has_values = any(v is not None for v in parsed_vals)

        # Règle 1 : rien du tout → supprimer
        if not has_label and not has_values:
            continue

        # Règle 2 : en-tête → supprimer
        if is_header(label) and not has_values:
            continue

        # Règle 3 : doublons → garder celui avec le plus de valeurs non-nulles
        key = re.sub(r'\W', '', label.lower())[:25]
        if key and key in seen_labels:
            idx = seen_labels[key]
            existing_vals = cleaned[idx][1]
            existing_count = sum(1 for v in existing_vals if v is not None)
            new_count = sum(1 for v in parsed_vals if v is not None)
            if new_count > existing_count:
                cleaned[idx] = (label, parsed_vals)
            continue

        if key:
            seen_labels[key] = len(cleaned)
        cleaned.append((label, parsed_vals))

    return cleaned


# ── Lecture d'une page ────────────────────────────────────────────────────────

def read_page_tables(page) -> list:
    """
    Lit tous les tableaux d'une page et retourne list of (label, [raw_vals]).

    Détection intelligente :
    - La colonne du label = celle avec le plus de textes non-numériques
    - Les colonnes de valeurs = celles avec >= 2 nombres, différentes du label
    """
    from collections import defaultdict
    rows = []
    tables = page.extract_tables()

    for table in tables:
        if not table: continue
        if len(table) < 2: continue
        if not table[0] or len(table[0]) < 2: continue

        # ── Détecter label_col et val_cols ────────────────────────────────
        text_counts = defaultdict(int)
        num_counts  = defaultdict(int)
        for row in table[3:]:  # skip 3 premières lignes (headers)
            if not row: continue
            for ci, cell in enumerate(row):
                s = str(cell or '').strip()
                if not s: continue
                if parse_num(s) is not None:
                    num_counts[ci] += 1
                elif len(s) > 2:
                    text_counts[ci] += 1

        if not num_counts: continue

        # Label = colonne avec le plus de textes
        label_col = max(text_counts, key=text_counts.get) if text_counts else 0

        # Valeurs = colonnes avec >= 2 nombres, différentes du label
        val_cols = sorted([ci for ci, cnt in num_counts.items()
                           if cnt >= 2 and ci != label_col])

        if not val_cols: continue

        # ── Extraire les lignes ───────────────────────────────────────────
        for row in table:
            if not row: continue
            cells = []
            for cell in row:
                s = str(cell or '').strip().replace('\n', ' ')
                s = re.sub(r'\s+', ' ', s).strip()
                cells.append(s)

            label = cells[label_col] if label_col < len(cells) else ''
            vals  = [cells[ci] if ci < len(cells) else '' for ci in val_cols]
            rows.append((label, vals))

    return rows


# ── Lecture multi-pages ───────────────────────────────────────────────────────

def read_section(pdf, page_indices: list) -> list:
    """
    Lit une section sur plusieurs pages et retourne les lignes nettoyées.
    Fusionne les pages et dédoublonne.
    """
    all_rows = []
    for idx in page_indices:
        if idx >= len(pdf.pages): continue
        page_rows = read_page_tables(pdf.pages[idx])
        all_rows.extend(page_rows)
    return clean_rows(all_rows)
