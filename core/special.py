"""
core/special.py
Traitement spécial pour PDFs avec cellules fusionnées (SGTM).

Méthode : grille lignes × colonnes
1. Détecter automatiquement les centres de colonnes numériques
2. Regrouper les tokens numériques fractionnés par (Y, col)
3. Reconstituer chaque nombre et l'associer à son label sur la même ligne Y
"""
import re
from collections import defaultdict
from core.reader import parse_num, clean_rows, _is_parasite


def _detect_col_centers(words, label_x_max: float, tolerance: float = 15.0) -> list:
    """
    Détecte les centres des colonnes de valeurs par clustering des positions X.
    Ignore les mots dans la zone label (x < label_x_max).
    """
    # Collecter les X des tokens numériques dans la zone valeurs
    num_x = [round(w['x0']) for w in words
              if w['x0'] >= label_x_max and parse_num(w['text']) is not None]
    if not num_x:
        return []

    # Clustering simple : trier et regrouper par gap > tolerance
    x_sorted = sorted(set(num_x))
    groups = [[x_sorted[0]]]
    for x in x_sorted[1:]:
        if x - groups[-1][-1] > tolerance:
            groups.append([x])
        else:
            groups[-1].append(x)

    centers = [sum(g) // len(g) for g in groups]
    return centers


def _assign_col(x: float, centers: list, tolerance: float = 45.0) -> int | None:
    """Retourne le centre de colonne le plus proche, ou None si trop loin."""
    if not centers:
        return None
    col = min(centers, key=lambda c: abs(c - x))
    return col if abs(col - x) < tolerance else None


def extract_grid(page) -> list:
    """
    Extraction par grille pour PDFs avec cellules fusionnées.

    Retourne list of (label, [val_col1, val_col2, val_col3, val_col4])
    dans l'ordre détecté des colonnes.
    """
    page_width  = page.width
    label_x_max = page_width * 0.40   # labels dans les 40% gauche

    words = page.extract_words(x_tolerance=3, y_tolerance=3)
    if not words:
        return []

    # Détecter les colonnes
    col_centers = _detect_col_centers(words, label_x_max)
    if not col_centers:
        return []

    # Garder au maximum 4 colonnes de valeurs (Brut, Amort, NetN, NetN1)
    col_centers = col_centers[-4:] if len(col_centers) > 4 else col_centers

    # Grouper les mots par ligne Y (tolérance 4pt)
    lines = defaultdict(list)
    for w in words:
        y = round(w['top'] / 4) * 4
        lines[y].append(w)

    # Construire les lignes label + valeurs
    row_data = []
    for y in sorted(lines.keys()):
        ws = sorted(lines[y], key=lambda w: w['x0'])

        # Séparer tokens label / tokens valeur
        label_tokens = []
        col_tokens   = defaultdict(list)

        for w in ws:
            x = w['x0']
            if x < label_x_max:
                # Zone label : ignorer lettres de rotation isolées
                t = w['text'].strip()
                if len(t) == 1 and t.isupper() and t.isalpha():
                    continue
                label_tokens.append(t)
            else:
                col = _assign_col(x, col_centers)
                if col is not None:
                    col_tokens[col].append(w['text'])

        # Reconstituer le label
        label = ' '.join(label_tokens).strip()
        label = re.sub(r'\s+', ' ', label)

        # Reconstituer les valeurs par colonne
        vals = []
        for col in col_centers:
            raw = ''.join(col_tokens.get(col, []))
            vals.append(parse_num(raw))

        # Compléter à 4 colonnes
        while len(vals) < 4:
            vals.append(None)

        row_data.append((y, label, vals[:4]))

    # Association décalée : si label sans valeurs, chercher Y+delta
    # Construire index {y: vals} pour lookup rapide
    y_to_vals = {y: vals for y, label, vals in row_data
                 if any(v is not None for v in vals)}

    result = []
    for y, label, vals in row_data:
        has_label  = bool(label and len(label) > 1 and parse_num(label) is None)
        has_values = any(v is not None for v in vals)

        if not has_label and not has_values:
            continue
        if _is_parasite(label):
            continue

        # Si label sans valeurs → chercher dans Y+1..Y+8
        if has_label and not has_values:
            for delta in range(1, 10):
                candidate = y + delta
                if candidate in y_to_vals:
                    vals = y_to_vals[candidate]
                    break

        result.append((label, vals))

    return clean_rows(result)


def read_section_special(pdf, page_indices: list) -> list:
    """
    Lit une section en mode traitement spécial (grille).
    Fusionne plusieurs pages et dédoublonne.
    """
    all_rows = []
    for idx in page_indices:
        if idx >= len(pdf.pages):
            continue
        all_rows.extend(extract_grid(pdf.pages[idx]))
    return clean_rows(all_rows)
