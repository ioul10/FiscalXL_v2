"""
core/extractor.py
Lit chaque page du PDF et retourne des données propres.
"""
import pdfplumber
import re
import unicodedata
from collections import defaultdict


def _parse_num(s) -> float | None:
    if not s: return None
    s = str(s).strip()
    s = re.sub(r'[\xa0\u202f\s]', '', s)
    s = s.replace("'", '').replace('\u2019', '')
    s = re.sub(r'[^\d,.\-]', '', s)
    if not s or s in ['-', '.', ',']: return None
    if ',' in s and '.' in s:
        s = s.replace('.', '').replace(',', '.')
    elif ',' in s:
        s = s.replace(',', '.')
    try:
        v = float(s)
        return v if abs(v) < 1e13 else None
    except:
        return None


def _clean_label(s) -> str:
    if not s: return ''
    s = str(s).strip()
    s = re.sub(r'^[\*\.\s]+', '', s)
    s = s.replace('\n', ' ')
    s = re.sub(r'\s+', ' ', s).strip()
    return s


def _detect_val_cols(table) -> list:
    if not table: return []
    n_cols = max(len(r) for r in table)
    counts = defaultdict(int)
    for row in table[2:]:
        for ci, cell in enumerate(row):
            if _parse_num(cell) is not None:
                counts[ci+1] += 1
    return [ci for ci, cnt in sorted(counts.items()) if cnt >= 2]


# ── Identification ────────────────────────────────────────────────────────────

def extract_identification(pdf, page_indices: list, extra_indices: list = None) -> dict:
    """
    Extrait les informations d'identification.
    Cherche dans pages_identification + extra_indices (actif/passif).
    """
    info = {
        'societe': '', 'identifiant_fiscal': '', 'article_is': '',
        'exercice': '', 'forme_juridique': '', 'activite': '',
        'siege': '', 'ice': '',
    }

    all_indices = list(page_indices) + (extra_indices or [])
    # Dédoublonner en gardant l'ordre
    seen_idx = set()
    search_indices = []
    for i in all_indices:
        if i not in seen_idx:
            seen_idx.add(i)
            search_indices.append(i)

    for idx in search_indices:
        if idx >= len(pdf.pages): continue
        page   = pdf.pages[idx]
        text   = page.extract_text() or ''
        lines  = text.split('\n')

        # ── Tableaux (format AMMC page 1) ────────────────────────────────
        for t in page.extract_tables():
            if not t: continue
            for row in t:
                if not row or len(row) < 2: continue
                for li in [0, 1]:
                    lbl = str(row[li] or '').lower().strip()
                    val = next((str(row[vi]).strip() for vi in range(li+1, min(len(row),6))
                                if row[vi] and str(row[vi]).strip() not in ['', ':']), '')
                    if not val: continue
                    if 'raison sociale' in lbl and not info['societe']:
                        info['societe'] = val.title()
                    elif 'identifiant fiscal' in lbl and not info['identifiant_fiscal']:
                        nums = re.findall(r'\d{4,}', val)
                        if nums: info['identifiant_fiscal'] = nums[0]
                    elif 'article' in lbl and 'i' in lbl and not info['article_is']:
                        nums = re.findall(r'\d+', val)
                        if nums: info['article_is'] = nums[0]

        # ── Texte brut ────────────────────────────────────────────────────
        for line in lines:
            ll = line.lower().strip()
            if not ll: continue

            # Société
            if ('raison sociale' in ll) and not info['societe']:
                m = re.search(r'raison\s+sociale\s*:?\s*(.+)', line, re.I)
                if m and len(m.group(1).strip()) > 2:
                    info['societe'] = m.group(1).strip().title()

            # Identifiant fiscal
            if 'identifiant fiscal' in ll and not info['identifiant_fiscal']:
                nums = re.findall(r'\d{4,}', line)
                if nums: info['identifiant_fiscal'] = nums[0]

            # Article IS  (ex: "ARTICLE IS : 20717256")
            if 'article' in ll and not info['article_is']:
                m = re.search(r'article\s+i\.?s\.?\s*:?\s*(\d+)', line, re.I)
                if m: info['article_is'] = m.group(1)

            # ICE
            if 'ice' in ll and ':' in line and not info['ice']:
                m = re.search(r'\bice\s*:?\s*(\d{10,})', line, re.I)
                if m: info['ice'] = m.group(1)

            # Exercice — priorité aux patterns explicites
            if not info['exercice']:
                # "période du X au Y" (DGI)
                m = re.search(
                    r'p[ée]riode\s+du\s+(\d{2}/\d{2}/\d{4})\s+au\s*\n?\s*(\d{2}/\d{2}/\d{4})',
                    text, re.I)
                if m:
                    d1, d2 = m.group(1), m.group(2)
                    y1, y2 = int(d1[-4:]), int(d2[-4:])
                    if y1 > y2: d1, d2 = d2, d1
                    info['exercice'] = f"Du {d1} au {d2}"
                else:
                    # Deux dates sur la même ligne
                    dates = re.findall(r'\d{2}/\d{2}/\d{4}', line)
                    if len(dates) >= 2:
                        d1, d2 = dates[0], dates[1]
                        y1, y2 = int(d1[-4:]), int(d2[-4:])
                        if y1 > y2: d1, d2 = d2, d1
                        info['exercice'] = f"Du {d1} au {d2}"
                    elif 'clos le' in ll and dates:
                        info['exercice'] = f"Clos le {dates[0]}"

            # Activité
            if 'activit' in ll and not info['activite']:
                m = re.search(r'activit[eé][^:]*:?\s*([^\n]+)', line, re.I)
                if m:
                    val = m.group(1).strip()
                    if val and len(val) > 2 and not val[0].isdigit():
                        info['activite'] = val[:80]

        # Société depuis première ligne non-système (BORJ/EtatsFiscaux style)
        if not info['societe']:
            skip_kw = ['tableau', 'bilan', 'pieces annexes', 'declaration',
                       'etat de synthese', 'modele', 'compte de produits', 'cpc']
            for l in [ln.strip() for ln in lines if ln.strip()]:
                ll = l.lower()
                if any(kw in ll for kw in skip_kw): continue
                if l.isupper() and len(l.split()) >= 2 and len(l) > 4:
                    info['societe'] = l.title()
                    break

    # Si IF vide → utiliser Article IS
    if not info['identifiant_fiscal'] and info['article_is']:
        info['identifiant_fiscal'] = info['article_is']

    return info


# ── Extraction tableau générique ──────────────────────────────────────────────

def extract_table_clean(pdf, page_indices: list, n_val_cols: int = 4) -> list:
    all_rows = []
    seen = set()

    SKIP = ['désignation', 'designation', 'nature', 'brut', 'amortissement',
            'net', 'exercice', 'precedent', 'propres', 'operations', 'totaux',
            'tableau', 'bilan', 'compte de produits', 'hors taxes', 'suite']

    for idx in page_indices:
        if idx >= len(pdf.pages): continue
        page = pdf.pages[idx]
        tables = page.extract_tables()
        good = [t for t in tables
                if t and len(t) >= 3 and len(t[0]) >= 3
                and sum(1 for r in t[2:] if any(c for c in r if c)) >= 2]
        if not good: continue

        for tab in good:
            val_cols = _detect_val_cols(tab)
            if not val_cols: continue

            for row in tab:
                if not row: continue
                label = ''
                for cell in row:
                    if cell:
                        cl = _clean_label(cell)
                        if cl and _parse_num(cell) is None and len(cl) > 2:
                            label = cl
                            break
                if not label: continue
                ll = label.lower()
                if any(s == ll or ll.startswith(s) for s in SKIP): continue

                vals = []
                for ci in val_cols[:n_val_cols]:
                    v = _parse_num(row[ci-1]) if ci-1 < len(row) else None
                    vals.append(v)
                while len(vals) < n_val_cols:
                    vals.append(None)
                vals = vals[:n_val_cols]

                key = re.sub(r'\W', '', ll)[:20]
                if key in seen:
                    # Remplacer si nouvelle version a plus de valeurs
                    new_count = sum(1 for v in vals if v is not None and v != 0)
                    for i, (ex_l, ex_v) in enumerate(all_rows):
                        if re.sub(r'\W', '', ex_l.lower())[:20] == key:
                            ex_count = sum(1 for v in ex_v if v is not None and v != 0)
                            if new_count > ex_count:
                                all_rows[i] = (label, vals)
                            break
                    continue

                seen.add(key)
                all_rows.append((label, vals))

    return all_rows


# ── Extraction actif ──────────────────────────────────────────────────────────

def extract_actif(pdf, page_indices: list) -> list:
    all_rows = []
    seen = set()

    SKIP = ['désignation', 'brut', 'amort', 'net', 'exercice', 'tableau',
            'actif immobilise', 'actif circulant']

    for idx in page_indices:
        if idx >= len(pdf.pages): continue
        page = pdf.pages[idx]
        tables = page.extract_tables()
        good = [t for t in tables
                if t and len(t[0]) >= 4
                and sum(1 for r in t[2:] if any(c for c in r if c)) >= 2]
        if not good: continue

        tab = good[0]
        val_cols = _detect_val_cols(tab)
        if not val_cols: continue

        nv = len(val_cols)
        if nv >= 4:
            col_map = [0, 1, 2, 3]
            use_none = [False, False, False, False]
        elif nv == 3:
            gap = val_cols[1] - val_cols[0]
            if gap > 1:
                col_map = [0, -1, 1, 2]   # -1 = None (Amort absent)
                use_none = [False, True, False, False]
            else:
                col_map = [0, 1, 2, -1]
                use_none = [False, False, False, True]
        else:
            col_map = [0, -1, 1, -1]
            use_none = [False, True, False, True]

        for row in tab:
            if not row: continue
            label = ''
            for cell in row:
                if cell:
                    cl = _clean_label(cell)
                    if cl and _parse_num(cell) is None and len(cl) > 2:
                        label = cl
                        break
            if not label: continue
            ll = label.lower()
            if any(s in ll for s in SKIP): continue

            vals = []
            for j, (ci_offset, is_none) in enumerate(zip(col_map, use_none)):
                if is_none:
                    vals.append(None)
                else:
                    ci = val_cols[ci_offset] - 1 if ci_offset >= 0 else -1
                    vals.append(_parse_num(row[ci]) if 0 <= ci < len(row) else None)

            key = re.sub(r'\W', '', ll)[:20]
            if key in seen:
                new_count = sum(1 for v in vals if v is not None and v != 0)
                for i, (ex_l, ex_v) in enumerate(all_rows):
                    if re.sub(r'\W', '', ex_l.lower())[:20] == key:
                        ex_count = sum(1 for v in ex_v if v is not None and v != 0)
                        if new_count > ex_count:
                            all_rows[i] = (label, vals)
                        break
                continue
            seen.add(key)
            all_rows.append((label, vals))

    return all_rows


def extract_esg(pdf, page_indices: list) -> list:
    return extract_table_clean(pdf, page_indices, n_val_cols=2)
