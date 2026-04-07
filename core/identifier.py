"""
core/identifier.py
Détecte le format (AMMC/DGI) et extrait les infos d'identification.
"""
import re
import pdfplumber


def detect_format(pdf) -> str:
    """AMMC ou DGI selon la page 1."""
    if not pdf.pages: return 'AMMC'
    text = (pdf.pages[0].extract_text() or '').lower()
    if 'etat de synthese conforme' in text or 'souscrite' in text:
        return 'DGI'
    return 'AMMC'


def get_page_indices(fmt: str, n_pages: int) -> dict:
    """
    Retourne les indices de pages pour chaque section.
    Pages fixes selon le format.
    """
    if fmt == 'DGI':
        return {
            'identification': [0],
            'actif':          [1, 2],
            'passif':         [3],
            'cpc':            [4, 5, 6],
            'esg':            [9] if n_pages > 9 else [],
        }
    else:  # AMMC
        return {
            'identification': [0],
            'actif':          [1],
            'passif':         [2],
            'cpc':            [3, 4],
            'esg':            [7] if n_pages > 7 else [],
        }


def extract_info(pdf, fmt: str) -> dict:
    """
    Extrait les informations d'identification depuis toutes les pages disponibles.
    Cherche dans page 1 + pages actif/passif comme source secondaire.
    """
    info = {
        'societe':            '',
        'identifiant_fiscal': '',
        'article_is':         '',
        'exercice':           '',
        'ice':                '',
        'activite':           '',
        'forme_juridique':    '',
    }

    # Pages à scanner : identification + actif page 1 + passif page 1
    pages_idx = get_page_indices(fmt, len(pdf.pages))
    scan = list(pages_idx['identification'])
    if pages_idx['actif']:  scan.append(pages_idx['actif'][0])
    if pages_idx['passif']: scan.append(pages_idx['passif'][0])

    # Dédoublonner
    seen = set()
    scan = [i for i in scan if not (i in seen or seen.add(i))]

    for idx in scan:
        if idx >= len(pdf.pages): continue
        page = pdf.pages[idx]
        text = page.extract_text() or ''

        # ── Tableaux ──────────────────────────────────────────────────────
        for t in page.extract_tables():
            if not t: continue
            for row in t:
                if not row: continue
                # Chercher label:valeur dans les colonnes
                for li in range(min(3, len(row))):
                    lbl = str(row[li] or '').lower().strip()
                    # Valeur = première cellule non-vide après le label
                    val = next(
                        (str(row[vi]).strip() for vi in range(li+1, len(row))
                         if row[vi] and str(row[vi]).strip() not in ['', ':']),
                        ''
                    )
                    if not val: continue

                    if 'raison sociale' in lbl and not info['societe']:
                        info['societe'] = val.title()
                    elif 'identifiant fiscal' in lbl and not info['identifiant_fiscal']:
                        nums = re.findall(r'\d{4,}', val)
                        if nums: info['identifiant_fiscal'] = nums[0]
                    elif 'article' in lbl and not info['article_is']:
                        nums = re.findall(r'\d{4,}', val)
                        if nums: info['article_is'] = nums[0]

        # ── Texte brut ────────────────────────────────────────────────────
        for line in text.split('\n'):
            ll = line.lower().strip()
            if not ll: continue

            # Société
            if 'raison sociale' in ll and not info['societe']:
                m = re.search(r'raison\s+sociale\s*:?\s*(.+)', line, re.I)
                if m and len(m.group(1).strip()) > 2:
                    info['societe'] = m.group(1).strip().title()

            # Identifiant fiscal
            if 'identifiant fiscal' in ll and not info['identifiant_fiscal']:
                nums = re.findall(r'\d{4,}', line)
                if nums: info['identifiant_fiscal'] = nums[0]

            # Article IS
            if 'article' in ll and 'i' in ll and not info['article_is']:
                m = re.search(r'article\s+i\.?s\.?\s*:?\s*(\d+)', line, re.I)
                if m: info['article_is'] = m.group(1)

            # ICE
            if 'ice' in ll and not info['ice']:
                m = re.search(r'\bice\s*:?\s*(\d{10,})', line, re.I)
                if m: info['ice'] = m.group(1)

            # Exercice — "période du X au Y"
            if not info['exercice']:
                m = re.search(
                    r'p[ée]riode\s+du\s+(\d{2}/\d{2}/\d{4})\s+au\s*\n?\s*(\d{2}/\d{2}/\d{4})',
                    text, re.I)
                if m:
                    d1, d2 = m.group(1), m.group(2)
                    if int(d1[-4:]) > int(d2[-4:]): d1, d2 = d2, d1
                    info['exercice'] = f"Du {d1} au {d2}"
                else:
                    dates = re.findall(r'\d{2}/\d{2}/\d{4}', line)
                    if len(dates) >= 2:
                        d1, d2 = dates[0], dates[1]
                        if int(d1[-4:]) > int(d2[-4:]): d1, d2 = d2, d1
                        info['exercice'] = f"Du {d1} au {d2}"
                    elif 'clos le' in ll and dates:
                        info['exercice'] = f"Clos le {dates[0]}"

            # Activité
            if 'activit' in ll and not info['activite']:
                m = re.search(r'activit[eé]\s*:?\s*([^\n]+)', line, re.I)
                if m:
                    v = m.group(1).strip()
                    if v and len(v) > 2 and not v[0].isdigit():
                        info['activite'] = v[:80]

        # Société depuis première ligne (BORJ/EtatsFiscaux style)
        if not info['societe']:
            skip = ['tableau', 'bilan', 'pieces annexes', 'declaration',
                    'etat de synthese', 'modele', 'compte de produits',
                    'identification', 'impots']
            for l in [ln.strip() for ln in text.split('\n') if ln.strip()]:
                if any(kw in l.lower() for kw in skip): continue
                if l.isupper() and len(l.split()) >= 2:
                    info['societe'] = l.title()
                    break

    # Fallback IF → Article IS
    if not info['identifiant_fiscal'] and info['article_is']:
        info['identifiant_fiscal'] = info['article_is']

    return info
