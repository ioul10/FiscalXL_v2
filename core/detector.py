"""
core/detector.py
Détecte le format (AMMC/DGI) et assigne chaque page à une section.
Approche : mots-clés dans le texte de chaque page.
"""
import pdfplumber
import re
import unicodedata


def _page_text(page) -> str:
    """Extrait le texte brut d'une page, normalisé en minuscules sans accents."""
    text = page.extract_text() or ''
    text = unicodedata.normalize('NFD', text)
    text = text.encode('ascii', 'ignore').decode().lower()
    return text


# Mots-clés pour identifier chaque section
KEYWORDS = {
    'identification': ['pieces annexes', 'declaration fiscale', 'etat de synthese conforme',
                       'modele 100', 'identification fiscale'],
    'actif':          ['bilan', 'actif', 'immobilisations', 'tresorerie actif'],
    'passif':         ['bilan', 'passif', 'capitaux propres', 'tresorerie passif'],
    'cpc':            ['compte de produits', 'charges hors taxes', 'produits exploitation',
                       'charges exploitation', 'resultat exploitation'],
    'esg':            ['etat des soldes de gestion', 'soldes de gestion', 'e.s.g',
                       'marge brute', 'valeur ajoutee', 'excedent brut',
                       'capacite d autofinancement', 'c.a.f'],
    'financement':    ['tableau de financement', 'emplois', 'ressources stables'],
}


def detect_format(pdf) -> str:
    """Détecte AMMC ou DGI en lisant la page 1."""
    if len(pdf.pages) == 0:
        return 'AMMC'
    text = _page_text(pdf.pages[0])
    if 'etat de synthese conforme' in text or 'souscrite' in text:
        return 'DGI'
    return 'AMMC'


def _score_page(text: str, section: str) -> int:
    """Compte combien de mots-clés d'une section sont présents dans le texte."""
    return sum(1 for kw in KEYWORDS[section] if kw in text)


def assign_pages(pdf) -> dict:
    """
    Assigne chaque page à une section.
    Retourne : {'identification': [0], 'actif': [1], 'passif': [2], 'cpc': [3,4], 'esg': [7]}
    """
    fmt = detect_format(pdf)
    n   = len(pdf.pages)

    # ── Approche 1 : pages fixes selon format ─────────────────────────────
    # C'est la méthode la plus fiable pour les formats connus
    if fmt == 'AMMC':
        pages = {
            'identification': [0],          # page 1
            'actif':          [1],           # page 2
            'passif':         [2],           # page 3
            'cpc':            [3, 4],        # pages 4-5
            'esg':            [7] if n > 7 else [],  # page 8 si existe
        }
    else:  # DGI
        pages = {
            'identification': [0],           # page 1
            'actif':          [1, 2],        # pages 2-3
            'passif':         [3],           # page 4
            'cpc':            [4, 5, 6],     # pages 5-6-7
            'esg':            [9] if n > 9 else [],  # page 10 si existe
        }

    # ── Approche 2 : vérification par mots-clés ───────────────────────────
    # Si une page assignée ne correspond pas à sa section attendue,
    # on cherche la vraie page par mots-clés
    verified = {}
    for section, pg_list in pages.items():
        if not pg_list:
            verified[section] = []
            continue
        # Vérifier que la première page assignée contient bien les bons mots-clés
        pg_idx = pg_list[0]
        if pg_idx < n:
            text = _page_text(pdf.pages[pg_idx])
            score = _score_page(text, section)
            if score > 0 or section == 'identification':
                verified[section] = pg_list
            else:
                # Chercher la vraie page par scan complet
                found = _find_section_pages(pdf, section, n)
                verified[section] = found if found else pg_list
        else:
            verified[section] = []

    # Chercher ESG par scan si pas trouvé par position fixe
    if not verified.get('esg'):
        found = _find_section_pages(pdf, 'esg', n)
        verified['esg'] = found

    return fmt, verified


def _find_section_pages(pdf, section: str, n: int) -> list:
    """Cherche toutes les pages d'une section par mots-clés."""
    result = []
    for i in range(n):
        text = _page_text(pdf.pages[i])
        if _score_page(text, section) >= 2:
            result.append(i)
    return result
