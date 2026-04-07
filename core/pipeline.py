"""
core/pipeline.py
Point d'entrée principal : PDF → Excel
"""
import pdfplumber
from core.detector  import detect_format, assign_pages
from core.extractor import (extract_identification, extract_actif,
                             extract_table_clean, extract_esg)
from core.writer    import build_excel


def convert(pdf_path: str, output_path: str) -> dict:
    """
    Convertit un PDF fiscal marocain en Excel 5 feuilles.
    Retourne un dict avec les stats de conversion.
    """
    pdf = pdfplumber.open(pdf_path)
    fmt, pages = assign_pages(pdf)

    # ── Extraction ─────────────────────────────────────────────────────────
    info   = extract_identification(pdf, pages['identification'], pages.get('actif', []) + pages.get('passif', []))
    actif  = extract_actif(pdf,          pages['actif'])
    passif = extract_table_clean(pdf,    pages['passif'],  n_val_cols=2)
    cpc    = extract_table_clean(pdf,    pages['cpc'],     n_val_cols=4)
    esg    = extract_esg(pdf,            pages['esg']) if pages.get('esg') else []

    pdf.close()

    # ── Construction Excel ─────────────────────────────────────────────────
    data = {
        'info':   info,
        'actif':  actif,
        'passif': passif,
        'cpc':    cpc,
        'esg':    esg,
    }
    build_excel(data, output_path)

    return {
        'format':    fmt,
        'societe':   info.get('societe', ''),
        'exercice':  info.get('exercice', ''),
        'n_actif':   len(actif),
        'n_passif':  len(passif),
        'n_cpc':     len(cpc),
        'n_esg':     len(esg),
        'has_esg':   len(esg) > 0,
        'pages':     pages,
    }
