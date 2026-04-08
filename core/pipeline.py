"""
core/pipeline.py — PDF → Excel
Paramètre traitement_special : activer pour PDFs fusionnés type SGTM.
"""
import pdfplumber
from core.identifier import detect_format, get_page_indices, extract_info
from core.reader     import read_section
from core.writer     import build_excel


def convert(pdf_path: str, output_path: str,
            traitement_special: bool = False) -> dict:
    """
    Convertit un PDF fiscal marocain en Excel 5 feuilles.
    traitement_special=True : utilise le mode SGTM (cellules fusionnées).
    """
    pdf = pdfplumber.open(pdf_path)
    fmt    = detect_format(pdf)
    n      = len(pdf.pages)
    pages  = get_page_indices(fmt, n)
    info   = extract_info(pdf, fmt)

    if traitement_special:
        from core.special import read_section_special as read_fn
    else:
        read_fn = read_section

    actif  = read_fn(pdf, pages['actif'])
    passif = read_fn(pdf, pages['passif'])
    cpc    = read_fn(pdf, pages['cpc'])
    esg    = read_fn(pdf, pages['esg']) if pages.get('esg') else []

    pdf.close()

    build_excel(
        {'info': info, 'actif': actif, 'passif': passif, 'cpc': cpc, 'esg': esg},
        output_path
    )

    return {
        'format':            fmt,
        'societe':           info.get('societe', ''),
        'exercice':          info.get('exercice', ''),
        'n_actif':           len(actif),
        'n_passif':          len(passif),
        'n_cpc':             len(cpc),
        'n_esg':             len(esg),
        'has_esg':           len(esg) > 0,
        'traitement_special': traitement_special,
    }
