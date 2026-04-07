"""
core/pipeline.py — PDF → Excel
"""
import pdfplumber
from core.identifier import detect_format, get_page_indices, extract_info
from core.reader     import read_section
from core.writer     import build_excel


def convert(pdf_path: str, output_path: str) -> dict:
    pdf = pdfplumber.open(pdf_path)
    fmt    = detect_format(pdf)
    n      = len(pdf.pages)
    pages  = get_page_indices(fmt, n)
    info   = extract_info(pdf, fmt)

    actif  = read_section(pdf, pages['actif'])
    passif = read_section(pdf, pages['passif'])
    cpc    = read_section(pdf, pages['cpc'])
    esg    = read_section(pdf, pages['esg']) if pages.get('esg') else []

    pdf.close()

    build_excel(
        {'info': info, 'actif': actif, 'passif': passif, 'cpc': cpc, 'esg': esg},
        output_path
    )

    return {
        'format':   fmt,
        'societe':  info.get('societe',''),
        'exercice': info.get('exercice',''),
        'n_actif':  len(actif),
        'n_passif': len(passif),
        'n_cpc':    len(cpc),
        'n_esg':    len(esg),
        'has_esg':  len(esg) > 0,
    }
