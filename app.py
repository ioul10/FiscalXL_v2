"""
FiscalXL v2 — Convertisseur PDF Fiscal Marocain → Excel
Interface Streamlit
"""
import streamlit as st
import tempfile, os, traceback
from core.pipeline import convert

st.set_page_config(
    page_title="FiscalXL v2",
    page_icon="📊",
    layout="centered"
)

# ── Style ─────────────────────────────────────────────────────────────────────
st.markdown("""
<style>
    .main { background: #f8fafc; }
    .stButton > button {
        background: #1F4E79; color: white;
        border-radius: 8px; font-weight: bold;
        width: 100%; height: 3rem;
    }
    .result-box {
        background: #e8f5e9; border-left: 4px solid #2e7d32;
        padding: 1rem; border-radius: 8px; margin: 1rem 0;
    }
    .error-box {
        background: #ffebee; border-left: 4px solid #c62828;
        padding: 1rem; border-radius: 8px; margin: 1rem 0;
    }
</style>
""", unsafe_allow_html=True)

# ── Interface ─────────────────────────────────────────────────────────────────
st.title("📊 FiscalXL v2")
st.caption("Convertisseur PDF Bilan Fiscal Marocain → Excel")
st.markdown("---")

st.markdown("### 📁 Importer votre PDF")
st.markdown("Formats supportés : **AMMC** (modèle normal) et **DGI** (état de synthèse conforme)")

uploaded = st.file_uploader(
    "Glissez votre PDF ici ou cliquez pour parcourir",
    type=['pdf'],
    help="PDF complet de la liasse fiscale IS (modèle normal MCN)"
)

if uploaded:
    st.success(f"✅ Fichier reçu : **{uploaded.name}** ({uploaded.size/1024:.0f} Ko)")

    col1, col2 = st.columns([2, 1])
    with col1:
        btn = st.button("🔄 Convertir en Excel", use_container_width=True)

    if btn:
        with st.spinner("Conversion en cours..."):
            try:
                # Sauvegarder le PDF temporairement
                with tempfile.NamedTemporaryFile(delete=False, suffix='.pdf') as tmp_in:
                    tmp_in.write(uploaded.read())
                    pdf_path = tmp_in.name

                # Fichier de sortie
                base_name = os.path.splitext(uploaded.name)[0]
                with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx') as tmp_out:
                    xlsx_path = tmp_out.name

                # Conversion
                result = convert(pdf_path, xlsx_path)

                # Afficher les résultats
                fmt_label = "🔵 AMMC" if result['format'] == 'AMMC' else "🟢 DGI"
                st.markdown(f"""
                <div class="result-box">
                    <b>✅ Conversion réussie !</b><br>
                    Format détecté : <b>{fmt_label}</b><br>
                    Société : <b>{result['societe']}</b><br>
                    Exercice : <b>{result['exercice']}</b><br>
                    <br>
                    📋 Actif : <b>{result['n_actif']} lignes</b> &nbsp;|&nbsp;
                    Passif : <b>{result['n_passif']} lignes</b> &nbsp;|&nbsp;
                    CPC : <b>{result['n_cpc']} lignes</b>
                    {'&nbsp;|&nbsp; ESG : <b>' + str(result['n_esg']) + ' lignes</b>' if result['has_esg'] else ''}
                </div>
                """, unsafe_allow_html=True)

                # Bouton téléchargement
                with open(xlsx_path, 'rb') as f:
                    xlsx_bytes = f.read()

                st.download_button(
                    label="📥 Télécharger l'Excel",
                    data=xlsx_bytes,
                    file_name=f"{base_name}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )

                # Nettoyage
                os.unlink(pdf_path)
                os.unlink(xlsx_path)

            except Exception as e:
                st.markdown(f"""
                <div class="error-box">
                    <b>❌ Erreur de conversion</b><br>
                    {str(e)}<br>
                    <small>{traceback.format_exc()}</small>
                </div>
                """, unsafe_allow_html=True)

st.markdown("---")
st.markdown("""
<div style="text-align:center; color:#888; font-size:0.8rem;">
FiscalXL v2 — 5 feuilles : Identification · Actif · Passif · CPC · ESG<br>
Supporte les PDFs complets avec ESG (page 8 AMMC / page 10 DGI)
</div>
""", unsafe_allow_html=True)
