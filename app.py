import streamlit as st
import tempfile, os, traceback
from core.pipeline import convert

st.set_page_config(page_title="FiscalXL", page_icon="📊", layout="centered")

st.markdown("""
<style>
.stButton>button{background:#1F4E79;color:white;border-radius:8px;
    font-weight:bold;width:100%;height:3rem;}
.ok{background:#e8f5e9;border-left:4px solid #2e7d32;
    padding:1rem;border-radius:8px;margin:1rem 0;}
.err{background:#ffebee;border-left:4px solid #c62828;
    padding:1rem;border-radius:8px;margin:1rem 0;}
.info-box{background:#e3f2fd;border-left:4px solid #1565c0;
    padding:0.75rem;border-radius:6px;margin:0.5rem 0;font-size:0.9rem;}
</style>
""", unsafe_allow_html=True)

st.title("📊 FiscalXL")
st.caption("Convertisseur PDF Bilan Fiscal Marocain → Excel")
st.divider()

# Upload
uploaded = st.file_uploader(
    "📁 Importer le PDF fiscal",
    type=['pdf'],
    help="PDF complet de la liasse fiscale IS — Format AMMC ou DGI"
)

# Options
col1, col2 = st.columns([2, 1])
with col2:
    special = st.checkbox(
        "🔧 Traitement spécial",
        value=False,
        help="Activer pour les PDFs avec cellules fusionnées multi-postes (ex: SGTM). "
             "Dans ce mode, le programme détecte chaque ligne de désignation "
             "et associe les montants correspondants."
    )
    if special:
        st.markdown("""<div class="info-box">
        Mode actif : les cellules fusionnées sont décomposées ligne par ligne.
        </div>""", unsafe_allow_html=True)

if uploaded:
    st.success(f"✅ **{uploaded.name}** — {uploaded.size/1024:.0f} Ko")

    if st.button("🔄 Convertir en Excel", use_container_width=True):
        with st.spinner("Conversion en cours..."):
            try:
                # Sauvegarder PDF
                with tempfile.NamedTemporaryFile(delete=False, suffix='.pdf') as f:
                    f.write(uploaded.read())
                    pdf_path = f.name

                with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx') as f:
                    xlsx_path = f.name

                # Convertir
                res = convert(pdf_path, xlsx_path, traitement_special=special)

                # Résultat
                fmt_label = "🔵 AMMC" if res['format'] == 'AMMC' else "🟢 DGI"
                mode_label = " · Mode spécial activé 🔧" if special else ""
                esg_info = f" · ESG : **{res['n_esg']}L**" if res['has_esg'] else ""

                st.markdown(f"""<div class="ok">
                    ✅ <b>Conversion réussie</b> — {fmt_label}{mode_label}<br>
                    <b>{res['societe']}</b><br>
                    {res['exercice']}<br><br>
                    Actif : <b>{res['n_actif']}L</b> &nbsp;·&nbsp;
                    Passif : <b>{res['n_passif']}L</b> &nbsp;·&nbsp;
                    CPC : <b>{res['n_cpc']}L</b>{esg_info}
                </div>""", unsafe_allow_html=True)

                # Téléchargement
                with open(xlsx_path, 'rb') as f:
                    data = f.read()

                base = os.path.splitext(uploaded.name)[0]
                st.download_button(
                    label="📥 Télécharger l'Excel",
                    data=data,
                    file_name=f"{base}.xlsx",
                    mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                    use_container_width=True
                )

                os.unlink(pdf_path)
                os.unlink(xlsx_path)

            except Exception as e:
                st.markdown(
                    f'<div class="err">❌ <b>Erreur :</b> {e}</div>',
                    unsafe_allow_html=True
                )
                with st.expander("Détails de l'erreur"):
                    st.code(traceback.format_exc())

st.divider()
st.markdown("""
<div style="text-align:center;color:#888;font-size:0.85rem;">
📋 5 feuilles générées : <b>Identification · Actif · Passif · CPC · ESG</b><br>
Formats supportés : AMMC (modèle normal) · DGI (état de synthèse conforme)<br>
Pour les PDFs avec cellules fusionnées multi-postes → cocher <b>Traitement spécial</b>
</div>
""", unsafe_allow_html=True)
