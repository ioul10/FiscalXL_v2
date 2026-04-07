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
</style>
""", unsafe_allow_html=True)

st.title("📊 FiscalXL")
st.caption("Convertisseur PDF Bilan Fiscal Marocain → Excel")
st.divider()

uploaded = st.file_uploader(
    "Importer le PDF",
    type=['pdf'],
    help="PDF complet de la liasse fiscale IS (AMMC ou DGI)"
)

if uploaded:
    st.success(f"✅ {uploaded.name} ({uploaded.size/1024:.0f} Ko)")
    if st.button("🔄 Convertir en Excel"):
        with st.spinner("Conversion en cours..."):
            try:
                with tempfile.NamedTemporaryFile(delete=False, suffix='.pdf') as f:
                    f.write(uploaded.read()); pdf_path = f.name
                with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx') as f:
                    xlsx_path = f.name

                res = convert(pdf_path, xlsx_path)

                esg_info = f" · ESG : **{res['n_esg']} lignes**" if res['has_esg'] else ""
                st.markdown(f"""<div class="ok">
                    ✅ <b>Conversion réussie</b> — Format : <b>{'🔵 AMMC' if res['format']=='AMMC' else '🟢 DGI'}</b><br>
                    <b>{res['societe']}</b> — {res['exercice']}<br><br>
                    Actif : <b>{res['n_actif']}</b> lignes &nbsp;·&nbsp;
                    Passif : <b>{res['n_passif']}</b> lignes &nbsp;·&nbsp;
                    CPC : <b>{res['n_cpc']}</b> lignes{esg_info}
                </div>""", unsafe_allow_html=True)

                with open(xlsx_path, 'rb') as f: data = f.read()
                st.download_button("📥 Télécharger l'Excel",
                    data=data,
                    file_name=os.path.splitext(uploaded.name)[0]+'.xlsx',
                    mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                    use_container_width=True)

                os.unlink(pdf_path); os.unlink(xlsx_path)

            except Exception as e:
                st.markdown(f'<div class="err">❌ <b>Erreur :</b> {e}</div>',
                            unsafe_allow_html=True)
                st.code(traceback.format_exc())

st.divider()
st.caption("5 feuilles : Identification · Actif · Passif · CPC · ESG")
