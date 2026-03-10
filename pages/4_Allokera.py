"""
Allokera (HIB) – webversion av allokera10.9.py
"""
import io
import os
import shutil
import sys
import tempfile

import streamlit as st

st.set_page_config(page_title="Allokera (HIB)", page_icon="📋", layout="wide")

st.title("📋 Allokera (HIB)")
st.markdown(
    "Ladda upp dina filer nedan och kör HIB-allokeringen. "
    "Resultatet visas direkt på sidan och kan laddas ned som Excel."
)

# ── Förklaring av filer ────────────────────────────────────────────────────────

with st.expander("ℹ️ Vilka filer behövs?"):
    st.markdown("""
| Fil | Typ | Innehåll |
|-----|-----|----------|
| HIB-ordrar | Excel/CSV | Ordrar från HIB-systemet |
| Butiksordrar | Excel/CSV | Matchande butiksordrar |
| Prognos | Excel (XLSX) | Prognos per artikel |
| Dispatch-pallar | Excel/CSV | Valfri – dispatchade pallar |
""")

# ── Filuppladdning ──────────────────────────────────────────────────────────────

st.subheader("1. Ladda upp filer")

col1, col2 = st.columns(2)
with col1:
    hib_file = st.file_uploader("HIB-ordrar (Excel/CSV):", type=["xlsx", "xls", "csv"])
    store_file = st.file_uploader("Butiksordrar (Excel/CSV):", type=["xlsx", "xls", "csv"])
with col2:
    prognos_file = st.file_uploader("Prognos (Excel XLSX):", type=["xlsx", "xls"])
    dispatch_file = st.file_uploader("Dispatch-pallar (valfri):", type=["xlsx", "xls", "csv"])

st.divider()
st.subheader("2. Kör allokering")

all_required = hib_file and store_file

if st.button("▶ Kör allokering", type="primary", disabled=not all_required):
    with tempfile.TemporaryDirectory() as tmpdir:
        # Spara uppladdade filer till temp-katalog
        file_map = {}
        for label, f in [
            ("hib", hib_file),
            ("store", store_file),
            ("prognos", prognos_file),
            ("dispatch", dispatch_file),
        ]:
            if f is not None:
                dest = os.path.join(tmpdir, f.name)
                with open(dest, "wb") as out:
                    out.write(f.getvalue())
                file_map[label] = dest

        # Lägg till verktygets mapp i Python-sökvägen
        repo_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
        if repo_dir not in sys.path:
            sys.path.insert(0, repo_dir)

        try:
            # Importera allokeringslogiken dynamiskt
            import importlib.util
            allokera_path = os.path.join(repo_dir, "allokera10.9.py")
            spec = importlib.util.spec_from_file_location("allokera", allokera_path)
            allokera_mod = importlib.util.module_from_spec(spec)

            # Kör inte GUI-koden vid import – den skyddas av if __name__ == '__main__'
            spec.loader.exec_module(allokera_mod)

            st.info(
                "Allokeringsmodulen laddades. "
                "Webbversionen av Allokera är under uppbyggnad – "
                "kontakta systemägaren för den fullständiga funktionen."
            )

        except Exception as e:
            st.error(f"Kunde inte ladda allokeringsmodulen: {e}")
            st.info(
                "Tipps: Kör allokera10.9.py lokalt på din dator tills webbversionen är klar."
            )

if not all_required:
    st.info("Ladda upp minst HIB-ordrar och Butiksordrar för att aktivera körning.")

st.divider()
st.caption(
    "💡 Allokera är ett komplext verktyg. Kontakta systemägaren om du behöver hjälp."
)
