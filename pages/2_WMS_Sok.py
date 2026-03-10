"""
WMS-sök – webversion av wms_sök79.py
Ladda upp CSV-loggfiler och sök på inköpsnr + artikelnr.
"""
import os
import shutil
import tempfile

import streamlit as st

from tools.wms_logic import WMSAnalyzer

st.set_page_config(page_title="WMS-sök", page_icon="🔍", layout="wide")

st.title("🔍 WMS-sök")
st.markdown(
    "Ladda upp dina WMS-loggfiler (CSV) och sök sedan på "
    "inköpsnummer och artikelnummer för att se pall-status, saldo och leveranser."
)

# ── Filuppladdning ──────────────────────────────────────────────────────────────

FILE_SLOTS = {
    "receive":  ("Mottagningslogg",          "v_ask_receive_log"),
    "booking":  ("Ej inlagrade (putaway)",   "v_ask_booking_putaway"),
    "buffert":  ("Buffertpall",              "v_ask_article_buffertpallet"),
    "trans":    ("Translogg",                "v_ask_trans_log"),
    "pick":     ("Plocklogg",                "v_ask_pick_log_full"),
    "correct":  ("Saldojustering (valfri)",  "v_ask_correct_log"),
}

st.subheader("1. Ladda upp loggfiler")
st.caption("Minst mottagningsloggen krävs. Övriga filer är valfria men ger mer information.")

uploaded: dict = {}
cols = st.columns(3)
for i, (key, (label, _prefix)) in enumerate(FILE_SLOTS.items()):
    with cols[i % 3]:
        f = st.file_uploader(label, type=["csv", "txt"], key=f"file_{key}")
        uploaded[key] = f
        if f:
            st.success(f"✓ {f.name}")

# ── Sökfält ────────────────────────────────────────────────────────────────────

st.divider()
st.subheader("2. Ange sökparametrar")

col_a, col_b = st.columns(2)
with col_a:
    purchase_number = st.text_input("Inköpsnummer:", placeholder="t.ex. 12345")
with col_b:
    article_number = st.text_input("Artikelnummer:", placeholder="t.ex. ABC-001")

# ── Analys ─────────────────────────────────────────────────────────────────────

st.divider()

if st.button("🔎 Analysera", type="primary", disabled=not uploaded.get("receive")):
    if not purchase_number.strip() or not article_number.strip():
        st.warning("Fyll i både inköpsnummer och artikelnummer.")
    else:
        with tempfile.TemporaryDirectory() as tmpdir:
            name_map = {
                "receive": "v_ask_receive_log.csv",
                "booking": "v_ask_booking_putaway.csv",
                "buffert": "v_ask_article_buffertpallet.csv",
                "trans":   "v_ask_trans_log.csv",
                "pick":    "v_ask_pick_log_full.csv",
                "correct": "v_ask_correct_log.csv",
            }
            for key, f in uploaded.items():
                if f is not None:
                    dest = os.path.join(tmpdir, name_map[key])
                    with open(dest, "wb") as out:
                        out.write(f.read())

            try:
                with st.spinner("Analyserar…"):
                    analyzer = WMSAnalyzer(data_path=tmpdir)
                    result = analyzer.analyze(purchase_number.strip(), article_number.strip())

                st.subheader("Resultat")
                st.text_area("Analysrapport:", value=result, height=500)
            except Exception as e:
                st.error(f"Fel vid analys: {e}")

if not uploaded.get("receive"):
    st.info("Ladda upp minst mottagningsloggen (Mottagningslogg) för att aktivera analysen.")
