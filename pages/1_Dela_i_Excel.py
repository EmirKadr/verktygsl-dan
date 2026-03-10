"""
Dela i Excel – webversion av 2000tal.py
Klistra in värden (en per rad), välj chunk-storlek och ladda ned Excel.
"""
import io
import streamlit as st
from openpyxl import Workbook

st.set_page_config(page_title="Dela i Excel", page_icon="📊", layout="centered")

st.title("📊 Dela i Excel")
st.markdown(
    "Klistra in dina värden (ett per rad) nedan. "
    "Verktyget delar upp dem i kolumner och skapar en Excel-fil att ladda ned."
)

with st.form("excel_form"):
    raw_text = st.text_area(
        "Klistra in dina värden (en per rad):",
        height=300,
        placeholder="Exempelvärde 1\nExempelvärde 2\n...",
    )
    chunk_size = st.number_input(
        "Antal rader per kolumn:", min_value=1, value=2000, step=100
    )
    submitted = st.form_submit_button("Skapa Excel-fil")

if submitted:
    lines = [r.strip() for r in raw_text.splitlines() if r.strip()]
    if not lines:
        st.warning("Inga värden att exportera. Klistra in dina värden först.")
    else:
        chunks = [lines[i : i + chunk_size] for i in range(0, len(lines), chunk_size)]

        wb = Workbook()
        ws = wb.active
        ws.title = "Delade värden"

        for col_idx, chunk in enumerate(chunks, start=1):
            for row_idx, val in enumerate(chunk, start=1):
                cell = ws.cell(row=row_idx, column=col_idx, value=str(val))
                cell.number_format = "@"

        buf = io.BytesIO()
        wb.save(buf)
        buf.seek(0)

        st.success(
            f"Klar! {len(lines)} värden delades upp i {len(chunks)} kolumn(er) "
            f"med max {chunk_size} rader var."
        )
        st.download_button(
            label="⬇️ Ladda ned Excel-filen",
            data=buf,
            file_name="delade_varden.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
