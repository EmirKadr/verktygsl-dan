"""
LPA Canvas – information om skrivbordsappen
"""
import streamlit as st

st.set_page_config(page_title="LPA Canvas", page_icon="🗺️", layout="centered")

st.title("🗺️ LPA Canvas")

st.info(
    "**LPA Canvas är ett interaktivt ritverktyg** som kräver en lokal installation "
    "och kan inte köras direkt i webbläsaren."
)

st.markdown("""
### Vad gör LPA Canvas?

LPA Canvas låter dig bygga en interaktiv lagerplatskarta:

- **Lägg till platser** (noder) på en canvas
- **Koppla platser** med avstånd (kanter)
- **Automatisk namnföljd** – baserat på en ordningslista (t.ex. AA66 → AA67)
- **Importera/exportera** karta som CSV eller JSON-projekt
- **Zoom och panorera** med mushjul och höger musknapp

### Hur kör jag LPA Canvas?

1. Se till att du har Python installerat på din dator
2. Installera beroenden:
   ```
   pip install pandas openpyxl
   ```
3. Kör programmet:
   ```
   python LPA1.2.py
   ```

### Varför finns det inte på webben?

LPA Canvas är ett grafiskt ritverktyg med en interaktiv canvas.
Det kräver realtidsinteraktion (drag, zoom, klick) som är svår att
replikera i en webbapp utan att bygga om hela verktyget i JavaScript.
Det fungerar bäst som en skrivbordsapp.
""")

st.divider()
st.caption("LPA Canvas · Kör lokalt med Python · LPA1.2.py")
