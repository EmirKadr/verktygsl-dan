"""
Bildklassificering – information om skrivbordsappen
"""
import streamlit as st

st.set_page_config(page_title="Bildklassificering", page_icon="🤖", layout="centered")

st.title("🤖 Bildklassificering")

st.info(
    "**Bildklassificering kräver en lokal AI-modell** och PyQt6. "
    "Det kan inte köras direkt i webbläsaren."
)

st.markdown("""
### Vad gör Bildklassificering?

Verktyget låter dig klassificera produktbilder med hjälp av AI:

- **Manuell klassificering** – bläddra bland bilder och tilldela kategorier
- **AI-klassificering** – använd en lokal språkmodell (t.ex. Qwen2.5-VL) för att
  automatiskt klassificera resterande bilder
- **Export** – spara klassificeringen som CSV/Excel
- **Kategorier** – skapa egna kategorier med färgkodning

### Krav för att köra lokalt

1. Python med PyQt6:
   ```
   pip install PyQt6 pillow openpyxl requests
   ```
2. En lokal AI-modell via **LM Studio** eller liknande
   (standardinställning: `http://localhost:1234/v1`)

3. Kör programmet:
   ```
   python classifier.py
   ```

### Varför finns det inte på webben?

Bildklassificering använder en lokal AI-modell och ett avancerat
grafiskt gränssnitt (PyQt6). Att flytta det till webb skulle kräva:
- En kraftfull server med GPU för AI-modellen
- Komplett omskrivning av gränssnittet

Det passar bäst som skrivbordsapp med en lokal AI-modell.
""")

st.divider()
st.caption("Bildklassificering · Kör lokalt med Python + LM Studio · classifier.py")
