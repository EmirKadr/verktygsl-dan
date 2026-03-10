"""
Verktygsl-dan – Samlad webbsida för Python-verktyg
====================================================
Startsida med navigering till alla tillgängliga verktyg.
"""
import streamlit as st

st.set_page_config(
    page_title="Verktygsl-dan",
    page_icon="🔧",
    layout="wide",
)

st.title("🔧 Verktygsl-dan")
st.subheader("Samlad webbsida för alla verktyg")
st.divider()

st.markdown("### Välj ett verktyg i menyn till vänster, eller klicka direkt nedan:")

col1, col2, col3 = st.columns(3)

with col1:
    st.markdown("#### 📊 Dela i Excel")
    st.markdown(
        "Klistra in ett lista med värden och dela upp dem "
        "i kolumner om valfritt antal rader. Exportera direkt som Excel."
    )
    st.page_link("pages/1_Dela_i_Excel.py", label="Öppna verktyget", icon="📊")

    st.divider()

    st.markdown("#### 🔍 WMS-sök")
    st.markdown(
        "Ladda upp WMS-loggar (CSV) och sök på inköpsnummer + "
        "artikelnummer för att se pallstatus, saldo och leveranser."
    )
    st.page_link("pages/2_WMS_Sok.py", label="Öppna verktyget", icon="🔍")

with col2:
    st.markdown("#### 📦 Order/Saldo-analys")
    st.markdown(
        "Ladda upp en orderfil (CSV) och identifiera kompletta ordrar, "
        "artiklar med underskott och enradsordrar sorterade efter antal."
    )
    st.page_link("pages/3_Order_Saldo.py", label="Öppna verktyget", icon="📦")

    st.divider()

    st.markdown("#### 📋 Allokera (HIB)")
    st.markdown(
        "Komplex allokeringslogik för HIB-ordrar. Ladda upp filer "
        "och kör allokeringen direkt i webbläsaren."
    )
    st.page_link("pages/4_Allokera.py", label="Öppna verktyget", icon="📋")

with col3:
    st.markdown("#### 🗺️ LPA Canvas")
    st.markdown(
        "Interaktivt verktyg för att rita lagerplatskarta med "
        "noder och avstånd. **Körs lokalt som skrivbordsapp.**"
    )
    st.page_link("pages/5_LPA_Canvas.py", label="Läs mer", icon="🗺️")

    st.divider()

    st.markdown("#### 🤖 Bildklassificering")
    st.markdown(
        "Klassificera produktbilder med hjälp av en lokal AI-modell. "
        "**Kräver PyQt6 och lokal AI – körs som skrivbordsapp.**"
    )
    st.page_link("pages/6_Bildklassificering.py", label="Läs mer", icon="🤖")

st.divider()
st.caption("Verktygsl-dan · Alla verktyg på ett ställe")
