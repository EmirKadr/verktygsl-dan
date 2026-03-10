"""
Order/Saldo-analys – webversion av OrderSaldo5.py
Ladda upp en CSV och analysera ordrar och saldo.
"""
import io

import pandas as pd
import streamlit as st

from tools.order_saldo_logic import analyze, auto_map_columns, read_csv_flex

st.set_page_config(page_title="Order/Saldo-analys", page_icon="📦", layout="wide")

st.title("📦 Order/Saldo-analys")
st.markdown(
    "Ladda upp din orderfil (CSV/TXT) för att identifiera:\n"
    "- **Kompletta ordrar** – saldo räcker för alla orderrader\n"
    "- **Artiklar att beställa** – underskott i saldo\n"
    "- **Enradsordrar** – ordrar med exakt 1 rad och 1–4 beställda"
)

# ── Filuppladdning ──────────────────────────────────────────────────────────────

uploaded_file = st.file_uploader("Välj CSV/TXT-fil:", type=["csv", "txt"])

if uploaded_file:
    # Spara tillfälligt till disk för read_csv_flex
    import tempfile, os
    with tempfile.NamedTemporaryFile(delete=False, suffix=".csv") as tmp:
        tmp.write(uploaded_file.read())
        tmp_path = tmp.name

    try:
        df = read_csv_flex(tmp_path)
    except Exception as e:
        st.error(f"Kunde inte läsa filen: {e}")
        os.unlink(tmp_path)
        st.stop()
    finally:
        try:
            os.unlink(tmp_path)
        except Exception:
            pass

    st.success(f"Inläst: {uploaded_file.name} ({len(df)} rader, {len(df.columns)} kolumner)")

    # Kolumnmappning
    try:
        mapping = auto_map_columns(df)
    except Exception as e:
        st.error(str(e))
        st.stop()

    m = mapping
    st.caption(
        f"Kolumner: Order=**{m['order']}** | Artikel=**{m['article']}** | "
        f"Plock=**{m['pick']}** | Beställt=**{m['demand']}** | "
        f"Plockat=**{m.get('pickedqty') or '–'}** | Namn=**{m.get('name') or '–'}**"
    )

    if st.button("▶ Analysera", type="primary"):
        with st.spinner("Analyserar…"):
            try:
                result = analyze(df, mapping)
            except Exception as e:
                st.error(f"Fel vid analys: {e}")
                st.stop()

        st.divider()

        tab1, tab2, tab3, tab4, tab5, tab6 = st.tabs([
            f"✅ Kompletta ordrar ({len(result['complete_orders'])})",
            f"⚠️ Artiklar att beställa ({len(result['holistic_short'])})",
            f"1×1 ({len(result['orders_1x1'])})",
            f"1×2 ({len(result['orders_1x2'])})",
            f"1×3 ({len(result['orders_1x3'])})",
            f"1×4 ({len(result['orders_1x4'])})",
        ])

        def order_list_tab(tab, orders: list, filename: str):
            with tab:
                if not orders:
                    st.info("Inga ordrar hittades.")
                    return
                st.metric("Antal ordrar", len(orders))
                col_a, col_b = st.columns(2)
                with col_a:
                    text = "\n".join(orders)
                    st.text_area("Ordernummer:", value=text, height=200)
                with col_b:
                    df_dl = pd.DataFrame({"Order nr": orders})
                    buf = io.StringIO()
                    df_dl.to_csv(buf, index=False)
                    st.download_button(
                        "⬇️ Ladda ned som CSV",
                        data=buf.getvalue().encode("utf-8-sig"),
                        file_name=filename,
                        mime="text/csv",
                    )

        order_list_tab(tab1, result["complete_orders"], "kompletta_ordrar.csv")

        with tab2:
            hs = result["holistic_short"]
            if hs.empty:
                st.info("Inga underskott hittades – saldo räcker för alla artiklar.")
            else:
                st.metric("Artiklar med underskott", len(hs))
                st.dataframe(hs.reset_index(), use_container_width=True)
                buf = io.StringIO()
                hs.reset_index().to_csv(buf, index=False)
                st.download_button(
                    "⬇️ Ladda ned som CSV",
                    data=buf.getvalue().encode("utf-8-sig"),
                    file_name="artiklar_att_bestalla.csv",
                    mime="text/csv",
                )

        order_list_tab(tab3, result["orders_1x1"], "ordrar_1rad_1st.csv")
        order_list_tab(tab4, result["orders_1x2"], "ordrar_1rad_2st.csv")
        order_list_tab(tab5, result["orders_1x3"], "ordrar_1rad_3st.csv")
        order_list_tab(tab6, result["orders_1x4"], "ordrar_1rad_4st.csv")
