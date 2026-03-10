"""Bildklassificering – manuell klassificering av produktbilder"""
import io
import json
import zipfile

import pandas as pd
import streamlit as st

st.set_page_config(page_title="Bildklassificering", page_icon="🖼️", layout="wide")

CATEGORY_DOT = ["🟢", "🔵", "🟠", "🟣", "🩵", "🔴", "🟤", "⚫", "🟡"]
_EMPTY = {"", "0", "0,00000", "0.00000", "0,0", "0.0"}


# ── state helpers ──────────────────────────────────────────────────────────────

def _init():
    defaults = {
        "phase": "setup",
        "test_name": "",
        "syfte": "",
        "categories": [],
        "items": [],
        "current_index": 0,
        "classifications": {},  # article_number → category_name
    }
    for k, v in defaults.items():
        if k not in st.session_state:
            st.session_state[k] = v


def _reset():
    for k in ["phase", "test_name", "syfte", "categories", "items",
              "current_index", "classifications", "cat_df"]:
        st.session_state.pop(k, None)
    st.rerun()


# ── phase: setup ──────────────────────────────────────────────────────────────

def phase_setup():
    st.title("🖼️ Bildklassificering")

    col, _ = st.columns([1, 1])
    with col:
        st.subheader("Skapa nytt test")
        with st.form("setup_form"):
            name = st.text_input("Namn på testet", placeholder="t.ex. Testomgång 1")
            syfte = st.text_area(
                "Syfte (valfritt)",
                placeholder="Vad ska klassificeringen användas till?",
                height=80,
            )
            submitted = st.form_submit_button("Gå vidare →", type="primary")

        if submitted:
            if not name.strip():
                st.error("Ange ett namn för testet.")
            else:
                st.session_state.test_name = name.strip()
                st.session_state.syfte = syfte.strip()
                st.session_state.phase = "categories"
                st.rerun()

        st.divider()
        st.subheader("eller öppna sparad session")
        zf = st.file_uploader("Ladda upp ZIP-session", type=["zip"], key="zip_upload")
        if zf and st.button("Öppna session", type="secondary"):
            _load_zip(zf)


def _load_zip(zf):
    try:
        with zipfile.ZipFile(io.BytesIO(zf.read())) as z:
            def _read(name, default):
                if name in z.namelist():
                    return json.loads(z.read(name).decode("utf-8"))
                return default

            session = _read("session.json", {})
            items = _read("csv_data.json", [])
            categorized = _read("categorized.json", [])

        st.session_state.test_name = session.get("test_name", "Import")
        st.session_state.syfte = session.get("syfte", "")
        st.session_state.categories = session.get("categories", [])
        st.session_state.items = [
            {"article_number": str(r.get("article_number", "")), "url": r.get("url", r.get("img_path", ""))}
            for r in items
            if r.get("url") or r.get("img_path")
        ]
        st.session_state.classifications = {
            str(c["article_number"]): c["category"]
            for c in categorized
            if c.get("article_number") and c.get("category")
        }
        st.session_state.current_index = len(st.session_state.items)
        st.session_state.phase = "done"
        st.rerun()
    except Exception as e:
        st.error(f"Kunde inte läsa ZIP: {e}")


# ── phase: categories ─────────────────────────────────────────────────────────

def phase_setup_categories():
    st.title(f"📂 Kategorier – {st.session_state.test_name}")
    st.caption('"Övrigt" läggs till automatiskt som sista kategori.')

    if "cat_df" not in st.session_state:
        st.session_state.cat_df = pd.DataFrame(
            {"Namn": ["", "", ""], "Beskrivning": ["", "", ""]}
        )

    edited = st.data_editor(
        st.session_state.cat_df,
        num_rows="dynamic",
        use_container_width=True,
        column_config={
            "Namn": st.column_config.TextColumn("Kategorinamn", width="medium"),
            "Beskrivning": st.column_config.TextColumn("Beskrivning (valfritt)", width="large"),
        },
        hide_index=True,
    )

    col_back, col_next, _ = st.columns([1, 2, 3])
    with col_back:
        if st.button("← Tillbaka"):
            st.session_state.phase = "setup"
            st.rerun()
    with col_next:
        if st.button("Välj bildkälla →", type="primary"):
            valid = edited[edited["Namn"].str.strip().astype(bool)]
            if valid.empty:
                st.error("Ange minst en kategori.")
            else:
                st.session_state.cat_df = edited
                cats = [
                    {"name": r["Namn"].strip(), "description": str(r.get("Beskrivning", "") or "").strip()}
                    for _, r in valid.iterrows()
                ]
                cats.append({"name": "Övrigt", "description": ""})
                st.session_state.categories = cats
                st.session_state.phase = "source"
                st.rerun()


# ── phase: source ─────────────────────────────────────────────────────────────

def phase_source():
    st.title(f"📁 Välj bildkälla – {st.session_state.test_name}")

    tab_csv, tab_manual = st.tabs(["📄 Ladda upp CSV", "✍️ Ange URL:er manuellt"])

    with tab_csv:
        st.markdown(
            "CSV-filen ska ha kolumnerna **`article_number`** (eller `Artikel`) "
            "och **`url`** (eller `URL`/`Bild`) med bild-URL:er."
        )
        uploaded = st.file_uploader("Välj CSV", type=["csv"], key="csv_upload")
        if uploaded:
            try:
                df = pd.read_csv(uploaded)
                col_map = {c.lower().strip(): c for c in df.columns}
                art_col = (
                    col_map.get("article_number")
                    or col_map.get("artikel")
                    or col_map.get("artikelnummer")
                )
                url_col = (
                    col_map.get("url")
                    or col_map.get("bild")
                    or col_map.get("bildurl")
                    or col_map.get("bild-url")
                )
                if not art_col or not url_col:
                    st.error(f"Kunde inte hitta kolumner. Hittade: {list(df.columns)}")
                    st.info("Filen behöver kolumnerna `article_number` och `url`.")
                else:
                    items = []
                    for _, row in df.iterrows():
                        art = str(row[art_col]).strip()
                        url = str(row[url_col]).strip()
                        if art and url and url.startswith("http"):
                            items.append({"article_number": art, "url": url})
                    st.success(f"✓ {len(items)} artiklar med giltiga URL:er.")
                    st.dataframe(df[[art_col, url_col]].head(5), use_container_width=True)
                    if items and st.button("Starta klassificering →", type="primary", key="csv_go"):
                        _start_classify(items)
            except Exception as e:
                st.error(f"Kunde inte läsa CSV: {e}")

    with tab_manual:
        st.markdown("En rad per artikel. Format: `artikelnummer,https://...` eller bara URL.")
        text = st.text_area(
            "URL:er",
            height=200,
            placeholder="12345,https://example.com/bild.jpg\n12346,https://example.com/bild2.jpg",
        )
        if st.button("Starta →", type="primary", key="manual_go"):
            items = []
            for i, line in enumerate(text.strip().splitlines()):
                line = line.strip()
                if not line:
                    continue
                if "," in line:
                    art, url = line.split(",", 1)
                    art, url = art.strip(), url.strip()
                else:
                    art, url = str(i + 1), line.strip()
                if url.startswith("http"):
                    items.append({"article_number": art, "url": url})
            if not items:
                st.error("Inga giltiga URL:er hittades.")
            else:
                _start_classify(items)

    st.divider()
    if st.button("← Tillbaka till kategorier"):
        st.session_state.phase = "categories"
        st.rerun()


def _start_classify(items):
    st.session_state.items = items
    st.session_state.current_index = 0
    st.session_state.classifications = {}
    st.session_state.phase = "classify"
    st.rerun()


# ── phase: classify ───────────────────────────────────────────────────────────

def phase_classify():
    items = st.session_state.items
    idx = st.session_state.current_index
    categories = st.session_state.categories
    clsf = st.session_state.classifications

    if idx >= len(items):
        st.session_state.phase = "done"
        st.rerun()
        return

    item = items[idx]
    art = item["article_number"]
    total = len(items)
    done_count = len(clsf)

    # ── header ────────────────────────────────────────────────────────────────
    col_title, col_stat = st.columns([3, 1])
    with col_title:
        st.markdown(f"### 🖼️ {st.session_state.test_name}")
    with col_stat:
        st.metric("Klassificerade", f"{done_count} / {total}")

    st.progress(idx / total, text=f"Artikel {idx + 1} av {total}  ·  nr {art}")

    # already classified label
    prev_cat = clsf.get(art)
    if prev_cat:
        st.info(f"Klassificerades tidigare som: **{prev_cat}**")

    # ── image + categories ────────────────────────────────────────────────────
    col_img, col_cat = st.columns([3, 1])

    with col_img:
        try:
            st.image(item["url"], use_container_width=True)
        except Exception:
            st.warning("Bilden kunde inte laddas.")
            st.code(item["url"])

    with col_cat:
        st.markdown("**Välj kategori**")
        for i, cat in enumerate(categories):
            dot = CATEGORY_DOT[i % len(CATEGORY_DOT)] if cat["name"] != "Övrigt" else "⬜"
            label = f"{dot} {cat['name']}"
            if st.button(label, key=f"cat_btn_{i}", use_container_width=True):
                clsf[art] = cat["name"]
                st.session_state.classifications = clsf
                st.session_state.current_index = idx + 1
                st.rerun()

    # ── navigation bar ────────────────────────────────────────────────────────
    st.divider()
    c1, c2, c3, c4 = st.columns([1, 1, 1, 2])
    with c1:
        if idx > 0:
            if st.button("← Föregående"):
                # un-classify previous item so it can be re-classified
                prev_art = items[idx - 1]["article_number"]
                clsf.pop(prev_art, None)
                st.session_state.classifications = clsf
                st.session_state.current_index = idx - 1
                st.rerun()
    with c2:
        if st.button("Hoppa över →"):
            st.session_state.current_index = idx + 1
            st.rerun()
    with c3:
        if st.button("⬇️ Avsluta & exportera"):
            st.session_state.phase = "done"
            st.rerun()
    with c4:
        remaining = total - done_count
        st.caption(f"{remaining} artiklar kvar att klassificera")


# ── phase: done ───────────────────────────────────────────────────────────────

def phase_done():
    clsf = st.session_state.classifications
    items = st.session_state.items
    test_name = st.session_state.test_name

    st.title(f"✅ {test_name} – exportera resultat")

    if not clsf:
        st.warning("Inga artiklar klassificerades.")
    else:
        # Build result DataFrame
        rows = []
        for item in items:
            art = item["article_number"]
            if art in clsf:
                rows.append({
                    "article_number": art,
                    "url": item["url"],
                    "category": clsf[art],
                })
        df = pd.DataFrame(rows)

        # Summary chart
        summary = df.groupby("category").size().reset_index(name="Antal")
        st.subheader(f"{len(df)} artiklar klassificerade")
        st.bar_chart(summary.set_index("category"))

        st.dataframe(df, use_container_width=True)

        col_csv, col_xlsx, col_zip = st.columns(3)

        with col_csv:
            st.download_button(
                "⬇️ Ladda ner CSV",
                data=df.to_csv(index=False),
                file_name=f"{test_name}_klassificering.csv",
                mime="text/csv",
            )

        with col_xlsx:
            try:
                buf = io.BytesIO()
                df.to_excel(buf, index=False, engine="openpyxl")
                buf.seek(0)
                st.download_button(
                    "⬇️ Ladda ner Excel",
                    data=buf,
                    file_name=f"{test_name}_klassificering.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )
            except ImportError:
                st.caption("openpyxl ej installerat – Excel ej tillgänglig.")

        with col_zip:
            zip_buf = _build_zip()
            st.download_button(
                "⬇️ Spara session (ZIP)",
                data=zip_buf,
                file_name=f"{test_name}_session.zip",
                mime="application/zip",
            )

    st.divider()
    col_back, col_restart = st.columns([1, 1])
    with col_back:
        if items and st.button("← Fortsätt klassificera"):
            # Return to first unclassified item
            clsf_arts = set(clsf.keys())
            for i, item in enumerate(items):
                if item["article_number"] not in clsf_arts:
                    st.session_state.current_index = i
                    break
            else:
                st.session_state.current_index = len(items)
            st.session_state.phase = "classify"
            st.rerun()
    with col_restart:
        if st.button("🔄 Starta om"):
            _reset()


def _build_zip() -> bytes:
    """Package session data as a ZIP compatible with the local classifier.py app.

    Format expected by classifier.py _import_zip():
      session.json    – {test_name, syfte, categories}
      csv_data.json   – [{article_number, url, img_path:""}]   ← img_path empty = use url
      categorized.json– [{article_number, category, url, image_path:""}]
    """
    clsf = st.session_state.classifications
    items = st.session_state.items
    item_by_art = {it["article_number"]: it for it in items}

    session = {
        "test_name": st.session_state.test_name,
        "syfte": st.session_state.syfte,
        "categories": st.session_state.categories,
    }

    # All items — local app needs img_path (empty → AI worker downloads via url)
    csv_data = [
        {"article_number": it["article_number"], "url": it["url"], "img_path": ""}
        for it in items
    ]

    # Only manually classified items — used by AI as training examples
    categorized = [
        {
            "article_number": art,
            "category": cat,
            "url": item_by_art.get(art, {}).get("url", ""),
            "image_path": "",   # no local image from web session
        }
        for art, cat in clsf.items()
    ]

    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("session.json", json.dumps(session, ensure_ascii=False, indent=2))
        zf.writestr("csv_data.json", json.dumps(csv_data, ensure_ascii=False, indent=2))
        zf.writestr("categorized.json", json.dumps(categorized, ensure_ascii=False, indent=2))
    buf.seek(0)
    return buf.read()


# ── router ────────────────────────────────────────────────────────────────────

_init()

phase = st.session_state.phase
if phase == "setup":
    phase_setup()
elif phase == "categories":
    phase_setup_categories()
elif phase == "source":
    phase_source()
elif phase == "classify":
    phase_classify()
elif phase == "done":
    phase_done()
