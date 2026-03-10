"""Bildklassificering – manuell klassificering av produktbilder"""
import csv
import io
import json
import zipfile
from io import StringIO

import pandas as pd
import streamlit as st

st.set_page_config(page_title="Bildklassificering", page_icon="🖼️", layout="wide")

CATEGORY_DOT = ["🟢", "🔵", "🟠", "🟣", "🩵", "🔴", "🟤", "⚫", "🟡"]
_EMPTY = {"", "0", "0,00000", "0.00000", "0,0", "0.0", "nan", "none"}


# ── DataManager (portad från classifier.py) ────────────────────────────────────

def _read_csv_bytes(data: bytes) -> list[dict]:
    """Auto-detect delimiter and parse CSV bytes → list of dicts."""
    text = data.decode("utf-8-sig", errors="replace")
    sample = text[:4096]
    try:
        dialect = csv.Sniffer().sniff(sample, delimiters=",;\t|")
    except csv.Error:
        dialect = csv.excel
    return list(csv.DictReader(StringIO(text), dialect=dialect))


def _load_attributes(rows: list[dict]) -> tuple[list[dict], dict]:
    """item_attribute CSV → (items_with_url, store_quantity_map)."""
    art_data: dict[tuple, dict] = {}
    for row in rows:
        art   = row.get("Artikel", "").strip()
        bolag = row.get("Bolag",   "").strip()
        namn  = row.get("Namn",    "").strip()
        val   = row.get("Värde",   "").strip()
        if not art:
            continue
        key = (art, bolag)
        if key not in art_data:
            art_data[key] = {"bolag": bolag}
        if namn == "IMG" and val.lower().startswith("http"):
            art_data[key]["url"] = val
        elif namn == "StoreQuantity":
            art_data[key]["store_quantity"] = val

    items = []
    store_qty: dict[tuple, str] = {}
    for (art, bolag), data in art_data.items():
        if "url" in data:
            items.append({"article_number": art, "url": data["url"], "bolag": bolag})
        if "store_quantity" in data:
            store_qty[(art, bolag)] = data["store_quantity"]
    return items, store_qty


def _load_alias(rows: list[dict]) -> dict:
    """item_alias CSV → dict[article_number → alias_dict]."""
    result = {}
    for row in rows:
        art = row.get("Artikel", "").strip()
        if not art or art in result:
            continue
        result[art] = {
            "ean":    row.get("Alias",  "").strip(),
            "enhet":  row.get("Enhet",  "").strip(),
            "faktor": row.get("Faktor", "").strip(),
            "langd":  row.get("Längd",  "").strip(),
            "bredd":  row.get("Bredd",  "").strip(),
            "hojd":   row.get("Höjd",   "").strip(),
            "bolag":  row.get("Bolag",  "").strip(),
        }
    return result


def _load_items(rows: list[dict]) -> dict:
    """item CSV → dict[article_number → item_dict]."""
    result = {}
    for row in rows:
        art = row.get("Artikel", "").strip()
        if not art:
            continue
        result[art] = {
            "beskrivning": row.get("Beskrivning", "").strip(),
            "un_nummer":   row.get("UN nummer",   "").strip(),
            "vikt_brutto": row.get("Vikt brutto", "").strip(),
            "vikt_netto":  row.get("Vikt netto",  "").strip(),
            "volym":       row.get("Volym",        "").strip(),
            "kategori":    row.get("Kategori",     "").strip(),
            "robot":       row.get("Robot",        "").strip(),
            "bolag":       row.get("Bolag",        "").strip(),
        }
    return result


def _load_main_category(rows: list[dict]) -> dict:
    """main_category CSV → dict[kategori_kod → huvudkategori]."""
    result = {}
    for row in rows:
        kat  = row.get("Kategori",      "").strip()
        hkat = row.get("Huvudkategori", "").strip()
        if kat and hkat:
            result[kat] = hkat
    return result


def _get_meta(art: str, bolag: str = "") -> dict | None:
    """Combine metadata from all loaded data sources for one article."""
    item_data  = st.session_state.get("meta_items", {})
    alias_data = st.session_state.get("meta_alias", {})
    cat_map    = st.session_state.get("meta_catmap", {})
    store_qty  = st.session_state.get("meta_storeqty", {})

    result: dict = {}
    if art in item_data:
        result.update(item_data[art])
    if art in alias_data:
        result.update(alias_data[art])
    cat_code = result.get("kategori", "")
    if cat_code and cat_code in cat_map:
        result["huvudkategori"] = cat_map[cat_code]
    sq = store_qty.get((art, bolag))
    if sq is None:
        sq = next((v for (a, _), v in store_qty.items() if a == art), None)
    if sq is not None:
        result["store_quantity"] = sq
    return result if result else None


def _detect_file_type(filename: str) -> str | None:
    """Detect which system file type a file is based on its name."""
    name = filename.lower()
    if name.startswith("item_attribute"):
        return "attribute"
    if name.startswith("item_alias"):
        return "alias"
    if name.startswith("item") and not name.startswith("item_"):
        return "item"
    if name.startswith("main_category"):
        return "main_category"
    return None


# ── state helpers ──────────────────────────────────────────────────────────────

def _init():
    defaults = {
        "phase": "setup",
        "test_name": "",
        "syfte": "",
        "categories": [],
        "items": [],
        "current_index": 0,
        "classifications": {},   # article_number → category_name
        "meta_items": {},
        "meta_alias": {},
        "meta_catmap": {},
        "meta_storeqty": {},
    }
    for k, v in defaults.items():
        if k not in st.session_state:
            st.session_state[k] = v


def _reset():
    for k in list(st.session_state.keys()):
        del st.session_state[k]
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

            session    = _read("session.json", {})
            items      = _read("csv_data.json", [])
            categorized = _read("categorized.json", [])

        st.session_state.test_name   = session.get("test_name", "Import")
        st.session_state.syfte       = session.get("syfte", "")
        st.session_state.categories  = session.get("categories", [])
        st.session_state.items = [
            {
                "article_number": str(r.get("article_number", "")),
                "url": r.get("url", r.get("img_path", "")),
                "bolag": r.get("bolag", ""),
            }
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
                    {
                        "name": r["Namn"].strip(),
                        "description": str(r.get("Beskrivning", "") or "").strip(),
                    }
                    for _, r in valid.iterrows()
                ]
                cats.append({"name": "Övrigt", "description": ""})
                st.session_state.categories = cats
                st.session_state.phase = "source"
                st.rerun()


# ── phase: source ─────────────────────────────────────────────────────────────

def phase_source():
    st.title(f"📁 Välj bildkälla – {st.session_state.test_name}")

    tab_sys, tab_csv, tab_manual = st.tabs([
        "📊 Systemfiler (item_attribute m.fl.)",
        "📄 Enkel CSV",
        "✍️ Manuella URL:er",
    ])

    # ── Tab 1: system files ───────────────────────────────────────────────────
    with tab_sys:
        st.markdown(
            "Ladda upp en eller flera av systemfilerna. "
            "Filtypen identifieras automatiskt via filnamnet."
        )

        col_info, col_legend = st.columns([2, 1])
        with col_legend:
            st.markdown("""
| Filnamn börjar med | Innehåll |
|---|---|
| `item_attribute` | **Bilder + StoreQuantity** ← källa |
| `item_alias` | EAN, mått, enhet |
| `item` | Beskrivning, vikt, volym |
| `main_category` | Kategori → Huvudkategori |
""")

        uploaded_files = st.file_uploader(
            "Välj CSV-filer",
            type=["csv"],
            accept_multiple_files=True,
            key="sys_files",
        )

        loaded_types: dict[str, str] = {}  # type → filename
        items_from_attr: list[dict] = []
        errors: list[str] = []

        if uploaded_files:
            for uf in uploaded_files:
                ftype = _detect_file_type(uf.name)
                if ftype is None:
                    errors.append(f"**{uf.name}** – okänd filtyp (ignoreras)")
                    continue
                try:
                    data = _read_csv_bytes(uf.read())
                    if ftype == "attribute":
                        items_from_attr, store_qty = _load_attributes(data)
                        st.session_state.meta_storeqty = store_qty
                    elif ftype == "alias":
                        st.session_state.meta_alias = _load_alias(data)
                    elif ftype == "item":
                        st.session_state.meta_items = _load_items(data)
                    elif ftype == "main_category":
                        st.session_state.meta_catmap = _load_main_category(data)
                    loaded_types[ftype] = uf.name
                except Exception as e:
                    errors.append(f"**{uf.name}** – fel: {e}")

            for err in errors:
                st.warning(err)

            if loaded_types:
                for ftype, fname in loaded_types.items():
                    icons = {
                        "attribute":     "✅ item_attribute",
                        "alias":         "✅ item_alias",
                        "item":          "✅ item",
                        "main_category": "✅ main_category",
                    }
                    st.success(f"{icons[ftype]}: {fname}")

            if "attribute" not in loaded_types:
                st.info(
                    "Ladda upp `item_attribute`-filen för att hämta artiklar och bild-URL:er. "
                    "Övriga filer är valfria och lägger till metadata."
                )
            else:
                st.success(f"**{len(items_from_attr)} artiklar** med bild-URL:er hittades.")

                # Show preview
                if items_from_attr:
                    prev = pd.DataFrame(items_from_attr[:5])
                    st.dataframe(prev, use_container_width=True)

                if st.button(
                    f"Starta klassificering med {len(items_from_attr)} artiklar →",
                    type="primary",
                    key="sys_go",
                ):
                    _start_classify(items_from_attr)

    # ── Tab 2: simple CSV ─────────────────────────────────────────────────────
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
                            items.append({"article_number": art, "url": url, "bolag": ""})
                    st.success(f"✓ {len(items)} artiklar med giltiga URL:er.")
                    st.dataframe(df[[art_col, url_col]].head(5), use_container_width=True)
                    if items and st.button("Starta klassificering →", type="primary", key="csv_go"):
                        _start_classify(items)
            except Exception as e:
                st.error(f"Kunde inte läsa CSV: {e}")

    # ── Tab 3: manual URLs ────────────────────────────────────────────────────
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
                    items.append({"article_number": art, "url": url, "bolag": ""})
            if not items:
                st.error("Inga giltiga URL:er hittades.")
            else:
                _start_classify(items)

    st.divider()
    if st.button("← Tillbaka till kategorier"):
        st.session_state.phase = "categories"
        st.rerun()


def _start_classify(items: list[dict]):
    st.session_state.items = items
    st.session_state.current_index = 0
    st.session_state.classifications = {}
    st.session_state.phase = "classify"
    st.rerun()


# ── phase: classify ───────────────────────────────────────────────────────────

def phase_classify():
    items      = st.session_state.items
    idx        = st.session_state.current_index
    categories = st.session_state.categories
    clsf       = st.session_state.classifications
    has_meta   = bool(
        st.session_state.get("meta_items")
        or st.session_state.get("meta_alias")
        or st.session_state.get("meta_catmap")
    )

    if idx >= len(items):
        st.session_state.phase = "done"
        st.rerun()
        return

    item  = items[idx]
    art   = item["article_number"]
    bolag = item.get("bolag", "")
    total = len(items)
    done_count = len(clsf)

    # ── header ────────────────────────────────────────────────────────────────
    col_title, col_stat = st.columns([3, 1])
    with col_title:
        st.markdown(f"### 🖼️ {st.session_state.test_name}")
    with col_stat:
        st.metric("Klassificerade", f"{done_count} / {total}")

    st.progress(idx / max(total, 1), text=f"Artikel {idx + 1} av {total}  ·  nr {art}")

    prev_cat = clsf.get(art)
    if prev_cat:
        st.info(f"Klassificerades tidigare som: **{prev_cat}**")

    # ── layout: image | categories (| metadata) ───────────────────────────────
    meta = _get_meta(art, bolag) if has_meta else None

    if meta:
        col_img, col_cat, col_meta = st.columns([3, 1, 1])
    else:
        col_img, col_cat = st.columns([3, 1])
        col_meta = None

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

    if col_meta and meta:
        with col_meta:
            st.markdown("**Artikelinfo**")
            fields = [
                ("Beskrivning",   meta.get("beskrivning")),
                ("Huvudkategori", meta.get("huvudkategori")),
                ("Kategori",      meta.get("kategori")),
                ("Robot",         meta.get("robot")),
                ("StoreQuantity", meta.get("store_quantity")),
                ("UN-nummer",     meta.get("un_nummer")),
                ("Vikt brutto",   meta.get("vikt_brutto")),
                ("Vikt netto",    meta.get("vikt_netto")),
                ("Volym",         meta.get("volym")),
                ("EAN",           meta.get("ean")),
                ("Enhet",         meta.get("enhet")),
                ("Längd",         meta.get("langd")),
                ("Bredd",         meta.get("bredd")),
                ("Höjd",          meta.get("hojd")),
            ]
            for label, value in fields:
                v = str(value).strip() if value else ""
                if v.lower() not in _EMPTY:
                    st.markdown(
                        f"<span style='color:#888;font-size:11px'>{label}</span><br>"
                        f"<span style='font-size:13px'>{v}</span>",
                        unsafe_allow_html=True,
                    )

    # ── navigation bar ────────────────────────────────────────────────────────
    st.divider()
    c1, c2, c3, c4 = st.columns([1, 1, 1, 2])
    with c1:
        if idx > 0 and st.button("← Föregående"):
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
        st.caption(f"{total - done_count} artiklar kvar att klassificera")


# ── phase: done ───────────────────────────────────────────────────────────────

def phase_done():
    clsf      = st.session_state.classifications
    items     = st.session_state.items
    test_name = st.session_state.test_name

    st.title(f"✅ {test_name} – exportera resultat")

    if not clsf:
        st.warning("Inga artiklar klassificerades.")
    else:
        rows = []
        for item in items:
            art = item["article_number"]
            if art in clsf:
                rows.append({
                    "article_number": art,
                    "url":            item["url"],
                    "category":       clsf[art],
                })
        df = pd.DataFrame(rows)

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
            st.download_button(
                "⬇️ Spara session (ZIP för lokal app)",
                data=_build_zip(),
                file_name=f"{test_name}_session.zip",
                mime="application/zip",
                help=(
                    "Öppna denna ZIP i den lokala classifier.py-appen. "
                    "De klassificerade artiklarna används som träningsexempel "
                    "när du kör AI-jobbet lokalt."
                ),
            )

    st.divider()
    col_back, col_restart = st.columns([1, 1])
    with col_back:
        if items and st.button("← Fortsätt klassificera"):
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


# ── ZIP export (kompatibel med lokal classifier.py) ───────────────────────────

def _build_zip() -> bytes:
    """Build a ZIP that classifier.py _import_zip() can open.

    session.json     – {test_name, syfte, categories}
    csv_data.json    – [{article_number, url, bolag, img_path:""}]
    categorized.json – [{article_number, category, url, image_path:""}]
    """
    clsf        = st.session_state.classifications
    items       = st.session_state.items
    item_by_art = {it["article_number"]: it for it in items}

    session = {
        "test_name":  st.session_state.test_name,
        "syfte":      st.session_state.syfte,
        "categories": st.session_state.categories,
    }
    csv_data = [
        {
            "article_number": it["article_number"],
            "url":            it["url"],
            "bolag":          it.get("bolag", ""),
            "img_path":       "",   # empty → AI worker downloads via url
        }
        for it in items
    ]
    categorized = [
        {
            "article_number": art,
            "category":       cat,
            "url":            item_by_art.get(art, {}).get("url", ""),
            "bolag":          item_by_art.get(art, {}).get("bolag", ""),
            "image_path":     "",   # no local image from web session
        }
        for art, cat in clsf.items()
    ]

    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("session.json",
                    json.dumps(session,     ensure_ascii=False, indent=2))
        zf.writestr("csv_data.json",
                    json.dumps(csv_data,    ensure_ascii=False, indent=2))
        zf.writestr("categorized.json",
                    json.dumps(categorized, ensure_ascii=False, indent=2))
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
