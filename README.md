# Verktygsl-dan

Samlad webbsida för alla Python-verktyg – byggd med [Streamlit](https://streamlit.io).

## Verktyg

| Verktyg | Fil | Webb |
|---------|-----|------|
| 📊 Dela i Excel | `2000tal.py` | ✅ Fungerar på webben |
| 🔍 WMS-sök | `wms_sök79.py` | ✅ Fungerar på webben |
| 📦 Order/Saldo-analys | `OrderSaldo5.py` | ✅ Fungerar på webben |
| 📋 Allokera (HIB) | `allokera10.9.py` | 🔧 Under uppbyggnad |
| 🗺️ LPA Canvas | `LPA1.2.py` | 💻 Kör lokalt |
| 🤖 Bildklassificering | `classifier.py` | 💻 Kör lokalt |

---

## Kör lokalt

### Krav
- Python 3.10 eller senare

### Installation

```bash
pip install -r requirements.txt
```

### Starta webbappen

```bash
streamlit run app.py
```

Öppna sedan webbläsaren på `http://localhost:8501`.

---

## Hosta på internet

### Alternativ 1 – Streamlit Community Cloud (gratis, enklast)

1. Ladda upp detta repo till GitHub (om det inte redan är där)
2. Gå till [share.streamlit.io](https://share.streamlit.io)
3. Logga in med GitHub
4. Välj ditt repo, branch `main`, och startfil `app.py`
5. Klicka **Deploy** – appen är online inom några minuter!

> ⚠️ Streamlit Community Cloud är **gratis** men appen sover efter
> ett tag av inaktivitet och vaknar igen när någon besöker den.
> Passar bra för interna verktyg med liten trafik.

### Alternativ 2 – Render.com (gratis/betald)

1. Skapa ett konto på [render.com](https://render.com)
2. Ny "Web Service" → koppla till ditt GitHub-repo
3. Build command: `pip install -r requirements.txt`
4. Start command: `streamlit run app.py --server.port $PORT --server.address 0.0.0.0`
5. Välj gratis-plan (begränsad) eller betald plan

### Alternativ 3 – Egen VPS (t.ex. DigitalOcean, Hetzner)

1. Köp en server (t.ex. DigitalOcean Droplet, ~6 $/mån)
2. Installera Python och beroenden
3. Kör appen med `streamlit run app.py --server.port 8501`
4. Sätt upp nginx som reverse proxy för att exponera port 80/443
5. Lägg till SSL-certifikat med Let's Encrypt

---

## Filstruktur

```
verktygsl-dan/
├── app.py                    # Startfil / landningssida
├── requirements.txt          # Python-beroenden
├── .streamlit/
│   └── config.toml           # Streamlit-konfiguration
├── pages/
│   ├── 1_Dela_i_Excel.py     # Dela värden i Excel-kolumner
│   ├── 2_WMS_Sok.py          # WMS-sök med CSV-uppladdning
│   ├── 3_Order_Saldo.py      # Order/Saldo-analys
│   ├── 4_Allokera.py         # HIB-allokering (under uppbyggnad)
│   ├── 5_LPA_Canvas.py       # Info om LPA Canvas (desktop)
│   └── 6_Bildklassificering.py # Info om Bildklassificering (desktop)
├── tools/
│   ├── wms_logic.py          # WMS-analyslogik (utan GUI)
│   └── order_saldo_logic.py  # Order/Saldo-logik (utan GUI)
└── (originala .py-filer)
```
