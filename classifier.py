#!/usr/bin/env python3
"""Bildklassificering med AI-stöd — PyQt6"""

import sys
import csv
import json
import base64
import random
import shutil
import tempfile
import threading
import urllib.request
from io import BytesIO
from pathlib import Path
from typing import Optional, Dict, List, Tuple

from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QLabel, QLineEdit, QPushButton, QFrame, QScrollArea, QTextEdit,
    QFileDialog, QMessageBox, QCheckBox, QStackedWidget, QGridLayout,
    QDialog, QDialogButtonBox, QProgressBar, QSizePolicy,
    QRadioButton, QButtonGroup, QScrollBar,
)
from PyQt6.QtCore import Qt, QThread, pyqtSignal, QTimer, QSize, QPoint, QMimeData, QByteArray
from PyQt6.QtGui import QPixmap, QKeySequence, QShortcut, QFont, QDrag

try:
    from PIL import Image as PILImage
    PIL_AVAILABLE = True
except ImportError:
    PIL_AVAILABLE = False

try:
    import openpyxl
    OPENPYXL_AVAILABLE = True
except ImportError:
    OPENPYXL_AVAILABLE = False

try:
    import requests as req
    REQUESTS_AVAILABLE = True
except ImportError:
    REQUESTS_AVAILABLE = False

# ── constants ──────────────────────────────────────────────────────────────────
IMAGE_DIR           = Path("bilder")
DATA_DIR            = Path("data")
SUPPORTED_EXT       = {".jpg", ".jpeg", ".png", ".gif", ".bmp", ".webp", ".tiff"}
CATEGORY_COLORS     = [
    "#4CAF50", "#2196F3", "#FF9800", "#9C27B0",
    "#00BCD4", "#E91E63", "#795548", "#607D8B", "#FF5722",
]
_EMPTY              = {"", "0", "0,00000", "0.00000", "0,0", "0.0"}
DEFAULT_MODEL       = "qwen2.5-vl-72b-instruct"
DEFAULT_AI_URL      = "http://localhost:1234/v1"
MAX_EXAMPLES_PER_CAT  = 10   # manually classified articles used per category in AI job (step 1)
MAX_OVRIGT_EXAMPLES   = 50   # Övrigt gets more examples since it's more diverse
AI_JOB_MIN_PER_CAT    = 1    # minimum examples per category to unlock AI job button

# ── global stylesheet ──────────────────────────────────────────────────────────
STYLE = """
QMainWindow, QWidget {
    background-color: #1e1e2e;
    color: #cdd6f4;
    font-family: 'Segoe UI', Arial, sans-serif;
    font-size: 13px;
}
QLabel { color: #cdd6f4; }
QLineEdit, QTextEdit {
    background-color: #313244;
    border: 1px solid #45475a;
    border-radius: 6px;
    color: #cdd6f4;
    padding: 5px 10px;
}
QLineEdit:focus, QTextEdit:focus { border: 1px solid #89b4fa; }
QPushButton {
    border-radius: 6px;
    padding: 8px 16px;
    font-weight: bold;
    border: none;
}
QPushButton:hover { opacity: 0.85; }
QPushButton:pressed { opacity: 0.7; }
QScrollArea { border: none; }
QCheckBox { color: #cdd6f4; }
QCheckBox::indicator { width: 16px; height: 16px; border-radius: 3px;
                       border: 1px solid #45475a; background: #313244; }
QCheckBox::indicator:checked { background: #89b4fa; }
QMessageBox { background-color: #1e1e2e; }
QDialog { background-color: #1e1e2e; }
"""


def mk_btn(text: str, bg: str = "#4CAF50", fg: str = "white",
           min_w: int = 0, h: int = 0) -> QPushButton:
    b = QPushButton(text)
    style = f"background-color:{bg}; color:{fg}; border-radius:6px; padding:8px 16px; font-weight:bold;"
    if min_w:
        style += f" min-width:{min_w}px;"
    b.setStyleSheet(style)
    if h:
        b.setFixedHeight(h)
    return b


def sep() -> QFrame:
    f = QFrame()
    f.setFrameShape(QFrame.Shape.HLine)
    f.setStyleSheet("color: #313244;")
    return f


# ── DataManager ────────────────────────────────────────────────────────────────
class DataManager:
    def __init__(self):
        self.builtin_attributes: List[Dict] = []
        self.store_quantity_data: Dict[Tuple[str, str], str] = {}  # (art, bolag) -> qty
        self.item_data:    Dict[str, Dict] = {}
        self.alias_data:   Dict[str, Dict] = {}
        self.category_map: Dict[str, str]  = {}
        self._load_all()

    def _load_all(self):
        if not DATA_DIR.exists():
            return
        for f in sorted(DATA_DIR.iterdir()):
            name = f.name.lower()
            if not name.endswith(".csv"):
                continue
            if name.startswith("item_attribute"):
                self._load_attributes(f)
            elif name.startswith("item_alias"):
                self._load_alias(f)
            elif name.startswith("item") and not name.startswith("item_"):
                self._load_items(f)
            elif name.startswith("main_category"):
                self._load_main_category(f)

    def _read_tsv(self, path) -> List[Dict]:
        try:
            with open(path, newline="", encoding="utf-8-sig") as fh:
                sample = fh.read(4096); fh.seek(0)
                try:
                    dialect = csv.Sniffer().sniff(sample, delimiters=",;\t|")
                except csv.Error:
                    dialect = csv.excel
                return list(csv.DictReader(fh, dialect=dialect))
        except Exception:
            return []

    def _load_attributes(self, path):
        self.builtin_attributes = []
        self.store_quantity_data = {}
        art_data: Dict[Tuple[str, str], Dict] = {}
        for row in self._read_tsv(path):
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
        for (art, bolag), data in art_data.items():
            if "url" in data:
                self.builtin_attributes.append({
                    "article_number": art,
                    "url": data["url"],
                    "bolag": bolag,
                })
            if "store_quantity" in data:
                self.store_quantity_data[(art, bolag)] = data["store_quantity"]

    def _load_alias(self, path):
        self.alias_data = {}
        for row in self._read_tsv(path):
            art = row.get("Artikel", "").strip()
            if not art or art in self.alias_data:
                continue
            self.alias_data[art] = {
                "ean":   row.get("Alias",  "").strip(),
                "enhet": row.get("Enhet",  "").strip(),
                "faktor":row.get("Faktor", "").strip(),
                "langd": row.get("Längd",  "").strip(),
                "bredd": row.get("Bredd",  "").strip(),
                "hojd":  row.get("Höjd",   "").strip(),
                "bolag": row.get("Bolag",  "").strip(),
            }

    def _load_items(self, path):
        self.item_data = {}
        for row in self._read_tsv(path):
            art = row.get("Artikel", "").strip()
            if not art:
                continue
            self.item_data[art] = {
                "beskrivning": row.get("Beskrivning", "").strip(),
                "un_nummer":   row.get("UN nummer",   "").strip(),
                "vikt_brutto": row.get("Vikt brutto", "").strip(),
                "vikt_netto":  row.get("Vikt netto",  "").strip(),
                "volym":       row.get("Volym",        "").strip(),
                "kategori":    row.get("Kategori",     "").strip(),
                "robot":       row.get("Robot",        "").strip(),
                "bolag":       row.get("Bolag",        "").strip(),
            }

    def _load_main_category(self, path):
        self.category_map = {}
        for row in self._read_tsv(path):
            kat  = row.get("Kategori",      "").strip()
            hkat = row.get("Huvudkategori", "").strip()
            if kat and hkat:
                self.category_map[kat] = hkat

    def get_meta(self, article_str: str, bolag: str = "") -> Optional[Dict]:
        art = article_str.strip()
        result: Dict = {}
        if art in self.item_data:
            result.update(self.item_data[art])
        if art in self.alias_data:
            result.update(self.alias_data[art])
        cat_code = result.get("kategori", "")
        if cat_code and cat_code in self.category_map:
            result["huvudkategori"] = self.category_map[cat_code]
        # Look up StoreQuantity: prefer matching bolag, fall back to any
        sq = self.store_quantity_data.get((art, bolag))
        if sq is None:
            sq = next((v for (a, _), v in self.store_quantity_data.items() if a == art), None)
        if sq is not None:
            result["store_quantity"] = sq
        return result or None


# ── AIJobWorker ────────────────────────────────────────────────────────────────
class AIJobWorker(QThread):
    """Two-step AI classification job.

    Step 1 — Category knowledge: For each non-Övrigt category, collect up to
    MAX_EXAMPLES_PER_CAT manually classified articles and ask the LLM to
    summarise what they have in common (text metadata + 1 representative image).

    Step 2 — Classify remaining: For every article in csv_data that has not yet
    been manually classified, download its image (if needed) and ask the LLM
    which category it belongs to, using the summaries from step 1.
    """
    progress           = pyqtSignal(str)
    knowledge_ready    = pyqtSignal(str, str)            # (category_name, knowledge_text)
    contrast_ready     = pyqtSignal(str)                 # (contrast_knowledge_text)
    article_classified = pyqtSignal(str, str, str, str)  # (article_number, category, url, image_path)
    finished_all       = pyqtSignal()
    error              = pyqtSignal(str)

    def __init__(self, categories, categorized, csv_data, syfte,
                 api_url, model, compress, data_mgr, parent=None):
        super().__init__(parent)
        self.categories  = categories   # list[{name, description, knowledge}]
        self.categorized = categorized  # already manually classified items
        self.csv_data    = csv_data     # full article list
        self.syfte       = syfte
        self.api_url     = api_url
        self.model       = model
        self.compress    = compress
        self.data_mgr    = data_mgr
        self._stop       = False
        self._paused     = False

    def stop(self):
        self._stop = True

    def pause(self):
        self._paused = True

    def resume(self):
        self._paused = False

    def _wait_if_paused(self):
        import time
        while self._paused and not self._stop:
            time.sleep(0.1)

    # ── main run ───────────────────────────────────────────────────────────────

    def run(self):
        if not REQUESTS_AVAILABLE:
            self.error.emit("requests ej installerat")
            return

        # ── Step 1: generate category knowledge summaries ──────────────────
        self.progress.emit("=== Steg 1: Genererar kategorikunskap ===")
        self.cat_knowledge: Dict[str, str] = {}
        cat_knowledge = self.cat_knowledge  # alias so rest of run() is unchanged
        by_cat: Dict[str, List[Dict]] = {}
        for item in self.categorized:
            cat = item.get("category", "")
            if cat and cat != "Övrigt":
                by_cat.setdefault(cat, []).append(item)

        for cat in self.categories:
            if self._stop:
                return
            name = cat["name"]
            if name == "Övrigt":
                ovrigt_items = by_cat.get("Övrigt", [])[:MAX_OVRIGT_EXAMPLES]
                if ovrigt_items:
                    self.progress.emit(
                        f"  Analyserar Övrigt ({len(ovrigt_items)} artiklar — bredare analys)…"
                    )
                    try:
                        knowledge = self._generate_ovrigt_knowledge(ovrigt_items)
                        cat_knowledge["Övrigt"] = knowledge
                        self.knowledge_ready.emit("Övrigt", knowledge)
                        self.progress.emit("  ✓ Övrigt klar")
                    except Exception as e:
                        self.progress.emit(f"  ✗ Övrigt: {e}")
                        cat_knowledge["Övrigt"] = cat.get("description", "")
                continue
            items = by_cat.get(name, [])[:MAX_EXAMPLES_PER_CAT]
            if not items:
                self.progress.emit(f"  Hoppar {name} — inga exempelartiklar")
                cat_knowledge[name] = cat.get("description", "")
                continue
            self.progress.emit(f"  Analyserar {name} ({len(items)} artiklar)…")
            try:
                knowledge = self._generate_knowledge(name, cat.get("description", ""), items)
                cat_knowledge[name] = knowledge
                self.knowledge_ready.emit(name, knowledge)
                self.progress.emit(f"  ✓ {name} klar")
            except Exception as e:
                self.progress.emit(f"  ✗ {name}: {e}")
                cat_knowledge[name] = cat.get("description", "")

        # ── Step 1.5: generate cross-category contrast rules ───────────────
        self.contrast_knowledge = ""
        known_cats = {k: v for k, v in cat_knowledge.items() if v and k != "Övrigt"}
        if len(known_cats) >= 2:
            self.progress.emit("\n=== Steg 1.5: Analyserar skillnader mellan kategorier ===")
            try:
                self.contrast_knowledge = self._generate_contrast_knowledge(cat_knowledge)
                self.contrast_ready.emit(self.contrast_knowledge)
                self.progress.emit("  ✓ Kontrastanalys klar")
            except Exception as e:
                self.progress.emit(f"  ✗ Kontrastanalys misslyckades: {e}")

        # ── Step 2: classify remaining articles ────────────────────────────
        self.progress.emit("\n=== Steg 2: Klassificerar återstående artiklar ===")
        classified_numbers = {
            e.get("article_number", "") for e in self.categorized
            if e.get("article_number")
        }
        remaining = [
            row for row in self.csv_data
            if str(row.get("article_number", "")) not in classified_numbers
        ]
        if not remaining:
            self.progress.emit("Inga återstående artiklar.")
            self.finished_all.emit()
            return

        self.progress.emit(f"  {len(remaining)} artiklar att klassificera…")
        for i, row in enumerate(remaining):
            if self._stop:
                return
            self._wait_if_paused()
            art_num = str(row.get("article_number", ""))
            url     = row.get("url", "")
            bolag   = row.get("bolag", "")
            img_path = row.get("img_path", "")

            # Download image if not already on disk
            if not img_path or not Path(img_path).exists():
                img_path = self._download_image(url)

            if not img_path:
                self.progress.emit(f"  [{i+1}/{len(remaining)}] {art_num}: bild saknas — hoppar")
                continue

            meta = self.data_mgr.get_meta(art_num, bolag) or {}
            try:
                category = self._classify_article(img_path, meta, cat_knowledge)
                self.article_classified.emit(art_num, category, url, img_path)
                if (i + 1) % 20 == 0 or i == len(remaining) - 1:
                    self.progress.emit(f"  [{i+1}/{len(remaining)}] klassificerade…")
            except Exception as e:
                self.progress.emit(f"  [{i+1}/{len(remaining)}] {art_num}: {e}")

        self.finished_all.emit()

    # ── Step 1 helper ──────────────────────────────────────────────────────────

    def _generate_knowledge(self, cat_name: str, cat_desc: str,
                            items: List[Dict]) -> str:
        """Ask LLM to summarise what's common across example articles."""
        article_lines = []
        representative_img: Optional[str] = None

        for idx, item in enumerate(items):
            art_num = str(item.get("article_number", ""))
            meta    = self.data_mgr.get_meta(art_num, "") or {} if art_num else {}
            parts   = [f"Artikel {idx + 1}:"]
            if meta.get("beskrivning"):
                parts.append(f"  Beskrivning: {meta['beskrivning']}")
            dims = []
            if meta.get("langd"): dims.append(f"längd {meta['langd']} mm")
            if meta.get("bredd"): dims.append(f"bredd {meta['bredd']} mm")
            if meta.get("hojd"):  dims.append(f"höjd {meta['hojd']} mm")
            if dims:
                parts.append(f"  Mått: {', '.join(dims)}")
            if meta.get("volym"):
                parts.append(f"  Volym: {meta['volym']}")
            vikt = []
            if meta.get("vikt_brutto"): vikt.append(f"brutto {meta['vikt_brutto']} kg")
            if meta.get("vikt_netto"):  vikt.append(f"netto {meta['vikt_netto']} kg")
            if vikt:
                parts.append(f"  Vikt: {', '.join(vikt)}")
            article_lines.append("\n".join(parts))

            if representative_img is None:
                p = item.get("image_path", "")
                if p and Path(p).exists():
                    representative_img = p

        prompt = "\n".join([
            f"Syfte: {self.syfte}", "",
            f"Kategori: {cat_name}",
            f"Beskrivning: {cat_desc}" if cat_desc else "",
            "",
            f"Nedan följer {len(items)} exempelartiklar i kategorin.",
            "\n\n".join(article_lines),
            "",
            "Sammanfatta vad som är gemensamt för artiklar i denna kategori.",
            "Fokusera på: produkttyp, typiska mått, volym, vikt och utseende.",
            "Svara på svenska med 3–5 meningar.",
        ])

        content: List[Dict] = []
        if representative_img:
            b64, mime = self._encode(representative_img)
            content.append({"type": "image_url",
                            "image_url": {"url": f"data:{mime};base64,{b64}"}})
        content.append({"type": "text", "text": prompt})

        payload = {"model": self.model,
                   "messages": [{"role": "user", "content": content}],
                   "max_tokens": 400, "temperature": 0.3}
        resp = req.post(f"{self.api_url}/chat/completions", json=payload, timeout=120)
        resp.raise_for_status()
        return resp.json()["choices"][0]["message"]["content"].strip()

    def _generate_ovrigt_knowledge(self, items: List[Dict]) -> str:
        """Generate a detailed description of what makes an article belong to Övrigt."""
        article_lines = []
        representative_imgs: List[str] = []

        for idx, item in enumerate(items):
            art_num = str(item.get("article_number", ""))
            meta    = self.data_mgr.get_meta(art_num, "") or {} if art_num else {}
            parts   = [f"Artikel {idx + 1}:"]
            if meta.get("beskrivning"):
                parts.append(f"  Beskrivning: {meta['beskrivning']}")
            dims = []
            if meta.get("langd"): dims.append(f"längd {meta['langd']} mm")
            if meta.get("bredd"): dims.append(f"bredd {meta['bredd']} mm")
            if meta.get("hojd"):  dims.append(f"höjd {meta['hojd']} mm")
            if dims:
                parts.append(f"  Mått: {', '.join(dims)}")
            article_lines.append("\n".join(parts))

            p = item.get("image_path", "")
            if p and Path(p).exists() and len(representative_imgs) < 3:
                representative_imgs.append(p)

        prompt = "\n".join([
            f"Syfte: {self.syfte}", "",
            "Kategori: Övrigt",
            "",
            f"Nedan följer {len(items)} artiklar som klassificerats som 'Övrigt' —",
            "dvs. artiklar som INTE passade in i någon annan specifik kategori.",
            "",
            "\n\n".join(article_lines),
            "",
            "Analysera dessa artiklar och beskriv:",
            "1. Vilka TYPER av artiklar som hamnar i Övrigt (t.ex. produktkategorier, storlekar, förpackningstyper).",
            "2. Vad som UTMÄRKER Övrigt-artiklar — varför passar de inte i de andra kategorierna?",
            "3. Konkreta VARNINGSSIGNALER — vilka egenskaper hos en artikel tyder på att den bör klassas som Övrigt",
            "   snarare än i en specifik kategori?",
            "",
            "Svara på svenska med 8–12 meningar. Var konkret och specifik.",
        ])

        content: List[Dict] = []
        for img_path in representative_imgs:
            try:
                b64, mime = self._encode(img_path)
                content.append({"type": "image_url",
                                "image_url": {"url": f"data:{mime};base64,{b64}"}})
            except Exception:
                pass
        content.append({"type": "text", "text": prompt})

        payload = {"model": self.model,
                   "messages": [{"role": "user", "content": content}],
                   "max_tokens": 900, "temperature": 0.3}
        resp = req.post(f"{self.api_url}/chat/completions", json=payload, timeout=180)
        resp.raise_for_status()
        return resp.json()["choices"][0]["message"]["content"].strip()

    def _generate_contrast_knowledge(self, cat_knowledge: Dict[str, str]) -> str:
        """Ask LLM to produce decision rules that distinguish categories from each other."""
        cats = [(name, desc) for name, desc in cat_knowledge.items()
                if desc and name != "Övrigt"]
        if len(cats) < 2:
            return ""
        cat_block = "\n\n".join(
            f"Kategori: {name}\nSammanfattning: {desc}"
            for name, desc in cats
        )
        prompt = "\n".join([
            f"Syfte: {self.syfte}", "",
            "Nedan följer sammanfattningar av alla produktkategorier i systemet.",
            "Din uppgift är att skriva konkreta beslutsregler som hjälper till att",
            "SKILJA kategorierna åt när en artikel kan verka passa in i flera.", "",
            cat_block, "",
            "Skriv för varje par av kategorier som kan förväxlas en regel på formen:",
            "  'Välj X framför Y om ...'",
            "Fokusera på de vanligaste förväxlingsriskerna. Max 10 regler.",
            "Svara på svenska. Var konkret — undvik vaga formuleringar.",
        ])
        payload = {"model": self.model,
                   "messages": [{"role": "user",
                                 "content": [{"type": "text", "text": prompt}]}],
                   "max_tokens": 600, "temperature": 0.3}
        resp = req.post(f"{self.api_url}/chat/completions", json=payload, timeout=120)
        resp.raise_for_status()
        return resp.json()["choices"][0]["message"]["content"].strip()

    # ── Step 2 helper ──────────────────────────────────────────────────────────

    def _classify_article(self, img_path: str, meta: Dict,
                          cat_knowledge: Dict[str, str],
                          hint: str = "") -> str:
        """Classify one article; returns Övrigt if uncertain."""
        cat_names = [c["name"] for c in self.categories if c["name"] != "Övrigt"]
        all_names = cat_names + ["Övrigt"]

        cat_block = "\n".join(
            f"- {name}: {cat_knowledge.get(name, '')}"
            if cat_knowledge.get(name)
            else f"- {name}"
            for name in cat_names
        )
        cat_block += "\n- Övrigt: Artikel som inte tydligt tillhör någon annan kategori."

        art_lines = []
        if meta.get("beskrivning"):
            art_lines.append(f"  Beskrivning: {meta['beskrivning']}")
        dims = []
        if meta.get("langd"): dims.append(f"längd {meta['langd']} mm")
        if meta.get("bredd"): dims.append(f"bredd {meta['bredd']} mm")
        if meta.get("hojd"):  dims.append(f"höjd {meta['hojd']} mm")
        if dims:
            art_lines.append(f"  Mått: {', '.join(dims)}")
        if meta.get("volym"):
            art_lines.append(f"  Volym: {meta['volym']}")
        vikt = []
        if meta.get("vikt_brutto"): vikt.append(f"brutto {meta['vikt_brutto']} kg")
        if meta.get("vikt_netto"):  vikt.append(f"netto {meta['vikt_netto']} kg")
        if vikt:
            art_lines.append(f"  Vikt: {', '.join(vikt)}")

        hint_block = (
            f"\nOBS: {hint}\n"
            if hint else ""
        )
        contrast = getattr(self, "contrast_knowledge", "")
        prompt = "\n".join([
            f"Syfte: {self.syfte}", "",
            "Klassificera artikeln nedan i en av följande kategorier.",
            "Välj 'Övrigt' om artikeln inte tydligt tillhör någon kategori.", "",
            "KATEGORIER:",
            cat_block, "",
            *(["SKILJ MELLAN KATEGORIER (använd dessa regler vid tvekan):",
               contrast, ""] if contrast else []),
            *([f"VIKTIGT SAMMANHANG:{hint_block}"] if hint else []),
            "ARTIKEL ATT KLASSIFICERA:",
            "\n".join(art_lines) if art_lines else "  (ingen metadata)",
            "",
            f"Svara ENDAST med exakt ett av dessa namn: {', '.join(all_names)}",
            "Inget annat — bara kategorinamnet.",
        ])

        b64, mime = self._encode(img_path)
        content = [
            {"type": "image_url", "image_url": {"url": f"data:{mime};base64,{b64}"}},
            {"type": "text", "text": prompt},
        ]
        payload = {"model": self.model,
                   "messages": [{"role": "user", "content": content}],
                   "max_tokens": 30, "temperature": 0.1}
        resp = req.post(f"{self.api_url}/chat/completions", json=payload, timeout=60)
        resp.raise_for_status()
        raw = resp.json()["choices"][0]["message"]["content"].strip()
        raw_lower = raw.lower()
        for name in all_names:
            if name.lower() in raw_lower:
                return name
        return "Övrigt"

    # ── utilities ──────────────────────────────────────────────────────────────

    def _download_image(self, url: str) -> Optional[str]:
        if not url:
            return None
        try:
            suffix = Path(url.split("?")[0]).suffix.lower()
            if suffix not in SUPPORTED_EXT:
                suffix = ".jpg"
            resp = req.get(url, timeout=30)
            resp.raise_for_status()
            tmp = tempfile.NamedTemporaryFile(delete=False, suffix=suffix)
            tmp.write(resp.content)
            tmp.close()
            return tmp.name
        except Exception:
            return None

    def _encode(self, path: str) -> Tuple[str, str]:
        suffix   = Path(path).suffix.lower()
        mime_map = {".jpg": "image/jpeg", ".jpeg": "image/jpeg", ".png": "image/png",
                    ".webp": "image/webp", ".gif": "image/gif", ".bmp": "image/bmp"}
        mime = mime_map.get(suffix, "image/jpeg")
        if self.compress and PIL_AVAILABLE:
            img = PILImage.open(path)
            img.thumbnail((600, 600), PILImage.LANCZOS)
            buf = BytesIO()
            img.save(buf, format="JPEG", quality=80)
            return base64.b64encode(buf.getvalue()).decode(), "image/jpeg"
        with open(path, "rb") as f:
            return base64.b64encode(f.read()).decode(), mime


# ── ImageDownloader ────────────────────────────────────────────────────────────
class ImageDownloader(QThread):
    image_ready = pyqtSignal(int, str)   # (index, local_path)

    def __init__(self, rows, temp_dir, parent=None):
        super().__init__(parent)
        self.rows     = rows
        self.temp_dir = temp_dir
        self._stop    = False

    def stop(self):
        self._stop = True

    def run(self):
        for i, row in enumerate(self.rows):
            if self._stop:
                break
            dest = self._download(i, row)
            if dest:
                self.image_ready.emit(i, str(dest))

    def _download(self, i: int, row: Dict) -> Optional[Path]:
        url      = row["url"]
        url_path = url.split("?")[0].rstrip("/")
        filename = url_path.split("/")[-1] or f"img_{i+1}"
        if not Path(filename).suffix:
            filename += ".jpg"
        dest = Path(self.temp_dir) / f"{i:05d}_{filename}"
        try:
            r = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"})
            with urllib.request.urlopen(r, timeout=15) as resp:
                dest.write_bytes(resp.read())
            return dest
        except Exception:
            return None


# ── HeaderBar ──────────────────────────────────────────────────────────────────
class HeaderBar(QFrame):
    def __init__(self, test_name: str = "", right_text: str = "", parent=None):
        super().__init__(parent)
        self.setStyleSheet("background-color:#181825; border-bottom:1px solid #313244;")
        self.setFixedHeight(48)
        lay = QHBoxLayout(self)
        lay.setContentsMargins(16, 0, 16, 0)
        self._left = QLabel(f"Test: {test_name}" if test_name else "Bildklassificering")
        self._left.setStyleSheet("font-size:15px; font-weight:bold; color:#89b4fa;")
        lay.addWidget(self._left)
        lay.addStretch()
        self._right = QLabel(right_text)
        self._right.setStyleSheet("font-size:12px; color:#6c7086;")
        lay.addWidget(self._right)

    def set_texts(self, left: str, right: str = ""):
        self._left.setText(left)
        self._right.setText(right)


# ══════════════════════════════════════════════════════════ Screen 1: Name ══════
class NameScreen(QWidget):
    go_next    = pyqtSignal(str, str)   # (test_name, syfte)
    load_zip   = pyqtSignal()

    def __init__(self, parent=None):
        super().__init__(parent)
        lay = QVBoxLayout(self)
        lay.setContentsMargins(0, 0, 0, 0)
        lay.addStretch()

        card = QFrame()
        card.setStyleSheet(
            "background-color:#313244; border-radius:14px;"
            "border: 1px solid #45475a;"
        )
        card.setFixedWidth(460)
        c = QVBoxLayout(card)
        c.setContentsMargins(36, 32, 36, 32)
        c.setSpacing(0)

        # ── Header ────────────────────────────────────────────────────────
        title = QLabel("Bildklassificering")
        title.setStyleSheet("font-size:24px; font-weight:bold; color:#89b4fa;"
                            "border:none;")
        title.setAlignment(Qt.AlignmentFlag.AlignCenter)
        c.addWidget(title)

        sub = QLabel("Skapa ett nytt klassificeringstest")
        sub.setStyleSheet("font-size:11px; color:#6c7086; border:none;")
        sub.setAlignment(Qt.AlignmentFlag.AlignCenter)
        c.addWidget(sub)
        c.addSpacing(24)

        # ── Namn ──────────────────────────────────────────────────────────
        lbl_name = QLabel("Namn på testet")
        lbl_name.setStyleSheet("font-size:11px; font-weight:600; color:#a6adc8;"
                               "border:none;")
        c.addWidget(lbl_name)
        c.addSpacing(4)
        self.name_edit = QLineEdit()
        self.name_edit.setPlaceholderText("t.ex. Testomgång 1")
        self.name_edit.setFixedHeight(36)
        c.addWidget(self.name_edit)
        c.addSpacing(16)

        # ── Syfte ─────────────────────────────────────────────────────────
        lbl_syfte = QLabel("Syfte med testet")
        lbl_syfte.setStyleSheet("font-size:11px; font-weight:600; color:#a6adc8;"
                                "border:none;")
        c.addWidget(lbl_syfte)
        c.addSpacing(4)
        hint = QLabel("AI:n använder detta för att förstå sammanhanget")
        hint.setStyleSheet("font-size:10px; color:#585b70; border:none;")
        c.addWidget(hint)
        c.addSpacing(4)
        self.syfte_edit = QTextEdit()
        self.syfte_edit.setPlaceholderText(
            'T.ex. "Kategorisera lagerartiklar för att förenkla lagerhållning.\n'
            'Fokus på att skilja farligt gods från övrigt."'
        )
        self.syfte_edit.setFixedHeight(90)
        c.addWidget(self.syfte_edit)
        c.addSpacing(24)

        # ── Button ────────────────────────────────────────────────────────
        go = mk_btn("Gå vidare  →", "#89b4fa", "#1e1e2e", h=40)
        go.clicked.connect(self._validate)
        c.addWidget(go)
        self.name_edit.returnPressed.connect(self._validate)

        sep = QFrame()
        sep.setFrameShape(QFrame.Shape.HLine)
        sep.setStyleSheet("color:#45475a; border:none; border-top:1px solid #45475a;")
        c.addSpacing(12)
        c.addWidget(sep)
        c.addSpacing(8)

        load_btn = mk_btn("📦  Öppna sparad session (.zip)", "#313244", "#585b70", h=36)
        load_btn.setStyleSheet(
            load_btn.styleSheet() +
            "border:1px solid #45475a; font-size:11px;"
        )
        load_btn.clicked.connect(self.load_zip.emit)
        c.addWidget(load_btn)

        lay.addWidget(card, 0, Qt.AlignmentFlag.AlignHCenter)
        lay.addStretch()

    def _validate(self):
        name  = self.name_edit.text().strip()
        syfte = self.syfte_edit.toPlainText().strip()
        if not name:
            QMessageBox.warning(self, "Fel", "Ange ett namn för testet.")
            return
        safe = "".join(c for c in name if c not in r'\/:*?"<>|').strip()
        if not safe:
            QMessageBox.warning(self, "Fel", "Namnet innehåller ogiltiga tecken.")
            return
        self.go_next.emit(safe, syfte)

    def reset(self):
        self.name_edit.clear()
        self.syfte_edit.clear()
        self.name_edit.setFocus()


# ══════════════════════════════════════════════════════ Screen 2: Categories ═══
class CategoryRow(QFrame):
    removed = pyqtSignal(object)

    def __init__(self, number: int, parent=None):
        super().__init__(parent)
        self.setStyleSheet("background:transparent;")
        lay = QHBoxLayout(self)
        lay.setContentsMargins(0, 0, 0, 0)
        lay.setSpacing(6)

        self.num_lbl = QLabel(f"{number}.")
        self.num_lbl.setFixedWidth(24)
        self.num_lbl.setStyleSheet("color:#6c7086;")
        lay.addWidget(self.num_lbl)

        self.name_edit = QLineEdit()
        self.name_edit.setPlaceholderText("Kategorinamn")
        self.name_edit.setFixedWidth(190)
        self.name_edit.setFixedHeight(34)
        lay.addWidget(self.name_edit)

        self.desc_edit = QLineEdit()
        self.desc_edit.setPlaceholderText("Beskrivning (valfritt — hjälper AI:n)")
        self.desc_edit.setFixedHeight(34)
        lay.addWidget(self.desc_edit)

        rm = QPushButton("✕")
        rm.setFixedSize(30, 30)
        rm.setStyleSheet("background:#f38ba8; color:#1e1e2e; border-radius:4px; font-weight:bold;")
        rm.clicked.connect(lambda: self.removed.emit(self))
        lay.addWidget(rm)

    def set_number(self, n: int):
        self.num_lbl.setText(f"{n}.")

    def get_data(self) -> Tuple[str, str]:
        return self.name_edit.text().strip(), self.desc_edit.text().strip()

    def is_empty(self) -> bool:
        return not self.name_edit.text().strip()


class CategoriesScreen(QWidget):
    go_next = pyqtSignal(list)   # [{name, description}]
    go_back = pyqtSignal()

    def __init__(self, parent=None):
        super().__init__(parent)
        self._rows: List[CategoryRow] = []

        outer = QVBoxLayout(self)
        outer.setContentsMargins(0, 0, 0, 0)
        outer.setSpacing(0)

        self.header = HeaderBar()
        outer.addWidget(self.header)

        body = QWidget()
        body_lay = QVBoxLayout(body)
        body_lay.setContentsMargins(48, 24, 48, 24)
        body_lay.setSpacing(10)

        title = QLabel("Kategorier")
        title.setStyleSheet("font-size:22px; font-weight:bold;")
        body_lay.addWidget(title)

        hint = QLabel(
            '"Övrigt" läggs alltid till automatiskt. '
            'Beskrivningarna är valfria men hjälper AI:n att gissa rätt.'
        )
        hint.setStyleSheet("color:#6c7086; font-size:12px;")
        hint.setWordWrap(True)
        body_lay.addWidget(hint)

        # Column headers
        col_hdr = QFrame()
        col_hdr.setStyleSheet("background:transparent;")
        ch = QHBoxLayout(col_hdr)
        ch.setContentsMargins(0, 0, 0, 0)
        ch.setSpacing(6)
        spacer = QLabel(); spacer.setFixedWidth(24); ch.addWidget(spacer)
        lbl_n = QLabel("Namn"); lbl_n.setStyleSheet("color:#6c7086; font-size:12px;"); lbl_n.setFixedWidth(190)
        ch.addWidget(lbl_n)
        lbl_d = QLabel("Beskrivning (hjälper AI:n)"); lbl_d.setStyleSheet("color:#6c7086; font-size:12px;")
        ch.addWidget(lbl_d)
        ch.addStretch()
        body_lay.addWidget(col_hdr)

        # Scrollable rows area
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setStyleSheet("background:transparent;")
        self.rows_widget = QWidget()
        self.rows_widget.setStyleSheet("background:transparent;")
        self.rows_lay = QVBoxLayout(self.rows_widget)
        self.rows_lay.setContentsMargins(0, 0, 0, 0)
        self.rows_lay.setSpacing(4)
        self.rows_lay.addStretch()
        scroll.setWidget(self.rows_widget)
        body_lay.addWidget(scroll, 1)

        for _ in range(3):
            self._add_row()

        btn_row = QHBoxLayout()
        add_btn = mk_btn("+ Lägg till rad", "#313244", "#cdd6f4")
        add_btn.clicked.connect(self._add_row)
        btn_row.addWidget(add_btn)
        btn_row.addStretch()
        back_btn = mk_btn("← Tillbaka", "#45475a", "#cdd6f4")
        back_btn.clicked.connect(self.go_back.emit)
        btn_row.addWidget(back_btn)
        next_btn = mk_btn("Starta klassificering  →", "#89b4fa", "#1e1e2e")
        next_btn.clicked.connect(self._validate)
        btn_row.addWidget(next_btn)
        body_lay.addLayout(btn_row)

        outer.addWidget(body)

    def _add_row(self):
        row = CategoryRow(len(self._rows) + 1)
        row.removed.connect(self._remove_row)
        self._rows.append(row)
        self.rows_lay.insertWidget(self.rows_lay.count() - 1, row)
        row.name_edit.setFocus()

    def _remove_row(self, row: CategoryRow):
        self._rows.remove(row)
        row.setParent(None)
        for i, r in enumerate(self._rows):
            r.set_number(i + 1)

    def _validate(self):
        cats = [{"name": n, "description": d}
                for r in self._rows
                for n, d in [r.get_data()] if n]
        if not cats:
            QMessageBox.warning(self, "Fel", "Ange minst en kategori.")
            return
        self.go_next.emit(cats)

    def set_test_name(self, name: str):
        self.header.set_texts(f"Test: {name}")


# ══════════════════════════════════════════════════════════ Screen 3: Source ════
class SourceScreen(QWidget):
    use_folder  = pyqtSignal()
    use_builtin = pyqtSignal()
    use_csv     = pyqtSignal()
    go_back     = pyqtSignal()

    def __init__(self, test_name: str, n_builtin: int, parent=None):
        super().__init__(parent)
        outer = QVBoxLayout(self)
        outer.setContentsMargins(0, 0, 0, 0)
        outer.addWidget(HeaderBar(test_name))

        center = QWidget()
        c = QVBoxLayout(center)
        c.setAlignment(Qt.AlignmentFlag.AlignCenter)

        card = QFrame()
        card.setStyleSheet("background-color:#313244; border-radius:12px;")
        card.setFixedWidth(420)
        cl = QVBoxLayout(card)
        cl.setContentsMargins(32, 32, 32, 32)
        cl.setSpacing(10)

        title = QLabel("Välj bildkälla")
        title.setStyleSheet("font-size:22px; font-weight:bold;")
        title.setAlignment(Qt.AlignmentFlag.AlignCenter)
        cl.addWidget(title)
        sub = QLabel("Varifrån ska bilderna hämtas?")
        sub.setStyleSheet("color:#6c7086;")
        sub.setAlignment(Qt.AlignmentFlag.AlignCenter)
        cl.addWidget(sub)
        cl.addSpacing(8)

        b1 = mk_btn("📁  Från mapp  (bilder/)", "#2196F3", h=48)
        b1.clicked.connect(self.use_folder.emit)
        cl.addWidget(b1)

        if n_builtin:
            b2 = mk_btn(f"📊  Inbyggd data  ({n_builtin} artiklar)", "#4CAF50", h=48)
            b2.clicked.connect(self.use_builtin.emit)
            cl.addWidget(b2)

        b3 = mk_btn("📄  Ladda upp CSV-fil", "#9C27B0", h=48)
        b3.clicked.connect(self.use_csv.emit)
        cl.addWidget(b3)

        cl.addSpacing(4)
        back = mk_btn("← Tillbaka", "#45475a", "#cdd6f4")
        back.clicked.connect(self.go_back.emit)
        cl.addWidget(back)

        c.addWidget(card)
        outer.addWidget(center)


# ════════════════════════════════════════════════════ Screen 3b: AI Settings ════
class AISettingsScreen(QWidget):
    go_next = pyqtSignal(dict)   # {model, api_url, compress_images} — empty dict = skip AI
    go_back = pyqtSignal()

    def __init__(self, test_name: str, parent=None):
        super().__init__(parent)
        outer = QVBoxLayout(self)
        outer.setContentsMargins(0, 0, 0, 0)
        outer.addWidget(HeaderBar(test_name))

        center = QWidget()
        c = QVBoxLayout(center)
        c.setAlignment(Qt.AlignmentFlag.AlignCenter)

        card = QFrame()
        card.setStyleSheet("background-color:#313244; border-radius:12px;")
        card.setFixedWidth(480)
        cl = QVBoxLayout(card)
        cl.setContentsMargins(36, 36, 36, 36)
        cl.setSpacing(12)

        title = QLabel("AI-inställningar")
        title.setStyleSheet("font-size:22px; font-weight:bold;")
        title.setAlignment(Qt.AlignmentFlag.AlignCenter)
        cl.addWidget(title)

        sub = QLabel(
            "Konfigurera LM Studio. Lämna fälten oförändrade för att använda standardvärden."
        )
        sub.setStyleSheet("color:#6c7086; font-size:12px;")
        sub.setWordWrap(True)
        sub.setAlignment(Qt.AlignmentFlag.AlignCenter)
        cl.addWidget(sub)
        cl.addSpacing(4)

        cl.addWidget(QLabel("LM Studio URL:"))
        self.url_edit = QLineEdit(DEFAULT_AI_URL)
        cl.addWidget(self.url_edit)

        cl.addWidget(QLabel("Modellnamn:"))
        self.model_edit = QLineEdit(DEFAULT_MODEL)
        cl.addWidget(self.model_edit)

        self.compress_cb = QCheckBox("Komprimera bilder (snabbare, marginellt sämre precision)")
        self.compress_cb.setChecked(True)
        cl.addWidget(self.compress_cb)

        cl.addSpacing(8)
        go = mk_btn("Använd AI  →", "#89b4fa", "#1e1e2e", h=44)
        go.clicked.connect(self._go)
        cl.addWidget(go)

        skip = mk_btn("Hoppa över AI", "#45475a", "#cdd6f4")
        skip.clicked.connect(lambda: self.go_next.emit({}))
        cl.addWidget(skip)

        back = mk_btn("← Tillbaka", "#45475a", "#cdd6f4")
        back.clicked.connect(self.go_back.emit)
        cl.addWidget(back)

        c.addWidget(card)
        outer.addWidget(center)

    def _go(self):
        self.go_next.emit({
            "api_url":         self.url_edit.text().strip() or DEFAULT_AI_URL,
            "model":           self.model_edit.text().strip() or DEFAULT_MODEL,
            "compress_images": self.compress_cb.isChecked(),
        })


# ═══════════════════════════════════════════════════════ Screen 3c: Filter ══════
class FilterScreen(QWidget):
    go_next = pyqtSignal(list)   # filtered rows
    go_back = pyqtSignal()

    def __init__(self, test_name: str, rows: List[Dict], data_mgr, parent=None):
        super().__init__(parent)
        self._all_rows = rows
        self._data_mgr = data_mgr

        # Pre-compute per-row metadata for fast filtering
        self._row_meta: List[Dict] = []
        for r in rows:
            meta = data_mgr.get_meta(str(r["article_number"]), r.get("bolag", "")) or {}
            self._row_meta.append({
                "bolag":       r.get("bolag", "") or "–",
                "hkat":        meta.get("huvudkategori", "") or "Okänd",
                "robot":       meta.get("robot", "N").upper() or "N",
            })

        bolags  = sorted({m["bolag"] for m in self._row_meta})
        hkats   = sorted({m["hkat"]  for m in self._row_meta})

        outer = QVBoxLayout(self)
        outer.setContentsMargins(0, 0, 0, 0)
        outer.addWidget(HeaderBar(test_name))

        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)

        content = QWidget()
        cl = QVBoxLayout(content)
        cl.setContentsMargins(40, 32, 40, 32)
        cl.setSpacing(20)

        title = QLabel("Filtrera artiklar")
        title.setStyleSheet("font-size:22px; font-weight:bold;")
        cl.addWidget(title)

        self._total_lbl = QLabel()
        self._total_lbl.setStyleSheet("color:#6c7086;")
        cl.addWidget(self._total_lbl)

        cl.addWidget(sep())

        # ── Bolag ────────────────────────────────────────────────────────────
        cl.addWidget(self._section_label("Bolag"))
        self._bolag_cbs: List[QCheckBox] = []
        bolag_all = QCheckBox("Alla bolag")
        bolag_all.setChecked(True)
        bolag_all.setStyleSheet("font-weight:bold;")
        cl.addWidget(bolag_all)
        bolag_grid = QWidget()
        bg = QGridLayout(bolag_grid)
        bg.setContentsMargins(16, 0, 0, 0)
        bg.setHorizontalSpacing(16)
        bg.setVerticalSpacing(4)
        for i, b in enumerate(bolags):
            cb = QCheckBox(b)
            cb.setChecked(True)
            cb.stateChanged.connect(self._update_count)
            self._bolag_cbs.append(cb)
            bg.addWidget(cb, i // 3, i % 3)
        cl.addWidget(bolag_grid)

        def _toggle_bolags(state):
            checked = state == Qt.CheckState.Checked.value
            for cb in self._bolag_cbs:
                cb.blockSignals(True)
                cb.setChecked(checked)
                cb.blockSignals(False)
            self._update_count()
        bolag_all.stateChanged.connect(_toggle_bolags)

        cl.addWidget(sep())

        # ── Huvudkategori ─────────────────────────────────────────────────────
        cl.addWidget(self._section_label("Huvudkategori"))
        self._hkat_cbs: List[QCheckBox] = []
        hkat_all = QCheckBox("Alla kategorier")
        hkat_all.setChecked(True)
        hkat_all.setStyleSheet("font-weight:bold;")
        cl.addWidget(hkat_all)
        hkat_grid = QWidget()
        hg = QGridLayout(hkat_grid)
        hg.setContentsMargins(16, 0, 0, 0)
        hg.setHorizontalSpacing(16)
        hg.setVerticalSpacing(4)
        for i, h in enumerate(hkats):
            cb = QCheckBox(h)
            cb.setChecked(True)
            cb.stateChanged.connect(self._update_count)
            self._hkat_cbs.append(cb)
            hg.addWidget(cb, i // 2, i % 2)
        cl.addWidget(hkat_grid)

        def _toggle_hkats(state):
            checked = state == Qt.CheckState.Checked.value
            for cb in self._hkat_cbs:
                cb.blockSignals(True)
                cb.setChecked(checked)
                cb.blockSignals(False)
            self._update_count()
        hkat_all.stateChanged.connect(_toggle_hkats)

        cl.addWidget(sep())

        # ── Robot ─────────────────────────────────────────────────────────────
        cl.addWidget(self._section_label("Robotartikel"))
        robot_row = QHBoxLayout()
        robot_row.setSpacing(20)
        self._robot_group = QButtonGroup(self)
        for i, (lbl, val) in enumerate([("Alla", "alla"), ("Ja (Y)", "Y"), ("Nej (N)", "N")]):
            rb = QRadioButton(lbl)
            rb.setProperty("robot_val", val)
            if i == 0:
                rb.setChecked(True)
            rb.toggled.connect(self._update_count)
            self._robot_group.addButton(rb, i)
            robot_row.addWidget(rb)
        robot_row.addStretch()
        cl.addLayout(robot_row)

        cl.addWidget(sep())

        # ── Artikelnummer (valfri lista) ───────────────────────────────────────
        cl.addWidget(self._section_label("Begränsa till artikelnummer (valfritt)"))
        art_hint = QLabel(
            "Klistra in ett artikelnummer per rad. Lämnas tomt används alla artiklar."
        )
        art_hint.setStyleSheet("color:#6c7086; font-size:11px;")
        cl.addWidget(art_hint)
        self._art_filter = QTextEdit()
        self._art_filter.setPlaceholderText("artikel1\nartikel2\nartikel3")
        self._art_filter.setFixedHeight(100)
        self._art_filter.setStyleSheet(
            "background:#11111b; color:#cdd6f4; font-family:monospace;"
            "border:1px solid #45475a; border-radius:4px;"
        )
        self._art_filter.textChanged.connect(self._update_count)
        cl.addWidget(self._art_filter)

        cl.addWidget(sep())

        # ── match count ───────────────────────────────────────────────────────
        self._match_lbl = QLabel()
        self._match_lbl.setStyleSheet("font-size:14px; font-weight:bold; color:#a6e3a1;")
        cl.addWidget(self._match_lbl)

        # ── buttons ───────────────────────────────────────────────────────────
        btn_row = QHBoxLayout()
        back_btn = mk_btn("← Tillbaka", "#45475a", "#cdd6f4")
        back_btn.clicked.connect(self.go_back.emit)
        btn_row.addWidget(back_btn)
        btn_row.addStretch()
        self._start_btn = mk_btn("Starta  →", "#89b4fa", "#1e1e2e", h=44)
        self._start_btn.clicked.connect(self._on_start)
        btn_row.addWidget(self._start_btn)
        cl.addLayout(btn_row)

        cl.addStretch()
        scroll.setWidget(content)
        outer.addWidget(scroll)

        self._update_count()

    # ── helpers ────────────────────────────────────────────────────────────────
    def _section_label(self, text: str) -> QLabel:
        lbl = QLabel(text)
        lbl.setStyleSheet("font-size:14px; font-weight:bold; color:#89b4fa;")
        return lbl

    def _selected_bolags(self) -> Optional[set]:
        sel = {cb.text() for cb in self._bolag_cbs if cb.isChecked()}
        return None if len(sel) == len(self._bolag_cbs) else sel

    def _selected_hkats(self) -> Optional[set]:
        sel = {cb.text() for cb in self._hkat_cbs if cb.isChecked()}
        return None if len(sel) == len(self._hkat_cbs) else sel

    def _robot_filter(self) -> str:
        checked = self._robot_group.checkedButton()
        return checked.property("robot_val") if checked else "alla"

    def _art_number_filter(self) -> Optional[set]:
        text = self._art_filter.toPlainText().strip()
        if not text:
            return None
        return {line.strip() for line in text.splitlines() if line.strip()}

    def _filtered_rows(self) -> List[Dict]:
        bolags   = self._selected_bolags()
        hkats    = self._selected_hkats()
        robot    = self._robot_filter()
        art_nums = self._art_number_filter()
        result = []
        for row, meta in zip(self._all_rows, self._row_meta):
            if art_nums and str(row.get("article_number", "")) not in art_nums:
                continue
            if bolags and meta["bolag"] not in bolags:
                continue
            if hkats and meta["hkat"] not in hkats:
                continue
            if robot != "alla" and meta["robot"] != robot:
                continue
            result.append(row)
        return result

    def _update_count(self):
        n = len(self._filtered_rows())
        total = len(self._all_rows)
        self._total_lbl.setText(f"Totalt {total} artiklar i källan")
        self._match_lbl.setText(f"{n} artikel{'er' if n != 1 else ''} matchar filtret")
        self._start_btn.setEnabled(n > 0)

    def _on_start(self):
        self.go_next.emit(self._filtered_rows())


# ══════════════════════════════════════════════════════ Screen 4: Classify ══════
class ClassifyScreen(QWidget):
    classified   = pyqtSignal(str)
    skipped      = pyqtSignal()
    go_back      = pyqtSignal()
    add_category = pyqtSignal()
    end_test     = pyqtSignal()
    run_ai_job   = pyqtSignal()

    def __init__(self, parent=None):
        super().__init__(parent)
        self._shortcuts: List[QShortcut] = []
        self._inner: Optional[QWidget]   = None
        self._main_lay = QVBoxLayout(self)
        self._main_lay.setContentsMargins(0, 0, 0, 0)
        self._main_lay.setSpacing(0)

    def show_image(self, test_name: str, categories: List[Dict],
                   image_path: str, meta: Optional[Dict],
                   current: int, total: int,
                   cat_counts: Optional[Dict[str, int]] = None,
                   threshold: int = 0,
                   ai_job_ready: bool = False,
                   prev_category: str = ""):
        self._clear()
        self._test_name    = test_name
        self._categories   = categories
        self._image_path   = image_path
        self._meta         = meta
        self._current      = current
        self._total        = total
        self._cat_counts   = cat_counts or {}
        self._threshold    = threshold
        self._ai_job_ready = ai_job_ready
        self._prev_category = prev_category
        self._build()

    def _clear(self):
        for sc in self._shortcuts:
            sc.setEnabled(False)
            sc.deleteLater()
        self._shortcuts.clear()
        if self._inner:
            self._main_lay.removeWidget(self._inner)
            self._inner.setParent(None)
            self._inner = None

    def _build(self):
        self._inner = QWidget()
        inner_lay = QVBoxLayout(self._inner)
        inner_lay.setContentsMargins(0, 0, 0, 0)
        inner_lay.setSpacing(0)

        # ── header
        prog = f"Bild {self._current + 1} av {self._total}"
        header = HeaderBar(self._test_name, prog)
        inner_lay.addWidget(header)

        # ── threshold progress bar (shown when AI settings configured)
        if self._threshold > 0:
            inner_lay.addWidget(self._build_threshold_bar())

        # ── image + meta
        content = QFrame()
        content.setStyleSheet("background-color:#11111b;")
        content_lay = QHBoxLayout(content)
        content_lay.setContentsMargins(0, 0, 0, 0)
        content_lay.setSpacing(0)

        self._img_lbl = QLabel()
        self._img_lbl.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self._img_lbl.setStyleSheet("background-color:#11111b;")
        content_lay.addWidget(self._img_lbl, 1)

        if self._meta:
            content_lay.addWidget(self._build_meta_panel())

        inner_lay.addWidget(content, 1)
        self._load_image()

        # ── file info bar
        info_bar = QFrame()
        info_bar.setStyleSheet("background:#181825; border-top:1px solid #313244;")
        info_bar.setFixedHeight(26)
        ib = QHBoxLayout(info_bar)
        ib.setContentsMargins(12, 0, 12, 0)
        ib.addWidget(QLabel(str(self._image_path)))
        inner_lay.addWidget(info_bar)

        # ── category buttons
        cat_frame = QFrame()
        cat_frame.setStyleSheet("background:#1e1e2e;")
        cf = QVBoxLayout(cat_frame)
        cf.setContentsMargins(12, 8, 12, 4)
        self._build_cat_buttons(cf)
        inner_lay.addWidget(cat_frame)

        # ── control bar
        ctrl = QFrame()
        ctrl.setStyleSheet("background:#1e1e2e; border-top:1px solid #313244;")
        ctrl_lay = QHBoxLayout(ctrl)
        ctrl_lay.setContentsMargins(12, 6, 12, 6)

        back_btn = mk_btn("← Tillbaka", "#45475a", "#cdd6f4")
        back_btn.setEnabled(self._current > 0)
        back_btn.clicked.connect(self.go_back.emit)
        ctrl_lay.addWidget(back_btn)

        skip_btn = mk_btn("Hoppa över  →", "#45475a", "#cdd6f4")
        skip_btn.clicked.connect(self.skipped.emit)
        ctrl_lay.addWidget(skip_btn)

        add_btn = mk_btn("+ Ny kategori", "#FF9800")
        add_btn.clicked.connect(self.add_category.emit)
        ctrl_lay.addWidget(add_btn)
        ctrl_lay.addStretch()

        if self._prev_category:
            prev_lbl = QLabel(f"Klassificerades som: {self._prev_category}")
            prev_lbl.setStyleSheet("color:#fab387; font-size:11px; font-style:italic;")
            ctrl_lay.addWidget(prev_lbl)

        if self._ai_job_ready:
            ai_btn = mk_btn("🤖  Kör AI jobb", "#1e3a5f", "#89b4fa", h=34)
            ai_btn.clicked.connect(self.run_ai_job.emit)
            ctrl_lay.addWidget(ai_btn)
        end_btn = mk_btn("Avsluta test", "#f38ba8", "#1e1e2e")
        end_btn.clicked.connect(self._confirm_end)
        ctrl_lay.addWidget(end_btn)
        inner_lay.addWidget(ctrl)

        # ── arrow shortcuts
        sc_back = QShortcut(QKeySequence(Qt.Key.Key_Left), self)
        if self._current > 0:
            sc_back.activated.connect(self.go_back.emit)
        else:
            sc_back.setEnabled(False)
        self._shortcuts.append(sc_back)

        sc_skip = QShortcut(QKeySequence(Qt.Key.Key_Right), self)
        sc_skip.activated.connect(self.skipped.emit)
        self._shortcuts.append(sc_skip)

        self._main_lay.addWidget(self._inner)

    def _build_threshold_bar(self) -> QFrame:
        """Shows per-category progress toward the AI job threshold."""
        bar = QFrame()
        bar.setStyleSheet("background:#181825; border-bottom:1px solid #313244;")
        bar.setFixedHeight(30)
        lay = QHBoxLayout(bar)
        lay.setContentsMargins(12, 0, 12, 0)
        lay.setSpacing(16)
        non_ovrigt = [c for c in self._categories if c["name"] != "Övrigt"]
        for cat in non_ovrigt:
            name  = cat["name"]
            count = self._cat_counts.get(name, 0)
            done  = count >= self._threshold
            color = "#a6e3a1" if done else "#f38ba8"
            lbl = QLabel(f"{name}: {count}/{self._threshold}")
            lbl.setStyleSheet(
                f"color:{color}; font-size:11px; font-weight:{'bold' if done else 'normal'};"
            )
            lay.addWidget(lbl)
        lay.addStretch()
        if self._ai_job_ready:
            hint = QLabel("Alla kategorier klara — klicka 'Kör AI jobb'")
            hint.setStyleSheet("color:#89b4fa; font-size:11px; font-style:italic;")
            lay.addWidget(hint)
        return bar

    def _build_meta_panel(self) -> QFrame:
        panel = QFrame()
        panel.setFixedWidth(220)
        panel.setStyleSheet("background:#181825; border-left:1px solid #313244;")
        lay = QVBoxLayout(panel)
        lay.setContentsMargins(12, 12, 12, 12)
        lay.setSpacing(5)

        title = QLabel("Artikelinfo")
        title.setStyleSheet("font-size:12px; font-weight:bold; color:#6c7086;")
        lay.addWidget(title)
        lay.addWidget(sep())

        fields = [
            ("Beskrivning",   self._meta.get("beskrivning")),
            ("Huvudkategori", self._meta.get("huvudkategori")),
            ("Kategori",      self._meta.get("kategori")),
            ("UN nummer",     self._meta.get("un_nummer")),
            ("StoreQuantity", self._meta.get("store_quantity")),
            ("Robot",         self._meta.get("robot")),
            ("Vikt brutto",   self._meta.get("vikt_brutto")),
            ("Vikt netto",    self._meta.get("vikt_netto")),
            ("Volym",         self._meta.get("volym")),
            ("EAN",           self._meta.get("ean")),
            ("Längd",         self._meta.get("langd")),
            ("Bredd",         self._meta.get("bredd")),
            ("Höjd",          self._meta.get("hojd")),
        ]
        for label, value in fields:
            if not value or value in _EMPTY:
                continue
            row = QFrame(); row.setStyleSheet("background:transparent;")
            rl = QHBoxLayout(row); rl.setContentsMargins(0, 0, 0, 0); rl.setSpacing(4)
            lbl_w = QLabel(f"{label}:"); lbl_w.setStyleSheet("color:#6c7086; font-size:11px;"); lbl_w.setFixedWidth(82)
            val_w = QLabel(str(value)); val_w.setStyleSheet("color:#cdd6f4; font-size:11px;"); val_w.setWordWrap(True)
            rl.addWidget(lbl_w); rl.addWidget(val_w, 1)
            lay.addWidget(row)

        lay.addStretch()
        return panel

    def _build_cat_buttons(self, parent_lay: QVBoxLayout):
        key_map: Dict[int, Tuple[str, str]] = {}
        for i, cat in enumerate(self._categories[:9]):
            key_map[i + 1] = (cat["name"], CATEGORY_COLORS[i % len(CATEGORY_COLORS)])
        key_map[0] = ("Övrigt", "#45475a")

        positions = {
            7: (0,0), 8: (0,1), 9: (0,2),
            4: (1,0), 5: (1,1), 6: (1,2),
            1: (2,0), 2: (2,1), 3: (2,2),
            0: (3,1),
        }
        grid_w = QWidget(); grid_w.setStyleSheet("background:transparent;")
        grid = QGridLayout(grid_w); grid.setSpacing(4)

        for key, (row, col) in positions.items():
            if key not in key_map:
                continue
            name, color = key_map[key]
            b = QPushButton(f"{name}  ({key})")
            b.setFixedSize(168, 40)
            b.setStyleSheet(
                f"background:{color}; color:white; border-radius:6px; "
                f"font-weight:bold; border:none;"
            )
            b.clicked.connect(lambda checked, c=name: self.classified.emit(c))
            grid.addWidget(b, row, col, Qt.AlignmentFlag.AlignCenter)

            sc = QShortcut(QKeySequence(str(key)), self)
            sc.activated.connect(lambda c=name: self.classified.emit(c))
            self._shortcuts.append(sc)

        parent_lay.addWidget(grid_w, 0, Qt.AlignmentFlag.AlignCenter)

    def _load_image(self):
        try:
            if PIL_AVAILABLE:
                img = PILImage.open(self._image_path)
                img.thumbnail((780, 370), PILImage.LANCZOS)
                buf = BytesIO(); img.save(buf, format="PNG"); buf.seek(0)
                px = QPixmap(); px.loadFromData(buf.read())
            else:
                px = QPixmap(self._image_path)
                px = px.scaled(780, 370, Qt.AspectRatioMode.KeepAspectRatio,
                               Qt.TransformationMode.SmoothTransformation)
            self._img_lbl.setPixmap(px)
        except Exception as e:
            self._img_lbl.setText(f"Kunde inte visa bild:\n{e}")
            self._img_lbl.setStyleSheet("color:#f38ba8;")

    def _confirm_end(self):
        if QMessageBox.question(self, "Avsluta", "Vill du avsluta testet?") == \
                QMessageBox.StandardButton.Yes:
            self.end_test.emit()


# ═══════════════════════════════════════════════ Screen 4b: AI Job Live View ════

_CARD_MIME = "application/x-article-card"


class ImageCard(QFrame):
    """Draggable thumbnail for one AI-classified article."""
    view_image            = pyqtSignal(str, str, str, str)  # (image_path, article_number, category, url)
    ctrl_clicked          = pyqtSignal(object)               # emits self
    context_menu_requested = pyqtSignal(object)              # emits self

    def __init__(self, article_number: str, image_path: str,
                 category: str, url: str = "",
                 meta: Optional[Dict] = None, parent=None):
        super().__init__(parent)
        self.article_number = article_number
        self.image_path     = image_path
        self.category       = category
        self.url            = url
        self._drag_start:   Optional[QPoint] = None
        self._selected:     bool = False

        self.setFixedHeight(120)
        self._normal_style   = "background:#313244; border-radius:6px; border:1px solid #45475a;"
        self._selected_style = "background:#313244; border-radius:6px; border:2px solid #89b4fa;"
        self.setStyleSheet(self._normal_style)
        self.setCursor(Qt.CursorShape.OpenHandCursor)
        self.setToolTip(article_number)

        lay = QHBoxLayout(self)
        lay.setContentsMargins(6, 6, 6, 6)
        lay.setSpacing(8)

        self._img_lbl = QLabel()
        self._img_lbl.setFixedSize(90, 108)
        self._img_lbl.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self._img_lbl.setStyleSheet("background:#11111b; border-radius:4px;")
        lay.addWidget(self._img_lbl)

        # ── info panel ──────────────────────────────────────────────────────
        info_lay = QVBoxLayout()
        info_lay.setContentsMargins(0, 2, 0, 2)
        info_lay.setSpacing(2)

        art_lbl = QLabel(article_number)
        art_lbl.setStyleSheet("color:#cdd6f4; font-size:10px; font-weight:bold;")
        info_lay.addWidget(art_lbl)

        m = meta or {}

        beskr = m.get("beskrivning", "")
        if beskr:
            d_lbl = QLabel(beskr[:70] + ("…" if len(beskr) > 70 else ""))
            d_lbl.setStyleSheet("color:#a6adc8; font-size:9px;")
            d_lbl.setWordWrap(True)
            info_lay.addWidget(d_lbl)

        # Dimensions row
        dims = []
        if m.get("langd"): dims.append(f"L {m['langd']} mm")
        if m.get("bredd"): dims.append(f"B {m['bredd']} mm")
        if m.get("hojd"):  dims.append(f"H {m['hojd']} mm")
        if dims:
            dim_lbl = QLabel("  ".join(dims))
            dim_lbl.setStyleSheet("color:#6c7086; font-size:9px;")
            info_lay.addWidget(dim_lbl)

        # Weight / volume row
        wv = []
        if m.get("vikt_brutto"): wv.append(f"Vikt {m['vikt_brutto']} kg")
        if m.get("volym"):       wv.append(f"Vol {m['volym']}")
        if wv:
            wv_lbl = QLabel("  ".join(wv))
            wv_lbl.setStyleSheet("color:#6c7086; font-size:9px;")
            info_lay.addWidget(wv_lbl)

        info_lay.addStretch()
        lay.addLayout(info_lay, 1)

        self._load_thumbnail()

    def set_selected(self, selected: bool):
        self._selected = selected
        self.setStyleSheet(self._selected_style if selected else self._normal_style)

    def _load_thumbnail(self):
        if not self.image_path or not Path(self.image_path).exists():
            self._img_lbl.setText("?")
            return
        try:
            if PIL_AVAILABLE:
                img = PILImage.open(self.image_path)
                img.thumbnail((90, 108), PILImage.LANCZOS)
                buf = BytesIO(); img.save(buf, format="PNG"); buf.seek(0)
                px = QPixmap(); px.loadFromData(buf.read())
            else:
                px = QPixmap(self.image_path)
                px = px.scaled(90, 108,
                               Qt.AspectRatioMode.KeepAspectRatio,
                               Qt.TransformationMode.SmoothTransformation)
            self._img_lbl.setPixmap(px)
        except Exception:
            self._img_lbl.setText("!")

    def mousePressEvent(self, event):
        if event.button() == Qt.MouseButton.LeftButton:
            self._drag_start = event.pos()

    def mouseMoveEvent(self, event):
        if (self._drag_start is not None and
                event.buttons() & Qt.MouseButton.LeftButton):
            if (event.pos() - self._drag_start).manhattanLength() > 8:
                self._start_drag()

    def mouseReleaseEvent(self, event):
        if event.button() == Qt.MouseButton.LeftButton and self._drag_start is not None:
            if (event.pos() - self._drag_start).manhattanLength() <= 8:
                if event.modifiers() & Qt.KeyboardModifier.ControlModifier:
                    self.ctrl_clicked.emit(self)
                else:
                    self.view_image.emit(self.image_path, self.article_number, self.category, self.url)
        self._drag_start = None

    def contextMenuEvent(self, event):
        self.context_menu_requested.emit(self)

    def _start_drag(self):
        import json
        self.setCursor(Qt.CursorShape.ClosedHandCursor)
        drag = QDrag(self)
        mime = QMimeData()
        mime.setData(
            _CARD_MIME,
            QByteArray(json.dumps({
                "article_number": self.article_number,
                "from_category":  self.category,
                "image_path":     self.image_path,
            }).encode()),
        )
        px = self._img_lbl.pixmap()
        if px and not px.isNull():
            drag.setPixmap(px.scaled(80, 60, Qt.AspectRatioMode.KeepAspectRatio))
        drag.setMimeData(mime)
        drag.exec(Qt.DropAction.MoveAction)
        self.setCursor(Qt.CursorShape.OpenHandCursor)
        self._drag_start = None


class CategoryColumn(QFrame):
    """Scrollable column for one category in the AI job live view."""
    card_dropped      = pyqtSignal(str, str, str)  # (article_number, from_cat, to_cat)
    header_clicked    = pyqtSignal(str)             # (category_name)
    threshold_reached = pyqtSignal(str, int)         # (category_name, count) – emitted at 1/5/10 cards
    analyze_requested = pyqtSignal(str)             # (category_name) – right-click → "Analysera kategori"

    def __init__(self, category_name: str, color: str, parent=None):
        super().__init__(parent)
        self.category_name = category_name
        self.setAcceptDrops(True)
        self._normal_style = "background:#1e1e2e; border-right:1px solid #313244;"
        self._hover_style  = (
            "background:#1e1e2e; border-right:1px solid #313244;"
            "border:2px solid #89b4fa;"
        )
        self.setStyleSheet(self._normal_style)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(0)

        # ── header (clickable to view/edit AI knowledge) ───────────────────
        header = QFrame()
        header.setFixedHeight(44)
        header.setStyleSheet("background:#181825; border-bottom:1px solid #313244;")
        header.setCursor(Qt.CursorShape.PointingHandCursor)
        header.setToolTip("Klicka för att visa/redigera AI-analysen")
        hl = QHBoxLayout(header)
        hl.setContentsMargins(10, 0, 10, 0)
        name_lbl = QLabel(category_name)
        name_lbl.setStyleSheet(
            f"color:{color}; font-size:12px; font-weight:bold;"
        )
        name_lbl.setAlignment(Qt.AlignmentFlag.AlignCenter)
        hl.addWidget(name_lbl, 1)
        self._count_lbl = QLabel("0")
        self._count_lbl.setStyleSheet("color:#6c7086; font-size:11px;")
        hl.addWidget(self._count_lbl)
        self._knowledge_dot = QLabel("●")
        self._knowledge_dot.setStyleSheet("color:#45475a; font-size:8px;")
        self._knowledge_dot.setToolTip("AI-analys ej klar ännu")
        hl.addWidget(self._knowledge_dot)
        layout.addWidget(header)

        # Make header clickable / right-clickable
        def _header_mouse(e):
            if e.button() == Qt.MouseButton.RightButton:
                self.analyze_requested.emit(self.category_name)
            else:
                self.header_clicked.emit(self.category_name)
        header.mousePressEvent = _header_mouse

        # ── scroll area ────────────────────────────────────────────────────
        self._scroll = QScrollArea()
        self._scroll.setWidgetResizable(True)
        self._scroll.setHorizontalScrollBarPolicy(
            Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        self._scroll.setStyleSheet("border:none;")

        self._container = QWidget()
        self._container.setStyleSheet("background:#1e1e2e;")
        self._cards_lay = QVBoxLayout(self._container)
        self._cards_lay.setContentsMargins(6, 6, 6, 6)
        self._cards_lay.setSpacing(5)
        self._cards_lay.addStretch()

        self._scroll.setWidget(self._container)
        layout.addWidget(self._scroll, 1)

        self._cards: List["ImageCard"] = []
        self._is_new_category = False
        self._thresholds_emitted: set = set()
        self._name_lbl = name_lbl  # keep ref for rename

    def mark_as_new_category(self):
        self._is_new_category = True

    def set_name(self, new_name: str, color: str = ""):
        """Rename this column (updates header label)."""
        self.category_name = new_name
        style = self._name_lbl.styleSheet()
        if color:
            import re
            style = re.sub(r"color:[^;]+;", f"color:{color};", style)
        self._name_lbl.setText(new_name)
        self._name_lbl.setStyleSheet(style)

    def set_knowledge_ready(self):
        """Green dot to indicate AI analysis is available."""
        self._knowledge_dot.setStyleSheet("color:#a6e3a1; font-size:8px;")
        self._knowledge_dot.setToolTip("AI-analys klar — klicka för att visa")

    def prepend_card(self, card: "ImageCard"):
        """Insert card at the top (newest first)."""
        self._cards_lay.insertWidget(0, card)
        self._cards.insert(0, card)
        n = len(self._cards)
        self._count_lbl.setText(str(n))
        QTimer.singleShot(30, lambda: self._scroll.verticalScrollBar().setValue(0))
        if self._is_new_category:
            for milestone in (1, 5, MAX_EXAMPLES_PER_CAT):
                if n == milestone and milestone not in self._thresholds_emitted:
                    self._thresholds_emitted.add(milestone)
                    self.threshold_reached.emit(self.category_name, milestone)

    def remove_card_by_article(self, article_number: str) -> Optional["ImageCard"]:
        for card in self._cards:
            if card.article_number == article_number:
                self._cards_lay.removeWidget(card)
                card.setParent(None)
                self._cards.remove(card)
                self._count_lbl.setText(str(len(self._cards)))
                return card
        return None

    # ── drag & drop ────────────────────────────────────────────────────────
    def dragEnterEvent(self, event):
        if event.mimeData().hasFormat(_CARD_MIME):
            import json
            data = json.loads(bytes(event.mimeData().data(_CARD_MIME)))
            if data.get("from_category") != self.category_name:
                event.acceptProposedAction()
                self.setStyleSheet(self._hover_style)
                return
        event.ignore()

    def dragLeaveEvent(self, event):
        self.setStyleSheet(self._normal_style)

    def dropEvent(self, event):
        self.setStyleSheet(self._normal_style)
        if event.mimeData().hasFormat(_CARD_MIME):
            import json
            data = json.loads(bytes(event.mimeData().data(_CARD_MIME)))
            from_cat = data.get("from_category", "")
            art_num  = data.get("article_number", "")
            if from_cat != self.category_name and art_num:
                event.acceptProposedAction()
                self.card_dropped.emit(art_num, from_cat, self.category_name)
                return
        event.ignore()


class NewCategoryWorker(AIJobWorker):
    """Generates knowledge for one newly added category, then re-classifies Övrigt."""
    article_reclassified = pyqtSignal(str, str, str)  # (article_number, new_cat, image_path)

    def __init__(self, new_cat_name: str, new_cat_desc: str,
                 example_cards: List[Dict],           # [{article_number, image_path}]
                 existing_knowledge: Dict[str, str],
                 ovrigt_cards: List[Dict],             # [{article_number, image_path}]
                 all_categories: List[Dict],
                 syfte: str, api_url: str, model: str,
                 compress: bool, data_mgr, parent=None):
        super().__init__(all_categories, [], [], syfte, api_url, model, compress, data_mgr)
        self._new_cat_name  = new_cat_name
        self._new_cat_desc  = new_cat_desc
        self._example_cards = example_cards
        self._ovrigt_cards  = ovrigt_cards
        self.cat_knowledge  = dict(existing_knowledge)

    def run(self):
        if not REQUESTS_AVAILABLE:
            self.error.emit("requests ej installerat")
            return

        # Step 1: generate knowledge for the new category
        self.progress.emit(f"=== Analyserar ny kategori: {self._new_cat_name} ===")
        items = [{"article_number": c["article_number"], "image_path": c["image_path"]}
                 for c in self._example_cards]
        try:
            knowledge = self._generate_knowledge(self._new_cat_name, self._new_cat_desc, items)
            self.cat_knowledge[self._new_cat_name] = knowledge
            self.knowledge_ready.emit(self._new_cat_name, knowledge)
            self.progress.emit("✓ Analys klar")
        except Exception as e:
            self.progress.emit(f"✗ Analys misslyckades: {e}")
            self.cat_knowledge[self._new_cat_name] = self._new_cat_desc
            self.knowledge_ready.emit(self._new_cat_name, self._new_cat_desc)

        # Step 1.5: regenerate contrast rules with updated knowledge
        known_cats = {k: v for k, v in self.cat_knowledge.items()
                      if v and k != "Övrigt"}
        if len(known_cats) >= 2:
            self.progress.emit("Uppdaterar kontrastanalys…")
            try:
                self.contrast_knowledge = self._generate_contrast_knowledge(self.cat_knowledge)
                self.contrast_ready.emit(self.contrast_knowledge)
                self.progress.emit("  ✓ Kontrastanalys uppdaterad")
            except Exception as e:
                self.progress.emit(f"  ✗ Kontrastanalys misslyckades: {e}")

        # Step 2: re-classify Övrigt cards with updated knowledge
        if not self._ovrigt_cards:
            self.finished_all.emit()
            return
        self.progress.emit(
            f"Omklassificerar {len(self._ovrigt_cards)} Övrigt-artiklar…"
        )
        for i, card in enumerate(self._ovrigt_cards):
            if self._stop:
                break
            img_path = card["image_path"]
            art_num  = card["article_number"]
            if not img_path or not Path(img_path).exists():
                continue
            meta = self.data_mgr.get_meta(art_num, "") or {}
            try:
                new_cat = self._classify_article(img_path, meta, self.cat_knowledge)
                if new_cat != "Övrigt":
                    self.article_reclassified.emit(art_num, new_cat, img_path)
                if (i + 1) % 10 == 0 or i == len(self._ovrigt_cards) - 1:
                    self.progress.emit(f"  [{i+1}/{len(self._ovrigt_cards)}] omklassificerade…")
            except Exception as e:
                self.progress.emit(f"  [{i+1}] {art_num}: {e}")

        self.finished_all.emit()


class ReClassifyWorker(AIJobWorker):
    """Re-classifies a specific list of articles using current knowledge."""

    def __init__(self, articles: List[Dict],   # [{article_number, image_path, url}]
                 cat_knowledge: Dict[str, str],
                 all_categories: List[Dict],
                 syfte: str, api_url: str, model: str,
                 compress: bool, data_mgr,
                 hint: str = "",
                 contrast_knowledge: str = "",
                 parent=None):
        super().__init__(all_categories, [], [], syfte, api_url, model, compress, data_mgr)
        self._articles        = articles
        self.cat_knowledge    = dict(cat_knowledge)
        self._hint            = hint
        self.contrast_knowledge = contrast_knowledge

    def run(self):
        if not REQUESTS_AVAILABLE:
            self.error.emit("requests ej installerat")
            return
        for i, art in enumerate(self._articles):
            if self._stop:
                break
            art_num  = art["article_number"]
            img_path = art.get("image_path", "")
            url      = art.get("url", "")
            if not img_path or not Path(img_path).exists():
                continue
            meta = self.data_mgr.get_meta(art_num, "") or {}
            try:
                cat = self._classify_article(img_path, meta, self.cat_knowledge, self._hint)
                self.article_classified.emit(art_num, cat, url, img_path)
                self.progress.emit(f"Gör om [{i+1}/{len(self._articles)}]: {art_num} → {cat}")
            except Exception as e:
                self.progress.emit(f"  [{i+1}] {art_num}: {e}")
        self.finished_all.emit()


class AIJobScreen(QWidget):
    """Full-screen live view while the AI job runs.

    Shows a kanban board — one column per category — updating in real time.
    Cards are draggable between columns to correct misclassifications.
    Clicking a card opens an enlarged image view.
    """
    article_added = pyqtSignal(str, str, str)   # (article_number, category, url)
    reclassified  = pyqtSignal(str, str)         # (article_number, new_category)
    finished      = pyqtSignal()

    def __init__(self, categories: List[Dict], categorized: List[Dict],
                 csv_data: List[Dict], syfte: str,
                 api_url: str, model: str, compress: bool,
                 data_mgr, test_name: str, parent=None):
        super().__init__(parent)
        self._categories = categories
        self._categorized = categorized
        self._csv_data    = csv_data
        self._syfte       = syfte
        self._api_url     = api_url
        self._model       = model
        self._compress    = compress
        self._data_mgr    = data_mgr
        self._test_name   = test_name
        self._worker: Optional[AIJobWorker] = None
        self._new_cat_workers: List[NewCategoryWorker] = []
        self._new_cat_workers_by_cat: Dict[str, NewCategoryWorker] = {}
        self._reclass_workers: List[ReClassifyWorker] = []
        self._columns: Dict[str, CategoryColumn] = {}
        self._total_classified = 0
        self._cat_knowledge: Dict[str, str] = {}  # editable knowledge per category
        self._contrast_knowledge: str = ""         # cross-category decision rules
        self._new_category_count = 0  # for color cycling
        self._selected_cards: set = set()  # Ctrl+click multi-select

        # How many articles remain (step 2)
        classified_numbers = {
            e.get("article_number", "") for e in categorized if e.get("article_number")
        }
        self._remaining_count = sum(
            1 for row in csv_data
            if str(row.get("article_number", "")) not in classified_numbers
        )

        self._build()

    # ── UI ─────────────────────────────────────────────────────────────────────

    def _build(self):
        main_lay = QVBoxLayout(self)
        main_lay.setContentsMargins(0, 0, 0, 0)
        main_lay.setSpacing(0)

        self._header = HeaderBar(
            self._test_name, "AI-jobb startar…"
        )
        main_lay.addWidget(self._header)

        # ── kanban columns ────────────────────────────────────────────────
        cols_widget = QWidget()
        cols_widget.setStyleSheet("background:#1e1e2e;")
        cols_lay = QHBoxLayout(cols_widget)
        cols_lay.setContentsMargins(0, 0, 0, 0)
        cols_lay.setSpacing(0)

        non_ovrigt = [c for c in self._categories if c["name"] != "Övrigt"]
        all_display = non_ovrigt + [{"name": "Övrigt"}]
        for i, cat in enumerate(all_display):
            name  = cat["name"]
            color = (CATEGORY_COLORS[i % len(CATEGORY_COLORS)]
                     if name != "Övrigt" else "#6c7086")
            col = CategoryColumn(name, color)
            col.card_dropped.connect(self._on_card_dropped)
            col.header_clicked.connect(self._show_knowledge_dialog)
            col.analyze_requested.connect(self._on_analyze_category_requested)
            cols_lay.addWidget(col)
            self._columns[name] = col
            self._new_category_count = len(all_display)

        self._cols_lay = cols_lay
        self._cols_widget = cols_widget
        main_lay.addWidget(cols_widget, 1)

        # ── footer ────────────────────────────────────────────────────────
        footer = QFrame()
        footer.setFixedHeight(44)
        footer.setStyleSheet(
            "background:#181825; border-top:1px solid #313244;"
        )
        fl = QHBoxLayout(footer)
        fl.setContentsMargins(16, 0, 16, 0)

        self._progress_lbl = QLabel("Steg 1: Genererar kategorikunskap…")
        self._progress_lbl.setStyleSheet("color:#6c7086; font-size:12px;")
        fl.addWidget(self._progress_lbl, 1)

        add_cat_btn = mk_btn("+ Ny kategori", "#FF9800", h=32)
        add_cat_btn.clicked.connect(self._open_add_category_dialog)
        fl.addWidget(add_cat_btn)

        self._stop_early_btn = mk_btn("⏹ Avsluta i förtid", "#b4637a", h=32)
        self._stop_early_btn.clicked.connect(self._stop_early)
        fl.addWidget(self._stop_early_btn)

        self._done_btn = mk_btn("💾  Exportera & Avsluta", "#1B5E20", h=32)
        self._done_btn.setVisible(False)
        self._done_btn.clicked.connect(self.finished.emit)
        fl.addWidget(self._done_btn)

        main_lay.addWidget(footer)

    # ── worker management ──────────────────────────────────────────────────────

    def start(self, skip_worker: bool = False):
        # Build lookup dicts from csv_data for pre-population
        self._url_by_art   = {str(r.get("article_number", "")): r.get("url", "")
                               for r in self._csv_data}
        self._bolag_by_art = {str(r.get("article_number", "")): r.get("bolag", "")
                               for r in self._csv_data}

        # Pre-populate columns with already manually classified articles
        for item in self._categorized:
            cat      = item.get("category", "")
            art_num  = str(item.get("article_number", ""))
            img_path = item.get("image_path", "")
            url      = self._url_by_art.get(art_num, "")
            meta     = self._data_mgr.get_meta(art_num, self._bolag_by_art.get(art_num, "")) or {}
            col = self._columns.get(cat) or self._columns.get("Övrigt")
            if col:
                card = ImageCard(art_num, img_path, cat, url, meta)
                card.view_image.connect(self._show_image_large)
                card.ctrl_clicked.connect(self._on_card_ctrl_clicked)
                card.context_menu_requested.connect(self._on_card_context_menu)
                col.prepend_card(card)

        if skip_worker:
            self._progress_lbl.setText("Session inläst — klar att redigera. Kör AI-jobb för att omklassificera.")
            self._stop_early_btn.setEnabled(False)
            self._done_btn.setVisible(True)
            return

        self._worker = AIJobWorker(
            self._categories, self._categorized, self._csv_data, self._syfte,
            self._api_url, self._model, self._compress, self._data_mgr,
        )
        self._worker.progress.connect(self._on_progress)
        self._worker.knowledge_ready.connect(self._on_knowledge_ready)
        self._worker.contrast_ready.connect(self._on_contrast_ready)
        self._worker.article_classified.connect(self._on_article_classified)
        self._worker.finished_all.connect(self._on_finished)
        self._worker.error.connect(
            lambda msg: self._progress_lbl.setText(f"FEL: {msg}")
        )
        self._worker.start()

    def stop_worker(self):
        if self._worker:
            self._worker.stop()
            self._worker.wait()
            self._worker = None
        for w in self._new_cat_workers:
            w.stop(); w.wait()
        self._new_cat_workers.clear()

    # ── add new category ───────────────────────────────────────────────────────

    def _open_add_category_dialog(self):
        dlg = QDialog(self)
        dlg.setWindowTitle("Ny kategori")
        dlg.setStyleSheet(STYLE)
        dlg.resize(420, 220)
        lay = QVBoxLayout(dlg)
        lay.addWidget(QLabel("Kategorinamn:"))
        name_edit = QLineEdit()
        name_edit.setPlaceholderText("t.ex. Kapslar")
        lay.addWidget(name_edit)
        lay.addWidget(QLabel("Beskrivning (valfri):"))
        desc_edit = QLineEdit()
        desc_edit.setPlaceholderText("Kort beskrivning av vad som hör hit")
        lay.addWidget(desc_edit)
        hint = QLabel(
            f"Dra minst {AI_JOB_MIN_PER_CAT} bild till den nya kolumnen "
            "så startar AI-analysen automatiskt."
        )
        hint.setWordWrap(True)
        hint.setStyleSheet("color:#6c7086; font-size:11px;")
        lay.addWidget(hint)
        btns = QDialogButtonBox(
            QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel
        )
        btns.accepted.connect(dlg.accept)
        btns.rejected.connect(dlg.reject)
        lay.addWidget(btns)
        name_edit.setFocus()
        if dlg.exec() != QDialog.DialogCode.Accepted:
            return
        name = name_edit.text().strip()
        if not name:
            return
        if name in self._columns or name == "Övrigt":
            QMessageBox.warning(self, "Dubblett", f'"{name}" finns redan.')
            return
        self._add_new_column(name, desc_edit.text().strip())

    def _add_new_column(self, name: str, desc: str):
        """Add a new category column to the kanban board."""
        color = CATEGORY_COLORS[self._new_category_count % len(CATEGORY_COLORS)]
        self._new_category_count += 1

        col = CategoryColumn(name, color)
        col.card_dropped.connect(self._on_card_dropped)
        col.header_clicked.connect(self._show_knowledge_dialog)
        col.threshold_reached.connect(self._on_new_cat_threshold)
        col.analyze_requested.connect(self._on_analyze_category_requested)
        col.mark_as_new_category()
        self._columns[name] = col
        self._categories.append({"name": name, "description": desc, "knowledge": ""})

        # Insert before Övrigt (last column)
        ovrigt_col = self._columns.get("Övrigt")
        if ovrigt_col:
            idx = self._cols_lay.indexOf(ovrigt_col)
            self._cols_lay.insertWidget(idx, col)
        else:
            self._cols_lay.addWidget(col)

        # Also update running main worker so new articles can be placed here
        if self._worker and hasattr(self._worker, "categories"):
            self._worker.categories = self._categories

    def _on_new_cat_threshold(self, category_name: str, count: int):
        """Triggered at 1, 5, and 10 cards — each time re-runs the analysis."""
        col = self._columns.get(category_name)
        if not col:
            return

        # Stop any previous analysis worker for this category
        prev = self._new_cat_workers_by_cat.get(category_name)
        if prev and prev.isRunning():
            prev.stop()

        example_cards = [
            {"article_number": c.article_number, "image_path": c.image_path}
            for c in col._cards[:MAX_EXAMPLES_PER_CAT]
        ]
        ovrigt_col = self._columns.get("Övrigt")
        ovrigt_cards = [
            {"article_number": c.article_number, "image_path": c.image_path}
            for c in (ovrigt_col._cards if ovrigt_col else [])
        ]
        cat_desc = next(
            (c.get("description", "") for c in self._categories if c["name"] == category_name),
            ""
        )
        label = {1: "1 bild", 5: "5 bilder", MAX_EXAMPLES_PER_CAT: "10 bilder (slutgiltig)"}.get(count, f"{count} bilder")
        w = NewCategoryWorker(
            category_name, cat_desc, example_cards,
            dict(self._cat_knowledge), ovrigt_cards,
            list(self._categories),
            self._syfte, self._api_url, self._model, self._compress, self._data_mgr,
        )
        w.progress.connect(self._on_progress)
        w.knowledge_ready.connect(self._on_knowledge_ready)
        w.knowledge_ready.connect(self._feed_knowledge_to_main_worker)
        w.contrast_ready.connect(self._on_contrast_ready)
        w.article_reclassified.connect(self._on_new_cat_article_reclassified)
        w.finished_all.connect(lambda: self._progress_lbl.setText(
            f"✓ Analys av '{category_name}' klar ({label})"
        ))
        self._new_cat_workers.append(w)
        self._new_cat_workers_by_cat[category_name] = w
        self._progress_lbl.setText(
            f"Analyserar ny kategori '{category_name}' ({label})…"
        )
        w.start()

    def _on_analyze_category_requested(self, category_name: str):
        """Right-click on category header → re-analyse with user-chosen example count."""
        from PyQt6.QtWidgets import QDialog, QVBoxLayout, QLabel, QSpinBox, QHBoxLayout
        col = self._columns.get(category_name)
        if not col:
            return
        max_n = len(col._cards)
        if max_n == 0:
            return

        dlg = QDialog(self)
        dlg.setWindowTitle(f"Analysera: {category_name}")
        dlg.setStyleSheet(STYLE)
        dlg.resize(400, 180)
        lay = QVBoxLayout(dlg)

        info = QLabel(
            f"Hur många bilder ska användas vid analysen av <b>{category_name}</b>?<br>"
            f"(Det finns {max_n} artiklar i kolumnen)"
        )
        info.setWordWrap(True)
        info.setStyleSheet("color:#cdd6f4; font-size:12px;")
        lay.addWidget(info)

        spin = QSpinBox()
        spin.setRange(1, max_n)
        spin.setValue(min(max_n, MAX_EXAMPLES_PER_CAT))
        spin.setStyleSheet(
            "background:#11111b; color:#cdd6f4; font-size:14px;"
            "border:1px solid #45475a; border-radius:4px; padding:4px;"
        )
        lay.addWidget(spin)

        btn_row = QHBoxLayout()
        ok_btn = mk_btn("Analysera", "#1B5E20")
        ok_btn.clicked.connect(dlg.accept)
        btn_row.addWidget(ok_btn)
        cancel_btn = mk_btn("Avbryt", "#45475a")
        cancel_btn.clicked.connect(dlg.reject)
        btn_row.addWidget(cancel_btn)
        lay.addLayout(btn_row)

        if dlg.exec() != QDialog.DialogCode.Accepted:
            return

        n = spin.value()
        example_cards = [
            {"article_number": c.article_number, "image_path": c.image_path}
            for c in col._cards[:n]
        ]
        ovrigt_col = self._columns.get("Övrigt")
        ovrigt_cards = [
            {"article_number": c.article_number, "image_path": c.image_path}
            for c in (ovrigt_col._cards if ovrigt_col else [])
        ]
        cat_desc = next(
            (c.get("description", "") for c in self._categories if c["name"] == category_name), ""
        )
        w = NewCategoryWorker(
            category_name, cat_desc, example_cards,
            dict(self._cat_knowledge), ovrigt_cards,
            list(self._categories),
            self._syfte, self._api_url, self._model, self._compress, self._data_mgr,
        )
        w.progress.connect(self._on_progress)
        w.knowledge_ready.connect(self._on_knowledge_ready)
        w.knowledge_ready.connect(self._feed_knowledge_to_main_worker)
        w.contrast_ready.connect(self._on_contrast_ready)
        w.article_reclassified.connect(self._on_new_cat_article_reclassified)
        w.finished_all.connect(lambda: self._progress_lbl.setText(
            f"✓ Analysering av '{category_name}' klar ({n} bilder)"
        ))
        self._new_cat_workers.append(w)
        self._progress_lbl.setText(f"Analyserar '{category_name}' med {n} bilder…")
        w.start()

    def _feed_knowledge_to_main_worker(self, category: str, knowledge: str):
        """Push new knowledge to the still-running main worker so it uses it."""
        if self._worker and hasattr(self._worker, "cat_knowledge"):
            self._worker.cat_knowledge[category] = knowledge

    def _on_new_cat_article_reclassified(self, article_number: str,
                                          new_category: str, image_path: str):
        """Move a card from Övrigt to the new category after re-classification."""
        self._on_card_dropped(article_number, "Övrigt", new_category)
        self.reclassified.emit(article_number, new_category)

    # ── slots ──────────────────────────────────────────────────────────────────

    def _on_progress(self, msg: str):
        # Show only the last meaningful line in the footer
        text = msg.strip()
        if text:
            self._progress_lbl.setText(text)

    def _on_article_classified(self, article_number: str, category: str,
                                url: str, image_path: str):
        self._total_classified += 1
        col = self._columns.get(category) or self._columns.get("Övrigt")
        if col:
            bolag = getattr(self, "_bolag_by_art", {}).get(article_number, "")
            meta  = self._data_mgr.get_meta(article_number, bolag) or {}
            card = ImageCard(article_number, image_path, category, url, meta)
            card.view_image.connect(self._show_image_large)
            card.ctrl_clicked.connect(self._on_card_ctrl_clicked)
            card.context_menu_requested.connect(self._on_card_context_menu)
            col.prepend_card(card)

        self._header.set_texts(
            self._test_name,
            f"Klassificerar… {self._total_classified}/{self._remaining_count}",
        )
        self.article_added.emit(article_number, category, url)

    def _stop_early(self):
        from PyQt6.QtWidgets import QMessageBox
        ans = QMessageBox.question(
            self, "Avsluta i förtid",
            "Vill du avbryta AI-jobbet?\n"
            "Artiklar som redan klassificerats sparas.",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
        )
        if ans != QMessageBox.StandardButton.Yes:
            return
        if self._worker and self._worker.isRunning():
            self._worker.stop()
        for w in self._new_cat_workers + list(self._new_cat_workers_by_cat.values()) + self._reclass_workers:
            if w and w.isRunning():
                w.stop()
        self._progress_lbl.setText(f"Avbrutet — {self._total_classified} artiklar klassificerade.")
        self._header.set_texts(self._test_name, "AI-jobb avbrutet")
        self._stop_early_btn.setEnabled(False)
        self._done_btn.setVisible(True)

    def _on_finished(self):
        self._progress_lbl.setText(
            f"✓ Klart!  {self._total_classified} artiklar klassificerade av AI."
        )
        self._header.set_texts(self._test_name, "AI-jobb klart")
        self._stop_early_btn.setEnabled(False)
        self._done_btn.setVisible(True)

    def _on_knowledge_ready(self, category: str, knowledge: str):
        self._cat_knowledge[category] = knowledge
        col = self._columns.get(category)
        if col:
            col.set_knowledge_ready()

    def _on_contrast_ready(self, contrast: str):
        self._contrast_knowledge = contrast
        # Push updated contrast to any still-running workers
        for wkr in [self._worker] + self._new_cat_workers:
            if wkr and wkr.isRunning():
                wkr.contrast_knowledge = contrast

    # ── card selection (Ctrl+click) ─────────────────────────────────────────

    def _on_card_ctrl_clicked(self, card):
        if card in self._selected_cards:
            self._selected_cards.discard(card)
            card.set_selected(False)
        else:
            self._selected_cards.add(card)
            card.set_selected(True)

    def _clear_selection(self):
        for c in list(self._selected_cards):
            c.set_selected(False)
        self._selected_cards.clear()

    # ── right-click context menu ────────────────────────────────────────────

    def _on_card_context_menu(self, card):
        from PyQt6.QtWidgets import QMenu
        from PyQt6.QtGui import QCursor

        # If right-clicked card is not in selection, use only it
        if card not in self._selected_cards:
            self._clear_selection()
            targets = [card]
        else:
            targets = list(self._selected_cards)

        menu = QMenu(self)
        menu.setStyleSheet(
            "QMenu { background:#313244; color:#cdd6f4; border:1px solid #45475a; }"
            "QMenu::item:selected { background:#45475a; }"
        )
        n = len(targets)
        label = f"Gör om ({n} artikel{'er' if n > 1 else ''})"
        action = menu.addAction(label)
        chosen = menu.exec(QCursor.pos())
        if chosen == action:
            self._prompt_and_reclassify(targets)

    def _prompt_and_reclassify(self, cards):
        """Show a dialog asking for a reclassify reason, then start the job."""
        from PyQt6.QtWidgets import QDialog, QVBoxLayout, QLabel, QTextEdit, QHBoxLayout
        n = len(cards)
        dlg = QDialog(self)
        dlg.setWindowTitle("Gör om")
        dlg.setStyleSheet(STYLE)
        dlg.resize(560, 280)
        lay = QVBoxLayout(dlg)

        info = QLabel(
            f"<b>{n} artikel{'er' if n > 1 else ''}</b> kommer omklassificeras av AI:n.<br>"
            "Ange gärna en orsak — det hjälper AI:n att välja rätt kategori."
        )
        info.setWordWrap(True)
        info.setStyleSheet("color:#cdd6f4; font-size:12px;")
        lay.addWidget(info)

        hint_lbl = QLabel("Orsak (valfritt):")
        hint_lbl.setStyleSheet("color:#6c7086; font-size:11px; margin-top:8px;")
        lay.addWidget(hint_lbl)

        hint_edit = QTextEdit()
        hint_edit.setPlaceholderText(
            'T.ex. "Kategori \'Säck\' delades upp i \'Säck max 15 kg\' och \'Säck minst 15 kg\'"'
        )
        hint_edit.setFixedHeight(80)
        hint_edit.setStyleSheet(
            "background:#11111b; color:#cdd6f4; font-size:12px;"
            "border:1px solid #45475a; border-radius:4px;"
        )
        lay.addWidget(hint_edit)

        btn_row = QHBoxLayout()
        run_btn = mk_btn("Kör om nu", "#1B5E20")
        run_btn.clicked.connect(dlg.accept)
        btn_row.addWidget(run_btn)
        cancel_btn = mk_btn("Avbryt", "#45475a")
        cancel_btn.clicked.connect(dlg.reject)
        btn_row.addWidget(cancel_btn)
        lay.addLayout(btn_row)

        if dlg.exec() != QDialog.DialogCode.Accepted:
            return

        hint = hint_edit.toPlainText().strip()
        self._reclassify_cards(cards, hint)

    def _reclassify_cards(self, cards, hint: str = ""):
        articles = [
            {"article_number": c.article_number, "image_path": c.image_path, "url": c.url}
            for c in cards
        ]
        # Remove cards from current columns
        for c in cards:
            col = self._columns.get(c.category)
            if col:
                col.remove_card_by_article(c.article_number)
        self._clear_selection()

        # Pause main worker so re-classify runs first
        if self._worker and self._worker.isRunning():
            self._worker.pause()

        w = ReClassifyWorker(
            articles, dict(self._cat_knowledge),
            list(self._categories),
            self._syfte, self._api_url, self._model, self._compress, self._data_mgr,
            hint=hint,
            contrast_knowledge=self._contrast_knowledge,
        )
        w.progress.connect(self._on_progress)
        w.article_classified.connect(self._on_article_classified)

        def _on_reclass_done():
            self._progress_lbl.setText(f"✓ Omklassificering klar ({len(articles)} artiklar)")
            # Resume main worker
            if self._worker and self._worker.isRunning():
                self._worker.resume()

        w.finished_all.connect(_on_reclass_done)
        self._reclass_workers.append(w)
        self._progress_lbl.setText(f"Pausar jobb — gör om {len(articles)} artiklar…")
        w.start()

    def _show_knowledge_dialog(self, category: str):
        from PyQt6.QtWidgets import QLineEdit
        knowledge = self._cat_knowledge.get(category, "")
        dlg = QDialog(self)
        dlg.setWindowTitle(f"AI-analys: {category}")
        dlg.setStyleSheet(STYLE)
        dlg.resize(700, 540)
        lay = QVBoxLayout(dlg)

        # ── Category name field (always visible) ──────────────────────────────
        name_row = QHBoxLayout()
        name_lbl = QLabel("Kategorinamn:")
        name_lbl.setStyleSheet("color:#6c7086; font-size:11px;")
        name_row.addWidget(name_lbl)
        name_edit = QLineEdit(category)
        name_edit.setStyleSheet(
            "background:#11111b; color:#cdd6f4; font-size:13px;"
            "border:1px solid #45475a; border-radius:4px; padding:4px 8px;"
        )
        name_row.addWidget(name_edit, 1)
        lay.addLayout(name_row)

        if not knowledge:
            info = QLabel("AI-analysen för denna kategori är inte klar ännu.")
            info.setStyleSheet("color:#6c7086; font-style:italic;")
            lay.addWidget(info)
        else:
            lbl = QLabel(
                "AI:ns analys av kategorin. "
                "Du kan justera texten — den används vid klassificering av återstående artiklar."
            )
            lbl.setWordWrap(True)
            lbl.setStyleSheet("color:#6c7086; font-size:11px;")
            lay.addWidget(lbl)

            editor = QTextEdit()
            editor.setPlainText(knowledge)
            editor.setStyleSheet(
                "background:#11111b; color:#cdd6f4; font-family:monospace;"
                "border:1px solid #45475a; border-radius:4px;"
            )
            lay.addWidget(editor)

        btn_row = QHBoxLayout()
        save_btn = mk_btn("Spara ändringar", "#1B5E20")

        def _save():
            new_name = name_edit.text().strip()
            new_knowledge = editor.toPlainText().strip() if knowledge else ""

            # ── rename if changed ──────────────────────────────────────────
            if new_name and new_name != category:
                col = self._columns.get(category)
                if col:
                    col.set_name(new_name)
                    # Update columns dict
                    self._columns[new_name] = self._columns.pop(category)
                    # Update cards in that column
                    for c in col._cards:
                        c.category = new_name
                    # Update categories list
                    for cat in self._categories:
                        if cat["name"] == category:
                            cat["name"] = new_name
                            break
                    # Move knowledge entry
                    self._cat_knowledge[new_name] = new_knowledge
                    self._cat_knowledge.pop(category, None)
                    # Update running workers
                    for wkr in [self._worker] + self._new_cat_workers + list(self._new_cat_workers_by_cat.values()):
                        if wkr and hasattr(wkr, "cat_knowledge"):
                            if category in wkr.cat_knowledge:
                                wkr.cat_knowledge[new_name] = wkr.cat_knowledge.pop(category)
                    dlg.accept()
                    return

            # ── save knowledge only ────────────────────────────────────────
            if knowledge:
                self._cat_knowledge[category] = new_knowledge
                for wkr in [self._worker] + self._new_cat_workers:
                    if wkr and hasattr(wkr, "cat_knowledge"):
                        wkr.cat_knowledge[category] = new_knowledge
            dlg.accept()

        save_btn.clicked.connect(_save)
        btn_row.addWidget(save_btn)
        cancel_btn = mk_btn("Stäng", "#45475a")
        cancel_btn.clicked.connect(dlg.reject)
        btn_row.addWidget(cancel_btn)
        lay.addLayout(btn_row)

        dlg.exec()

    def _on_card_dropped(self, article_number: str, from_cat: str, to_cat: str):
        from_col = self._columns.get(from_cat)
        to_col   = self._columns.get(to_cat)
        if from_col and to_col:
            card = from_col.remove_card_by_article(article_number)
            if card:
                card.category = to_cat
                to_col.prepend_card(card)
        self.reclassified.emit(article_number, to_cat)

    def _show_image_large(self, image_path: str, article_number: str = "",
                          category: str = "", url: str = ""):
        if not image_path or not Path(image_path).exists():
            return
        dlg = QDialog(self)
        dlg.setWindowTitle(f"Bildvisning — {article_number}" if article_number else "Bildvisning")
        dlg.setStyleSheet(STYLE)
        dlg.setMinimumWidth(900)

        main_lay = QHBoxLayout(dlg)
        main_lay.setContentsMargins(12, 12, 12, 12)
        main_lay.setSpacing(16)

        # ── Left: image ───────────────────────────────────────────────────────
        img_lbl = QLabel()
        img_lbl.setAlignment(Qt.AlignmentFlag.AlignCenter)
        img_lbl.setStyleSheet("background:#11111b; border-radius:8px;")
        img_lbl.setMinimumSize(500, 500)
        try:
            if PIL_AVAILABLE:
                img = PILImage.open(image_path)
                img.thumbnail((700, 600), PILImage.LANCZOS)
                buf = BytesIO(); img.save(buf, format="PNG"); buf.seek(0)
                px = QPixmap(); px.loadFromData(buf.read())
            else:
                px = QPixmap(image_path)
                px = px.scaled(700, 600,
                               Qt.AspectRatioMode.KeepAspectRatio,
                               Qt.TransformationMode.SmoothTransformation)
            img_lbl.setPixmap(px)
        except Exception as e:
            img_lbl.setText(str(e))
        main_lay.addWidget(img_lbl, 3)

        # ── Right: article info ───────────────────────────────────────────────
        info_widget = QWidget()
        info_widget.setStyleSheet("background:#313244; border-radius:8px;")
        info_lay = QVBoxLayout(info_widget)
        info_lay.setContentsMargins(16, 16, 16, 16)
        info_lay.setSpacing(10)

        def add_field(label: str, value: str):
            if not value:
                return
            lbl = QLabel(f"<span style='color:#6c7086;font-size:11px;'>{label}</span><br>"
                         f"<span style='color:#cdd6f4;font-size:13px;'>{value}</span>")
            lbl.setWordWrap(True)
            lbl.setTextFormat(Qt.TextFormat.RichText)
            lbl.setStyleSheet("background:transparent;")
            info_lay.addWidget(lbl)
            sep = QFrame()
            sep.setFrameShape(QFrame.Shape.HLine)
            sep.setStyleSheet("color:#45475a;")
            info_lay.addWidget(sep)

        add_field("Artikelnummer", article_number)
        add_field("AI-kategori", category)

        # Look up article metadata
        meta = {}
        if article_number and self._data_mgr:
            meta = self._data_mgr.get_meta(article_number) or {}

        add_field("Beskrivning", meta.get("beskrivning", ""))
        add_field("Kategori (original)", meta.get("kategori", ""))
        add_field("Huvudkategori", meta.get("huvudkategori", ""))
        add_field("Vikt brutto", meta.get("vikt_brutto", ""))
        add_field("Vikt netto", meta.get("vikt_netto", ""))
        add_field("Volym", meta.get("volym", ""))
        add_field("Bolag", meta.get("bolag", ""))
        add_field("UN-nummer", meta.get("un_nummer", ""))
        if url:
            add_field("URL", url)

        info_lay.addStretch()

        close_btn = mk_btn("Stäng", "#45475a")
        close_btn.clicked.connect(dlg.accept)
        info_lay.addWidget(close_btn)

        main_lay.addWidget(info_widget, 2)
        dlg.exec()


# ═══════════════════════════════════════════════════════════ Screen 5: Done ════
class DoneScreen(QWidget):
    new_test      = pyqtSignal()
    retest_ovrigt = pyqtSignal()
    export_excel  = pyqtSignal()
    export_zip    = pyqtSignal()
    resume_job    = pyqtSignal()   # open AI job screen to continue editing
    quit_app      = pyqtSignal()

    def __init__(self, parent=None):
        super().__init__(parent)
        self._lay = QVBoxLayout(self)
        self._lay.setContentsMargins(0, 0, 0, 0)

    def show_results(self, test_name: str, categories: List[Dict],
                     n_processed: int, csv_mode: bool, has_results: bool,
                     ovrigt_count: int):
        # Clear old content
        while self._lay.count():
            item = self._lay.takeAt(0)
            if item.widget():
                item.widget().setParent(None)

        self._lay.addWidget(HeaderBar(test_name))

        center = QWidget()
        c = QVBoxLayout(center)
        c.setAlignment(Qt.AlignmentFlag.AlignCenter)

        card = QFrame()
        card.setStyleSheet("background-color:#313244; border-radius:12px;")
        card.setFixedWidth(500)
        cl = QVBoxLayout(card)
        cl.setContentsMargins(40, 40, 40, 40)
        cl.setSpacing(8)

        ok_lbl = QLabel("✓  Test avslutat!")
        ok_lbl.setStyleSheet("font-size:28px; font-weight:bold; color:#a6e3a1;")
        ok_lbl.setAlignment(Qt.AlignmentFlag.AlignCenter)
        cl.addWidget(ok_lbl)

        processed_lbl = QLabel(f"Behandlade bilder: {n_processed}")
        processed_lbl.setStyleSheet("color:#6c7086;")
        processed_lbl.setAlignment(Qt.AlignmentFlag.AlignCenter)
        cl.addWidget(processed_lbl)
        cl.addSpacing(8)

        for cat in categories + [{"name": "Övrigt"}]:
            folder = Path(f"{test_name}.{cat['name']}")
            if folder.exists():
                count = len(list(folder.iterdir()))
                row = QLabel(f"📁  {folder.name}  —  {count} bild(er)")
                cl.addWidget(row)

        cl.addSpacing(12)

        if csv_mode and has_results:
            ex = mk_btn("💾  Exportera Excel", "#1B5E20", h=40)
            ex.clicked.connect(self.export_excel.emit)
            cl.addWidget(ex)

        if ovrigt_count:
            ov = mk_btn(f"Testa Övrigt igen  ({ovrigt_count} bilder)", "#FF9800", h=40)
            ov.clicked.connect(self.retest_ovrigt.emit)
            cl.addWidget(ov)

        resume_b = mk_btn("🔀  Fortsätt redigera i AI-vyn", "#6c7086", "#cdd6f4", h=40)
        resume_b.clicked.connect(self.resume_job.emit)
        cl.addWidget(resume_b)

        zip_b = mk_btn("📦  Ladda ner session (.zip)", "#45475a", "#cdd6f4", h=40)
        zip_b.clicked.connect(self.export_zip.emit)
        cl.addWidget(zip_b)

        cl.addSpacing(4)
        nav = QHBoxLayout()
        nav.setSpacing(8)
        new_b = mk_btn("Nytt test", "#2196F3"); new_b.clicked.connect(self.new_test.emit)
        quit_b = mk_btn("Avsluta", "#f38ba8", "#1e1e2e"); quit_b.clicked.connect(self.quit_app.emit)
        nav.addWidget(new_b); nav.addWidget(quit_b)
        cl.addLayout(nav)

        c.addWidget(card)
        self._lay.addWidget(center)


# ══════════════════════════════════════════════════════════ MainApp ════════════
class MainApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Bildklassificering")
        self.resize(1000, 700)
        self.setMinimumSize(820, 600)
        self.setStyleSheet(STYLE)

        # ── Session state
        self.test_name    = ""
        self.syfte        = ""
        self.categories: List[Dict] = []
        self.images: List[Optional[Path]] = []
        self.current_index = 0
        self.csv_mode     = False
        self.csv_data:    List[Dict] = []
        self.results:     List[Dict] = []
        self.temp_dir:    Optional[str] = None
        self.retesting_ovrigt = False
        self.categorized: List[Dict] = []

        # ── AI state
        self.ai_settings: Dict = {}
        self.ai_enabled   = False

        # ── Data
        self.data_mgr = DataManager()

        # ── Download worker
        self.dl_worker:       Optional[ImageDownloader] = None
        self._ready_images:   set = set()

        # ── Stack
        self.stack = QStackedWidget()
        self.setCentralWidget(self.stack)

        self._name_scr  = NameScreen()
        self._cat_scr   = CategoriesScreen()
        self._cl_scr    = ClassifyScreen()
        self._done_scr  = DoneScreen()

        self.stack.addWidget(self._name_scr)   # 0
        self.stack.addWidget(self._cat_scr)    # 1
        # indices 2+ are dynamic (source, ai-settings, wait screens)
        self.stack.addWidget(self._cl_scr)     # added as needed
        self.stack.addWidget(self._done_scr)

        # ── Connections
        self._name_scr.go_next.connect(self._on_name_done)
        self._name_scr.load_zip.connect(self._import_zip)
        self._cat_scr.go_next.connect(self._on_cats_done)
        self._cat_scr.go_back.connect(lambda: self.stack.setCurrentWidget(self._name_scr))

        self._cl_scr.classified.connect(self._on_classified)
        self._cl_scr.skipped.connect(self._on_skip)
        self._cl_scr.go_back.connect(self._on_go_back)
        self._cl_scr.add_category.connect(self._add_cat_during_test)
        self._cl_scr.end_test.connect(self._show_done)
        self._cl_scr.run_ai_job.connect(self._run_ai_job)

        self._done_scr.new_test.connect(self._on_new_test)
        self._done_scr.retest_ovrigt.connect(self._retest_ovrigt)
        self._done_scr.export_excel.connect(self._export_excel)
        self._done_scr.export_zip.connect(self._export_zip)
        self._done_scr.resume_job.connect(self._open_resumed_session)
        self._done_scr.quit_app.connect(self.close)

        self.stack.setCurrentWidget(self._name_scr)
        self.showMaximized()

    # ── helpers ────────────────────────────────────────────────────────────────

    def _push_screen(self, widget: QWidget):
        """Add widget to stack and show it."""
        self.stack.addWidget(widget)
        self.stack.setCurrentWidget(widget)

    def _replace_top(self, new_widget: QWidget, old_widget: Optional[QWidget]):
        """Replace the top (last) dynamic screen."""
        if old_widget and self.stack.indexOf(old_widget) >= 0:
            self.stack.removeWidget(old_widget)
            old_widget.setParent(None)
        self._push_screen(new_widget)

    # ── navigation ─────────────────────────────────────────────────────────────

    def _on_name_done(self, name: str, syfte: str):
        self.test_name = name
        self.syfte     = syfte
        self._cat_scr.set_test_name(name)
        self.stack.setCurrentWidget(self._cat_scr)

    def _on_cats_done(self, cats: List[Dict]):
        self.categories = [dict(c, knowledge="") for c in cats]
        self._show_source_screen()

    def _show_source_screen(self):
        src = SourceScreen(self.test_name, len(self.data_mgr.builtin_attributes))
        src.use_folder.connect(self._load_folder)
        src.use_builtin.connect(self._show_filter_screen)
        src.use_csv.connect(self._load_csv)
        src.go_back.connect(lambda: self.stack.setCurrentWidget(self._cat_scr))
        self._src_scr = src
        self._push_screen(src)

    def _show_filter_screen(self):
        flt = FilterScreen(self.test_name, list(self.data_mgr.builtin_attributes), self.data_mgr)
        flt.go_next.connect(self._download_images)
        flt.go_back.connect(lambda: self.stack.setCurrentWidget(self._src_scr))
        self._flt_scr = flt
        self._push_screen(flt)

    def _show_ai_settings(self):
        ai = AISettingsScreen(self.test_name)
        ai.go_next.connect(self._on_ai_done)
        ai.go_back.connect(lambda: self.stack.setCurrentWidget(self._src_scr))
        self._ai_scr = ai
        self._push_screen(ai)

    def _on_ai_done(self, settings: Dict):
        self.ai_settings = settings
        self.ai_enabled  = bool(settings)
        self._show_classify()

    # ── image loading ──────────────────────────────────────────────────────────

    def _load_folder(self):
        self.csv_mode = False
        if not IMAGE_DIR.exists():
            QMessageBox.critical(self, "Mapp saknas", f'Mappen "{IMAGE_DIR}" hittades inte.')
            return
        imgs = [f for f in IMAGE_DIR.iterdir() if f.suffix.lower() in SUPPORTED_EXT]
        if not imgs:
            QMessageBox.warning(self, "Inga bilder", f'Inga bilder i "{IMAGE_DIR}".')
            return
        random.shuffle(imgs)
        self.images = imgs
        self.current_index = 0
        self._show_classify()

    def _load_csv(self):
        path, _ = QFileDialog.getOpenFileName(
            self, "Välj CSV-fil", "", "CSV-filer (*.csv);;Alla filer (*)"
        )
        if not path:
            return
        rows = self._parse_csv(path)
        if rows:
            self._download_images(rows)

    def _parse_csv(self, path: str) -> Optional[List[Dict]]:
        try:
            with open(path, newline="", encoding="utf-8-sig") as f:
                sample = f.read(4096); f.seek(0)
                try:
                    dialect = csv.Sniffer().sniff(sample, delimiters=",;\t|")
                except csv.Error:
                    dialect = csv.excel
                all_rows = list(csv.reader(f, dialect))
            url_col = None
            for row in all_rows[:5]:
                for i, cell in enumerate(row):
                    if cell.strip().lower().startswith("http"):
                        url_col = i; break
                if url_col is not None:
                    break
            if url_col is None:
                QMessageBox.warning(self, "Ingen URL-kolumn",
                                    "Kunde inte hitta kolumn med URL:er.")
                return None
            rows = []
            for row in all_rows:
                if len(row) <= url_col: continue
                art = row[0].strip(); url = row[url_col].strip()
                if art and url.lower().startswith("http"):
                    rows.append({"article_number": art, "url": url})
            if not rows:
                QMessageBox.warning(self, "Inga rader", "Inga giltiga rader i filen.")
                return None
            return rows
        except Exception as e:
            QMessageBox.critical(self, "CSV-fel", f"Kunde inte läsa filen:\n{e}")
            return None

    def _download_images(self, rows: List[Dict]):
        random.shuffle(rows)
        self.csv_mode  = True
        self.csv_data  = [{"article_number": r["article_number"], "url": r["url"],
                           "bolag": r.get("bolag", ""), "img_path": None} for r in rows]
        self.images    = [None] * len(rows)
        self.results   = []
        self.current_index  = 0
        self._ready_images  = set()

        self.temp_dir = tempfile.mkdtemp(prefix="bildklassificering_")
        if self.dl_worker:
            self.dl_worker.stop(); self.dl_worker.wait()
        self.dl_worker = ImageDownloader(rows, self.temp_dir)
        self.dl_worker.image_ready.connect(self._on_image_ready)
        self.dl_worker.start()

        # Show a loading screen until image 0 is ready
        self._loading_scr = self._make_loading_screen(len(rows))
        self._push_screen(self._loading_scr)

        def poll():
            if 0 in self._ready_images:
                self.stack.removeWidget(self._loading_scr)
                self._loading_scr.setParent(None)
                self._show_classify()
            else:
                QTimer.singleShot(200, poll)
        QTimer.singleShot(200, poll)

    def _make_loading_screen(self, total: int) -> QWidget:
        w = QWidget()
        w.setStyleSheet("background:#1e1e2e;")
        lay = QVBoxLayout(w)
        lay.setAlignment(Qt.AlignmentFlag.AlignCenter)
        lbl = QLabel("Hämtar bilder…")
        lbl.setStyleSheet("font-size:20px; font-weight:bold;")
        lbl.setAlignment(Qt.AlignmentFlag.AlignCenter)
        lay.addWidget(lbl)
        sub = QLabel(f"{total} bilder totalt — resten hämtas i bakgrunden")
        sub.setStyleSheet("color:#6c7086;")
        sub.setAlignment(Qt.AlignmentFlag.AlignCenter)
        lay.addWidget(sub)
        return w

    def _on_image_ready(self, index: int, path: str):
        self._ready_images.add(index)
        self.images[index] = Path(path)
        if index < len(self.csv_data):
            self.csv_data[index]["img_path"] = path

    def _get_meta(self, index: int) -> Optional[Dict]:
        if not self.csv_mode or index >= len(self.csv_data):
            return None
        entry = self.csv_data[index]
        return self.data_mgr.get_meta(str(entry["article_number"]), entry.get("bolag", ""))

    # ── classify screen ────────────────────────────────────────────────────────

    def _show_classify(self):
        if self.current_index >= len(self.images):
            self._show_done()
            return

        # Wait for download
        if self.csv_mode and self.current_index not in self._ready_images:
            self._show_wait_screen()
            return

        img_path = self.images[self.current_index]
        if img_path is None:
            self.current_index += 1
            self._show_classify()
            return

        meta = self._get_meta(self.current_index)
        cat_counts, threshold, ai_job_ready = self._get_threshold_data()

        # Find previous classification for this article (shown when going back)
        prev_cat = ""
        if self.csv_mode and self.csv_data:
            art_num = str(self.csv_data[self.current_index].get("article_number", ""))
            for e in self.categorized:
                if str(e.get("article_number", "")) == art_num:
                    prev_cat = e.get("category", "")
                    break
        else:
            for e in self.categorized:
                if e.get("image_path") == str(img_path):
                    prev_cat = e.get("category", "")
                    break

        self._cl_scr.show_image(
            self.test_name, self.categories,
            str(img_path), meta,
            self.current_index, len(self.images),
            cat_counts, threshold, ai_job_ready,
            prev_category=prev_cat,
        )
        self.stack.setCurrentWidget(self._cl_scr)

    def _get_threshold_data(self) -> Tuple[Dict[str, int], int, bool]:
        """Return (counts_per_cat, threshold, ai_job_ready).

        ai_job_ready = True when every non-Övrigt category has >= AI_JOB_MIN_PER_CAT
        items AND AI settings have been configured.
        """
        non_ovrigt = [c["name"] for c in self.categories if c["name"] != "Övrigt"]
        if not non_ovrigt or not self.ai_enabled:
            return {}, 0, False
        threshold = AI_JOB_MIN_PER_CAT
        counts: Dict[str, int] = {name: 0 for name in non_ovrigt}
        for entry in self.categorized:
            cat = entry.get("category", "")
            if cat in counts:
                counts[cat] += 1
        ready = all(counts[name] >= threshold for name in non_ovrigt)
        return counts, threshold, ready

    def _show_wait_screen(self):
        w = QWidget(); w.setStyleSheet("background:#1e1e2e;")
        lay = QVBoxLayout(w); lay.setAlignment(Qt.AlignmentFlag.AlignCenter)
        lbl = QLabel("Väntar på nedladdning…")
        lbl.setStyleSheet("font-size:18px;"); lbl.setAlignment(Qt.AlignmentFlag.AlignCenter)
        lay.addWidget(lbl)
        sub = QLabel(f"{len(self._ready_images)} av {len(self.images)} klara")
        sub.setStyleSheet("color:#6c7086;"); sub.setAlignment(Qt.AlignmentFlag.AlignCenter)
        lay.addWidget(sub)
        self.stack.addWidget(w); self.stack.setCurrentWidget(w)

        def poll():
            if self.current_index in self._ready_images:
                self.stack.removeWidget(w); w.setParent(None)
                self._show_classify()
            else:
                QTimer.singleShot(300, poll)
        QTimer.singleShot(300, poll)

    # ── classify logic ─────────────────────────────────────────────────────────

    def _on_classified(self, category: str):
        if self.current_index >= len(self.images):
            return
        img_path = self.images[self.current_index]

        # ── detect re-classification (user went back) ─────────────────────────
        old_category: str = ""
        if self.csv_mode and self.csv_data:
            art_num = str(self.csv_data[self.current_index]["article_number"])
            for e in self.categorized:
                if str(e.get("article_number", "")) == art_num:
                    old_category = e["category"]
                    e["category"] = category
                    break
            else:
                self.categorized.append({
                    "image_path":     str(img_path),
                    "category":       category,
                    "article_number": art_num,
                })
            # update or insert results entry
            for r in self.results:
                if str(r.get("article_number", "")) == art_num:
                    r["category"] = category
                    break
            else:
                self.results.append({
                    "article_number": art_num,
                    "url":            self.csv_data[self.current_index]["url"],
                    "category":       category,
                })
        else:
            for e in self.categorized:
                if e.get("image_path") == str(img_path):
                    old_category = e["category"]
                    e["category"] = category
                    break
            else:
                self.categorized.append({"image_path": str(img_path), "category": category})

        # Övrigt retest — don't move files
        if self.retesting_ovrigt and category == "Övrigt":
            self.current_index += 1
            self._show_classify()
            return

        # ── move file if category changed, copy if new ────────────────────────
        if self.csv_mode and self.csv_data:
            meta = self.csv_data[self.current_index]
            base_name = f"{meta['article_number']}{img_path.suffix or '.jpg'}"
        else:
            base_name = img_path.name

        dest_dir = Path(f"{self.test_name}.{category}")
        dest_dir.mkdir(exist_ok=True)
        dest = dest_dir / base_name
        counter = 1
        while dest.exists() and dest != Path(f"{self.test_name}.{old_category}") / base_name:
            stem, suf = Path(base_name).stem, Path(base_name).suffix
            dest = dest_dir / f"{stem}_{counter}{suf}"
            counter += 1

        if old_category and old_category != category:
            old_file = Path(f"{self.test_name}.{old_category}") / base_name
            if old_file.exists():
                try:
                    shutil.move(str(old_file), dest)
                except Exception:
                    pass
        elif not old_category:
            try:
                if self.retesting_ovrigt:
                    shutil.move(str(img_path), dest)
                else:
                    shutil.copy2(img_path, dest)
            except Exception:
                pass

        self.current_index += 1
        self._show_classify()

    def _on_skip(self):
        self.current_index += 1
        self._show_classify()

    def _on_go_back(self):
        if self.current_index <= 0:
            return
        self.current_index -= 1
        self._show_classify()

    def _add_cat_during_test(self):
        if len(self.categories) >= 9:
            QMessageBox.warning(self, "Max antal", "Max 9 kategorier (tangent 1–9).")
            return
        dlg = QDialog(self)
        dlg.setWindowTitle("Ny kategori")
        dlg.setStyleSheet(STYLE)
        lay = QVBoxLayout(dlg)
        lay.addWidget(QLabel("Kategorinamn:"))
        edit = QLineEdit(); lay.addWidget(edit)
        btns = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok |
                                QDialogButtonBox.StandardButton.Cancel)
        btns.accepted.connect(dlg.accept); btns.rejected.connect(dlg.reject)
        lay.addWidget(btns); edit.setFocus()
        if dlg.exec() != QDialog.DialogCode.Accepted:
            return
        name = edit.text().strip()
        if not name:
            return
        if any(c["name"] == name for c in self.categories) or name == "Övrigt":
            QMessageBox.warning(self, "Dubblett", f'"{name}" finns redan.')
            return
        self.categories.append({"name": name, "description": "", "knowledge": ""})
        self._show_classify()

    # ── done screen ────────────────────────────────────────────────────────────

    def _show_done(self):
        self._cleanup_workers()
        ovrigt_dir = Path(f"{self.test_name}.Övrigt")
        ov_count = len([f for f in ovrigt_dir.iterdir()
                        if f.suffix.lower() in SUPPORTED_EXT]) \
                   if ovrigt_dir.exists() else 0
        self._done_scr.show_results(
            self.test_name, self.categories, self.current_index,
            self.csv_mode, bool(self.results), ov_count
        )
        self.stack.setCurrentWidget(self._done_scr)

    def _on_new_test(self):
        self._cleanup_workers()
        self._cleanup_temp()
        self._reset_state()
        self._name_scr.reset()
        self.stack.setCurrentWidget(self._name_scr)

    def _retest_ovrigt(self):
        ovrigt_dir = Path(f"{self.test_name}.Övrigt")
        imgs = sorted([f for f in ovrigt_dir.iterdir()
                       if f.suffix.lower() in SUPPORTED_EXT])
        if not imgs:
            QMessageBox.information(self, "Inga bilder", "Inga bilder i Övrigt-mappen.")
            return
        self.images = imgs
        self.current_index = 0
        self.retesting_ovrigt = True
        self.csv_mode = False
        self._show_classify()

    # ── AI job ─────────────────────────────────────────────────────────────────

    def _run_ai_job(self):
        if not self.ai_enabled:
            # Show settings dialog inline instead of a separate screen
            dlg = QDialog(self)
            dlg.setWindowTitle("AI-inställningar")
            dlg.setStyleSheet(STYLE)
            dlg.setFixedWidth(440)
            lay = QVBoxLayout(dlg)
            lay.setContentsMargins(28, 24, 28, 24)
            lay.setSpacing(0)

            t = QLabel("AI-inställningar")
            t.setStyleSheet("font-size:18px; font-weight:bold; color:#89b4fa;")
            t.setAlignment(Qt.AlignmentFlag.AlignCenter)
            lay.addWidget(t)
            sub = QLabel("Konfigurera LM Studio. Lämna fälten oförändrade för standardvärden.")
            sub.setStyleSheet("font-size:10px; color:#6c7086;")
            sub.setAlignment(Qt.AlignmentFlag.AlignCenter)
            sub.setWordWrap(True)
            lay.addWidget(sub)
            lay.addSpacing(20)

            lay.addWidget(QLabel("LM Studio URL:"))
            lay.addSpacing(3)
            url_edit = QLineEdit(self.ai_settings.get("api_url", DEFAULT_AI_URL))
            url_edit.setFixedHeight(34)
            lay.addWidget(url_edit)
            lay.addSpacing(12)

            lay.addWidget(QLabel("Modellnamn:"))
            lay.addSpacing(3)
            model_edit = QLineEdit(self.ai_settings.get("model", DEFAULT_MODEL))
            model_edit.setFixedHeight(34)
            lay.addWidget(model_edit)
            lay.addSpacing(10)

            compress_cb = QCheckBox("Komprimera bilder (snabbare, marginellt sämre precision)")
            compress_cb.setChecked(self.ai_settings.get("compress_images", True))
            lay.addWidget(compress_cb)
            lay.addSpacing(20)

            go = mk_btn("Kör AI jobb  →", "#89b4fa", "#1e1e2e", h=40)
            go.clicked.connect(dlg.accept)
            lay.addWidget(go)
            lay.addSpacing(6)
            cancel = mk_btn("Avbryt", "#45475a", "#cdd6f4", h=34)
            cancel.clicked.connect(dlg.reject)
            lay.addWidget(cancel)

            if dlg.exec() != QDialog.DialogCode.Accepted:
                return
            self.ai_settings = {
                "api_url":         url_edit.text().strip() or DEFAULT_AI_URL,
                "model":           model_edit.text().strip() or DEFAULT_MODEL,
                "compress_images": compress_cb.isChecked(),
            }
            self.ai_enabled = True

        if not self.categorized:
            QMessageBox.information(self, "Inga data",
                                    "Inga manuellt klassificerade artiklar att utgå från.")
            return

        scr = AIJobScreen(
            self.categories, self.categorized, self.csv_data, self.syfte,
            self.ai_settings.get("api_url", DEFAULT_AI_URL),
            self.ai_settings.get("model", DEFAULT_MODEL),
            self.ai_settings.get("compress_images", True),
            self.data_mgr, self.test_name,
        )
        scr.article_added.connect(self._on_ai_article_classified)
        scr.reclassified.connect(self._on_ai_reclassified)
        scr.finished.connect(self._show_done)
        self._push_screen(scr)
        scr.start()

    def _on_ai_article_classified(self, article_number: str, category: str, url: str):
        """Add an AI-classified article to results (if not already there)."""
        existing = {r["article_number"] for r in self.results}
        if article_number not in existing:
            bolag = next(
                (r.get("bolag", "") for r in self.csv_data
                 if str(r.get("article_number", "")) == article_number),
                ""
            )
            self.results.append({
                "article_number": article_number,
                "category":       category,
                "url":            url,
                "bolag":          bolag,
            })

    def _on_ai_reclassified(self, article_number: str, new_category: str):
        """Update result when user drags a card to a different column."""
        for r in self.results:
            if r["article_number"] == article_number:
                r["category"] = new_category
                break

    # ── ZIP session export ─────────────────────────────────────────────────────

    def _export_zip(self):
        import zipfile as _zip, json as _json
        path, _ = QFileDialog.getSaveFileName(
            self, "Spara session som ZIP",
            f"{self.test_name}_session.zip", "ZIP-filer (*.zip)"
        )
        if not path:
            return
        try:
            # Collect all unique image paths
            img_srcs: Dict[str, str] = {}  # orig_path → archive name

            def _register(src: str):
                if not src or src in img_srcs:
                    return
                p = Path(src)
                if not p.exists():
                    return
                name = p.name
                taken = set(img_srcs.values())
                if name in taken:
                    i = 1
                    while f"{p.stem}_{i}{p.suffix}" in taken:
                        i += 1
                    name = f"{p.stem}_{i}{p.suffix}"
                img_srcs[src] = f"images/{name}"

            for item in self.categorized:
                _register(item.get("image_path", ""))
            for row in self.csv_data:
                _register(row.get("img_path", ""))

            def _rel(src: str) -> str:
                return img_srcs.get(src, src)

            session = {
                "test_name": self.test_name,
                "syfte":     self.syfte,
                "categories": self.categories,
            }
            csv_export = [
                {**r, "img_path": _rel(r.get("img_path", ""))}
                for r in self.csv_data
            ]
            cat_export = [
                {**c, "image_path": _rel(c.get("image_path", ""))}
                for c in self.categorized
            ]

            with _zip.ZipFile(path, "w", _zip.ZIP_DEFLATED) as zf:
                zf.writestr("session.json",
                            _json.dumps(session, ensure_ascii=False, indent=2))
                zf.writestr("csv_data.json",
                            _json.dumps(csv_export, ensure_ascii=False, indent=2))
                zf.writestr("categorized.json",
                            _json.dumps(cat_export, ensure_ascii=False, indent=2))
                if self.results:
                    zf.writestr("results.json",
                                _json.dumps(self.results, ensure_ascii=False, indent=2))
                for orig, arc in img_srcs.items():
                    zf.write(orig, arc)

            QMessageBox.information(self, "Session sparad", f"ZIP sparad:\n{path}")
        except Exception as e:
            QMessageBox.critical(self, "Fel", f"Kunde inte skapa ZIP:\n{e}")

    # ── ZIP session import ─────────────────────────────────────────────────────

    def _import_zip(self):
        import zipfile as _zip, json as _json
        path, _ = QFileDialog.getOpenFileName(
            self, "Öppna sparad session", "", "ZIP-filer (*.zip)"
        )
        if not path:
            return
        try:
            extract_dir = Path(tempfile.mkdtemp(prefix="bildklassificering_session_"))
            with _zip.ZipFile(path, "r") as zf:
                zf.extractall(extract_dir)

            def _load(name: str, default):
                p = extract_dir / name
                if p.exists():
                    with open(p, encoding="utf-8") as f:
                        return _json.load(f)
                return default

            session    = _load("session.json", {})
            csv_data   = _load("csv_data.json", [])
            categorized = _load("categorized.json", [])
            results    = _load("results.json", [])

            def _fix(p: str) -> str:
                if not p:
                    return p
                full = extract_dir / p
                return str(full) if full.exists() else p

            for item in categorized:
                item["image_path"] = _fix(item.get("image_path", ""))
            for row in csv_data:
                row["img_path"] = _fix(row.get("img_path", ""))

            self._cleanup_workers()
            self._cleanup_temp()
            self._reset_state()
            self.temp_dir = str(extract_dir)

            self.test_name   = session.get("test_name", "Import")
            self.syfte       = session.get("syfte", "")
            self.categories  = session.get("categories", [])
            self.csv_data    = csv_data
            self.csv_mode    = bool(csv_data)
            self.categorized = categorized
            self.results     = results
            self.images      = [Path(r["img_path"]) if r.get("img_path") else None
                                 for r in csv_data]
            self.current_index = len(self.images)  # past end → no manual classify

            # Update static screens
            self._name_scr.name_edit.setText(self.test_name)
            self._cat_scr.set_test_name(self.test_name)

            if results:
                self._open_resumed_session()
            else:
                self._show_classify()

        except Exception as e:
            QMessageBox.critical(self, "Fel", f"Kunde inte läsa session:\n{e}")

    def _open_resumed_session(self):
        """Open AI job screen pre-populated with all results, no worker started."""
        img_by_art = {str(r.get("article_number", "")): r.get("img_path", "")
                      for r in self.csv_data}
        art_in_cat = {str(c.get("article_number", "")) for c in self.categorized}

        merged = list(self.categorized)
        for r in self.results:
            art = str(r.get("article_number", ""))
            if art not in art_in_cat:
                merged.append({
                    "article_number": art,
                    "category":   r.get("category", "Övrigt"),
                    "image_path": img_by_art.get(art, ""),
                    "url":        r.get("url", ""),
                    "bolag":      r.get("bolag", ""),
                })

        scr = AIJobScreen(
            self.categories, merged, self.csv_data, self.syfte,
            self.ai_settings.get("api_url", DEFAULT_AI_URL),
            self.ai_settings.get("model", DEFAULT_MODEL),
            self.ai_settings.get("compress_images", True),
            self.data_mgr, self.test_name,
        )
        scr.article_added.connect(self._on_ai_article_classified)
        scr.reclassified.connect(self._on_ai_reclassified)
        scr.finished.connect(self._show_done)
        self._push_screen(scr)
        scr.start(skip_worker=True)

    # ── Excel export ───────────────────────────────────────────────────────────

    def _export_excel(self):
        if not OPENPYXL_AVAILABLE:
            QMessageBox.critical(self, "openpyxl saknas",
                                 "Installera openpyxl:\n  pip install openpyxl")
            return
        path, _ = QFileDialog.getSaveFileName(
            self, "Spara Excel", f"{self.test_name}_resultat.xlsx", "Excel (*.xlsx)"
        )
        if not path:
            return
        wb = openpyxl.Workbook(); ws = wb.active; ws.title = "Resultat"
        headers = [
            "Artikelnummer", "Resultat kategori", "Huvudkategori",
            "Beskrivning", "Längd (mm)", "Bredd (mm)", "Höjd (mm)",
            "Volym", "Vikt brutto (kg)", "Vikt netto (kg)",
            "Robot (Y/N)", "StoreQuantity", "Bild (URL)",
        ]
        ws.append(headers)
        for row in self.results:
            art = str(row.get("article_number", ""))
            meta = self.data_mgr.get_meta(art, row.get("bolag", "")) or {}
            ws.append([
                art,
                row.get("category", ""),
                meta.get("huvudkategori", ""),
                meta.get("beskrivning", ""),
                meta.get("langd", ""),
                meta.get("bredd", ""),
                meta.get("hojd", ""),
                meta.get("volym", ""),
                meta.get("vikt_brutto", ""),
                meta.get("vikt_netto", ""),
                meta.get("robot", ""),
                meta.get("store_quantity", ""),
                row.get("url", ""),
            ])
        col_widths = [20, 25, 25, 40, 12, 12, 12, 12, 18, 18, 12, 15, 60]
        for i, w in enumerate(col_widths, 1):
            ws.column_dimensions[openpyxl.utils.get_column_letter(i)].width = w
        try:
            wb.save(path)
            QMessageBox.information(self, "Exporterat", f"Sparad:\n{path}")
        except Exception as e:
            QMessageBox.critical(self, "Fel", f"Kunde inte spara:\n{e}")

    # ── cleanup ────────────────────────────────────────────────────────────────

    def _cleanup_workers(self):
        if self.dl_worker:
            self.dl_worker.stop(); self.dl_worker.wait(); self.dl_worker = None

    def _cleanup_temp(self):
        if self.temp_dir and Path(self.temp_dir).exists():
            shutil.rmtree(self.temp_dir, ignore_errors=True)
        self.temp_dir = None

    def _reset_state(self):
        self.test_name = ""; self.syfte = ""; self.categories = []
        self.images = []; self.current_index = 0
        self.csv_mode = False; self.csv_data = []; self.results = []
        self.retesting_ovrigt = False; self.categorized = []
        self.ai_settings = {}; self.ai_enabled = False
        self._ready_images = set()

    def closeEvent(self, event):
        self._cleanup_workers()
        self._cleanup_temp()
        super().closeEvent(event)


# ── entry point ────────────────────────────────────────────────────────────────
def main():
    app = QApplication(sys.argv)
    app.setStyle("Fusion")
    win = MainApp()
    win.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
