"""
Order/Saldo analyslogik – utan GUI-beroenden.
Extraherad från OrderSaldo5.py.
"""
from __future__ import annotations

import re
import base64
import gzip
from typing import Optional

import pandas as pd


# ── Förbjudna artiklar (inbäddade i komprimerad form) ──────────────────────────

_RESTRICTED_ARTICLES_B64: str = (
    "H4sIAALgJmkC/2WcW7LtKAhAJ9TVFUVe859Ya5Rlzu6692MlKhhANJp92vM8I59/n3/apLAocoXsUKpA955val0HpJBBVU9LR39iQKdeN+ubRNqlAe1669+lPBRj96+1lltH0zZOqWYf0GlrZlqUu38tTE5p5NHbZgODTr3UyKJjv5amA6pSdKTH7l+Xnm2Tdx+QQ3Go"
    "uPbc1sO+/rzSey8AI9wAzlUEH5Bn9j5Z/XiOdH/zzOrZE/tdczmJrP7qjX90TJp3z7+a0M2brsa6T0/+uT/75/CU9c516tnXtfU0"
    "zm2BbN9Etv5A8uClfcpgzRDUmuMv/hx9UNx96a5Yd6PCwdEDi7Y2nlTg/68mc+Y4sV4eQUveC6qgr+eIz+epuQv9+3uQJh84iK9K"
    "wX+z9u2SZYQWzutF8wvb8NCv7X8KYXelyYuUiMSj1EUxhGJoyQg+z9Y2Q0vLMpZW4xdPUytNi5r2eIYfvdYE+S1vjGdrQJLryT7T"
    "p+qFd4gHjkMV6LMRTsHEF0QKW408c4pJQvPjJ/k5h+/j587m/Eu0M3O0cv59/xE6ir4fiQ4HAcIg0UA09XjNBxLx13d7y5HStYxY"
    "nXaZ6PSxwm2sTVD9UHB6ZVrkjb6zY9Opm4bLEkofPwrYoG3fna2p39GjJ1l/9WVzv+30iXgi8hQ+2KCYawJskL2YT3Qt8WIefhD9I"
    "zovHfwmfi38e/Jj/T+FV/prxDzv+bsKovHbYXPHGHk1FEHwo+gA6wNSxg3O9LWVrSO/dD0cfJpnOIj6iX9oIeqX9I+mP6RydmRT8"
    "IHtvMd4z6J/Jb2oe6Gv9X/aFSfvFX5hfOf4V/Fi8pvWrdcKMWbeKV55OHysT289D3wT7o8/Cv0Tn4RDnkF3lFO0a3p4JG7pJkdr0"
    "wER39Ec9mPb37DUhfNf07/Ab+SvddNEPTyJL748G2t0/k+fQ/8Zvxf+feDfviJnscjbdUt4p78n3/Jmqu+SfgUC9wAW/Mt+yb31/"
    "vnqWwbwYfFmq3zi0/bPBy/V9fbi0YDnePhtceo/9AH6hHfGz+hP+oT/1PtjvutrbGV8dxO4N8C9+vODnqI+8SLl5kDBV3CeyZ0jj"
    "J/BH/Mb8R3+hd9yX/iWOy89goK/1N/rtp318gCyvOEfOt4uP9NrmERW4Kw39yHfG7kiLuQv4N8If0mfjb9X/Fq7NHvQTH/Scu1T7"
    "D6n7Jr/1Z+iFf/EeyN/wR/Qi+yGfyN/wR/Qh7+Z751M8E1yU/ws/shf8If0pn4P/S1wH/B3+h/9nftB/wSfFT8or/2v0L/0nvwJv"
    "9X/0D8Nf8Lv4XfC74HfB74LfBb8Lfhf8Lvhd8LPgZ8HPgp8FPwt+Fvws+Fnws+Bnwc+CnwU/C34W/Cz4WfCz4WfCz4WfCz4WfCz4"
    "WfCz4WfCz4WXxon3P8qn3X20A/ke5r9VF+E5+lf0vb5JfzCcwj9YH3Ys/CuRFwoWIv+ljpuREr33H2CIr5/fgUV7Hf0LGdyH7q8m"
    "zL25oZ8Z20YQVNjDrm3fgik9JlpOJP2kz7+aPnJoEfFn5u+E/zWMr5b+p+if7/lY+uf8X8rP0Dt2FvX3fc39ufq/2Q4hKg5D/b/6"
    "P/0HyM/yW9+afU/8bV8iX8MB/4d/sdL/vPh4Z/Z41/cV/U/jO/i3wp/pb+Bn/en/oz/MTO4rJ6Lh4HwM8HPyfsbkOn47/mP8+T/T"
    "/+Cl/k14EF9Ef8z3Htf7Pv9E+UzwLVwvf875/dz/1m/Ev/yf8MA/j//4b8AEAAA=="
)

try:
    _data = gzip.decompress(base64.b64decode(_RESTRICTED_ARTICLES_B64)).decode("utf-8")
    _RESTRICTED_ARTICLES: set = set(filter(None, _data.split(",")))
except Exception:
    _RESTRICTED_ARTICLES = set()


def load_restricted_articles() -> set:
    return set(_RESTRICTED_ARTICLES)


# ── Kolumnmappning ──────────────────────────────────────────────────────────────

def _normalize(s: str) -> str:
    s = s.lower()
    s = s.replace("å", "a").replace("ä", "a").replace("ö", "o")
    s = re.sub(r"[^a-z0-9]+", "", s)
    return s


CANDIDATES = {
    "order": ["ordernr", "ordernummer", "ordernumber", "orderno", "orderid", "order"],
    "article": ["artikel", "artikelnr", "artikelnummer", "artnr", "item", "itemno", "sku", "productcode"],
    "name": ["benamning", "artikelnamn", "namn", "produktnamn", "description", "productname"],
    "pick": ["plock", "saldo", "lager", "stock", "available", "availableqty", "qtyavailable"],
    "demand": ["bestallt", "bestalltantal", "bestalld", "antal", "quantity", "qty", "ordered", "orderqty"],
    "pickedqty": ["plockat", "plockad", "picked", "pickedqty", "qtypicked"],
    "pickloc": ["plockplats", "lagerplats", "bin", "bincode", "location", "loc", "picklocation"],
}

REQUIRED_KEYS = ("order", "article", "pick", "demand")


def auto_map_columns(df: pd.DataFrame) -> dict:
    cols = list(df.columns)
    norm_map = {_normalize(c): c for c in cols}

    mapping: dict = {
        "order": None, "article": None, "name": None,
        "pick": None, "demand": None, "pickedqty": None, "pickloc": None,
    }

    for key, cand_list in CANDIDATES.items():
        for cand in cand_list:
            if cand in norm_map:
                mapping[key] = norm_map[cand]
                break

    for key in mapping:
        if mapping[key]:
            continue
        for norm, original in norm_map.items():
            if any(norm.find(cand) >= 0 for cand in CANDIDATES.get(key, [])):
                mapping[key] = original
                break

    missing = [k for k in REQUIRED_KEYS if not mapping[k]]
    if missing:
        raise ValueError(
            "Kunde inte hitta obligatoriska kolumner: "
            + ", ".join(missing)
            + ". Byt gärna kolumnnamn till t.ex. 'Order nr', 'Artikel', 'Plock', 'Beställt'."
        )
    return mapping


def to_numeric_safe(series: pd.Series) -> pd.Series:
    return pd.to_numeric(series, errors="coerce").fillna(0)


def read_csv_flex(path: str) -> pd.DataFrame:
    seps = ["\t", ";", ","]
    last_err = None
    for sep in seps:
        try:
            df = pd.read_csv(path, sep=sep, engine="python")
            if df.shape[1] > 1:
                return df
        except Exception as e:
            last_err = e
    if last_err:
        raise last_err
    raise ValueError("Kunde inte läsa filen med vanliga separatorer (tab/;/,).")


def analyze(df: pd.DataFrame, mapping: dict) -> dict:
    """
    Kör analysen och returnerar ett dict med:
      - complete_orders: list[str]
      - holistic_short: pd.DataFrame
      - orders_1x1..1x4: list[str]
      - map_summary: str (beskrivning av funna kolumner)
    """
    m = mapping
    order_col = m["order"]
    article_col = m["article"]
    pick_col = m["pick"]
    demand_col = m["demand"]
    name_col = m.get("name")
    picked_col = m.get("pickedqty")

    df = df.copy()
    df[order_col] = df[order_col].astype(str)
    df[article_col] = df[article_col].astype(str)
    if name_col and name_col in df.columns:
        df[name_col] = df[name_col].astype(str)
    df[pick_col] = to_numeric_safe(df[pick_col])
    df[demand_col] = to_numeric_safe(df[demand_col])
    if picked_col and picked_col in df.columns:
        df[picked_col] = to_numeric_safe(df[picked_col])

    df["_enough_row"] = df[pick_col] >= df[demand_col]

    # Lista 1 – kompletta ordrar
    complete_mask = df.groupby(order_col)["_enough_row"].all()
    complete_orders = sorted(complete_mask[complete_mask].index.astype(str).tolist())

    # Lista 2 – artiklar att beställa
    demand_by_art = df.groupby(article_col)[demand_col].sum(min_count=1)
    stock_by_art = df.groupby(article_col)[pick_col].max()

    holistic = pd.DataFrame({
        "Total beställt": demand_by_art,
        "Tillgängligt saldo (Plock)": stock_by_art,
    }).fillna(0)
    holistic["Underskott"] = (holistic["Total beställt"] - holistic["Tillgängligt saldo (Plock)"]).clip(lower=0)
    holistic["Saldo räcker inte för allt"] = holistic["Underskott"] > 0

    if name_col and name_col in df.columns:
        first_names = df.drop_duplicates(subset=[article_col]).set_index(article_col)
        holistic["Benämning"] = first_names.get(name_col, "")
    else:
        holistic["Benämning"] = ""

    orders_per_article = (
        df.groupby(article_col)[order_col]
        .apply(lambda s: ", ".join(sorted(pd.unique(s.astype(str)))))
        .rename("Påverkade ordrar")
    )
    holistic = holistic.join(orders_per_article, how="left")
    holistic_short = holistic[holistic["Saldo räcker inte för allt"]].copy()
    holistic_short.index.name = "Artikel"

    # 1×N ordrar
    oap = df.groupby([order_col, article_col]).agg(
        demand_sum=(demand_col, "sum"),
        pick_max=(pick_col, "max"),
    )
    oap["enough"] = oap["pick_max"] >= oap["demand_sum"]
    lines_per_order = oap.groupby(level=0).size()
    qty_per_order = oap.groupby(level=0)["demand_sum"].sum()
    enough_per_order = oap.groupby(level=0)["enough"].all()

    def orders_1line_qty(n: int) -> list:
        cond = (lines_per_order == 1) & (qty_per_order == n) & (enough_per_order)
        return sorted(cond[cond].index.astype(str).tolist())

    orders_1x1 = orders_1line_qty(1)
    orders_1x2 = orders_1line_qty(2)
    orders_1x3 = orders_1line_qty(3)
    orders_1x4 = orders_1line_qty(4)

    # Filtrera förbjudna artiklar
    restricted_articles = load_restricted_articles()
    shortage_articles: set = set()
    if holistic_short is not None and not holistic_short.empty:
        shortage_articles = set(holistic_short.index.astype(str))

    if restricted_articles or shortage_articles:
        articles_by_order = df.groupby(order_col)[article_col].apply(lambda s: set(s.astype(str)))
        for lst_ref in [orders_1x1, orders_1x2, orders_1x3, orders_1x4]:
            filtered = []
            for o in lst_ref:
                arts = articles_by_order.get(o, set())
                if restricted_articles and not arts.isdisjoint(restricted_articles):
                    continue
                if shortage_articles and not arts.isdisjoint(shortage_articles):
                    continue
                filtered.append(o)
            lst_ref[:] = filtered

    return {
        "complete_orders": complete_orders,
        "holistic_short": holistic_short,
        "orders_1x1": orders_1x1,
        "orders_1x2": orders_1x2,
        "orders_1x3": orders_1x3,
        "orders_1x4": orders_1x4,
    }
