# lpa1.2.py — LPA Canvas med "nästa-lagerplats"-förslag via ordningslista
# Funktioner:
#  - Interaktiv canvas: lägg till platser, koppla med avstånd, lägg till "nästa" i kedja
#  - Zoom (mushjul), panorera (höger musknapp), flytta noder (vänster-dra), centera vy
#  - Import/Export: CSV (Från lokation, Till lokation, Avstånd), spara/öppna projekt (JSON)
#  - Dubbelklick på nod för att byta namn
#  - NYTT: Ladda "Ordningslista…" (CSV) och få förslag på nästa platsnamn i kedja (AA66 → AA67, AA2 → AA3, …)
#           Fallback: automatisk inkrement av sifferdel (behåller nollpadding)

import tkinter as tk
from tkinter import ttk, filedialog, messagebox, simpledialog
import json
import math
import re
import pandas as pd

# ===================== Modell =====================

class GraphModel:
    def __init__(self):
        self.nodes = {}         # nid -> {"name": str, "x": float, "y": float}
        self.edges = {}         # eid -> {"a": nid, "b": nid, "dist": float}
        self._nid = 0
        self._eid = 0
        self.name_to_nid = {}

    def add_node(self, name, x, y):
        if name in self.name_to_nid:
            raise ValueError(f"Plats '{name}' finns redan.")
        nid = self._nid; self._nid += 1
        self.nodes[nid] = {"name": name, "x": float(x), "y": float(y)}
        self.name_to_nid[name] = nid
        return nid

    def rename_node(self, nid, new_name):
        old = self.nodes[nid]["name"]
        if new_name != old and new_name in self.name_to_nid:
            raise ValueError(f"Plats '{new_name}' finns redan.")
        del self.name_to_nid[old]
        self.nodes[nid]["name"] = new_name
        self.name_to_nid[new_name] = nid

    def delete_node(self, nid):
        to_del = [eid for eid,e in self.edges.items() if e["a"]==nid or e["b"]==nid]
        for eid in to_del: del self.edges[eid]
        name = self.nodes[nid]["name"]
        del self.nodes[nid]
        if name in self.name_to_nid: del self.name_to_nid[name]

    def add_edge(self, nid_a, nid_b, dist):
        if nid_a == nid_b:
            raise ValueError("Kan inte koppla en nod till sig själv.")
        # ersätt ev. existerande koppling mellan samma noder
        for eid, e in list(self.edges.items()):
            if {e["a"], e["b"]} == {nid_a, nid_b}:
                del self.edges[eid]
        eid = self._eid; self._eid += 1
        self.edges[eid] = {"a": nid_a, "b": nid_b, "dist": float(dist)}
        return eid

    def delete_edge_between(self, nid_a, nid_b):
        for eid, e in list(self.edges.items()):
            if {e["a"], e["b"]} == {nid_a, nid_b}:
                del self.edges[eid]

    def to_csv_dataframe(self):
        rows = []
        for e in self.edges.values():
            a = self.nodes[e["a"]]["name"]
            b = self.nodes[e["b"]]["name"]
            rows.append({"Från lokation": a, "Till lokation": b, "Avstånd": e["dist"]})
        return pd.DataFrame(rows)

    def clear(self):
        self.nodes.clear(); self.edges.clear()
        self._nid = 0; self._eid = 0; self.name_to_nid.clear()

    def to_json(self):
        return json.dumps({
            "nodes": self.nodes,
            "edges": self.edges,
            "_nid": self._nid,
            "_eid": self._eid
        }, ensure_ascii=False, indent=2)

    def from_json(self, s):
        d = json.loads(s)
        self.nodes = {int(k):v for k,v in d["nodes"].items()}
        self.edges = {int(k):v for k,v in d["edges"].items()}
        self._nid = int(d.get("_nid", len(self.nodes)))
        self._eid = int(d.get("_eid", len(self.edges)))
        self.name_to_nid = {v["name"]: int(k) for k,v in self.nodes.items()}

# ===================== Canvas-vy =====================

class CanvasView(tk.Frame):
    R = 18
    FONT = ("Segoe UI", 10, "bold")
    EDGE_FONT = ("Segoe UI", 9)
    HIT_RADIUS = 16

    def __init__(self, master, model: GraphModel, suggester=None):
        super().__init__(master)
        self.model = model
        # callback som ger förslag på nästa namn (kan vara None)
        self.suggester = suggester

        # världs->skärm transform
        self.scale = 1.0
        self.tx = 80
        self.ty = 200

        self.mode = tk.StringVar(value="select")
        self.snap_line = tk.BooleanVar(value=False)      # håll y=0
        self.respect_length = tk.BooleanVar(value=True)  # “kedja” på exakt avstånd

        # Toolbar
        toolbar = tk.Frame(self); toolbar.pack(side="top", fill="x")
        def add_btn(text, mode):
            ttk.Radiobutton(toolbar, text=text, value=mode, variable=self.mode).pack(side="left", padx=3)
        add_btn("Markera/Flytta", "select")
        add_btn("Lägg till plats", "add_node")
        add_btn("Koppla (ange avstånd)", "link")
        add_btn("Lägg till nästa (kedja)", "add_next")
        add_btn("Radera", "delete")
        ttk.Checkbutton(toolbar, text="Håll på linje (y=0)", variable=self.snap_line).pack(side="left", padx=12)
        ttk.Checkbutton(toolbar, text="Respektera avstånd vid kedja", variable=self.respect_length).pack(side="left")

        right = tk.Frame(self); right.pack(side="bottom", fill="x")
        ttk.Button(right, text="Centera vy", command=self.reset_view).pack(side="left", padx=8)

        self.canvas = tk.Canvas(self, bg="#fafafa", width=1100, height=600)
        self.canvas.pack(fill="both", expand=True)

        self.node_draw = {}  # nid -> {"oval": id, "text": id}
        self.edge_draw = {}  # eid -> {"line": id, "text": id}

        # interaktionstillstånd
        self.dragging_nid = None
        self.drag_offset = (0,0)
        self.selected_nid = None
        self.link_first_nid = None
        self.pan_active = False
        self.pan_start = (0,0)
        self.last_added_index = 1  # auto-namn vid fri "Lägg till plats"

        # binds
        c = self.canvas
        c.bind("<Button-1>", self.on_left_down)
        c.bind("<B1-Motion>", self.on_left_drag)
        c.bind("<ButtonRelease-1>", self.on_left_up)
        c.bind("<Button-3>", self.on_right_down)
        c.bind("<B3-Motion>", self.on_right_drag)
        c.bind("<ButtonRelease-3>", self.on_right_up)
        c.bind("<MouseWheel>", self.on_wheel)        # Windows
        c.bind("<Button-4>", self.on_wheel_linux)    # Linux scroll up
        c.bind("<Button-5>", self.on_wheel_linux)    # Linux scroll down
        c.bind("<Double-Button-1>", self.on_double_click)

        self.pack(fill="both", expand=True)
        self.redraw_all()

    # ----- koordinater -----
    def world_to_screen(self, x, y):
        return (self.scale * x + self.tx, self.scale * y + self.ty)

    def screen_to_world(self, sx, sy):
        return ((sx - self.tx) / self.scale, (sy - self.ty) / self.scale)

    # ----- rita -----
    def redraw_all(self):
        c = self.canvas
        c.delete("all")
        self.node_draw.clear(); self.edge_draw.clear()
        # kanter
        for eid, e in self.model.edges.items():
            ax, ay = self.model.nodes[e["a"]]["x"], self.model.nodes[e["a"]]["y"]
            bx, by = self.model.nodes[e["b"]]["x"], self.model.nodes[e["b"]]["y"]
            sa, sb = self.world_to_screen(ax, ay), self.world_to_screen(bx, by)
            lid = c.create_line(*sa, *sb, width=2, fill="#333")
            mx = (sa[0] + sb[0]) / 2; my = (sa[1] + sb[1]) / 2
            tid = c.create_text(mx, my - 10, text=self._fmt(e["dist"]), font=self.EDGE_FONT)
            self.edge_draw[eid] = {"line": lid, "text": tid}
        # noder
        for nid, n in self.model.nodes.items():
            sx, sy = self.world_to_screen(n["x"], n["y"])
            oid = c.create_oval(sx-self.R, sy-self.R, sx+self.R, sy+self.R, fill="#9BD3F3", outline="#366")
            tid = c.create_text(sx, sy, text=n["name"], font=self.FONT)
            self.node_draw[nid] = {"oval": oid, "text": tid}
        if self.selected_nid is not None and self.selected_nid in self.node_draw:
            self._highlight(self.selected_nid)

    def _fmt(self, v):
        try:
            return f"{float(v):.2f}".rstrip("0").rstrip(".")
        except:
            return str(v)

    def _highlight(self, nid):
        c = self.canvas
        sx, sy = self.world_to_screen(self.model.nodes[nid]["x"], self.model.nodes[nid]["y"])
        c.create_oval(sx-self.R-4, sy-self.R-4, sx+self.R+4, sy+self.R+4, outline="#ff5a00", width=2)

    # ----- hit-test -----
    def pick_node(self, sx, sy):
        best = None; best_d2 = (self.HIT_RADIUS+1)**2
        for nid, n in self.model.nodes.items():
            x, y = self.world_to_screen(n["x"], n["y"])
            d2 = (sx-x)**2 + (sy-y)**2
            if d2 < best_d2:
                best_d2 = d2; best = nid
        return best

    def pick_edge(self, sx, sy):
        best = None; best_dist = 8.0
        for eid, e in self.model.edges.items():
            a = self.model.nodes[e["a"]]; b = self.model.nodes[e["b"]]
            ax, ay = self.world_to_screen(a["x"], a["y"])
            bx, by = self.world_to_screen(b["x"], b["y"])
            d = self.point_segment_distance(sx, sy, ax, ay, bx, by)
            if d < best_dist:
                best_dist = d; best = eid
        return best

    @staticmethod
    def point_segment_distance(px, py, x1, y1, x2, y2):
        dx, dy = x2-x1, y2-y1
        if dx==dy==0: return math.hypot(px-x1, py-y1)
        t = ((px-x1)*dx + (py-y1)*dy) / (dx*dx + dy*dy)
        t = max(0, min(1, t))
        cx, cy = x1 + t*dx, y1 + t*dy
        return math.hypot(px-cx, py-cy)

    # ----- interaktion -----
    def on_left_down(self, e):
        sx, sy = e.x, e.y
        mode = self.mode.get()
        nid = self.pick_node(sx, sy)

        if mode == "select":
            if nid is not None:
                self.selected_nid = nid
                self.dragging_nid = nid
                nx, ny = self.model.nodes[nid]["x"], self.model.nodes[nid]["y"]
                wx, wy = self.screen_to_world(sx, sy)
                self.drag_offset = (nx - wx, ny - wy)
            else:
                self.selected_nid = None
            self.redraw_all()

        elif mode == "add_node":
            wx, wy = self.screen_to_world(sx, sy)
            if self.snap_line.get(): wy = 0.0
            default_name = f"LP{self.last_added_index}"; self.last_added_index += 1
            name = simpledialog.askstring("Ny plats", "Namn på plats:", initialvalue=default_name, parent=self)
            if not name: return
            try:
                self.model.add_node(name.strip(), wx, wy)
            except ValueError as ex:
                messagebox.showerror("Fel", str(ex)); return
            self.selected_nid = self.model.name_to_nid[name.strip()]
            self.redraw_all()

        elif mode == "link":
            if nid is None: return
            if self.link_first_nid is None:
                self.link_first_nid = nid
                self.selected_nid = nid
                self.redraw_all()
            else:
                if nid == self.link_first_nid: return
                dist = simpledialog.askfloat(
                    "Avstånd",
                    f"Ange avstånd mellan {self.model.nodes[self.link_first_nid]['name']} och {self.model.nodes[nid]['name']}:",
                    minvalue=0.0, parent=self
                )
                if dist is None:
                    self.link_first_nid = None; return
                try:
                    self.model.add_edge(self.link_first_nid, nid, dist)
                except ValueError as ex:
                    messagebox.showerror("Fel", str(ex))
                self.link_first_nid = None
                self.redraw_all()

        elif mode == "add_next":
            if nid is None: return
            anchor = nid
            # ============= NYTT: förslag på nästa namn =============
            suggested = None
            if callable(self.suggester):
                suggested = self.suggester(self.model.nodes[anchor]["name"], used=set(self.model.name_to_nid.keys()))
            if not suggested:
                suggested = f"LP{self.last_added_index}"
            # ======================================================
            name = simpledialog.askstring("Ny plats", "Namn på nya platsen:", initialvalue=suggested, parent=self)
            if not name:
                return
            self.last_added_index += 1
            dist = simpledialog.askfloat("Avstånd", f"Avstånd från {self.model.nodes[anchor]['name']} till {name}:", minvalue=0.0, parent=self)
            if dist is None:
                return
            ax, ay = self.model.nodes[anchor]["x"], self.model.nodes[anchor]["y"]
            nx = ax + (dist if self.respect_length.get() else dist*0.5)
            ny = 0.0 if self.snap_line.get() else ay
            try:
                new_id = self.model.add_node(name.strip(), nx, ny)
                self.model.add_edge(anchor, new_id, dist)
            except ValueError as ex:
                messagebox.showerror("Fel", str(ex)); return
            self.selected_nid = new_id
            self.redraw_all()

        elif mode == "delete":
            if nid is not None:
                if messagebox.askyesno("Radera plats", f"Radera {self.model.nodes[nid]['name']} + dess kopplingar?"):
                    self.model.delete_node(nid)
                    if self.selected_nid == nid: self.selected_nid = None
                    self.redraw_all()
            else:
                eid = self.pick_edge(sx, sy)
                if eid is not None:
                    a = self.model.nodes[self.model.edges[eid]["a"]]["name"]
                    b = self.model.nodes[self.model.edges[eid]["b"]]["name"]
                    if messagebox.askyesno("Radera kant", f"Radera kopplingen {a} — {b}?"):
                        del self.model.edges[eid]; self.redraw_all()

    def on_left_drag(self, e):
        if self.mode.get() != "select": return
        if self.dragging_nid is None: return
        wx, wy = self.screen_to_world(e.x, e.y)
        nx = wx + self.drag_offset[0]
        ny = 0.0 if self.snap_line.get() else (wy + self.drag_offset[1])
        self.model.nodes[self.dragging_nid]["x"] = nx
        self.model.nodes[self.dragging_nid]["y"] = ny
        self.redraw_all()

    def on_left_up(self, e):
        self.dragging_nid = None

    def on_double_click(self, e):
        """Byt namn på nod via dubbelklick."""
        nid = self.pick_node(e.x, e.y)
        if nid is None:
            return
        current = self.model.nodes[nid]["name"]
        new_name = simpledialog.askstring(
            "Byt namn", f"Nytt namn för {current}:", initialvalue=current, parent=self
        )
        if not new_name or new_name.strip() == current:
            return
        try:
            self.model.rename_node(nid, new_name.strip())
            self.redraw_all()
        except ValueError as ex:
            messagebox.showerror("Kan inte byta namn", str(ex))

    # pan/zoom (höger musknapp + mushjul)
    def on_right_down(self, e):
        self.pan_active = True
        self.pan_start = (e.x, e.y)

    def on_right_drag(self, e):
        if not self.pan_active: return
        dx = e.x - self.pan_start[0]
        dy = e.y - self.pan_start[1]
        self.tx += dx; self.ty += dy
        self.pan_start = (e.x, e.y)
        self.redraw_all()

    def on_right_up(self, e):
        self.pan_active = False

    def on_wheel(self, e):
        factor = 1.1 if e.delta > 0 else 1/1.1
        self.zoom_at(e.x, e.y, factor)

    def on_wheel_linux(self, e):
        factor = 1.1 if e.num == 4 else 1/1.1
        self.zoom_at(e.x, e.y, factor)

    def zoom_at(self, sx, sy, factor):
        wx, wy = self.screen_to_world(sx, sy)
        self.scale *= factor
        sx2, sy2 = self.world_to_screen(wx, wy)
        self.tx += (sx - sx2)
        self.ty += (sy - sy2)
        self.redraw_all()

    def reset_view(self):
        if not self.model.nodes:
            return
        xs = [n["x"] for n in self.model.nodes.values()]
        ys = [n["y"] for n in self.model.nodes.values()]
        xmin, xmax = min(xs), max(xs); ymin, ymax = min(ys), max(ys)
        pad = 1.0
        w = (xmax - xmin) + pad*2
        h = (ymax - ymin) + pad*2
        cw = self.canvas.winfo_width() or 1100
        ch = self.canvas.winfo_height() or 600
        self.scale = min((cw*0.8)/max(w, 1e-6), (ch*0.6)/max(h, 1e-6))
        self.tx = cw/2 - self.scale * ((xmin+xmax)/2)
        self.ty = ch*0.7 - self.scale * ((ymin+ymax)/2)
        self.redraw_all()

# ===================== App / Meny =====================

def read_csv_safely(path):
    seps = [None, ",", ";", "\t", "|"]
    last_err = None
    for sep in seps:
        try:
            df = pd.read_csv(path, sep=sep, engine="python")
            if df.shape[1] >= 1:
                return df
        except Exception as e:
            last_err = e
            continue
    raise last_err if last_err else RuntimeError("Kunde inte läsa CSV.")

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("LPA Canvas – bygg din lagerplatskarta (v1.2)")
        self.geometry("1200x760")

        self.model = GraphModel()
        self.order_list = []      # ordningslista (valfri)
        self.order_index = {}     # name -> index i ordningslista

        self.view = CanvasView(self, self.model, suggester=self.suggest_next_name)

        menubar = tk.Menu(self)
        filem = tk.Menu(menubar, tearoff=0)
        filem.add_command(label="Ny", command=self.new_project)
        filem.add_command(label="Öppna CSV…", command=self.import_csv)
        filem.add_command(label="Spara CSV…", command=self.export_csv)
        filem.add_separator()
        filem.add_command(label="Spara projekt (JSON)…", command=self.save_project)
        filem.add_command(label="Öppna projekt (JSON)…", command=self.load_project)
        filem.add_separator()
        filem.add_command(label="Avsluta", command=self.destroy)
        menubar.add_cascade(label="Arkiv", menu=filem)

        toolm = tk.Menu(menubar, tearoff=0)
        toolm.add_command(label="Öppna ordningslista…", command=self.load_order_list)
        toolm.add_command(label="Visa ordningslista", command=self.show_order_list)
        menubar.add_cascade(label="Verktyg", menu=toolm)

        helpm = tk.Menu(menubar, tearoff=0)
        helpm.add_command(label="Hjälp", command=self.show_help)
        menubar.add_cascade(label="Hjälp", menu=helpm)

        self.config(menu=menubar)

    # -------- Ordningslista & förslag --------
    def load_order_list(self):
        path = filedialog.askopenfilename(
            title="Öppna ordningslista (CSV)",
            filetypes=[("CSV-filer", "*.csv"), ("Alla filer", "*.*")]
        )
        if not path: return
        try:
            df = read_csv_safely(path)
            # Hitta lämplig kolumn – prio: 'plats','lagerplats','location','name','namn', annars första
            cand = {c.lower(): c for c in df.columns}
            pick = None
            for k in ["plats", "lagerplats", "location", "name", "namn"]:
                if k in cand: pick = cand[k]; break
            if pick is None:
                pick = df.columns[0]
            seq = [str(x).strip() for x in df[pick].dropna().astype(str).tolist() if str(x).strip()]
            if not seq:
                raise ValueError("Hittade inga namn i vald kolumn.")
            self.order_list = seq
            self.order_index = {name: i for i, name in enumerate(self.order_list)}
            messagebox.showinfo("Ordningslista", f"Laddade {len(self.order_list)} namn.\nExempel: {self.order_list[:5]} …")
        except Exception as e:
            messagebox.showerror("Fel vid läsning", str(e))

    def show_order_list(self):
        if not self.order_list:
            messagebox.showinfo("Ordningslista", "Ingen ordningslista är laddad.")
        else:
            preview = ", ".join(self.order_list[:20]) + ("…" if len(self.order_list) > 20 else "")
            messagebox.showinfo("Ordningslista", f"Antal: {len(self.order_list)}\nFörsta: {preview}")

    @staticmethod
    def _increment_name_pattern(name: str) -> str:
        """
        Fallback: öka sista siffergruppen med bibehållen noll-padding.
        AA66 -> AA67, SK009 -> SK010, X -> X1
        """
        m = re.search(r'(.*?)(\d+)([^0-9]*)$', name)
        if not m:
            return name + "1"
        prefix, digits, suffix = m.group(1), m.group(2), m.group(3)
        newnum = str(int(digits) + 1).zfill(len(digits))
        return f"{prefix}{newnum}{suffix}"

    def suggest_next_name(self, current: str, used=None) -> str | None:
        """
        Förslag på nästa namn:
         1) Om ordningslista laddad och 'current' finns i den: ta nästa i listan.
         2) Annars: incrementa sifferdelen (AA66 -> AA67 etc.).
         3) Om förslaget redan används i bilden: hoppa vidare tills ledigt (max 200 steg).
        """
        if used is None: used = set()

        # 1) Ordningslista
        if current in self.order_index:
            idx = self.order_index[current] + 1
            if idx < len(self.order_list):
                cand = self.order_list[idx]
            else:
                cand = None
        else:
            cand = None

        # 2) Fallback
        if not cand:
            cand = self._increment_name_pattern(current)

        # 3) Gör förslaget ledigt genom att hoppa vidare
        steps = 0
        while cand in used and steps < 200:
            if cand in self.order_index:
                nxt_i = self.order_index[cand] + 1
                cand = self.order_list[nxt_i] if nxt_i < len(self.order_list) else self._increment_name_pattern(cand)
            else:
                cand = self._increment_name_pattern(cand)
            steps += 1

        return cand

    # -------------- Arkiv / projekt --------------
    def new_project(self):
        if messagebox.askyesno("Ny", "Rensa allt?"):
            self.model.clear()
            self.view.redraw_all()

    def import_csv(self):
        path = filedialog.askopenfilename(
            title="Öppna CSV",
            filetypes=[("CSV-filer", "*.csv"), ("Alla filer", "*.*")]
        )
        if not path: return
        try:
            df = read_csv_safely(path)
            cols = {c.lower(): c for c in df.columns}
            def pick(names):
                for nm in names:
                    if nm in cols: return cols[nm]
                for c in df.columns:
                    lc = c.lower()
                    for nm in names:
                        if nm in lc: return c
                return None
            s_col = pick(["från lokation","fran lokation","plats1","from","source"])
            t_col = pick(["till lokation","plats2","to","target"])
            d_col = pick(["avstånd","avstand","distance","distans","vikt","weight","kostnad","cost"])
            if not (s_col and t_col and d_col):
                raise ValueError("Hittade inte kolumnerna för från/till/avstånd i CSV.")
            df = df[[s_col, t_col, d_col]].copy()
            df.columns = ["Från lokation","Till lokation","Avstånd"]
            df["Avstånd"] = pd.to_numeric(df["Avstånd"], errors="coerce")
            df = df.dropna(subset=["Avstånd"])
        except Exception as e:
            messagebox.showerror("Fel vid läsning", str(e)); return

        self.model.clear()
        # skapa noder
        for name in pd.unique(pd.concat([df["Från lokation"], df["Till lokation"]], ignore_index=True)):
            try:
                self.model.add_node(str(name), 0, 0)
            except ValueError:
                pass
        # lägg kanter
        for _, r in df.iterrows():
            a = self.model.name_to_nid[str(r["Från lokation"])]
            b = self.model.name_to_nid[str(r["Till lokation"])]
            self.model.add_edge(a, b, float(r["Avstånd"]))

        # enkel linjär placering längs x
        deg = {nid:0 for nid in self.model.nodes}
        for e in self.model.edges.values():
            deg[e["a"]] += 1; deg[e["b"]] += 1
        start = None
        for nid, d in deg.items():
            if d == 1: start = nid; break
        if start is None:
            start = next(iter(self.model.nodes))
        placed = {start: (0.0, 0.0)}
        queue = [start]
        while queue:
            u = queue.pop(0)
            ux, uy = placed[u]
            for e in self.model.edges.values():
                v = None
                if e["a"] == u: v = e["b"]
                elif e["b"] == u: v = e["a"]
                if v is None or v in placed: continue
                placed[v] = (ux + e["dist"], 0.0)
                queue.append(v)
        for nid, (x,y) in placed.items():
            self.model.nodes[nid]["x"] = x; self.model.nodes[nid]["y"] = y
        self.view.reset_view()

    def export_csv(self):
        if not self.model.edges:
            messagebox.showwarning("Tomt", "Det finns inga kopplingar att exportera.")
            return
        path = filedialog.asksaveasfilename(
            title="Spara CSV", defaultextension=".csv",
            filetypes=[("CSV-filer", "*.csv")]
        )
        if not path: return
        try:
            df = self.model.to_csv_dataframe()
            df.to_csv(path, index=False)
            messagebox.showinfo("Sparad", f"Sparade {len[df]} rader till:\n{path}")
        except Exception as e:
            messagebox.showerror("Kunde inte spara", str(e))

    def save_project(self):
        path = filedialog.asksaveasfilename(
            title="Spara projekt (JSON)", defaultextension=".json",
            filetypes=[("JSON", "*.json")]
        )
        if not path: return
        try:
            with open(path, "w", encoding="utf-8") as f:
                f.write(self.model.to_json())
            messagebox.showinfo("Sparad", f"Projekt sparat till:\n{path}")
        except Exception as e:
            messagebox.showerror("Kunde inte spara", str(e))

    def load_project(self):
        path = filedialog.askopenfilename(
            title="Öppna projekt (JSON)", filetypes=[("JSON", "*.json"), ("Alla filer", "*.*")]
        )
        if not path: return
        try:
            with open(path, "r", encoding="utf-8") as f:
                self.model.from_json(f.read())
            self.view.reset_view()
        except Exception as e:
            messagebox.showerror("Fel vid öppning", str(e))

    def show_help(self):
        messagebox.showinfo(
            "Hjälp",
            "Lägestyper:\n"
            "• Markera/Flytta: vänster-dra för att flytta noder. Höger-dra för att panorera. Mushjul zoomar.\n"
            "• Lägg till plats: klicka för att skapa nod. Dubbelklicka på nod för att byta namn.\n"
            "• Koppla (ange avstånd): klicka nod A, klicka nod B, skriv avstånd.\n"
            "• Lägg till nästa (kedja): klicka en nod → namn föreslås via ordningslista (om laddad) eller AA66→AA67.\n"
            "• Radera: klicka nod (raderar även kopplingar) eller klicka nära en linje för att radera den.\n\n"
            "Arkiv: Importera/Exportera CSV samt Spara/Öppna projekt (JSON).\n"
            "Verktyg: Öppna ordningslista (CSV med en kolumn med namn i rätt ordning)."
        )

if __name__ == "__main__":
    App().mainloop()
