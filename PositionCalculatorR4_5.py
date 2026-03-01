#Position Calculator R4.5

"""
HOW TO ENABLE / DISABLE DEBUG LOGGING
-------------------------------------
Windows PowerShell:
  # enable for this PowerShell session
  $env:PCALC_DEBUG_LOG = "1"
  # disable (unset)
  Remove-Item Env:PCALC_DEBUG_LOG

Windows Command Prompt (cmd.exe):
  REM enable for this cmd session
  set PCALC_DEBUG_LOG=1
  REM disable (unset)
  set PCALC_DEBUG_LOG=

Linux / macOS (bash/zsh):
  # enable for this shell
  export PCALC_DEBUG_LOG=1
  # disable (unset)
  unset PCALC_DEBUG_LOG
"""
import os
import csv
import json
import math
import traceback
from datetime import datetime
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from typing import List, Optional, Tuple, Dict

APP_TITLE = "Position Calculator R4.5"

# ---------------------------------
# Logging & Config (robust)
# ---------------------------------
import tempfile

def _script_dir() -> str:
    try:
        return os.path.dirname(os.path.abspath(__file__))
    except Exception:
        return os.getcwd()

# Enable only when env var is set (default OFF)
ENABLE_LOG = str(os.environ.get("PCALC_DEBUG_LOG", "0")).strip().lower() in ("1", "true", "yes", "on")

# ----------------------------
# Config & log persistence
# ----------------------------
CONFIG_DIR = os.path.join(os.environ.get("LOCALAPPDATA", os.path.expanduser("~")), "PositionCalculator")
os.makedirs(CONFIG_DIR, exist_ok=True)

CONFIG_PATH = os.path.join(CONFIG_DIR, "config.json")
LOG_PATH    = os.path.join(CONFIG_DIR, "debug_log.txt")

CONFIG: Dict[str, str] = {}


def _ts() -> str:
    from datetime import datetime
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S.%f")[:-3]

def init_logger():
    if not ENABLE_LOG:
        return
    try:
        with open(LOG_PATH, "w", encoding="utf-8") as f:
            f.write(f"[{_ts()}] === APP START ===\n")
            f.write(f"[{_ts()}] Running from: {_script_dir()}\n")
            f.write(f"[{_ts()}] Working dir: {os.getcwd()}\n")
            f.write(f"[{_ts()}] Log path: {LOG_PATH}\n")
            f.write(f"[{_ts()}] Config path: {CONFIG_PATH}\n")
            f.flush()
    except Exception:
        # last-gasp: try temp
        try:
            tmp = os.path.join(tempfile.gettempdir(), "PositionCalculator_debug_log.txt")
            with open(tmp, "a", encoding="utf-8") as f:
                f.write(f"[{_ts()}] Logger fallback active\n")
        except Exception:
            pass

def log(msg: str):
    if not ENABLE_LOG:
        return
    try:
        with open(LOG_PATH, "a", encoding="utf-8") as f:
            f.write(f"[{_ts()}] {msg}\n")
            f.flush()
    except Exception:
        pass

def load_config():
    global CONFIG
    try:
        with open(CONFIG_PATH, "r", encoding="utf-8") as f:
            CONFIG = json.load(f)
    except Exception:
        CONFIG = {}

def save_config():
    try:
        with open(CONFIG_PATH, "w", encoding="utf-8") as f:
            json.dump(CONFIG, f, indent=2)
    except Exception:
        pass

# ----------------------------
# Export order persistence
# ----------------------------
def get_export_order_default() -> str:
    # 'top_down' (N..1) or 'bottom_up' (1..N)
    v = (CONFIG.get("export_order") or "top_down").strip().lower()
    return v if v in ("top_down", "bottom_up") else "top_down"

def set_export_order(v: str):
    v = (v or "").strip().lower()
    if v not in ("top_down", "bottom_up"):
        v = "top_down"
    CONFIG["export_order"] = v
    save_config()

def get_seed_start_md_default(mode_key: str) -> str:
    """
    Persisted optional start MD (from collar) used to seed instrument placement.
    If blank, the mode will compute start_md from the current Top Instrument logic.
    """
    k = f"{mode_key}_seed_start_md"
    return (CONFIG.get(k) or "").strip()

def set_seed_start_md(mode_key: str, v: str):
    k = f"{mode_key}_seed_start_md"
    CONFIG[k] = (v or "").strip()
    save_config()


# ---------------------------------
# Utils
# ---------------------------------
def column_letters(n: int) -> List[str]:
    letters = []
    i = 0
    while len(letters) < n:
        q = i
        s = ""
        while True:
            q, r = divmod(q, 26)
            s = chr(ord('A') + r) + s
            if q == 0:
                break
            q -= 1
        letters.append(s)
        i += 1
    return letters[:n]

def safe_float(x: str) -> Optional[float]:
    try:
        return float(x)
    except Exception:
        return None

# ---------------------------------
# Data
# ---------------------------------
class SurveyData:
    def __init__(self, headers: List[str], rows: List[List[str]]):
        self.headers = headers
        self.rows = rows

    @classmethod
    def from_csv(cls, path: str) -> "SurveyData":
        log(f"Reading survey file: {path}")
        with open(path, "r", encoding="utf-8", errors="ignore", newline="") as f:
            sample = f.read(8192)
            f.seek(0)
            try:
                reader = csv.reader(f, dialect=csv.Sniffer().sniff(sample))
                log("CSV dialect sniffed successfully")
            except Exception:
                f.seek(0)
                reader = csv.reader(f)
            rows = list(reader)
        if not rows:
            raise ValueError("CSV/Text file appears to be empty.")
        headers = rows[0]
        data = rows[1:]
        log(f"Loaded rows: {len(rows)} (headers + {len(data)} data rows)")
        return cls(headers, data)

# ---------------------------------
# Trajectory math
# ---------------------------------
def rad(d: float) -> float:
    return math.radians(d)

def wrap_angle_rad(a: float) -> float:
    twopi = 2 * math.pi
    a = a % twopi
    if a < 0:
        a += twopi
    return a

def _rf_min_curve(beta: float) -> float:
    if abs(beta) < 1e-9:
        return 1.0
    return (2.0 / beta) * math.tan(beta / 2.0)

def _segment_displacement(method: str,
                          dmd: float,
                          inc1: float, az1: float,
                          inc2: float, az2: float) -> Tuple[float, float, float]:
    if method == 'tan':
        dE = dmd * math.sin(inc1) * math.sin(az1)
        dN = dmd * math.sin(inc1) * math.cos(az1)
        dTVD = dmd * math.cos(inc1)
        return dE, dN, dTVD
    if method == 'min':
        cos_beta = (math.cos(inc1) * math.cos(inc2) +
                    math.sin(inc1) * math.sin(inc2) * math.cos(az2 - az1))
        cos_beta = max(-1.0, min(1.0, cos_beta))
        beta = math.acos(cos_beta)
        rf = _rf_min_curve(beta)
    else:
        rf = 1.0
    dE = 0.5 * dmd * ((math.sin(inc1) * math.sin(az1)) + (math.sin(inc2) * math.sin(az2))) * rf
    dN = 0.5 * dmd * ((math.sin(inc1) * math.cos(az1)) + (math.sin(inc2) * math.cos(az2))) * rf
    dTVD = 0.5 * dmd * (math.cos(inc1) + math.cos(inc2)) * rf
    return dE, dN, dTVD

class Trajectory:
    def __init__(self, mds: List[float], incs: List[float], azs: List[float], method: str = 'min'):
        self.method = method
        zipped = sorted(zip(mds, incs, azs), key=lambda t: t[0])
        self.mds = [t[0] for t in zipped]
        self.incs = [t[1] for t in zipped]
        self.azs = [wrap_angle_rad(t[2]) for t in zipped]

        self.E = [0.0]; self.N = [0.0]; self.TVD = [0.0]
        for i in range(1, len(self.mds)):
            dmd = self.mds[i] - self.mds[i-1]
            dE, dN, dTVD = _segment_displacement(self.method, dmd, self.incs[i-1], self.azs[i-1], self.incs[i], self.azs[i])
            self.E.append(self.E[-1] + dE)
            self.N.append(self.N[-1] + dN)
            self.TVD.append(self.TVD[-1] + dTVD)

    def tangent_at_start(self) -> Tuple[float, float, float]:
        inc0, az0 = self.incs[0], self.azs[0]
        return math.sin(inc0) * math.sin(az0), math.sin(inc0) * math.cos(az0), math.cos(inc0)

    def tangent_at_end(self) -> Tuple[float, float, float]:
        incL, azL = self.incs[-1], self.azs[-1]
        return math.sin(incL) * math.sin(azL), math.sin(incL) * math.cos(azL), math.cos(incL)

    def pos_at_md_rel(self, md: float) -> Tuple[float, float, float]:
        if md <= self.mds[0]:
            if md < 0 or len(self.mds) == 1:
                tE, tN, tTVD = self.tangent_at_start()
                return md * tE, md * tN, md * tTVD
            dmd = md - self.mds[0]
            tE, tN, tTVD = self.tangent_at_start()
            return self.E[0] + dmd * tE, self.N[0] + dmd * tN, self.TVD[0] + dmd * tTVD

        if md >= self.mds[-1]:
            tE, tN, tTVD = self.tangent_at_end()
            dmd = md - self.mds[-1]
            return self.E[-1] + dmd * tE, self.N[-1] + dmd * tN, self.TVD[-1] + dmd * tTVD

        for i in range(len(self.mds) - 1):
            if self.mds[i] <= md <= self.mds[i+1]:
                seg_md0, seg_md1 = self.mds[i], self.mds[i+1]
                dmd_full = seg_md1 - seg_md0
                if dmd_full <= 0:
                    return self.E[i], self.N[i], self.TVD[i]
                frac = (md - seg_md0) / dmd_full
                incf = self.incs[i] + frac * (self.incs[i+1] - self.incs[i])
                azf = wrap_angle_rad(self.azs[i] + frac * (self.azs[i+1] - self.azs[i]))
                dE, dN, dTVD = _segment_displacement(self.method, md - seg_md0, self.incs[i], self.azs[i], incf, azf)
                return self.E[i] + dE, self.N[i] + dN, self.TVD[i] + dTVD

        return self.E[-1], self.N[-1], self.TVD[-1]

# ---------------------------------
# Tooltip
# ---------------------------------
class ToolTip:
    def __init__(self, widget, text: str, delay_ms: int = 350):
        self.widget = widget
        self.text = text
        self.delay = delay_ms
        self.tip = None
        self._id = None
        widget.bind("<Enter>", self._enter)
        widget.bind("<Leave>", self._leave)
        widget.bind("<ButtonPress>", self._leave)

    def _enter(self, _):
        self._id = self.widget.after(self.delay, self._show)

    def _leave(self, _):
        if self._id:
            try: self.widget.after_cancel(self._id)
            except Exception: pass
            self._id = None
        if self.tip:
            try: self.tip.destroy()
            except Exception: pass
            self.tip = None

    def _show(self):
        if self.tip or not self.text:
            return
        x = self.widget.winfo_rootx() + 10
        y = self.widget.winfo_rooty() + self.widget.winfo_height() + 2
        self.tip = tk.Toplevel(self.widget)
        self.tip.wm_overrideredirect(True)
        self.tip.wm_geometry(f"+{x}+{y}")
        label = tk.Label(
            self.tip, text=self.text, justify=tk.LEFT,
            relief=tk.SOLID, borderwidth=1,
            background="#ffffe0", foreground="#000",
            padx=6, pady=4, wraplength=360
        )
        label.pack(ipadx=1)

# ---------------------------------
# Preview & Mapping mixin (used by Modes 1 & 2)
# ---------------------------------
class PreviewAndMappingMixin:
    def _init_preview_state(self):
        self.show_all_cols = tk.BooleanVar(value=False)
        self.persist_depth_idx = int(CONFIG.get("last_depth_col_idx", -1))
        self.persist_az_idx    = int(CONFIG.get("last_az_col_idx", -1))
        self.persist_dip_idx   = int(CONFIG.get("last_dip_col_idx", -1))

    def _apply_persisted_mapping(self, combobox: ttk.Combobox, idx: int, options_len: int):
        if 0 <= idx < options_len:
            combobox.current(idx)
        elif options_len:
            combobox.current(0)

    def _remember_mapping_indices(self, depth_idx: int, az_idx: int, dip_idx: int):
        CONFIG["last_depth_col_idx"] = depth_idx
        CONFIG["last_az_col_idx"]    = az_idx
        CONFIG["last_dip_col_idx"]   = dip_idx
        save_config()
        log(f"Saved mapping indices D/Az/Dip = {depth_idx}/{az_idx}/{dip_idx}")

    def _bind_mapping_change(self):
        def on_any_change(_evt=None):
            d = self._get_col_index_from_choice(self.selected_depth_col.get())
            a = self._get_col_index_from_choice(self.selected_azimuth_col.get())
            p = self._get_col_index_from_choice(self.selected_dip_col.get())
            if None not in (d, a, p):
                self._remember_mapping_indices(d, a, p)
                self._render_preview()
        for cb in (self.depth_combo, self.az_combo, self.dip_combo):
            cb.bind("<<ComboboxSelected>>", on_any_change)

    def _add_all_columns_toggle(self, parent_frame: ttk.Frame):
        chk = ttk.Checkbutton(parent_frame, text="All Columns", variable=self.show_all_cols, command=self._render_preview)
        chk.pack(anchor=tk.W, padx=10, pady=4)

    def _render_preview_table(self, container: ttk.Frame, headers: List[str], first_row: List[str]):
        letters = column_letters(len(headers))
        if not self.show_all_cols.get():
            d = self._get_col_index_from_choice(self.selected_depth_col.get())
            a = self._get_col_index_from_choice(self.selected_azimuth_col.get())
            p = self._get_col_index_from_choice(self.selected_dip_col.get())
            picked = [i for i in (d, a, p) if i is not None and 0 <= i < len(headers)]
        else:
            picked = list(range(len(headers)))
        grid = container
        ttk.Label(grid, text="").grid(row=0, column=0, sticky=tk.W, padx=6)
        for col, i in enumerate(picked, start=1):
            ttk.Label(grid, text=letters[i], font=("TkDefaultFont", 9, "bold")).grid(row=0, column=col, padx=6, pady=4)
        ttk.Label(grid, text="Header:").grid(row=1, column=0, sticky=tk.W, padx=6)
        for col, i in enumerate(picked, start=1):
            ttk.Label(grid, text=headers[i], wraplength=220).grid(row=1, column=col, padx=6, pady=4, sticky=tk.W)
        ttk.Label(grid, text="Row 1:").grid(row=2, column=0, sticky=tk.W, padx=6)
        for col, i in enumerate(picked, start=1):
            v = first_row[i] if i < len(first_row) else ""
            ttk.Label(grid, text=v, foreground="#333").grid(row=2, column=col, padx=6, pady=4, sticky=tk.W)
        for i in range(len(picked) + 1):
            grid.grid_columnconfigure(i, weight=1)

# ---------------------------------
# Mode 1 Window
# ---------------------------------
class Mode1Window(tk.Toplevel, PreviewAndMappingMixin):
    """Mode1: Using survey data, Fixed spacing, Specified number from collar XYZ.
       “Top Instrument” = closest to collar; “Toe Instrument” = deepest (largest MD).
    """
    def __init__(self, master, export_dir: str):
        super().__init__(master)
        self.title("Mode1 — Using survey data, Fixed spacing, Specified number from collar XYZ")
        self.geometry("1150x960")
        self.export_dir = export_dir
        self._mode_key = "mode1"
        try: self.transient(master)
        except Exception: pass

        self.survey: Optional[SurveyData] = None
        self.survey_path: Optional[str] = None
        self.summary_cache: str = ""

        self.selected_depth_col = tk.StringVar()
        self.selected_azimuth_col = tk.StringVar()
        self.selected_dip_col = tk.StringVar()

        self.dip_positive_var = tk.BooleanVar(value=True)
        self.method_var = tk.StringVar(value='min')

        self.hole_name_var = tk.StringVar(value="")
        self.fixed_spacing_var = tk.StringVar(value="2.000")
        self.num_instruments_var = tk.StringVar(value="10")
        self.seed_start_md_var = tk.StringVar(value=get_seed_start_md_default("mode1"))
        self.collar_x_var = tk.StringVar(value="0.000")
        self.collar_y_var = tk.StringVar(value="0.000")
        self.collar_z_var = tk.StringVar(value="0.000")

        self.top_x_var = tk.StringVar(value="0.000")
        self.top_y_var = tk.StringVar(value="0.000")
        self.top_z_var = tk.StringVar(value="0.000")

        self.computed_rows: List[Tuple[int, float, float, float, float]] = []
        self.summary_text = tk.StringVar(value="Export Summary:\n(Compute to populate)")
        self.preview_frame = None

        # NEW: export order (persisted)
        self.export_order_var = tk.StringVar(value=get_export_order_default())

        self._init_preview_state()
        self._build_ui()

    def _build_ui(self):
        file_row = ttk.Frame(self); file_row.pack(fill=tk.X, padx=12, pady=10)
        ttk.Button(file_row, text="Select Survey CSV", command=self.on_select_csv).pack(side=tk.LEFT)
        self.file_label = ttk.Label(file_row, text="No file selected"); self.file_label.pack(side=tk.LEFT, padx=10)

        cfg = ttk.LabelFrame(self, text="Configure CSV format"); cfg.pack(fill=tk.X, padx=12, pady=10)
        dd = ttk.Frame(cfg); dd.pack(fill=tk.X, padx=10, pady=8)
        ttk.Label(dd, text="Depth column:").grid(row=0, column=0, sticky=tk.W, padx=4, pady=2)
        self.depth_combo = ttk.Combobox(dd, textvariable=self.selected_depth_col, width=22, state="readonly"); self.depth_combo.grid(row=0, column=1, sticky=tk.W, padx=6)
        ttk.Label(dd, text="Azimuth column:").grid(row=0, column=2, sticky=tk.W, padx=12)
        self.az_combo = ttk.Combobox(dd, textvariable=self.selected_azimuth_col, width=22, state="readonly"); self.az_combo.grid(row=0, column=3, sticky=tk.W, padx=6)
        ttk.Label(dd, text="Dip column:").grid(row=0, column=4, sticky=tk.W, padx=12)
        self.dip_combo = ttk.Combobox(dd, textvariable=self.selected_dip_col, width=22, state="readonly"); self.dip_combo.grid(row=0, column=5, sticky=tk.W, padx=6)

        self.dip_toggle_btn = ttk.Button(cfg, text="Dip Positive", command=self.toggle_dip_sign); self.dip_toggle_btn.pack(anchor=tk.W, padx=10, pady=6)
        method_frame = ttk.Frame(cfg); method_frame.pack(fill=tk.X, padx=10, pady=6)
        ttk.Label(method_frame, text="Computation method:").pack(side=tk.LEFT)
        self.rb_min = ttk.Radiobutton(method_frame, text="Minimum Curvature", variable=self.method_var, value='min')
        self.rb_avg = ttk.Radiobutton(method_frame, text="Average Angle", variable=self.method_var, value='avg')
        self.rb_tan = ttk.Radiobutton(method_frame, text="Tangential", variable=self.method_var, value='tan')
        self.rb_min.pack(side=tk.LEFT, padx=8)
        self.rb_avg.pack(side=tk.LEFT, padx=8)
        self.rb_tan.pack(side=tk.LEFT, padx=8)
        ToolTip(self.rb_min,
                "Minimum Curvature: Smoothly arcs between survey stations using a curvature factor.\n"
                "• Best overall accuracy for curved holes.\n"
                "• Produces smooth E/N/Z progression and realistic toe location.\n"
                "• Recommended default.")
        ToolTip(self.rb_avg,
                "Average Angle: Uses the average of the two station directions across each segment.\n"
                "• Simpler than Min Curv; small bias where curvature is high.\n"
                "• Positions are close to Min Curv for gentle trajectories.")
        ToolTip(self.rb_tan,
                "Tangential: Extends from each station using only the station's own angle.\n"
                "• Quick approximation; may over/under-shoot E/N/Z in curved segments.\n"
                "• Use when legacy systems expect tangential math.")

        self._add_all_columns_toggle(cfg)

        params = ttk.LabelFrame(self, text="Mode parameters"); params.pack(fill=tk.X, padx=12, pady=10)
        pgrid = ttk.Frame(params); pgrid.pack(fill=tk.X, padx=10, pady=8)
        ttk.Label(pgrid, text="Hole Name:").grid(row=0, column=0, sticky=tk.W, padx=4, pady=2); ttk.Entry(pgrid, textvariable=self.hole_name_var, width=32).grid(row=0, column=1, sticky=tk.W, padx=6)
        ttk.Label(pgrid, text="Instrument spacing (m):").grid(row=0, column=2, sticky=tk.W, padx=12); ttk.Entry(pgrid, textvariable=self.fixed_spacing_var, width=12).grid(row=0, column=3, sticky=tk.W, padx=6)
        ttk.Label(pgrid, text="Number of instruments:").grid(row=0, column=4, sticky=tk.W, padx=12); ttk.Entry(pgrid, textvariable=self.num_instruments_var, width=12).grid(row=0, column=5, sticky=tk.W, padx=6)
        ttk.Label(pgrid, text="Seed start depth from collar (MD, m):").grid(row=0, column=6, sticky=tk.W, padx=12)
        self.seed_start_md_entry = ttk.Entry(pgrid, textvariable=self.seed_start_md_var, width=14)
        self.seed_start_md_entry.grid(row=0, column=7, sticky=tk.W, padx=6)
        self.seed_start_md_entry.bind("<FocusOut>", lambda _e: set_seed_start_md(self._mode_key, self.seed_start_md_var.get()))
        ttk.Label(pgrid, text="(leave blank to auto-calc)").grid(row=0, column=8, sticky=tk.W, padx=6)
        ttk.Label(pgrid, text="Collar X:").grid(row=1, column=0, sticky=tk.W, padx=4, pady=2); ttk.Entry(pgrid, textvariable=self.collar_x_var, width=12).grid(row=1, column=1, sticky=tk.W, padx=6)
        ttk.Label(pgrid, text="Collar Y:").grid(row=1, column=2, sticky=tk.W, padx=12); ttk.Entry(pgrid, textvariable=self.collar_y_var, width=12).grid(row=1, column=3, sticky=tk.W, padx=6)
        ttk.Label(pgrid, text="Collar Z:").grid(row=1, column=4, sticky=tk.W, padx=12); ttk.Entry(pgrid, textvariable=self.collar_z_var, width=12).grid(row=1, column=5, sticky=tk.W, padx=6)
        ttk.Label(pgrid, text="Top Instrument X:").grid(row=2, column=0, sticky=tk.W, padx=4, pady=2)
        self.top_x_entry = ttk.Entry(pgrid, textvariable=self.top_x_var, width=12); self.top_x_entry.grid(row=2, column=1, sticky=tk.W, padx=6)
        ttk.Label(pgrid, text="Top Instrument Y:").grid(row=2, column=2, sticky=tk.W, padx=12)
        self.top_y_entry = ttk.Entry(pgrid, textvariable=self.top_y_var, width=12); self.top_y_entry.grid(row=2, column=3, sticky=tk.W, padx=6)
        ttk.Label(pgrid, text="Top Instrument Z:").grid(row=2, column=4, sticky=tk.W, padx=12)
        self.top_z_entry = ttk.Entry(pgrid, textvariable=self.top_z_var, width=12); self.top_z_entry.grid(row=2, column=5, sticky=tk.W, padx=6)
        self.top_x_entry.bind("<KeyRelease>", self._on_top_xyz_edited)
        self.top_y_entry.bind("<KeyRelease>", self._on_top_xyz_edited)
        self.top_z_entry.bind("<KeyRelease>", self._on_top_xyz_edited)
        ttk.Button(pgrid, text="Same as Collar values", command=self.copy_top_from_collar).grid(row=2, column=6, sticky=tk.W, padx=6)

        self.preview_container = ttk.LabelFrame(self, text="Survey preview"); self.preview_container.pack(fill=tk.BOTH, expand=True, padx=12, pady=10)

        # NEW: Export order
        order_frame = ttk.LabelFrame(self, text="Export order")
        order_frame.pack(fill=tk.X, padx=12, pady=6)
        def _on_order_change():
            set_export_order(self.export_order_var.get())
        ttk.Radiobutton(order_frame, text="Top Down (N → 1)",
                        variable=self.export_order_var, value="top_down",
                        command=_on_order_change).pack(side=tk.LEFT, padx=10, pady=6)
        ttk.Radiobutton(order_frame, text="Bottom Up (1 → N)",
                        variable=self.export_order_var, value="bottom_up",
                        command=_on_order_change).pack(side=tk.LEFT, padx=10, pady=6)

        self.summary_box = ttk.LabelFrame(self, text="Summary"); self.summary_box.pack(fill=tk.X, padx=12, pady=8)
        self.summary_text_widget = tk.Text(self.summary_box, height=12, wrap='word'); self.summary_text_widget.pack(fill=tk.X, padx=10, pady=8)
        self._set_summary_text(self.summary_text.get())

        btns = ttk.Frame(self); btns.pack(fill=tk.X, padx=12, pady=10)
        ttk.Button(btns, text="Calculate", command=self.on_compute).pack(side=tk.LEFT)
        ttk.Button(btns, text="Export to CSV", command=self.on_export_csv).pack(side=tk.LEFT, padx=8)
        ttk.Button(btns, text="Cancel", command=self.on_cancel).pack(side=tk.RIGHT)

        self._bind_mapping_change()

    def _set_summary_text(self, text: str):
        self.summary_text_widget.configure(state='normal'); self.summary_text_widget.delete('1.0', tk.END)
        self.summary_text_widget.insert(tk.END, text); self.summary_text_widget.configure(state='disabled')

    def _clear_seed_start_md(self):
        if hasattr(self, 'seed_start_md_var') and self.seed_start_md_var.get() != "":
            self.seed_start_md_var.set("")
            try:
                set_seed_start_md(self._mode_key, "")
            except Exception as e:
                log(f"Error saving seed start md: {e}")

    def _on_top_xyz_edited(self, event=None):
        self._clear_seed_start_md()

    def copy_top_from_collar(self):
        self.top_x_var.set(self.collar_x_var.get())
        self.top_y_var.set(self.collar_y_var.get())
        self.top_z_var.set(self.collar_z_var.get())
        self._clear_seed_start_md()
        log("Top Instrument XYZ copied from Collar XYZ")

    def toggle_dip_sign(self):
        self.dip_positive_var.set(not self.dip_positive_var.get())
        self.dip_toggle_btn.config(text="Dip Positive" if self.dip_positive_var.get() else "Dip Negative")

    def on_select_csv(self):
        initial = CONFIG.get("last_survey_dir") or CONFIG.get("last_export_dir") or os.getcwd()
        path = filedialog.askopenfilename(parent=self, title="Select survey CSV or TXT",
                                          filetypes=[("CSV or text", "*.csv *.txt"), ("All files", "*.*")],
                                          initialdir=initial)
        if not path: return
        try:
            survey = SurveyData.from_csv(path)
        except Exception as e:
            messagebox.showerror("File Error", f"Failed to load file.\n\n{e}", parent=self); return
        CONFIG["last_survey_dir"] = os.path.dirname(path); save_config()
        self.survey = survey; self.survey_path = path
        self.file_label.config(text=os.path.basename(path))

        letters = column_letters(len(survey.headers))
        options = [f"{letters[i]}: {survey.headers[i]}" for i in range(len(survey.headers))]
        for combo in (self.depth_combo, self.az_combo, self.dip_combo):
            combo['values'] = options
        self._apply_persisted_mapping(self.depth_combo, self.persist_depth_idx, len(options))
        self._apply_persisted_mapping(self.az_combo,    self.persist_az_idx,    len(options))
        self._apply_persisted_mapping(self.dip_combo,   self.persist_dip_idx,   len(options))
        if options and self.depth_combo.get() == "":
            self.depth_combo.current(0); self.az_combo.current(0); self.dip_combo.current(0)
        self._render_preview()

    def _render_preview(self):
        if hasattr(self, 'preview_frame') and self.preview_frame is not None:
            self.preview_frame.destroy()
        self.preview_frame = ttk.Frame(self.preview_container); self.preview_frame.pack(fill=tk.BOTH, expand=True)
        if not self.survey:
            ttk.Label(self.preview_frame, text="No data loaded.").pack(padx=8, pady=8); return
        headers = self.survey.headers; first_row = self.survey.rows[0] if self.survey.rows else []
        self._render_preview_table(self.preview_frame, headers, first_row)

    def _get_col_index_from_choice(self, choice: str) -> Optional[int]:
        if not choice: return None
        try:
            letter = choice.split(":", 1)[0].strip()
            idx = 0
            for ch in letter:
                idx = idx * 26 + (ord(ch.upper()) - ord('A') + 1)
            return idx - 1
        except Exception:
            return None

    def _read_numeric_columns(self, depth_idx: int, az_idx: int, dip_idx: int):
        mds, azs, dips = [], [], []
        for r in self.survey.rows:
            try:
                d = safe_float(r[depth_idx]); az = safe_float(r[az_idx]); dp = safe_float(r[dip_idx])
            except IndexError:
                d, az, dp = None, None, None
            if None not in (d, az, dp):
                mds.append(d); azs.append(az); dips.append(dp)
        return mds, azs, dips

    def _build_trajectory(self, mds: List[float], azs_deg: List[float], dips_deg: List[float]) -> 'Trajectory':
        eff_dip = [dp if self.dip_positive_var.get() else -dp for dp in dips_deg]
        incs = [rad(90.0 - max(-90.0, min(90.0, dp))) for dp in eff_dip]
        azs = [rad(a % 360.0) for a in azs_deg]

        # Patch 01 (Critical): Ensure a collar station exists at MD=0
        # Some surveys start at non-zero depth. Without an MD=0 station, MD=0 is not at the collar,
        # which can shift "instrument at collar" downhole and inflate installed depth.
        if mds and mds[0] > 0.0:
            mds = [0.0] + list(mds)
            incs = [incs[0]] + list(incs)
            azs  = [azs[0]]  + list(azs)

        return Trajectory(mds, incs, azs, method=self.method_var.get())

    def on_compute(self):
        log("Calculate pressed (Mode1)")
        if not self.survey:
            messagebox.showwarning("No Data", "Please select a survey CSV first.", parent=self); return

        depth_idx = self._get_col_index_from_choice(self.selected_depth_col.get())
        az_idx = self._get_col_index_from_choice(self.selected_azimuth_col.get())
        dip_idx = self._get_col_index_from_choice(self.selected_dip_col.get())
        if None in (depth_idx, az_idx, dip_idx):
            messagebox.showwarning("Missing mapping", "Please choose Depth, Azimuth, and Dip columns.", parent=self); return
        self._remember_mapping_indices(depth_idx, az_idx, dip_idx)

        spacing = safe_float(self.fixed_spacing_var.get())
        try:
            n_inst = int(float(self.num_instruments_var.get()))
        except Exception:
            n_inst = None
        cx = safe_float(self.collar_x_var.get()); cy = safe_float(self.collar_y_var.get()); cz = safe_float(self.collar_z_var.get())
        tx = safe_float(self.top_x_var.get()); ty = safe_float(self.top_y_var.get()); tz = safe_float(self.top_z_var.get())
        if spacing is None or spacing <= 0:
            messagebox.showwarning("Invalid spacing", "Instrument spacing must be a positive number.", parent=self); return
        if n_inst is None or n_inst <= 0:
            messagebox.showwarning("Invalid count", "Number of instruments must be a positive integer.", parent=self); return
        if None in (cx, cy, cz, tx, ty, tz):
            messagebox.showwarning("Invalid XYZ", "Collar and Top Instrument XYZ must be numbers.", parent=self); return

        mds, azs, dips = self._read_numeric_columns(depth_idx, az_idx, dip_idx)
        if not mds:
            messagebox.showerror("Data error", "No valid numeric rows found for the selected columns.", parent=self); return

        traj = self._build_trajectory(mds, azs, dips)
        toe_md = max(traj.mds)  # hole depth (collar→toe)

        # Estimate start MD nearest to the Top instrument (closest to collar)
        def estimate_start_md():
            vE = tx - cx; vN = ty - cy; vTVD = (cz - tz)
            tE, tN, tTVD = traj.tangent_at_start()
            md_proj = vE * tE + vN * tN + vTVD * tTVD
            if md_proj < 0:
                return md_proj
            last_md = traj.mds[-1]
            step = max(0.1, spacing) / 5.0
            best_md, best_d2, s = 0.0, float('inf'), 0.0
            while s <= last_md:
                e, n, tvd = traj.pos_at_md_rel(s)
                x = cx + e; y = cy + n; z = cz - tvd
                d2 = (x - tx)**2 + (y - ty)**2 + (z - tz)**2
                if d2 < best_d2: best_d2, best_md = d2, s
                s += step
            return best_md

        seed_txt = (self.seed_start_md_var.get() or "").strip()
        if seed_txt:
            seed_md = safe_float(seed_txt)
            if seed_md is None:
                messagebox.showwarning("Invalid seed start depth", "Seed start depth from collar must be a number (or blank for auto-calc).", parent=self); return
            start_md = float(seed_md)
        else:
            start_md = estimate_start_md()
            # Persist the auto-calculated start MD for convenience
            self.seed_start_md_var.set(f"{start_md:.3f}")
        set_seed_start_md(self._mode_key, self.seed_start_md_var.get())

        rows = []
        for i in range(n_inst):
            md = start_md + i * spacing
            e, n, tvd = traj.pos_at_md_rel(md)
            x = cx + e; y = cy + n; z = cz - tvd
            rows.append((i + 1, md, x, y, z))
        self.computed_rows = rows

        # Closest vs deepest by MD
        top_closest   = min(rows, key=lambda r: r[1])
        toe_deepest   = max(rows, key=lambda r: r[1])
        hole = (self.hole_name_var.get() or "").strip()
        installed_offset = min(spacing, 2.0)
        installed_depth_md = toe_deepest[1] + installed_offset  # deepest MD + one spacing (capped at +2.0m)

        summary = (
            "Export Summary:\n"
            f"Hole Name: {hole if hole else '(unnamed)'}\n"
            f"Collar Coordinates: X:{cx:.3f}, Y:{cy:.3f}, Z:{cz:.3f}\n"
            f"Hole depth (MD collar→toe): {toe_md:.3f} m\n"
            f"Installed depth (MD collar→Toe Instrument): {installed_depth_md:.3f} m\n"
            f"Top Instrument (closest to collar): X:{top_closest[2]:.3f}, Y:{top_closest[3]:.3f}, Z:{top_closest[4]:.3f}\n"
            f"Toe Instrument (deepest): X:{toe_deepest[2]:.3f}, Y:{toe_deepest[3]:.3f}, Z:{toe_deepest[4]:.3f}\n"
            f"Number of instruments: {len(rows)}"
        )
        self.summary_cache = summary
        self._set_summary_text(summary)
        messagebox.showinfo("Calculated", "Instrument positions have been calculated. Review the summary, then Export or Calculate again.", parent=self)

    def on_export_csv(self):
        if not self.computed_rows:
            self.on_compute()
            if not self.computed_rows: return

        # NEW: build export rows using export order
        order = self.export_order_var.get()
        ordered = sorted(self.computed_rows, key=lambda r: r[1], reverse=(order == "top_down"))

        export_rows = []
        if order == "top_down":
            n = len(ordered)
            for i, (_old_idx, md, x, y, z) in enumerate(ordered, start=0):
                idx = n - i  # N..1
                export_rows.append((idx, md, x, y, z))
        else:
            for i, (_old_idx, md, x, y, z) in enumerate(ordered, start=1):
                idx = i      # 1..N
                export_rows.append((idx, md, x, y, z))

        hole = (self.hole_name_var.get() or "").strip()
        filename = (hole if hole else "Hole").strip()
        for ch in '\\/:*?"<>|': filename = filename.replace(ch, "_")
        out_path = os.path.join(self.export_dir, f"{filename} Instrument Positions.csv")

        method_map = {"min": "Minimum Curvature", "avg": "Average Angle", "tan": "Tangential"}
        method_name = method_map.get(self.method_var.get(), self.method_var.get())
        summary_text = self.summary_cache or self.summary_text_widget.get("1.0", "end").strip()
        export_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

        try:
            with open(out_path, 'w', encoding='utf-8-sig', newline='') as f:
                w = csv.writer(f)
                w.writerow(["# Exported", export_time])
                w.writerow(["# Survey CSV", os.path.basename(self.survey_path) if self.survey_path else ""])
                w.writerow(["# Method", method_name])
                for line in summary_text.splitlines():
                    w.writerow(["#", line])
                w.writerow([])
                w.writerow(["Instrument #", "MD (m)", "X", "Y", "Z", "DistToPrev_m"])
                for idx, md, x, y, z in export_rows:
                    w.writerow([idx, f"{md:.3f}", f"{x:.3f}", f"{y:.3f}", f"{z:.3f}"])
            messagebox.showinfo("Exported", f"Wrote: {out_path}", parent=self)
        except Exception as e:
            messagebox.showerror("Write error", f"Could not write output.\n\n{e}", parent=self)

    def on_cancel(self):
        self.destroy()



# ---------------------------------
# Mode 2 Window
# ---------------------------------
class Mode2Window(tk.Toplevel, PreviewAndMappingMixin):
    """Mode2: Using survey data with instrument depth CSV."""
    def __init__(self, master, export_dir: str):
        super().__init__(master)
        self.title("Mode2 — Using survey data with instrument depth CSV")
        self.geometry("1180x920")
        self.export_dir = export_dir
        try: self.transient(master)
        except Exception: pass

        self.survey: Optional[SurveyData] = None
        self.survey_path: Optional[str] = None
        self.summary_cache: str = ""

        self.selected_depth_col = tk.StringVar()
        self.selected_azimuth_col = tk.StringVar()
        self.selected_dip_col = tk.StringVar()

        self.dip_positive_var = tk.BooleanVar(value=True)
        self.method_var = tk.StringVar(value='min')

        self.hole_name_var = tk.StringVar(value="")
        self.fixed_spacing_var = tk.StringVar(value="")   # disabled
        self.num_instruments_var = tk.StringVar(value="0")
        self.collar_x_var = tk.StringVar(value="0.000")
        self.collar_y_var = tk.StringVar(value="0.000")
        self.collar_z_var = tk.StringVar(value="0.000")

        self.toe_x_var = tk.StringVar(value="")
        self.toe_y_var = tk.StringVar(value="")
        self.toe_z_var = tk.StringVar(value="")

        self.first_x_var = tk.StringVar(value="0.000")  # UI label will say Toe Instrument
        self.first_y_var = tk.StringVar(value="0.000")
        self.first_z_var = tk.StringVar(value="0.000")

        self.dist_from_toe_var = tk.StringVar(value="")
        self.dist_from_toe_manual = False
        self.tie_first_to_toe = False
        self.offset_source_var = tk.StringVar(value="")  # Calculated from Depth File / Manual Override

        self.measured_depths: List[Tuple[int, float]] = []
        self.computed_rows: List[Tuple[int, float, float, float, float]] = []
        self.summary_text = tk.StringVar(value="Export Summary:\n(Import MD CSV and Calculate)")
        self.preview_frame = None

        # NEW: export order (persisted)
        self.export_order_var = tk.StringVar(value=get_export_order_default())

        self._init_preview_state()
        self._build_ui()

    def _build_ui(self):
        file_row = ttk.Frame(self); file_row.pack(fill=tk.X, padx=12, pady=10)
        ttk.Button(file_row, text="Select Survey CSV", command=self.on_select_csv).pack(side=tk.LEFT)
        self.file_label = ttk.Label(file_row, text="No file selected"); self.file_label.pack(side=tk.LEFT, padx=10)

        cfg = ttk.LabelFrame(self, text="Configure CSV format"); cfg.pack(fill=tk.X, padx=12, pady=10)
        dd = ttk.Frame(cfg); dd.pack(fill=tk.X, padx=10, pady=8)
        ttk.Label(dd, text="Depth column:").grid(row=0, column=0, sticky=tk.W, padx=4, pady=2)
        self.depth_combo = ttk.Combobox(dd, textvariable=self.selected_depth_col, width=22, state="readonly"); self.depth_combo.grid(row=0, column=1, sticky=tk.W, padx=6)
        ttk.Label(dd, text="Azimuth column:").grid(row=0, column=2, sticky=tk.W, padx=12)
        self.az_combo = ttk.Combobox(dd, textvariable=self.selected_azimuth_col, width=22, state="readonly"); self.az_combo.grid(row=0, column=3, sticky=tk.W, padx=6)
        ttk.Label(dd, text="Dip column:").grid(row=0, column=4, sticky=tk.W, padx=12)
        self.dip_combo = ttk.Combobox(dd, textvariable=self.selected_dip_col, width=22, state="readonly"); self.dip_combo.grid(row=0, column=5, sticky=tk.W, padx=6)

        self.dip_toggle_btn = ttk.Button(cfg, text="Dip Positive", command=self.toggle_dip_sign); self.dip_toggle_btn.pack(anchor=tk.W, padx=10, pady=6)
        method_frame = ttk.Frame(cfg); method_frame.pack(fill=tk.X, padx=10, pady=6)
        ttk.Label(method_frame, text="Computation method:").pack(side=tk.LEFT)
        self.rb_min = ttk.Radiobutton(method_frame, text="Minimum Curvature", variable=self.method_var, value='min')
        self.rb_avg = ttk.Radiobutton(method_frame, text="Average Angle", variable=self.method_var, value='avg')
        self.rb_tan = ttk.Radiobutton(method_frame, text="Tangential", variable=self.method_var, value='tan')
        self.rb_min.pack(side=tk.LEFT, padx=8)
        self.rb_avg.pack(side=tk.LEFT, padx=8)
        self.rb_tan.pack(side=tk.LEFT, padx=8)
        ToolTip(self.rb_min,
                "Minimum Curvature: Smoothly arcs between survey stations using a curvature factor.\n"
                "• Best overall accuracy for curved holes.\n"
                "• Produces smooth E/N/Z progression and realistic toe location.\n"
                "• Recommended default.")
        ToolTip(self.rb_avg,
                "Average Angle: Uses the average of the two station directions across each segment.\n"
                "• Simpler than Min Curv; small bias where curvature is high.\n"
                "• Positions are close to Min Curv for gentle trajectories.")
        ToolTip(self.rb_tan,
                "Tangential: Extends from each station using only the station's own angle.\n"
                "• Quick approximation; may over/under-shoot E/N/Z in curved segments.\n"
                "• Use when legacy systems expect tangential math.")

        self._add_all_columns_toggle(cfg)

        params = ttk.LabelFrame(self, text="Mode parameters"); params.pack(fill=tk.X, padx=12, pady=10)
        pgrid = ttk.Frame(params); pgrid.pack(fill=tk.X, padx=10, pady=8)
        ttk.Label(pgrid, text="Hole Name:").grid(row=0, column=0, sticky=tk.W, padx=4, pady=2)
        ttk.Entry(pgrid, textvariable=self.hole_name_var, width=32).grid(row=0, column=1, sticky=tk.W, padx=6)
        ttk.Label(pgrid, text="Instrument spacing (m):").grid(row=0, column=2, sticky=tk.W, padx=12)
        ttk.Entry(pgrid, textvariable=self.fixed_spacing_var, width=12, state="disabled").grid(row=0, column=3, sticky=tk.W, padx=6)
        ttk.Label(pgrid, text="Number of instruments:").grid(row=0, column=4, sticky=tk.W, padx=12)
        ttk.Entry(pgrid, textvariable=self.num_instruments_var, width=12, state="disabled").grid(row=0, column=5, sticky=tk.W, padx=6)

        ttk.Label(pgrid, text="Collar X:").grid(row=1, column=0, sticky=tk.W, padx=4, pady=2)
        ttk.Entry(pgrid, textvariable=self.collar_x_var, width=12).grid(row=1, column=1, sticky=tk.W, padx=6)
        ttk.Label(pgrid, text="Collar Y:").grid(row=1, column=2, sticky=tk.W, padx=12)
        ttk.Entry(pgrid, textvariable=self.collar_y_var, width=12).grid(row=1, column=3, sticky=tk.W, padx=6)
        ttk.Label(pgrid, text="Collar Z:").grid(row=1, column=4, sticky=tk.W, padx=12)
        ttk.Entry(pgrid, textvariable=self.collar_z_var, width=12).grid(row=1, column=5, sticky=tk.W, padx=6)

        ttk.Label(pgrid, text="Hole Toe X:").grid(row=2, column=0, sticky=tk.W, padx=4, pady=2)
        ttk.Entry(pgrid, textvariable=self.toe_x_var, width=12, state="readonly").grid(row=2, column=1, sticky=tk.W, padx=6)
        ttk.Label(pgrid, text="Hole Toe Y:").grid(row=2, column=2, sticky=tk.W, padx=12)
        ttk.Entry(pgrid, textvariable=self.toe_y_var, width=12, state="readonly").grid(row=2, column=3, sticky=tk.W, padx=6)
        ttk.Label(pgrid, text="Hole Toe Z:").grid(row=2, column=4, sticky=tk.W, padx=12)
        ttk.Entry(pgrid, textvariable=self.toe_z_var, width=12, state="readonly").grid(row=2, column=5, sticky=tk.W, padx=6)

        ttk.Label(pgrid, text="Toe Instrument X:").grid(row=3, column=0, sticky=tk.W, padx=4, pady=2)
        self.first_x_entry = ttk.Entry(pgrid, textvariable=self.first_x_var, width=12); self.first_x_entry.grid(row=3, column=1, sticky=tk.W, padx=6)
        ttk.Label(pgrid, text="Toe Instrument Y:").grid(row=3, column=2, sticky=tk.W, padx=12)
        self.first_y_entry = ttk.Entry(pgrid, textvariable=self.first_y_var, width=12); self.first_y_entry.grid(row=3, column=3, sticky=tk.W, padx=6)
        ttk.Label(pgrid, text="Toe Instrument Z:").grid(row=3, column=4, sticky=tk.W, padx=12)
        self.first_z_entry = ttk.Entry(pgrid, textvariable=self.first_z_var, width=12); self.first_z_entry.grid(row=3, column=5, sticky=tk.W, padx=6)
        ttk.Button(pgrid, text="Same as Toe coordinate", command=self.copy_first_from_toe).grid(row=3, column=6, sticky=tk.W, padx=6)

        ttk.Label(pgrid, text="Offset from toe (m):").grid(row=4, column=0, sticky=tk.W, padx=4, pady=2)
        self.dist_entry = ttk.Entry(pgrid, textvariable=self.dist_from_toe_var, width=12); self.dist_entry.grid(row=4, column=1, sticky=tk.W, padx=6)
        self.dist_entry.bind("<KeyRelease>", self._on_dist_manual_edit)
        self.offset_source_label = ttk.Label(pgrid, textvariable=self.offset_source_var, foreground="#555")
        self.offset_source_label.grid(row=4, column=2, columnspan=4, sticky=tk.W, padx=12)

        self.preview_container = ttk.LabelFrame(self, text="Survey preview")
        self.preview_container.pack(fill=tk.BOTH, expand=True, padx=12, pady=10)

        # NEW: Export order
        order_frame = ttk.LabelFrame(self, text="Export order")
        order_frame.pack(fill=tk.X, padx=12, pady=6)
        def _on_order_change():
            set_export_order(self.export_order_var.get())
        ttk.Radiobutton(order_frame, text="Top Down (N → 1)",
                        variable=self.export_order_var, value="top_down",
                        command=_on_order_change).pack(side=tk.LEFT, padx=10, pady=6)
        ttk.Radiobutton(order_frame, text="Bottom Up (1 → N)",
                        variable=self.export_order_var, value="bottom_up",
                        command=_on_order_change).pack(side=tk.LEFT, padx=10, pady=6)

        self.summary_box = ttk.LabelFrame(self, text="Summary"); self.summary_box.pack(fill=tk.X, padx=12, pady=8)
        self.summary_text_widget = tk.Text(self.summary_box, height=12, wrap='word'); self.summary_text_widget.pack(fill=tk.X, padx=10, pady=8)

        btns = ttk.Frame(self); btns.pack(fill=tk.X, padx=12, pady=10)
        ttk.Button(btns, text="Calculate", command=self.on_compute).pack(side=tk.LEFT)
        ttk.Button(btns, text="Export to CSV", command=self.on_export_csv).pack(side=tk.LEFT, padx=8)
        ttk.Button(btns, text="Import Measured Depth CSV", command=self.on_import_md_csv).pack(side=tk.LEFT, padx=8)
        ttk.Button(btns, text="Cancel", command=self.on_cancel).pack(side=tk.RIGHT)

        self._bind_mapping_change()

    def _set_summary_text(self, text: str):
        self.summary_text_widget.configure(state='normal'); self.summary_text_widget.delete('1.0', tk.END)
        self.summary_text_widget.insert(tk.END, text); self.summary_text_widget.configure(state='disabled')

    def _on_dist_manual_edit(self, *_):
        txt = (self.dist_from_toe_var.get() or "").strip()
        self.dist_from_toe_manual = (txt != "")
        self.offset_source_var.set("Manual Override" if self.dist_from_toe_manual else "Calculated from Depth File")
        log(f"Distance-from-toe manual override: {self.dist_from_toe_manual} (value='{txt}')")

    def copy_first_from_toe(self):
        if not all([self.toe_x_var.get(), self.toe_y_var.get(), self.toe_z_var.get()]):
            messagebox.showwarning("Toe not available", "Calculate first to determine the toe coordinates.", parent=self); return
        self.first_x_var.set(self.toe_x_var.get()); self.first_y_var.set(self.toe_y_var.get()); self.first_z_var.set(self.toe_z_var.get())
        self.tie_first_to_toe = True
        self.dist_from_toe_manual = False
        self.dist_from_toe_var.set("0.000")
        self.offset_source_var.set("Calculated from Depth File")
        log("Tie to toe enabled: Toe Instrument XYZ set to Toe XYZ; distance=0")

    def toggle_dip_sign(self):
        self.dip_positive_var.set(not self.dip_positive_var.get())
        self.dip_toggle_btn.config(text="Dip Positive" if self.dip_positive_var.get() else "Dip Negative")

    def on_select_csv(self):
        initial = CONFIG.get("last_survey_dir") or CONFIG.get("last_export_dir") or os.getcwd()
        path = filedialog.askopenfilename(parent=self, title="Select survey CSV or TXT",
                                          filetypes=[("CSV or text", "*.csv *.txt"), ("All files", "*.*")],
                                          initialdir=initial)
        if not path: return
        try:
            survey = SurveyData.from_csv(path)
        except Exception as e:
            messagebox.showerror("File Error", f"Failed to load file.\n\n{e}", parent=self); return
        CONFIG["last_survey_dir"] = os.path.dirname(path); save_config()
        self.survey = survey; self.survey_path = path
        self.file_label.config(text=os.path.basename(path))

        letters = column_letters(len(survey.headers))
        options = [f"{letters[i]}: {survey.headers[i]}" for i in range(len(survey.headers))]
        for combo in (self.depth_combo, self.az_combo, self.dip_combo):
            combo['values'] = options
        self._apply_persisted_mapping(self.depth_combo, self.persist_depth_idx, len(options))
        self._apply_persisted_mapping(self.az_combo,    self.persist_az_idx,    len(options))
        self._apply_persisted_mapping(self.dip_combo,   self.persist_dip_idx,   len(options))
        if options and self.depth_combo.get() == "":
            self.depth_combo.current(0); self.az_combo.current(0); self.dip_combo.current(0)

        # Reset state for new survey
        self.toe_x_var.set(""); self.toe_y_var.set(""); self.toe_z_var.set("")
        self.dist_from_toe_var.set(""); self.dist_from_toe_manual = False; self.tie_first_to_toe = False
        self.offset_source_var.set("")

        self._render_preview()

    def on_import_md_csv(self):
        initial = CONFIG.get("last_md_dir") or CONFIG.get("last_survey_dir") or CONFIG.get("last_export_dir") or os.getcwd()
        path = filedialog.askopenfilename(parent=self, title="Select measured-depth CSV",
                                          filetypes=[("CSV or text", "*.csv *.txt"), ("All files", "*.*")],
                                          initialdir=initial)
        if not path: return
        CONFIG["last_md_dir"] = os.path.dirname(path); save_config()
        try:
            with open(path, "r", encoding="utf-8", errors="ignore", newline="") as f:
                reader = csv.reader(f); rows = list(reader)
            if not rows: raise ValueError("Empty file.")
            header = [h.strip().lower() for h in rows[0]]

            def col_idx(name_options):
                for i, h in enumerate(header):
                    for opt in name_options:
                        if opt in h:
                            return i
                return None

            idx_nr = col_idx(["instrument nr", "instrument#", "instrument", "inst", "nr", "no"])
            idx_md = col_idx(["measured depth (m)", "measured depth", "measured_depth", "md"])
            if idx_nr is None or idx_md is None:
                raise ValueError("Columns not found. Expect 'Instrument Nr' and 'Measured Depth (m)'.")

            md_list = []
            for r in rows[1:]:
                try:
                    nr = int(float((r[idx_nr] or "").strip()))
                    md = float((r[idx_md] or "").strip())
                except Exception:
                    continue
                md_list.append((nr, md))
            if not md_list: raise ValueError("No valid rows found.")

            md_list.sort(key=lambda t: t[0])  # 1 (deepest) upwards
            self.measured_depths = md_list
            self.num_instruments_var.set(str(len(md_list)))

            # Reset auto/manual state
            self.tie_first_to_toe = False
            self.dist_from_toe_manual = False
            self.dist_from_toe_var.set("")
            self.offset_source_var.set("Calculated from Depth File")

            messagebox.showinfo("Imported", f"Loaded {len(md_list)} instruments from MD CSV.", parent=self)
        except Exception as e:
            messagebox.showerror("Import error", f"Failed to import MD CSV.\n\n{e}", parent=self)

    def _render_preview(self):
        if self.preview_frame is not None:
            self.preview_frame.destroy()
        self.preview_frame = ttk.Frame(self.preview_container); self.preview_frame.pack(fill=tk.BOTH, expand=True)
        if not self.survey:
            ttk.Label(self.preview_frame, text="No data loaded.").pack(padx=8, pady=8); return
        headers = self.survey.headers
        first_row = self.survey.rows[0] if self.survey.rows else []
        self._render_preview_table(self.preview_frame, headers, first_row)

    def _get_col_index_from_choice(self, choice: str) -> Optional[int]:
        if not choice: return None
        try:
            letter = choice.split(":", 1)[0].strip()
            idx = 0
            for ch in letter:
                idx = idx * 26 + (ord(ch.upper()) - ord('A') + 1)
            return idx - 1
        except Exception:
            return None

    def _read_numeric_columns(self, depth_idx: int, az_idx: int, dip_idx: int):
        mds, azs, dips = [], [], []
        for r in self.survey.rows:
            try:
                d = safe_float(r[depth_idx]); az = safe_float(r[az_idx]); dp = safe_float(r[dip_idx])
            except IndexError:
                d, az, dp = None, None, None
            if None not in (d, az, dp):
                mds.append(d); azs.append(az); dips.append(dp)
        log(f"Parsed numeric rows: {len(mds)}")
        return mds, azs, dips

    def _build_trajectory(self, mds: List[float], azs_deg: List[float], dips_deg: List[float]) -> 'Trajectory':
        eff_dip = [dp if self.dip_positive_var.get() else -dp for dp in dips_deg]
        incs = [rad(90.0 - max(-90.0, min(90.0, dp))) for dp in eff_dip]
        azs = [rad(a % 360.0) for a in azs_deg]
        method = self.method_var.get()
        log(f"Building trajectory with method={method}")
        # Patch 03 (Critical): Ensure a collar station exists at MD=0 (Mode2 also uses collar-referenced MDs)
        # Some surveys start at non-zero depth (e.g. first station at 6.03m). Without an MD=0 station,
        # any instrument depths "from collar" are effectively offset by that first station depth.
        if mds and mds[0] > 0.0:
            mds = [0.0] + list(mds)
            incs = [incs[0]] + list(incs)
            azs  = [azs[0]]  + list(azs)        
        return Trajectory(mds, incs, azs, method=method)

    def _compute_toe(self, traj: 'Trajectory', cx: float, cy: float, cz: float) -> Tuple[float, float, float, float]:
        toe_md = traj.mds[-1]
        e_toe, n_toe, tvd_toe = traj.pos_at_md_rel(toe_md)
        return toe_md, cx + e_toe, cy + n_toe, cz - tvd_toe

    def _vertical_fallback(self, prev_pts: List[Tuple[int, float, float, float, float]]) -> Tuple[float, float, float]:
        if not prev_pts:
            return float('nan'), float('nan'), float('nan')
        last = prev_pts[-1]
        dz_step = (prev_pts[-1][4] - prev_pts[-2][4]) if len(prev_pts) >= 2 else 0.0
        return last[2], last[3], last[4] + dz_step

    def _maybe_autofill_distance(self, toe_md: float):
        if not self.measured_depths:
            return
        md_inst1 = sorted(self.measured_depths, key=lambda t: t[0])[0][1]
        auto_dist = max(0.0, toe_md - md_inst1)
        if not self.dist_from_toe_manual or (self.dist_from_toe_var.get() or "").strip() == "":
            self.dist_from_toe_var.set(f"{auto_dist:.3f}")
            self.dist_from_toe_manual = False
            self.offset_source_var.set("Calculated from Depth File")
            log(f"[AUTO] Distance-from-toe set to {auto_dist:.3f} m (toe_md={toe_md:.3f}, md1={md_inst1:.3f})")
        else:
            log(f"[AUTO-SKIP] User manual distance retained: {self.dist_from_toe_var.get()}")

    def on_compute(self):
        if not self.survey:
            messagebox.showwarning("No Data", "Please select a survey CSV first.", parent=self); return
        if not self.measured_depths:
            messagebox.showwarning("No MDs", "Import a Measured Depth CSV first.", parent=self); return

        depth_idx = self._get_col_index_from_choice(self.selected_depth_col.get())
        az_idx = self._get_col_index_from_choice(self.selected_azimuth_col.get())
        dip_idx = self._get_col_index_from_choice(self.selected_dip_col.get())
        if None in (depth_idx, az_idx, dip_idx):
            messagebox.showwarning("Missing mapping", "Please choose Depth, Azimuth, and Dip columns.", parent=self); return
        self._remember_mapping_indices(depth_idx, az_idx, dip_idx)

        cx = safe_float(self.collar_x_var.get()); cy = safe_float(self.collar_y_var.get()); cz = safe_float(self.collar_z_var.get())
        if None in (cx, cy, cz):
            messagebox.showwarning("Invalid XYZ", "Enter numeric Collar X/Y/Z.", parent=self); return

        mds, azs, dips = self._read_numeric_columns(depth_idx, az_idx, dip_idx)
        if not mds:
            messagebox.showerror("Data error", "No valid numeric rows found for the selected columns.", parent=self); return

        traj = self._build_trajectory(mds, azs, dips)
        min_md, max_md = min(traj.mds), max(traj.mds)

        # Toe
        toe_md, toe_x, toe_y, toe_z = self._compute_toe(traj, cx, cy, cz)
        self.toe_x_var.set(f"{toe_x:.3f}"); self.toe_y_var.set(f"{toe_y:.3f}"); self.toe_z_var.set(f"{toe_z:.3f}")

        # Distance auto-fill unless manual
        self._maybe_autofill_distance(toe_md)

        # Build MD list (Instrument # asc: 1 deepest)
        md_list = sorted(self.measured_depths, key=lambda t: t[0])
        md_inst1 = md_list[0][1]

        # Determine MD shift policy
        if self.tie_first_to_toe:
            delta = toe_md - md_inst1
            delta_mode = "tie_to_toe"
        else:
            if self.dist_from_toe_manual:
                manual = safe_float(self.dist_from_toe_var.get())
                if manual is None or manual < 0:
                    messagebox.showwarning("Invalid distance", "Distance from toe must be a non-negative number.", parent=self); return
                target_md1 = max(0.0, toe_md - manual)
                delta = target_md1 - md_inst1
                delta_mode = f"manual({manual:.3f} m)"
            else:
                delta = 0.0
                delta_mode = "auto_no_shift"

        log(f"[MODE2] min_md={min_md:.3f} max_md={max_md:.3f} toe_md={toe_md:.3f} md1={md_inst1:.3f} "
            f"delta={delta:.3f} mode={delta_mode} tie={self.tie_first_to_toe} manual={self.dist_from_toe_manual}")

        adj_mds = [(nr, md + delta) for (nr, md) in md_list]

        out_rows: List[Tuple[int, float, float, float, float]] = []
        for nr, md in adj_mds:
            if min_md <= md <= max_md:
                e, n, tvd = traj.pos_at_md_rel(md)
                x, y, z = cx + e, cy + n, cz - tvd
            else:
                x, y, z = self._vertical_fallback(out_rows)
                if any(map(math.isnan, (x, y, z))):
                    if md > max_md:
                        x, y, z = toe_x, toe_y, toe_z
                    else:
                        x, y, z = cx, cy, cz
            out_rows.append((nr, md, x, y, z))

        # Toe Instrument XYZ preview
        if self.tie_first_to_toe:
            self.first_x_var.set(f"{toe_x:.3f}"); self.first_y_var.set(f"{toe_y:.3f}"); self.first_z_var.set(f"{toe_z:.3f}")
        else:
            x1, y1, z1 = out_rows[0][2], out_rows[0][3], out_rows[0][4]
            self.first_x_var.set(f"{x1:.3f}"); self.first_y_var.set(f"{y1:.3f}"); self.first_z_var.set(f"{z1:.3f}")

        self.computed_rows = out_rows

        # Summary (by MD: collar -> toe)
        topmost = min(out_rows, key=lambda r: r[1])   # closest to collar (smallest MD)
        deepest = max(out_rows, key=lambda r: r[1])   # closest to toe (largest MD)
        hole = (self.hole_name_var.get() or "").strip()
        # Patch 03 (Mode2): installed depth = deepest MD + spacing between deepest and 2nd deepest (capped at +2.0m)
        # spacing_last = deepest_md - second_deepest_md (when at least 2 instruments); otherwise 0.0
        mds_sorted = sorted(r[1] for r in out_rows)
        if len(mds_sorted) >= 2:
            spacing_last = mds_sorted[-1] - mds_sorted[-2]
        else:
            spacing_last = 0.0
        installed_offset = min(spacing_last, 2.0)
        installed_depth_md = deepest[1] + installed_offset
        summary = (
            "Export Summary:\n"
            f"Hole Name: {hole if hole else '(unnamed)'}\n"
            f"Collar Coordinates: X:{cx:.3f}, Y:{cy:.3f}, Z:{cz:.3f}\n"
            f"Hole depth (MD collar→toe): {toe_md:.3f} m\n"
            f"Installed depth (MD collar→Toe Instrument): {installed_depth_md:.3f} m\n"
            f"Toe Instrument (deepest): X:{deepest[2]:.3f}, Y:{deepest[3]:.3f}, Z:{deepest[4]:.3f}\n"
            f"Top Instrument (closest to collar): X:{topmost[2]:.3f}, Y:{topmost[3]:.3f}, Z:{topmost[4]:.3f}\n"
            f"Number of instruments: {len(out_rows)}"
        )
        self.summary_cache = summary
        self._set_summary_text(summary)

        outside_ct = sum(1 for _, md, *_ in out_rows if (md < min_md or md > max_md))
        if outside_ct >= max(1, len(out_rows)//2):
            messagebox.showwarning("Outside survey range",
                                   f"{outside_ct} of {len(out_rows)} instruments lie outside the survey MD range "
                                   f"({min_md:.3f}–{max_md:.3f}). Vertical fallback used for those points.",
                                   parent=self)

        messagebox.showinfo("Calculated", "Calculation complete. Review the summary, then Export or Calculate again.", parent=self)

    def on_export_csv(self):
        if not self.computed_rows:
            self.on_compute()
            if not self.computed_rows: return

        # NEW: order by instrument number (do NOT renumber)
        order = self.export_order_var.get()
        export_rows = sorted(self.computed_rows, key=lambda r: r[0], reverse=(order == "top_down"))

        hole = (self.hole_name_var.get() or "").strip()
        filename = (hole if hole else "Hole").strip()
        for ch in '\\/:*?"<>|': filename = filename.replace(ch, "_")
        out_path = os.path.join(self.export_dir, f"{filename} Instrument Positions.csv")

        method_map = {"min": "Minimum Curvature", "avg": "Average Angle", "tan": "Tangential"}
        method_name = method_map.get(self.method_var.get(), self.method_var.get())
        summary_text = self.summary_cache or self.summary_text_widget.get("1.0", "end").strip()
        export_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

        try:
            with open(out_path, 'w', encoding='utf-8-sig', newline='') as f:
                w = csv.writer(f)
                w.writerow(["# Exported", export_time])
                w.writerow(["# Survey CSV", os.path.basename(self.survey_path) if self.survey_path else ""])
                w.writerow(["# Method", method_name])
                for line in summary_text.splitlines():
                    w.writerow(["#", line])
                w.writerow([])
                w.writerow(["Instrument #", "Measured Depth Used (m)", "X", "Y", "Z"])
                prev = None
                for nr, md, x, y, z in export_rows:
                    if prev is None:
                        dist = ""
                    else:
                        dist = f"{math.dist(prev, (x, y, z)):.6f}"
                    w.writerow([nr, f"{md:.3f}", f"{x:.6f}", f"{y:.6f}", f"{z:.6f}", dist])
                    prev = (x, y, z)
            messagebox.showinfo("Exported", f"Wrote: {out_path}", parent=self)
        except Exception as e:
            messagebox.showerror("Write error", f"Could not write output.\n\n{e}", parent=self)

    def on_cancel(self):
        self.destroy()


# ---------------------------------
# Mode 3 Window (UPDATED: no survey button)
# ---------------------------------
class Mode3Window(tk.Toplevel):
    """Mode3: Using collar and toe coordinates, Fixed spacing, Specified number from collar XYZ.
       Straight-line trajectory between collar and toe. “Top Instrument” = closest to collar.
       “Toe Instrument” = deepest (closest to toe).
    """
    def __init__(self, master, export_dir: str):
        super().__init__(master)
        self.title("Mode3 — Using collar & toe coordinates, Fixed spacing, Specified number from collar XYZ")
        self.geometry("1150x700")
        self.export_dir = export_dir
        self._mode_key = "mode3"
        try: self.transient(master)
        except Exception: pass

        self.hole_name_var = tk.StringVar(value="")
        self.fixed_spacing_var = tk.StringVar(value="2.000")
        self.num_instruments_var = tk.StringVar(value="10")
        self.seed_start_md_var = tk.StringVar(value=get_seed_start_md_default("mode3"))

        self.collar_x_var = tk.StringVar(value="0.000")
        self.collar_y_var = tk.StringVar(value="0.000")
        self.collar_z_var = tk.StringVar(value="0.000")

        self.toe_x_var = tk.StringVar(value="0.000")
        self.toe_y_var = tk.StringVar(value="0.000")
        self.toe_z_var = tk.StringVar(value="0.000")

        self.top_x_var = tk.StringVar(value="0.000")
        self.top_y_var = tk.StringVar(value="0.000")
        self.top_z_var = tk.StringVar(value="0.000")

        self.computed_rows: List[Tuple[int, float, float, float, float]] = []
        self.summary_cache: str = ""

        # NEW: export order (persisted)
        self.export_order_var = tk.StringVar(value=get_export_order_default())

        self._build_ui()

    def _build_ui(self):
        # NOTE: No "Select Survey CSV" button in Mode3 by design (R4.5 change).

        params = ttk.LabelFrame(self, text="Mode parameters"); params.pack(fill=tk.X, padx=12, pady=10)
        pgrid = ttk.Frame(params); pgrid.pack(fill=tk.X, padx=10, pady=8)

        ttk.Label(pgrid, text="Hole Name:").grid(row=0, column=0, sticky=tk.W, padx=4, pady=2)
        ttk.Entry(pgrid, textvariable=self.hole_name_var, width=32).grid(row=0, column=1, sticky=tk.W, padx=6)
        ttk.Label(pgrid, text="Instrument spacing (m):").grid(row=0, column=2, sticky=tk.W, padx=12)
        ttk.Entry(pgrid, textvariable=self.fixed_spacing_var, width=12).grid(row=0, column=3, sticky=tk.W, padx=6)
        ttk.Label(pgrid, text="Number of instruments:").grid(row=0, column=4, sticky=tk.W, padx=12)
        ttk.Label(pgrid, text="Seed start depth from collar (MD, m):").grid(row=0, column=6, sticky=tk.W, padx=12)
        self.seed_start_md_entry = ttk.Entry(pgrid, textvariable=self.seed_start_md_var, width=14)
        self.seed_start_md_entry.grid(row=0, column=7, sticky=tk.W, padx=6)
        self.seed_start_md_entry.bind("<FocusOut>", lambda _e: set_seed_start_md(self._mode_key, self.seed_start_md_var.get()))
        ttk.Label(pgrid, text="(leave blank to auto-calc)").grid(row=0, column=8, sticky=tk.W, padx=6)
        ttk.Entry(pgrid, textvariable=self.num_instruments_var, width=12).grid(row=0, column=5, sticky=tk.W, padx=6)

        ttk.Label(pgrid, text="Collar X:").grid(row=1, column=0, sticky=tk.W, padx=4, pady=2)
        ttk.Entry(pgrid, textvariable=self.collar_x_var, width=12).grid(row=1, column=1, sticky=tk.W, padx=6)
        ttk.Label(pgrid, text="Collar Y:").grid(row=1, column=2, sticky=tk.W, padx=12)
        ttk.Entry(pgrid, textvariable=self.collar_y_var, width=12).grid(row=1, column=3, sticky=tk.W, padx=6)
        ttk.Label(pgrid, text="Collar Z:").grid(row=1, column=4, sticky=tk.W, padx=12)
        ttk.Entry(pgrid, textvariable=self.collar_z_var, width=12).grid(row=1, column=5, sticky=tk.W, padx=6)

        # Toe row
        ttk.Label(pgrid, text="Toe X:").grid(row=2, column=0, sticky=tk.W, padx=4, pady=2)
        ttk.Entry(pgrid, textvariable=self.toe_x_var, width=12).grid(row=2, column=1, sticky=tk.W, padx=6)
        ttk.Label(pgrid, text="Toe Y:").grid(row=2, column=2, sticky=tk.W, padx=12)
        ttk.Entry(pgrid, textvariable=self.toe_y_var, width=12).grid(row=2, column=3, sticky=tk.W, padx=6)
        ttk.Label(pgrid, text="Toe Z:").grid(row=2, column=4, sticky=tk.W, padx=12)
        ttk.Entry(pgrid, textvariable=self.toe_z_var, width=12).grid(row=2, column=5, sticky=tk.W, padx=6)

        # Top Instrument row (same behavior as Mode1)
        ttk.Label(pgrid, text="Top Instrument X:").grid(row=3, column=0, sticky=tk.W, padx=4, pady=2)
        self.top_x_entry = ttk.Entry(pgrid, textvariable=self.top_x_var, width=12); self.top_x_entry.grid(row=3, column=1, sticky=tk.W, padx=6)
        ttk.Label(pgrid, text="Top Instrument Y:").grid(row=3, column=2, sticky=tk.W, padx=12)
        self.top_y_entry = ttk.Entry(pgrid, textvariable=self.top_y_var, width=12); self.top_y_entry.grid(row=3, column=3, sticky=tk.W, padx=6)
        ttk.Label(pgrid, text="Top Instrument Z:").grid(row=3, column=4, sticky=tk.W, padx=12)
        self.top_z_entry = ttk.Entry(pgrid, textvariable=self.top_z_var, width=12); self.top_z_entry.grid(row=3, column=5, sticky=tk.W, padx=6)
        self.top_x_entry.bind("<KeyRelease>", self._on_top_xyz_edited)
        self.top_y_entry.bind("<KeyRelease>", self._on_top_xyz_edited)
        self.top_z_entry.bind("<KeyRelease>", self._on_top_xyz_edited)
        ttk.Button(pgrid, text="Same as Collar values", command=self.copy_top_from_collar).grid(row=3, column=6, sticky=tk.W, padx=6)

        # NEW: Export order
        order_frame = ttk.LabelFrame(self, text="Export order")
        order_frame.pack(fill=tk.X, padx=12, pady=6)
        def _on_order_change():
            set_export_order(self.export_order_var.get())
        ttk.Radiobutton(order_frame, text="Top Down (N → 1)",
                        variable=self.export_order_var, value="top_down",
                        command=_on_order_change).pack(side=tk.LEFT, padx=10, pady=6)
        ttk.Radiobutton(order_frame, text="Bottom Up (1 → N)",
                        variable=self.export_order_var, value="bottom_up",
                        command=_on_order_change).pack(side=tk.LEFT, padx=10, pady=6)

        self.summary_box = ttk.LabelFrame(self, text="Summary"); self.summary_box.pack(fill=tk.X, padx=12, pady=8)
        self.summary_text_widget = tk.Text(self.summary_box, height=12, wrap='word'); self.summary_text_widget.pack(fill=tk.X, padx=10, pady=8)
        self._set_summary_text("Export Summary:\n(Compute to populate)")

        btns = ttk.Frame(self); btns.pack(fill=tk.X, padx=12, pady=10)
        ttk.Button(btns, text="Calculate", command=self.on_compute).pack(side=tk.LEFT)
        ttk.Button(btns, text="Export to CSV", command=self.on_export_csv).pack(side=tk.LEFT, padx=8)
        ttk.Button(btns, text="Cancel", command=self.on_cancel).pack(side=tk.RIGHT)

    def _set_summary_text(self, text: str):
        self.summary_text_widget.configure(state='normal'); self.summary_text_widget.delete('1.0', tk.END)
        self.summary_text_widget.insert(tk.END, text); self.summary_text_widget.configure(state='disabled')

    def _clear_seed_start_md(self):
        if hasattr(self, 'seed_start_md_var') and self.seed_start_md_var.get() != "":
            self.seed_start_md_var.set("")
            try:
                set_seed_start_md(self._mode_key, "")
            except Exception as e:
                log(f"Error saving seed start md: {e}")

    def _on_top_xyz_edited(self, event=None):
        self._clear_seed_start_md()

    def copy_top_from_collar(self):
        self.top_x_var.set(self.collar_x_var.get())
        self.top_y_var.set(self.collar_y_var.get())
        self.top_z_var.set(self.collar_z_var.get())
        self._clear_seed_start_md()
        log("Mode3: Top Instrument XYZ copied from Collar XYZ")

    def _project_top_to_start_md(self, cx, cy, cz, tx, ty, tz, ux, uy, uz) -> float:
        vX, vY, vZ = (tx - cx), (ty - cy), (tz - cz)
        return vX * ux + vY * uy + vZ * uz  # MD along the line

    def on_compute(self):
        spacing = safe_float(self.fixed_spacing_var.get())
        try:
            n_inst = int(float(self.num_instruments_var.get()))
        except Exception:
            n_inst = None

        cx = safe_float(self.collar_x_var.get()); cy = safe_float(self.collar_y_var.get()); cz = safe_float(self.collar_z_var.get())
        tx = safe_float(self.top_x_var.get());    ty = safe_float(self.top_y_var.get());    tz = safe_float(self.top_z_var.get())
        tox = safe_float(self.toe_x_var.get());   toy = safe_float(self.toe_y_var.get());   toz = safe_float(self.toe_z_var.get())

        if spacing is None or spacing <= 0:
            messagebox.showwarning("Invalid spacing", "Instrument spacing must be a positive number.", parent=self); return
        if n_inst is None or n_inst <= 0:
            messagebox.showwarning("Invalid count", "Number of instruments must be a positive integer.", parent=self); return
        if None in (cx, cy, cz, tx, ty, tz, tox, toy, toz):
            messagebox.showwarning("Invalid XYZ", "All Collar, Toe and Top Instrument XYZ must be numbers.", parent=self); return

        dx, dy, dz = (tox - cx), (toy - cy), (toz - cz)
        L = math.sqrt(dx*dx + dy*dy + dz*dz)
        if L <= 1e-9:
            messagebox.showerror("Invalid geometry", "Collar and Toe coordinates are identical (zero length).", parent=self); return
        ux, uy, uz = dx / L, dy / L, dz / L  # unit vector from collar to toe

        # MD where the Top Instrument projects onto the line
        seed_txt = (self.seed_start_md_var.get() or "").strip()
        if seed_txt:
            seed_md = safe_float(seed_txt)
            if seed_md is None:
                messagebox.showwarning("Invalid seed start depth", "Seed start depth from collar must be a number (or blank for auto-calc).", parent=self); return
            start_md = float(seed_md)
        else:
            start_md = self._project_top_to_start_md(cx, cy, cz, tx, ty, tz, ux, uy, uz)
            self.seed_start_md_var.set(f"{start_md:.3f}")
        set_seed_start_md(self._mode_key, self.seed_start_md_var.get())

        # Build raw points as (md, x, y, z)
        # Build raw points with cumulative stepping to ensure exact spacing
        # Anchor at the projected Top Instrument point
        x0 = cx + (start_md / L) * dx
        y0 = cy + (start_md / L) * dy
        z0 = cz + (start_md / L) * dz
        px, py, pz = x0, y0, z0
        raw_pts = []
        for i in range(n_inst):
            md = start_md + i * spacing
            if i > 0:
                px += spacing * ux
                py += spacing * uy
                pz += spacing * uz
            raw_pts.append((md, px, py, pz))
        # Numbering rule: 1 = deepest (largest MD), highest # = closest to collar
        ordered_by_depth = sorted(raw_pts, key=lambda r: r[0], reverse=True)  # deepest → shallowest
        self.computed_rows = [(i + 1, md, x, y, z) for i, (md, x, y, z) in enumerate(ordered_by_depth)]

        # For summary (by MD, not by assigned number)
        top_by_md = min(raw_pts, key=lambda r: r[0])  # closest to collar
        toe_by_md = max(raw_pts, key=lambda r: r[0])  # closest to toe

        hole = (self.hole_name_var.get() or "").strip()
        installed_offset = min(spacing, 2.0)
        installed_depth_md = toe_by_md[0] + installed_offset  # deepest MD + one spacing (capped at +2.0m)

        summary = (
            "Export Summary:\n"
            f"Hole Name: {hole if hole else '(unnamed)'}\n"
            f"Collar Coordinates: X:{cx:.3f}, Y:{cy:.3f}, Z:{cz:.3f}\n"
            f"Toe Coordinates: X:{tox:.3f}, Y:{toy:.3f}, Z:{toz:.3f}\n"
            f"Hole depth (MD collar→toe): {L:.3f} m\n"
            f"Installed depth (MD collar→Toe Instrument): {installed_depth_md:.3f} m\n"
            f"Top Instrument (closest to collar): X:{top_by_md[1]:.3f}, Y:{top_by_md[2]:.3f}, Z:{top_by_md[3]:.3f}\n"
            f"Toe Instrument (deepest): X:{toe_by_md[1]:.3f}, Y:{toe_by_md[2]:.3f}, Z:{toe_by_md[3]:.3f}\n"
            f"Number of instruments: {len(self.computed_rows)}"
        )

        self.summary_cache = summary
        self._set_summary_text(summary)
        messagebox.showinfo("Calculated", "Instrument positions have been calculated. Review the summary, then Export or Calculate again.", parent=self)


    def on_export_csv(self):
        if not self.computed_rows:
            self.on_compute()
            if not self.computed_rows:
                return

        # Sort by instrument number only (do NOT renumber)
        order = self.export_order_var.get()
        export_rows = sorted(self.computed_rows, key=lambda r: r[0], reverse=(order == "top_down"))

        # Build output path (sanitize hole name)
        hole = (self.hole_name_var.get() or "").strip()
        filename = (hole if hole else "Hole").strip()
        for ch in '\\/:*?"<>|':
            filename = filename.replace(ch, "_")
        out_path = os.path.join(self.export_dir, f"{filename} Instrument Positions.csv")

        # Summary + timestamp
        summary_text = self.summary_cache or self.summary_text_widget.get("1.0", "end").strip()
        export_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

        try:
            # UTF-8 with BOM so Excel renders "→"
            with open(out_path, 'w', encoding='utf-8-sig', newline='') as f:
                w = csv.writer(f)
                w.writerow(["# Exported", export_time])
                w.writerow(["# Method", "Straight line (Collar→Toe)"])
                for line in summary_text.splitlines():
                    w.writerow(["#", line])
                w.writerow([])
                w.writerow(["Instrument #", "MD (m)", "X", "Y", "Z", "DistToPrev_m"])
                prev = None
                for nr, md, x, y, z in export_rows:
                    if prev is None:
                        dist = ""
                    else:
                        dist = f"{math.dist(prev, (x, y, z)):.6f}"
                    w.writerow([nr, f"{md:.3f}", f"{x:.6f}", f"{y:.6f}", f"{z:.6f}", dist])
                    prev = (x, y, z)

            try:
                log(f"Mode3 export OK: {out_path} ({len(export_rows)} rows, order={order})")
            except Exception:
                pass

            messagebox.showinfo("Exported", f"Wrote: {out_path}", parent=self)
        except Exception as e:
            try:
                log(f"Mode3 export error: {e}")
            except Exception:
                pass
            messagebox.showerror("Write error", f"Could not write output.\n\n{e}", parent=self)


    def on_cancel(self):
        self.destroy()


# ---------------------------------
# Mode 4 Window
# ---------------------------------
class Mode4Window(tk.Toplevel):
    """Mode4: Using collar coordinates, Azimuth, Dip, and Hole Length (straight line).
       - No Toe XYZ inputs. Toe is computed from az/dip/length and shown in Summary.
       - Instruments are placed along the straight line from Collar→Toe, starting from
         the Top Instrument XYZ projected onto that line, using fixed spacing and count.
    """
    def __init__(self, master, export_dir: str):
        super().__init__(master)
        self.title("Mode4 — Using collar XYZ, Azimuth, Dip and Hole Length")
        self.geometry("1150x740")
        self.export_dir = export_dir
        self._mode_key = "mode4"
        try: self.transient(master)
        except Exception: pass

        self.hole_name_var = tk.StringVar(value="")
        self.fixed_spacing_var = tk.StringVar(value="2.000")
        self.num_instruments_var = tk.StringVar(value="10")
        self.seed_start_md_var = tk.StringVar(value=get_seed_start_md_default("mode4"))

        self.collar_x_var = tk.StringVar(value="0.000")
        self.collar_y_var = tk.StringVar(value="0.000")
        self.collar_z_var = tk.StringVar(value="0.000")

        self.azimuth_deg_var = tk.StringVar(value="0.0")   # degrees, 0=N, 90=E (clockwise from North)
        self.dip_deg_var = tk.StringVar(value="0.0")       # degrees; positive = down-dip
        self.length_var = tk.StringVar(value="10.0")       # metres

        self.top_x_var = tk.StringVar(value="0.000")
        self.top_y_var = tk.StringVar(value="0.000")
        self.top_z_var = tk.StringVar(value="0.000")

        self.computed_rows: List[Tuple[int, float, float, float, float]] = []
        self.summary_cache: str = ""

        # NEW: export order (persisted)
        self.export_order_var = tk.StringVar(value=get_export_order_default())

        self._build_ui()

    def _build_ui(self):
        params = ttk.LabelFrame(self, text="Mode parameters"); params.pack(fill=tk.X, padx=12, pady=10)
        pgrid = ttk.Frame(params); pgrid.pack(fill=tk.X, padx=10, pady=8)

        ttk.Label(pgrid, text="Hole Name:").grid(row=0, column=0, sticky=tk.W, padx=4, pady=2)
        ttk.Entry(pgrid, textvariable=self.hole_name_var, width=32).grid(row=0, column=1, sticky=tk.W, padx=6)
        ttk.Label(pgrid, text="Instrument spacing (m):").grid(row=0, column=2, sticky=tk.W, padx=12)
        ttk.Entry(pgrid, textvariable=self.fixed_spacing_var, width=12).grid(row=0, column=3, sticky=tk.W, padx=6)
        ttk.Label(pgrid, text="Number of instruments:").grid(row=0, column=4, sticky=tk.W, padx=12)
        ttk.Label(pgrid, text="Seed start depth from collar (MD, m):").grid(row=0, column=6, sticky=tk.W, padx=12)
        self.seed_start_md_entry = ttk.Entry(pgrid, textvariable=self.seed_start_md_var, width=14)
        self.seed_start_md_entry.grid(row=0, column=7, sticky=tk.W, padx=6)
        self.seed_start_md_entry.bind("<FocusOut>", lambda _e: set_seed_start_md(self._mode_key, self.seed_start_md_var.get()))
        ttk.Label(pgrid, text="(leave blank to auto-calc)").grid(row=0, column=8, sticky=tk.W, padx=6)
        ttk.Entry(pgrid, textvariable=self.num_instruments_var, width=12).grid(row=0, column=5, sticky=tk.W, padx=6)

        ttk.Label(pgrid, text="Collar X:").grid(row=1, column=0, sticky=tk.W, padx=4, pady=2)
        ttk.Entry(pgrid, textvariable=self.collar_x_var, width=12).grid(row=1, column=1, sticky=tk.W, padx=6)
        ttk.Label(pgrid, text="Collar Y:").grid(row=1, column=2, sticky=tk.W, padx=12)
        ttk.Entry(pgrid, textvariable=self.collar_y_var, width=12).grid(row=1, column=3, sticky=tk.W, padx=6)
        ttk.Label(pgrid, text="Collar Z:").grid(row=1, column=4, sticky=tk.W, padx=12)
        ttk.Entry(pgrid, textvariable=self.collar_z_var, width=12).grid(row=1, column=5, sticky=tk.W, padx=6)

        ttk.Label(pgrid, text="Azimuth (° from North, CW):").grid(row=2, column=0, sticky=tk.W, padx=4, pady=2)
        ttk.Entry(pgrid, textvariable=self.azimuth_deg_var, width=12).grid(row=2, column=1, sticky=tk.W, padx=6)
        ttk.Label(pgrid, text="Dip (° down):").grid(row=2, column=2, sticky=tk.W, padx=12)
        ttk.Entry(pgrid, textvariable=self.dip_deg_var, width=12).grid(row=2, column=3, sticky=tk.W, padx=6)
        ttk.Label(pgrid, text="Hole Length (m):").grid(row=2, column=4, sticky=tk.W, padx=12)
        ttk.Entry(pgrid, textvariable=self.length_var, width=12).grid(row=2, column=5, sticky=tk.W, padx=6)

        # Top Instrument row
        ttk.Label(pgrid, text="Top Instrument X:").grid(row=3, column=0, sticky=tk.W, padx=4, pady=2)
        self.top_x_entry = ttk.Entry(pgrid, textvariable=self.top_x_var, width=12); self.top_x_entry.grid(row=3, column=1, sticky=tk.W, padx=6)
        ttk.Label(pgrid, text="Top Instrument Y:").grid(row=3, column=2, sticky=tk.W, padx=12)
        self.top_y_entry = ttk.Entry(pgrid, textvariable=self.top_y_var, width=12); self.top_y_entry.grid(row=3, column=3, sticky=tk.W, padx=6)
        ttk.Label(pgrid, text="Top Instrument Z:").grid(row=3, column=4, sticky=tk.W, padx=12)
        self.top_z_entry = ttk.Entry(pgrid, textvariable=self.top_z_var, width=12); self.top_z_entry.grid(row=3, column=5, sticky=tk.W, padx=6)
        self.top_x_entry.bind("<KeyRelease>", self._on_top_xyz_edited)
        self.top_y_entry.bind("<KeyRelease>", self._on_top_xyz_edited)
        self.top_z_entry.bind("<KeyRelease>", self._on_top_xyz_edited)
        ttk.Button(pgrid, text="Same as Collar values", command=self.copy_top_from_collar).grid(row=3, column=6, sticky=tk.W, padx=6)

        # NEW: Export order
        order_frame = ttk.LabelFrame(self, text="Export order")
        order_frame.pack(fill=tk.X, padx=12, pady=6)
        def _on_order_change():
            set_export_order(self.export_order_var.get())
        ttk.Radiobutton(order_frame, text="Top Down (N → 1)",
                        variable=self.export_order_var, value="top_down",
                        command=_on_order_change).pack(side=tk.LEFT, padx=10, pady=6)
        ttk.Radiobutton(order_frame, text="Bottom Up (1 → N)",
                        variable=self.export_order_var, value="bottom_up",
                        command=_on_order_change).pack(side=tk.LEFT, padx=10, pady=6)

        self.summary_box = ttk.LabelFrame(self, text="Summary"); self.summary_box.pack(fill=tk.X, padx=12, pady=8)
        self.summary_text_widget = tk.Text(self.summary_box, height=12, wrap='word'); self.summary_text_widget.pack(fill=tk.X, padx=10, pady=8)
        self._set_summary_text("Export Summary:\n(Compute to populate)")

        btns = ttk.Frame(self); btns.pack(fill=tk.X, padx=12, pady=10)
        ttk.Button(btns, text="Calculate", command=self.on_compute).pack(side=tk.LEFT)
        ttk.Button(btns, text="Export to CSV", command=self.on_export_csv).pack(side=tk.LEFT, padx=8)
        ttk.Button(btns, text="Cancel", command=self.on_cancel).pack(side=tk.RIGHT)

    def _set_summary_text(self, text: str):
        self.summary_text_widget.configure(state='normal'); self.summary_text_widget.delete('1.0', tk.END)
        self.summary_text_widget.insert(tk.END, text); self.summary_text_widget.configure(state='disabled')

    def _clear_seed_start_md(self):
        if hasattr(self, 'seed_start_md_var') and self.seed_start_md_var.get() != "":
            self.seed_start_md_var.set("")
            try:
                set_seed_start_md(self._mode_key, "")
            except Exception as e:
                log(f"Error saving seed start md: {e}")

    def _on_top_xyz_edited(self, event=None):
        self._clear_seed_start_md()

    def copy_top_from_collar(self):
        self.top_x_var.set(self.collar_x_var.get())
        self.top_y_var.set(self.collar_y_var.get())
        self.top_z_var.set(self.collar_z_var.get())
        self._clear_seed_start_md()
        log("Mode4: Top Instrument XYZ copied from Collar XYZ")

    def _dir_unit_from_az_dip(self, az_deg: float, dip_deg: float) -> Tuple[float, float, float]:
        """Return unit vector (ux, uy, uz) in XYZ from azimuth & dip."""
        az = rad((az_deg % 360.0 + 360.0) % 360.0)
        dip = max(-90.0, min(90.0, dip_deg))
        inc = rad(90.0 - dip)
        ux = math.sin(inc) * math.sin(az)   # East
        uy = math.sin(inc) * math.cos(az)   # North
        uz = -math.cos(inc)                 # Down increases TVD => Z decreases => negative uz
        L = math.sqrt(ux*ux + uy*uy + uz*uz)
        if L <= 1e-12: return 1.0, 0.0, 0.0
        return ux / L, uy / L, uz / L

    def on_compute(self):
        spacing = safe_float(self.fixed_spacing_var.get())
        try:
            n_inst = int(float(self.num_instruments_var.get()))
        except Exception:
            n_inst = None

        cx = safe_float(self.collar_x_var.get()); cy = safe_float(self.collar_y_var.get()); cz = safe_float(self.collar_z_var.get())
        az_deg = safe_float(self.azimuth_deg_var.get()); dip_deg = safe_float(self.dip_deg_var.get()); length = safe_float(self.length_var.get())
        tx = safe_float(self.top_x_var.get());    ty = safe_float(self.top_y_var.get());    tz = safe_float(self.top_z_var.get())

        if spacing is None or spacing <= 0:
            messagebox.showwarning("Invalid spacing", "Instrument spacing must be a positive number.", parent=self); return
        if n_inst is None or n_inst <= 0:
            messagebox.showwarning("Invalid count", "Number of instruments must be a positive integer.", parent=self); return
        if None in (cx, cy, cz, az_deg, dip_deg, length, tx, ty, tz):
            messagebox.showwarning("Invalid inputs", "All Collar XYZ, Azimuth, Dip, Length, and Top Instrument XYZ must be numbers.", parent=self); return
        if length <= 0:
            messagebox.showwarning("Invalid length", "Hole length must be a positive number.", parent=self); return

        ux, uy, uz = self._dir_unit_from_az_dip(az_deg, dip_deg)

        # Toe coordinate (for summary only)
        tox = cx + length * ux
        toy = cy + length * uy
        toz = cz + length * uz

        # Start MD via projection of top
        vX, vY, vZ = (tx - cx), (ty - cy), (tz - cz)
        seed_txt = (self.seed_start_md_var.get() or "").strip()
        if seed_txt:
            seed_md = safe_float(seed_txt)
            if seed_md is None:
                messagebox.showwarning("Invalid seed start depth", "Seed start depth from collar must be a number (or blank for auto-calc).", parent=self); return
            start_md = float(seed_md)
        else:
            start_md = vX * ux + vY * uy + vZ * uz
            self.seed_start_md_var.set(f"{start_md:.3f}")
        set_seed_start_md(self._mode_key, self.seed_start_md_var.get())

        # Build raw points with MD; we'll assign instrument numbers by depth later
        # Build raw points with cumulative stepping to ensure exact spacing
        # Anchor at the projected Top Instrument point
        x0 = cx + start_md * ux
        y0 = cy + start_md * uy
        z0 = cz + start_md * uz
        px, py, pz = x0, y0, z0
        raw_pts = []
        for i in range(n_inst):
            md = start_md + i * spacing
            if i > 0:
                px += spacing * ux
                py += spacing * uy
                pz += spacing * uz
            raw_pts.append((md, px, py, pz))
        # Numbering rule: Instrument 1 = deepest (largest MD), top = smallest MD
        ordered_by_depth = sorted(raw_pts, key=lambda r: r[0], reverse=True)  # deepest → shallowest
        self.computed_rows = [(i + 1, md, x, y, z) for i, (md, x, y, z) in enumerate(ordered_by_depth)]

        # For summary: identify top/deepest by MD (not by index)
        top_by_md = min(raw_pts, key=lambda r: r[0])
        toe_by_md = max(raw_pts, key=lambda r: r[0])
        hole = (self.hole_name_var.get() or "").strip()
        installed_offset = min(spacing, 2.0)
        installed_depth_md = toe_by_md[0] + installed_offset  # deepest MD + one spacing (capped at +2.0m)

        summary = (
            "Export Summary:\n"
            f"Hole Name: {hole if hole else '(unnamed)'}\n"
            f"Collar Coordinates: X:{cx:.3f}, Y:{cy:.3f}, Z:{cz:.3f}\n"
            f"Azimuth / Dip / Length: {az_deg:.3f}° / {dip_deg:.3f}° / {length:.3f} m\n"
            f"Toe Coordinates (computed): X:{tox:.3f}, Y:{toy:.3f}, Z:{toz:.3f}\n"
            f"Hole depth (MD collar→toe): {length:.3f} m\n"
            f"Installed depth (MD collar→Toe Instrument): {installed_depth_md:.3f} m\n"
            f"Top Instrument (closest to collar): X:{top_by_md[1]:.3f}, Y:{top_by_md[2]:.3f}, Z:{top_by_md[3]:.3f}\n"
            f"Toe Instrument (deepest): X:{toe_by_md[1]:.3f}, Y:{toe_by_md[2]:.3f}, Z:{toe_by_md[3]:.3f}\n"
            f"Number of instruments: {len(self.computed_rows)}"
        )
        self.summary_cache = summary
        self._set_summary_text(summary)
        messagebox.showinfo("Calculated", "Instrument positions have been calculated. Review the summary, then Export or Calculate again.", parent=self)


    def on_export_csv(self):
        if not self.computed_rows:
            self.on_compute()
            if not self.computed_rows:
                return

        # Sort by instrument number only (do NOT renumber)
        order = self.export_order_var.get()
        export_rows = sorted(self.computed_rows, key=lambda r: r[0], reverse=(order == "top_down"))

        # Build output path (sanitize hole name)
        hole = (self.hole_name_var.get() or "").strip()
        filename = (hole if hole else "Hole").strip()
        for ch in '\\/:*?"<>|':
            filename = filename.replace(ch, "_")
        out_path = os.path.join(self.export_dir, f"{filename} Instrument Positions.csv")

        # Summary + timestamp
        summary_text = self.summary_cache or self.summary_text_widget.get("1.0", "end").strip()
        export_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

        try:
            # UTF-8 with BOM so Excel renders "→"
            with open(out_path, 'w', encoding='utf-8-sig', newline='') as f:
                w = csv.writer(f)
                w.writerow(["# Exported", export_time])
                w.writerow(["# Method", "Straight line (Collar→Toe from Az/Dip/Length)"])
                for line in summary_text.splitlines():
                    w.writerow(["#", line])
                w.writerow([])
                w.writerow(["Instrument #", "MD (m)", "X", "Y", "Z", "DistToPrev_m"])
                prev = None
                for nr, md, x, y, z in export_rows:
                    if prev is None:
                        dist = ""
                    else:
                        dist = f"{math.dist(prev, (x, y, z)):.6f}"
                    w.writerow([nr, f"{md:.3f}", f"{x:.6f}", f"{y:.6f}", f"{z:.6f}", dist])
                    prev = (x, y, z)

            try:
                log(f"Mode4 export OK: {out_path} ({len(export_rows)} rows, order={order})")
            except Exception:
                pass

            messagebox.showinfo("Exported", f"Wrote: {out_path}", parent=self)
        except Exception as e:
            try:
                log(f"Mode4 export error: {e}")
            except Exception:
                pass
            messagebox.showerror("Write error", f"Could not write output.\n\n{e}", parent=self)


    def on_cancel(self):
        self.destroy()



# ============================================================================
# Main Application Window
# ============================================================================
class PositionCalculatorApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(APP_TITLE)
        self.geometry("880x600")
        self.export_dir: Optional[str] = None
        self._init_export_dir()
        self._build_main_ui()
        log("Main window created")

    def _init_export_dir(self):
        initial = CONFIG.get("last_export_dir")
        if not initial:
            initial = os.getcwd()
        self.export_dir = initial
        CONFIG["last_export_dir"] = self.export_dir
        save_config()
        log(f"Initial export folder set to: {self.export_dir}")

    def _build_main_ui(self):
        container = ttk.Frame(self); container.pack(fill=tk.BOTH, expand=True, padx=12, pady=12)

        exp = ttk.Frame(container); exp.pack(fill=tk.X)
        ttk.Label(exp, text="Export folder for generated files:").pack(side=tk.LEFT)
        self.export_label = ttk.Label(exp, text=self.export_dir, foreground="#444"); self.export_label.pack(side=tk.LEFT, padx=8)
        ttk.Button(exp, text="Output Folder", command=self.choose_folder).pack(side=tk.LEFT, padx=10)

        modes = ttk.LabelFrame(container, text="Mode Selection"); modes.pack(fill=tk.X, pady=16)
        self.mode_var = tk.StringVar(value="mode1")
        ttk.Radiobutton(modes, text="Mode1: Using survey data, Fixed spacing, Specified number from collar XYZ",
                        variable=self.mode_var, value="mode1").pack(anchor=tk.W, padx=8, pady=6)
        ttk.Radiobutton(modes, text="Mode2: Using survey data with instrument depth CSV",
                        variable=self.mode_var, value="mode2").pack(anchor=tk.W, padx=8, pady=6)
        ttk.Radiobutton(modes, text="Mode3: Using collar & toe coordinates, Fixed spacing, Specified number from collar XYZ",
                        variable=self.mode_var, value="mode3").pack(anchor=tk.W, padx=8, pady=6)
        ttk.Radiobutton(modes, text="Mode4: Using collar XYZ, Azimuth, Dip, and Hole Length",
                        variable=self.mode_var, value="mode4").pack(anchor=tk.W, padx=8, pady=6)

        ttk.Label(container, text="(Select the folder where CSVs will be exported to, then select the required Mode and click Open)").pack(anchor=tk.W, pady=4)

        btns = ttk.Frame(container); btns.pack(fill=tk.X, pady=10)
        ttk.Button(btns, text="Open Selected Mode", command=self.open_selected_mode).pack(side=tk.LEFT)
        ttk.Button(btns, text="Exit", command=self.destroy).pack(side=tk.RIGHT)

    def choose_folder(self):
        new = filedialog.askdirectory(title="Select export folder", initialdir=self.export_dir or os.getcwd())
        if new:
            self.export_dir = new
            CONFIG["last_export_dir"] = new
            save_config()
            self.export_label.config(text=new)

    def open_selected_mode(self):
        if not self.export_dir:
            messagebox.showwarning("No export folder", "Please select an export folder first."); return
        if self.mode_var.get() == "mode1":
            Mode1Window(self, self.export_dir)
        elif self.mode_var.get() == "mode2":
            Mode2Window(self, self.export_dir)
        elif self.mode_var.get() == "mode3":
            Mode3Window(self, self.export_dir)
        elif self.mode_var.get() == "mode4":
            Mode4Window(self, self.export_dir)
        else:
            messagebox.showinfo("Coming soon", "This mode is not implemented yet.")


# ============================================================================
# Entrypoint
# ============================================================================
if __name__ == "__main__":
    load_config()
    init_logger()
    log("Launching application")
    try:
        app = PositionCalculatorApp()
        app.mainloop()
        log("Application closed normally")
    except Exception as e:
        tb = traceback.format_exc()
        log(f"Fatal error: {e}\n{tb}")
        try:
            messagebox.showerror("Fatal Error", f"The application encountered an error and must close.\n\n{e}")
        except Exception:
            pass
