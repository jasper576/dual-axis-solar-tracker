"""
Solar Tracker Monitor  v3.0
============================
Leest de Arduino Solar Tracker via seriële poort en toont:
  - Realtime grafieken van spanning (V), stroom (mA) en vermogen (mW)
  - LDR sensorwaarden (Top / Right / Bottom / Left)
  - Servostatus en verbindingsstatus
  - Meting-modus: instelbare duur, voortgangsbalk, automatisch stoppen
  - Export: CSV-data + PNG-grafieken downloaden

Installeer vereisten:
  pip install pyserial matplotlib
"""

import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import threading
import time
import collections
import re
import os
import csv
import datetime

try:
    import serial
    import serial.tools.list_ports
    SERIAL_AVAILABLE = True
except ImportError:
    SERIAL_AVAILABLE = False

try:
    import matplotlib
    matplotlib.use("TkAgg")
    import matplotlib.pyplot as plt
    from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
    from matplotlib.figure import Figure
    import matplotlib.animation as animation
    MATPLOTLIB_AVAILABLE = True
except ImportError:
    MATPLOTLIB_AVAILABLE = False

# ── Kleuren ──────────────────────────────────────────────────────────────────
BG         = "#0d1117"
PANEL      = "#161b22"
BORDER     = "#30363d"
ACCENT     = "#f0b429"
ACCENT2    = "#58a6ff"
ACCENT3    = "#3fb950"
ACCENT4    = "#ff7b72"
ACCENT5    = "#c084fc"
TEXT       = "#e6edf3"
TEXT_MUTED = "#8b949e"
GRAPH_BG   = "#0d1117"

MAX_POINTS = 600  # 10 minuten bij 1/s

# ── Regex voor seriële data ───────────────────────────────────────────────────
PATTERN = re.compile(
    r"T:\s*(\d+).*?L:\s*(\d+).*?R:\s*(\d+).*?B:\s*(\d+)"
    r".*?U=([\d.]+)V.*?I=([\d.]+)mA.*?P=([\d.]+)mW"
    r".*?rot=(\d+).*?hoek=(\d+)",
    re.IGNORECASE
)


# ─────────────────────────────────────────────────────────────────────────────
class SolarTrackerApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Solar Tracker Monitor  v3.0")
        self.configure(bg=BG)
        self.minsize(1200, 780)
        self.geometry("1340x860")

        self._setup_style()
        self._build_ui()

        # ── Data buffers ──
        self.timestamps  = collections.deque(maxlen=MAX_POINTS)
        self.voltages    = collections.deque(maxlen=MAX_POINTS)
        self.currents    = collections.deque(maxlen=MAX_POINTS)
        self.powers      = collections.deque(maxlen=MAX_POINTS)
        self.energies    = collections.deque(maxlen=MAX_POINTS)
        self.ldr_top     = collections.deque(maxlen=MAX_POINTS)
        self.ldr_right   = collections.deque(maxlen=MAX_POINTS)
        self.ldr_bottom  = collections.deque(maxlen=MAX_POINTS)
        self.ldr_left    = collections.deque(maxlen=MAX_POINTS)
        self.rot_pos     = collections.deque(maxlen=MAX_POINTS)
        self.hoek_pos    = collections.deque(maxlen=MAX_POINTS)

        # Alle ruwe rijen voor CSV-export (ook buiten meting)
        self._all_rows: list[dict] = []

        # Meting-modus buffers (worden gevuld tijdens actieve meting)
        self._meting_rows: list[dict] = []
        self._meting_active   = False
        self._meting_start_ts = None
        self._meting_duur_s   = 60
        self._meting_timer_id = None

        # Energie
        self._total_energy_mWh = 0.0
        self._last_sample_time = None

        # Verbinding
        self.serial_port  = None
        self.reading      = False
        self._read_thread = None
        self.start_time   = None
        self._demo_running = False

        self._setup_graphs()
        self._refresh_ports()
        self.protocol("WM_DELETE_WINDOW", self._on_close)

    # ─────────────────────────────────────────────────────────────────────────
    # Stijl
    # ─────────────────────────────────────────────────────────────────────────
    def _setup_style(self):
        s = ttk.Style(self)
        s.theme_use("clam")
        s.configure("TFrame",      background=BG)
        s.configure("Panel.TFrame",background=PANEL, relief="flat")
        s.configure("TLabel",      background=BG,    foreground=TEXT,
                    font=("Courier New", 10))
        s.configure("Muted.TLabel",background=PANEL, foreground=TEXT_MUTED,
                    font=("Courier New", 9))
        s.configure("Header.TLabel",background=BG,   foreground=ACCENT,
                    font=("Courier New", 11, "bold"))
        s.configure("Value.TLabel", background=PANEL, foreground=TEXT,
                    font=("Courier New", 13, "bold"))
        for name, bg in [
            ("Accent.TButton",  ACCENT),
            ("Danger.TButton",  ACCENT4),
            ("Purple.TButton",  ACCENT5),
            ("Green.TButton",   ACCENT3),
            ("Blue.TButton",    ACCENT2),
        ]:
            s.configure(name, background=bg, foreground=BG,
                        font=("Courier New", 10, "bold"),
                        relief="flat", padding=6)
        s.configure("TCombobox", fieldbackground=PANEL, background=PANEL,
                    foreground=TEXT, selectbackground=BORDER)
        s.configure("Horizontal.TProgressbar",
                    troughcolor=BORDER, background=ACCENT3,
                    thickness=14)

    # ─────────────────────────────────────────────────────────────────────────
    # UI bouwen
    # ─────────────────────────────────────────────────────────────────────────
    def _build_ui(self):
        # Topbalk
        topbar = tk.Frame(self, bg=PANEL, height=52)
        topbar.pack(fill="x")
        topbar.pack_propagate(False)
        tk.Label(topbar, text="☀  SOLAR TRACKER MONITOR  v3.0",
                 bg=PANEL, fg=ACCENT,
                 font=("Courier New", 15, "bold")).pack(side="left", padx=18, pady=10)
        self._status_dot = tk.Label(topbar, text="●", bg=PANEL, fg=ACCENT4,
                                    font=("Courier New", 18))
        self._status_dot.pack(side="right", padx=6)
        self._status_lbl = tk.Label(topbar, text="NIET VERBONDEN",
                                    bg=PANEL, fg=TEXT_MUTED,
                                    font=("Courier New", 9, "bold"))
        self._status_lbl.pack(side="right", padx=2)

        # Verbindingsbalk
        conn = tk.Frame(self, bg=BG)
        conn.pack(fill="x", padx=12, pady=(8, 2))
        tk.Label(conn, text="Poort:", bg=BG, fg=TEXT_MUTED,
                 font=("Courier New", 10)).pack(side="left", padx=(0, 4))
        self._port_var = tk.StringVar()
        self._port_combo = ttk.Combobox(conn, textvariable=self._port_var,
                                        width=16, state="readonly")
        self._port_combo.pack(side="left", padx=(0, 6))
        tk.Label(conn, text="Baud:", bg=BG, fg=TEXT_MUTED,
                 font=("Courier New", 10)).pack(side="left", padx=(0, 4))
        self._baud_var = tk.StringVar(value="9600")
        ttk.Combobox(conn, textvariable=self._baud_var,
                     values=["9600", "115200"], width=8,
                     state="readonly").pack(side="left", padx=(0, 10))
        self._connect_btn = ttk.Button(conn, text="Verbinden",
                                       style="Accent.TButton",
                                       command=self._toggle_connect)
        self._connect_btn.pack(side="left", padx=(0, 6))
        ttk.Button(conn, text="↻ Vernieuw",
                   command=self._refresh_ports).pack(side="left", padx=(0, 10))
        self._demo_btn = ttk.Button(conn, text="▶ Demo",
                                    command=self._toggle_demo)
        self._demo_btn.pack(side="left", padx=(0, 10))

        # Download-knop in topbalk
        ttk.Button(conn, text="⬇ Exporteer CSV",
                   style="Blue.TButton",
                   command=self._export_csv).pack(side="right", padx=(6, 0))
        ttk.Button(conn, text="⬇ Exporteer PNG",
                   style="Purple.TButton",
                   command=self._export_png).pack(side="right", padx=(6, 0))

        # Hoofdlayout
        main = tk.Frame(self, bg=BG)
        main.pack(fill="both", expand=True, padx=12, pady=(4, 4))

        # Linkerkolom
        left = tk.Frame(main, bg=BG, width=240)
        left.pack(side="left", fill="y", padx=(0, 10))
        left.pack_propagate(False)
        self._build_sensor_cards(left)

        # Rechterkolom: meting-balk + grafieken
        right = tk.Frame(main, bg=BG)
        right.pack(side="left", fill="both", expand=True)
        self._build_meting_panel(right)
        self._graph_frame = right

    # ─────────────────────────────────────────────────────────────────────────
    # Meting-paneel
    # ─────────────────────────────────────────────────────────────────────────
    def _build_meting_panel(self, parent):
        panel = tk.Frame(parent, bg=PANEL, bd=0,
                         highlightthickness=1, highlightbackground=BORDER)
        panel.pack(fill="x", pady=(0, 8))

        # Rij 1: titel + duur-instelling + knoppen
        row1 = tk.Frame(panel, bg=PANEL)
        row1.pack(fill="x", padx=10, pady=(8, 4))

        tk.Label(row1, text="METING", bg=PANEL, fg=ACCENT,
                 font=("Courier New", 11, "bold")).pack(side="left", padx=(0, 16))

        tk.Label(row1, text="Duur:", bg=PANEL, fg=TEXT_MUTED,
                 font=("Courier New", 10)).pack(side="left", padx=(0, 4))

        self._duur_var = tk.StringVar(value="60")
        duur_entry = tk.Entry(row1, textvariable=self._duur_var,
                              width=6, bg=BORDER, fg=TEXT,
                              font=("Courier New", 11), insertbackground=TEXT,
                              relief="flat", bd=4)
        duur_entry.pack(side="left", padx=(0, 2))

        tk.Label(row1, text="s", bg=PANEL, fg=TEXT_MUTED,
                 font=("Courier New", 10)).pack(side="left", padx=(0, 16))

        # Snelkeuze-knoppen
        for label, sec in [("10s", 10), ("30s", 30), ("60s", 60),
                            ("5m", 300), ("10m", 600)]:
            def _set(s=sec):
                self._duur_var.set(str(s))
            tk.Button(row1, text=label, bg=BORDER, fg=TEXT_MUTED,
                      font=("Courier New", 9), relief="flat", padx=6,
                      activebackground=ACCENT, activeforeground=BG,
                      command=_set).pack(side="left", padx=2)

        self._meting_btn = ttk.Button(row1, text="▶ Start meting",
                                      style="Green.TButton",
                                      command=self._toggle_meting)
        self._meting_btn.pack(side="right", padx=(10, 0))

        # Rij 2: voortgangsbalk + timer + statistieken
        row2 = tk.Frame(panel, bg=PANEL)
        row2.pack(fill="x", padx=10, pady=(0, 6))

        self._meting_progress = ttk.Progressbar(
            row2, orient="horizontal", length=300, mode="determinate",
            style="Horizontal.TProgressbar")
        self._meting_progress.pack(side="left", padx=(0, 10))

        self._meting_timer_lbl = tk.Label(
            row2, text="--:--  /  --:--", bg=PANEL, fg=TEXT_MUTED,
            font=("Courier New", 10))
        self._meting_timer_lbl.pack(side="left", padx=(0, 20))

        # Live statistieken tijdens meting
        for attr, label, color in [
            ("_mst_gem", "Gem:", ACCENT),
            ("_mst_max", "Max:", ACCENT4),
            ("_mst_min", "Min:", ACCENT2),
            ("_mst_e",   "Energie:", ACCENT5),
        ]:
            tk.Label(row2, text=label, bg=PANEL, fg=TEXT_MUTED,
                     font=("Courier New", 9)).pack(side="left", padx=(0, 2))
            lbl = tk.Label(row2, text="---", bg=PANEL, fg=color,
                           font=("Courier New", 10, "bold"))
            lbl.pack(side="left", padx=(0, 12))
            setattr(self, attr, lbl)

    # ─────────────────────────────────────────────────────────────────────────
    # Sensor-kaarten (linkerkolom)
    # ─────────────────────────────────────────────────────────────────────────
    def _build_sensor_cards(self, parent):
        def section(title):
            tk.Label(parent, text=title, bg=BG, fg=ACCENT,
                     font=("Courier New", 10, "bold")).pack(anchor="w", pady=(8, 2))
            card = tk.Frame(parent, bg=PANEL, bd=0,
                            highlightthickness=1, highlightbackground=BORDER)
            card.pack(fill="x", pady=(0, 6))
            return card

        ina = section("INA219")
        self._v_lbl = self._val_row(ina, "Spanning",  "-- V",    ACCENT2)
        self._i_lbl = self._val_row(ina, "Stroom",    "-- mA",   ACCENT3)
        self._p_lbl = self._val_row(ina, "Vermogen",  "-- mW",   ACCENT)
        self._e_lbl = self._val_row(ina, "Energie",   "0.000 mWh", ACCENT5)
        self._ina_st = self._stat_row(ina, "INA219")

        ldr = section("LDR SENSOREN")
        self._ldr_bars = {}
        for name in ("Top", "Rechts", "Onder", "Links"):
            self._ldr_bars[name] = self._ldr_row(ldr, name)
        self._ldr_st = self._stat_row(ldr, "LDR x4")

        srv = section("SERVO")
        self._rot_lbl  = self._val_row(srv, "Rotatie", "--°", TEXT)
        self._hoek_lbl = self._val_row(srv, "Hoek",    "--°", TEXT)

        misc = section("SYSTEEM")
        self._serial_st  = self._stat_row(misc, "Serieel")
        self._samples_lbl = self._val_row(misc, "Samples", "0", TEXT_MUTED)

        ttk.Button(parent, text="Wis alle data",
                   command=self._clear_data).pack(fill="x", pady=(8, 0))

    def _val_row(self, parent, label, initial, color):
        row = tk.Frame(parent, bg=PANEL)
        row.pack(fill="x", padx=8, pady=3)
        tk.Label(row, text=label, bg=PANEL, fg=TEXT_MUTED,
                 font=("Courier New", 9), width=9, anchor="w").pack(side="left")
        lbl = tk.Label(row, text=initial, bg=PANEL, fg=color,
                       font=("Courier New", 12, "bold"), anchor="e")
        lbl.pack(side="right")
        return lbl

    def _stat_row(self, parent, label):
        row = tk.Frame(parent, bg=PANEL)
        row.pack(fill="x", padx=8, pady=2)
        tk.Label(row, text=label, bg=PANEL, fg=TEXT_MUTED,
                 font=("Courier New", 9), anchor="w").pack(side="left")
        lbl = tk.Label(row, text="● ONBEKEND", bg=PANEL, fg=TEXT_MUTED,
                       font=("Courier New", 9, "bold"), anchor="e")
        lbl.pack(side="right")
        return lbl

    def _ldr_row(self, parent, label):
        row = tk.Frame(parent, bg=PANEL)
        row.pack(fill="x", padx=8, pady=3)
        tk.Label(row, text=label, bg=PANEL, fg=TEXT_MUTED,
                 font=("Courier New", 9), width=6, anchor="w").pack(side="left")
        bg_f = tk.Frame(row, bg=BORDER, height=10, width=100)
        bg_f.pack(side="left", padx=4)
        bg_f.pack_propagate(False)
        fill = tk.Frame(bg_f, bg=ACCENT3, height=10)
        fill.place(x=0, y=0, relheight=1, width=0)
        val = tk.Label(row, text="---", bg=PANEL, fg=TEXT,
                       font=("Courier New", 9), width=4, anchor="e")
        val.pack(side="right")
        return {"bar": fill, "bg": bg_f, "lbl": val}

    # ─────────────────────────────────────────────────────────────────────────
    # Grafieken
    # ─────────────────────────────────────────────────────────────────────────
    def _setup_graphs(self):
        if not MATPLOTLIB_AVAILABLE:
            tk.Label(self._graph_frame,
                     text="matplotlib niet gevonden.\npip install matplotlib",
                     bg=BG, fg=ACCENT4,
                     font=("Courier New", 13)).pack(expand=True)
            return

        plt.rcParams.update({
            "axes.facecolor":   GRAPH_BG,
            "figure.facecolor": GRAPH_BG,
            "axes.edgecolor":   BORDER,
            "axes.labelcolor":  TEXT_MUTED,
            "xtick.color":      TEXT_MUTED,
            "ytick.color":      TEXT_MUTED,
            "grid.color":       BORDER,
            "grid.alpha":       0.4,
            "text.color":       TEXT,
            "lines.linewidth":  1.8,
        })

        self.fig = Figure(figsize=(9, 6.5), dpi=96, facecolor=GRAPH_BG)
        self.fig.subplots_adjust(hspace=0.55, top=0.95, bottom=0.07,
                                 left=0.09, right=0.97)

        self.ax_v = self.fig.add_subplot(4, 1, 1)
        self.ax_i = self.fig.add_subplot(4, 1, 2)
        self.ax_p = self.fig.add_subplot(4, 1, 3)
        self.ax_e = self.fig.add_subplot(4, 1, 4)

        for ax, title, color, unit in [
            (self.ax_v, "Spanning",       ACCENT2, "V"),
            (self.ax_i, "Stroom",         ACCENT3, "mA"),
            (self.ax_p, "Vermogen",       ACCENT,  "mW"),
            (self.ax_e, "Totale energie", ACCENT5, "mWh"),
        ]:
            ax.set_title(title, color=color, fontsize=9,
                         fontweight="bold", loc="left", pad=3)
            ax.set_ylabel(unit, fontsize=8)
            ax.grid(True, linestyle="--")
            ax.tick_params(labelsize=7)

        self.line_v, = self.ax_v.plot([], [], color=ACCENT2, lw=1.8)
        self.line_i, = self.ax_i.plot([], [], color=ACCENT3, lw=1.8)
        self.line_p, = self.ax_p.plot([], [], color=ACCENT,  lw=1.8)
        self.line_e, = self.ax_e.plot([], [], color=ACCENT5, lw=1.8)

        # Markeerlijnen voor meting-start (worden getekend bij start meting)
        self._vlines = []
        self._energy_fill = None

        self.canvas = FigureCanvasTkAgg(self.fig, master=self._graph_frame)
        self.canvas.get_tk_widget().pack(fill="both", expand=True)

        self._anim = animation.FuncAnimation(
            self.fig, self._animate, interval=500,
            blit=False, cache_frame_data=False
        )

    def _animate(self, _frame):
        if not self.timestamps:
            return
        t = list(self.timestamps)
        for line, buf, ax in [
            (self.line_v, self.voltages, self.ax_v),
            (self.line_i, self.currents, self.ax_i),
            (self.line_p, self.powers,   self.ax_p),
            (self.line_e, self.energies, self.ax_e),
        ]:
            data = list(buf)
            line.set_data(t, data)
            ax.set_xlim(max(0, t[-1] - 120), t[-1] + 2)
            if data:
                mn, mx = min(data), max(data)
                pad = (mx - mn) * 0.15 or 0.5
                ax.set_ylim(mn - pad, mx + pad)

        if self._energy_fill:
            try:
                self._energy_fill.remove()
            except Exception:
                pass
            self._energy_fill = None
        e_data = list(self.energies)
        if e_data:
            self._energy_fill = self.ax_e.fill_between(
                t, e_data, alpha=0.18, color=ACCENT5)

    # ─────────────────────────────────────────────────────────────────────────
    # Meting-modus
    # ─────────────────────────────────────────────────────────────────────────
    def _toggle_meting(self):
        if self._meting_active:
            self._stop_meting(manual=True)
        else:
            self._start_meting()

    def _start_meting(self):
        try:
            duur = int(self._duur_var.get())
            if duur < 1:
                raise ValueError
        except ValueError:
            messagebox.showerror("Ongeldige duur",
                                 "Vul een geheel getal in seconden in (minimaal 1).")
            return

        if not (self.reading or self._demo_running):
            messagebox.showwarning("Niet verbonden",
                                   "Verbind eerst met de Arduino of start de demo-modus.")
            return

        self._meting_duur_s   = duur
        self._meting_active   = True
        self._meting_start_ts = time.time()
        self._meting_rows     = []

        self._meting_btn.configure(text="■ Stop meting", style="Danger.TButton")
        self._meting_progress["maximum"] = duur
        self._meting_progress["value"]   = 0

        # Verticale lijn in grafieken bij meting-start
        if self.timestamps:
            t_now = list(self.timestamps)[-1]
            for ax in (self.ax_v, self.ax_i, self.ax_p, self.ax_e):
                vl = ax.axvline(t_now, color=ACCENT3, lw=1.2,
                                linestyle="--", alpha=0.7)
                self._vlines.append(vl)

        self._set_status(f"Meting actief  ({duur}s)", error=False)
        self._tick_meting()

    def _tick_meting(self):
        if not self._meting_active:
            return
        elapsed = time.time() - self._meting_start_ts
        remaining = self._meting_duur_s - elapsed

        self._meting_progress["value"] = min(elapsed, self._meting_duur_s)

        # Timer label
        el_str  = time.strftime("%M:%S", time.gmtime(int(elapsed)))
        tot_str = time.strftime("%M:%S", time.gmtime(self._meting_duur_s))
        self._meting_timer_lbl.configure(
            text=f"{el_str}  /  {tot_str}", fg=ACCENT if remaining > 5 else ACCENT4)

        # Live statistieken
        if self._meting_rows:
            powers = [r["power_mW"] for r in self._meting_rows]
            gem = sum(powers) / len(powers)
            mx  = max(powers)
            mn  = min(powers)
            # Energie van de meting (mWh)
            e = sum(r["power_mW"] * r["dt_s"] / 3600 for r in self._meting_rows)
            self._mst_gem.configure(text=f"{gem:.1f} mW")
            self._mst_max.configure(text=f"{mx:.1f} mW")
            self._mst_min.configure(text=f"{mn:.1f} mW")
            self._mst_e.configure(text=f"{e:.4f} mWh")

        if remaining <= 0:
            self._stop_meting(manual=False)
            return

        self._meting_timer_id = self.after(200, self._tick_meting)

    def _stop_meting(self, manual=False):
        self._meting_active = False
        if self._meting_timer_id:
            self.after_cancel(self._meting_timer_id)
            self._meting_timer_id = None

        self._meting_btn.configure(text="▶ Start meting", style="Green.TButton")
        self._meting_progress["value"] = self._meting_duur_s if not manual else \
            self._meting_progress["value"]

        reden = "handmatig gestopt" if manual else "voltooid"
        self._set_status(f"Meting {reden}  ({len(self._meting_rows)} samples)", error=False)

        # Resultaat tonen
        if self._meting_rows:
            powers = [r["power_mW"] for r in self._meting_rows]
            gem = sum(powers) / len(powers)
            mx  = max(powers)
            mn  = min(powers)
            e   = sum(r["power_mW"] * r["dt_s"] / 3600 for r in self._meting_rows)
            msg = (f"Meting {reden}\n\n"
                   f"Samples  : {len(self._meting_rows)}\n"
                   f"Gemiddeld: {gem:.2f} mW\n"
                   f"Maximum  : {mx:.2f} mW\n"
                   f"Minimum  : {mn:.2f} mW\n"
                   f"Energie  : {e:.5f} mWh\n\n"
                   f"Wil je de meetdata exporteren?")
            if messagebox.askyesno("Meting klaar", msg):
                self._export_csv(meting_only=True)

    # ─────────────────────────────────────────────────────────────────────────
    # Seriële verbinding
    # ─────────────────────────────────────────────────────────────────────────
    def _refresh_ports(self):
        if not SERIAL_AVAILABLE:
            self._port_combo["values"] = ["(pyserial niet gevonden)"]
            return
        ports = [p.device for p in serial.tools.list_ports.comports()]
        self._port_combo["values"] = ports or ["(geen poorten)"]
        if ports:
            self._port_combo.current(0)

    def _toggle_connect(self):
        if self.reading:
            self._disconnect()
        else:
            self._connect()

    def _connect(self):
        if not SERIAL_AVAILABLE:
            self._set_status("pyserial niet geïnstalleerd!", error=True)
            return
        port = self._port_var.get()
        baud = int(self._baud_var.get())
        try:
            self.serial_port = serial.Serial(port, baud, timeout=1)
            time.sleep(1.5)
            self.reading    = True
            self.start_time = time.time()
            self._connect_btn.configure(text="Verbreken", style="Danger.TButton")
            self._set_status(f"Verbonden  {port} @ {baud}", error=False)
            self._set_stat(self._serial_st, True)
            self._read_thread = threading.Thread(target=self._read_loop, daemon=True)
            self._read_thread.start()
        except Exception as e:
            self._set_status(str(e), error=True)

    def _disconnect(self):
        self.reading = False
        if self.serial_port:
            try:
                self.serial_port.close()
            except Exception:
                pass
            self.serial_port = None
        self._connect_btn.configure(text="Verbinden", style="Accent.TButton")
        self._set_status("Verbinding verbroken", error=True)
        self._set_stat(self._serial_st, False)

    def _read_loop(self):
        while self.reading and self.serial_port:
            try:
                raw = self.serial_port.readline().decode("utf-8", errors="ignore").strip()
                if raw:
                    self._parse_line(raw)
            except Exception:
                break
        self.reading = False

    # ─────────────────────────────────────────────────────────────────────────
    # Demo-modus
    # ─────────────────────────────────────────────────────────────────────────
    def _toggle_demo(self):
        if self._demo_running:
            self._demo_running = False
            self._demo_btn.configure(text="▶ Demo")
            self._set_status("Demo gestopt", error=True)
            self._set_stat(self._serial_st, False)
        else:
            self._demo_running = True
            self._demo_btn.configure(text="■ Stop demo")
            self.start_time = time.time()
            self.reading    = True
            self._set_status("Demo-modus actief", error=False)
            self._set_stat(self._serial_st, True)
            threading.Thread(target=self._demo_loop, daemon=True).start()

    def _demo_loop(self):
        import math, random
        t = 0
        while self._demo_running:
            t += 1
            v = 5.0 + 0.5 * math.sin(t / 10) + random.uniform(-0.05, 0.05)
            i = 80  + 20  * math.sin(t / 8 + 1) + random.uniform(-2, 2)
            p = v * i
            top    = 500 + int(200 * math.sin(t / 15)) + random.randint(-10, 10)
            right  = 480 + int(200 * math.cos(t / 15)) + random.randint(-10, 10)
            bottom = 1023 - top  + random.randint(-10, 10)
            left   = 1023 - right + random.randint(-10, 10)
            rot  = int(90 + 30 * math.sin(t / 20))
            hoek = int(90 + 20 * math.sin(t / 30))
            self.after(0, self._update_ui, {
                "voltage": v, "current": i, "power": p,
                "top": top, "right": right, "bottom": bottom, "left": left,
                "rot": rot, "hoek": hoek,
            })
            time.sleep(0.5)

    # ─────────────────────────────────────────────────────────────────────────
    # Data verwerken
    # ─────────────────────────────────────────────────────────────────────────
    def _parse_line(self, line):
        m = PATTERN.search(line)
        if not m:
            return
        self.after(0, self._update_ui, {
            "top":     int(m.group(1)),
            "right":   int(m.group(3)),
            "bottom":  int(m.group(4)),
            "left":    int(m.group(2)),
            "voltage": float(m.group(5)),
            "current": float(m.group(6)),
            "power":   float(m.group(7)),
            "rot":     int(m.group(8)),
            "hoek":    int(m.group(9)),
        })

    def _update_ui(self, d):
        now_ts = time.time()
        t      = now_ts - (self.start_time or now_ts)

        # Energie
        dt_s = 0.0
        if self._last_sample_time is not None:
            dt_s = now_ts - self._last_sample_time
            self._total_energy_mWh += d["power"] * dt_s / 3600
        self._last_sample_time = now_ts

        # Buffers
        self.timestamps.append(t)
        self.voltages.append(d["voltage"])
        self.currents.append(d["current"])
        self.powers.append(d["power"])
        self.energies.append(self._total_energy_mWh)
        self.ldr_top.append(d["top"])
        self.ldr_right.append(d["right"])
        self.ldr_bottom.append(d["bottom"])
        self.ldr_left.append(d["left"])
        self.rot_pos.append(d["rot"])
        self.hoek_pos.append(d["hoek"])

        # Alle rijen voor CSV
        row = {
            "tijd_s":      round(t, 2),
            "timestamp":   datetime.datetime.now().isoformat(timespec="milliseconds"),
            "voltage_V":   round(d["voltage"],  3),
            "current_mA":  round(d["current"],  3),
            "power_mW":    round(d["power"],     3),
            "energie_mWh": round(self._total_energy_mWh, 6),
            "ldr_top":     d["top"],
            "ldr_right":   d["right"],
            "ldr_bottom":  d["bottom"],
            "ldr_left":    d["left"],
            "rot_deg":     d["rot"],
            "hoek_deg":    d["hoek"],
            "in_meting":   self._meting_active,
        }
        self._all_rows.append(row)

        # Meting-buffer
        if self._meting_active:
            mrow = dict(row)
            mrow["dt_s"] = dt_s
            self._meting_rows.append(mrow)

        # Labels bijwerken
        self._v_lbl.configure(text=f'{d["voltage"]:.2f} V')
        self._i_lbl.configure(text=f'{d["current"]:.1f} mA')
        self._p_lbl.configure(text=f'{d["power"]:.1f} mW')
        self._e_lbl.configure(text=f'{self._total_energy_mWh:.4f} mWh')
        self._samples_lbl.configure(text=str(len(self.timestamps)))
        self._rot_lbl.configure(text=f'{d["rot"]}°')
        self._hoek_lbl.configure(text=f'{d["hoek"]}°')

        ina_ok = d["voltage"] > 0.05 or d["current"] > 0.5
        self._set_stat(self._ina_st, ina_ok)

        for key, name in [("top","Top"),("right","Rechts"),("bottom","Onder"),("left","Links")]:
            val = d[key]
            bi  = self._ldr_bars[name]
            bi["lbl"].configure(text=str(val))
            bw = bi["bg"].winfo_width() or 100
            bi["bar"].place(x=0, y=0, relheight=1, width=int(val / 1023 * bw))

        ldr_ok = all(0 < d[k] < 1023 for k in ("top","right","bottom","left"))
        self._set_stat(self._ldr_st, ldr_ok)

    def _set_stat(self, lbl, ok, on="OK", off="FOUT", c_on=None, c_off=None):
        c_on  = c_on  or ACCENT3
        c_off = c_off or ACCENT4
        lbl.configure(text=f"● {on}" if ok else f"● {off}",
                      fg=c_on if ok else c_off)

    def _set_status(self, msg, error=False):
        c = ACCENT4 if error else ACCENT3
        self._status_dot.configure(fg=c)
        self._status_lbl.configure(text=msg, fg=c)

    # ─────────────────────────────────────────────────────────────────────────
    # Export
    # ─────────────────────────────────────────────────────────────────────────
    def _export_csv(self, meting_only=False):
        rows = self._meting_rows if meting_only else self._all_rows
        if not rows:
            messagebox.showinfo("Geen data", "Er zijn nog geen samples om te exporteren.")
            return

        ts    = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        label = "meting" if meting_only else "alle_data"
        init  = f"solar_tracker_{label}_{ts}.csv"

        path = filedialog.asksaveasfilename(
            defaultextension=".csv",
            initialfile=init,
            filetypes=[("CSV bestand", "*.csv"), ("Alle bestanden", "*.*")],
            title="Sla CSV op"
        )
        if not path:
            return

        fieldnames = list(rows[0].keys())
        # dt_s niet in export als het er niet in zit (all_rows hebben het niet)
        fieldnames = [f for f in fieldnames if f != "dt_s"]

        try:
            with open(path, "w", newline="", encoding="utf-8") as f:
                writer = csv.DictWriter(f, fieldnames=fieldnames, extrasaction="ignore")
                writer.writeheader()
                writer.writerows(rows)
            messagebox.showinfo("Geëxporteerd", f"CSV opgeslagen:\n{path}")
        except Exception as e:
            messagebox.showerror("Fout", str(e))

    def _export_png(self):
        if not MATPLOTLIB_AVAILABLE or not self.timestamps:
            messagebox.showinfo("Geen data", "Geen grafiekdata beschikbaar.")
            return

        ts   = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        path = filedialog.asksaveasfilename(
            defaultextension=".png",
            initialfile=f"solar_tracker_grafieken_{ts}.png",
            filetypes=[("PNG afbeelding", "*.png"), ("Alle bestanden", "*.*")],
            title="Sla grafiek op"
        )
        if not path:
            return

        try:
            # Aparte exportfiguur (niet de live figuur, zodat die blijft draaien)
            fig, axes = plt.subplots(4, 1, figsize=(14, 10),
                                     facecolor=GRAPH_BG)
            fig.subplots_adjust(hspace=0.55, top=0.93, bottom=0.07,
                                left=0.09, right=0.97)
            fig.suptitle(
                f"Solar Tracker  —  Export {datetime.datetime.now().strftime('%d-%m-%Y %H:%M:%S')}",
                color=TEXT, fontsize=12, fontweight="bold"
            )

            t = list(self.timestamps)
            datasets = [
                (list(self.voltages),  "Spanning (V)",       ACCENT2),
                (list(self.currents),  "Stroom (mA)",        ACCENT3),
                (list(self.powers),    "Vermogen (mW)",      ACCENT),
                (list(self.energies),  "Totale energie (mWh)", ACCENT5),
            ]

            for ax, (data, title, color) in zip(axes, datasets):
                ax.set_facecolor(GRAPH_BG)
                ax.plot(t, data, color=color, lw=1.8)
                ax.set_title(title, color=color, fontsize=9,
                             fontweight="bold", loc="left", pad=3)
                ax.grid(True, linestyle="--", color=BORDER, alpha=0.4)
                ax.tick_params(colors=TEXT_MUTED, labelsize=7)
                for spine in ax.spines.values():
                    spine.set_edgecolor(BORDER)
                ax.set_xlabel("Tijd (s)", fontsize=7, color=TEXT_MUTED)
                if data:
                    mn, mx = min(data), max(data)
                    pad = (mx - mn) * 0.15 or 0.5
                    ax.set_ylim(mn - pad, mx + pad)
                # Annotaties: min/max
                if len(data) > 1:
                    i_max = data.index(max(data))
                    i_min = data.index(min(data))
                    ax.annotate(f"max {max(data):.2f}",
                                xy=(t[i_max], data[i_max]),
                                fontsize=7, color=color, ha="center",
                                xytext=(0, 8), textcoords="offset points")
                    ax.annotate(f"min {min(data):.2f}",
                                xy=(t[i_min], data[i_min]),
                                fontsize=7, color=TEXT_MUTED, ha="center",
                                xytext=(0, -12), textcoords="offset points")

                # Markeer meting-rijen
                if self._meting_rows:
                    m_t0 = self._meting_rows[0]["tijd_s"]
                    m_t1 = self._meting_rows[-1]["tijd_s"]
                    ax.axvspan(m_t0, m_t1, alpha=0.08, color=ACCENT3)

            fig.savefig(path, dpi=150, facecolor=GRAPH_BG, bbox_inches="tight")
            plt.close(fig)
            messagebox.showinfo("Geëxporteerd", f"Grafiek opgeslagen:\n{path}")
        except Exception as e:
            messagebox.showerror("Fout", str(e))

    # ─────────────────────────────────────────────────────────────────────────
    # Hulp
    # ─────────────────────────────────────────────────────────────────────────
    def _clear_data(self):
        for buf in (self.timestamps, self.voltages, self.currents, self.powers,
                    self.energies, self.ldr_top, self.ldr_right,
                    self.ldr_bottom, self.ldr_left, self.rot_pos, self.hoek_pos):
            buf.clear()
        self._all_rows.clear()
        self._meting_rows.clear()
        self._total_energy_mWh = 0.0
        self._last_sample_time = None
        self.start_time        = time.time()

    def _on_close(self):
        self._demo_running = False
        self._meting_active = False
        self.reading = False
        if self.serial_port:
            try:
                self.serial_port.close()
            except Exception:
                pass
        self.destroy()


# ─────────────────────────────────────────────────────────────────────────────
if __name__ == "__main__":
    app = SolarTrackerApp()
    app.mainloop()