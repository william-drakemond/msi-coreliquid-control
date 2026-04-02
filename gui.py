"""
MSI MPG Coreliquid K360 - Control Panel
Requires: pip install liquidctl pillow pystray
"""

import tkinter as tk
from tkinter import ttk, messagebox
import threading
import time
import json
import os
import sys

try:
    import pystray
    from PIL import Image, ImageDraw
    _TRAY_AVAILABLE = True
except ImportError:
    _TRAY_AVAILABLE = False

# When frozen as .exe, save settings next to the exe, not in the temp extract dir
if getattr(sys, "frozen", False):
    _BASE_DIR = os.path.dirname(sys.executable)
else:
    _BASE_DIR = os.path.dirname(__file__)
SETTINGS_FILE = os.path.join(_BASE_DIR, "settings.json")

try:
    from liquidctl.driver.msi import MpgCooler
except ImportError:
    messagebox.showerror("Error", "liquidctl not installed.\nRun: pip install liquidctl")
    raise SystemExit

# ── colours ───────────────────────────────────────────────────────────────────
BG      = "#1a1a2e"
CARD    = "#16213e"
CARD2   = "#0f1829"
ACCENT  = "#e94560"
TEXT    = "#eaeaea"
SUBTEXT = "#8888aa"
GREEN   = "#00d4aa"
YELLOW  = "#f5a623"
RED     = "#e94560"
BLUE    = "#4fc3f7"

MODES = ["silent", "balance", "game", "smart"]
MODE_COLORS = {"silent": "#4fc3f7", "balance": "#00d4aa", "game": "#f5a623", "smart": "#e94560"}

# Preset speeds per mode: {key: duty%}  (radfans, waterblock, pump)
MODE_PRESETS = {
    "silent":  {"radfans": 60,  "waterblock": 60,  "pump": 60},
    "balance": {"radfans": 70,  "waterblock": 70,  "pump": 70},
    "game":    {"radfans": 90,  "waterblock": 90,  "pump": 100},
    "smart":   {"radfans": 80,  "waterblock": 80,  "pump": 75},
}

CHANNELS = [
    # key,         label,         lctl_ch,          spd_key,             duty_key,           min, pump
    ("radfans",    "Rad Fans",    "fans",            "Fan 1 speed",       "Fan 1 duty",        0,  False),
    ("waterblock", "Water Block", "waterblock-fan", "Water block speed", "Water block duty",  0,  False),
    ("pump",       "Pump",        "pump",           "Pump speed",        "Pump duty",         30, True),
]

# extra status keys to average for the Rad Fans card
RAD_FAN_KEYS = [
    ("Fan 1 speed", "Fan 1 duty"),
    ("Fan 2 speed", "Fan 2 duty"),
    ("Fan 3 speed", "Fan 3 duty"),
]

REFRESH_MS   = 2000
TEMP_FEED_MS = 3000   # how often to send CPU temp to device when curve is active

# curve canvas geometry
CW, CH   = 460, 240   # canvas width/height
PAD      = 36          # axis padding
TMIN, TMAX = 20, 90   # temp range °C
DMIN, DMAX =  0, 100  # duty range %
PT_R     = 7           # point radius


# ── custom slider ─────────────────────────────────────────────────────────────
class Slider(tk.Canvas):
    """A clean, visible horizontal slider."""
    H, W = 28, 280
    TRACK_H = 4
    THUMB_W, THUMB_H = 12, 20

    def __init__(self, parent, from_=0, to=100, value=50,
                 color=GREEN, command=None, **kwargs):
        super().__init__(parent, width=self.W, height=self.H,
                         bg=BG, highlightthickness=0, **kwargs)
        self._min   = from_
        self._max   = to
        self._value = value
        self._color = color
        self._cmd   = command
        self._draw()
        self.bind("<ButtonPress-1>",  self._click)
        self.bind("<B1-Motion>",      self._click)

    def _val_to_x(self, v):
        frac = (v - self._min) / (self._max - self._min)
        return int(self.THUMB_W // 2 + frac * (self.W - self.THUMB_W))

    def _x_to_val(self, x):
        frac = (x - self.THUMB_W // 2) / (self.W - self.THUMB_W)
        return round(self._min + max(0.0, min(1.0, frac)) * (self._max - self._min))

    def _draw(self):
        self.delete("all")
        cy  = self.H // 2
        tx  = self._val_to_x(self._value)
        # track background
        self.create_rectangle(self.THUMB_W // 2, cy - self.TRACK_H // 2,
                              self.W - self.THUMB_W // 2, cy + self.TRACK_H // 2,
                              fill="#2a3550", outline="")
        # filled portion
        self.create_rectangle(self.THUMB_W // 2, cy - self.TRACK_H // 2,
                              tx, cy + self.TRACK_H // 2,
                              fill=self._color, outline="")
        # thumb
        self.create_rectangle(tx - self.THUMB_W // 2, cy - self.THUMB_H // 2,
                              tx + self.THUMB_W // 2, cy + self.THUMB_H // 2,
                              fill=TEXT, outline="", tags="thumb")

    def _click(self, ev):
        self._value = self._x_to_val(ev.x)
        self._draw()
        if self._cmd:
            self._cmd(self._value)

    def get(self):
        return self._value

    def set(self, v):
        self._value = max(self._min, min(self._max, v))
        self._draw()


def rpm_color(duty: int, is_pump: bool) -> str:
    if is_pump: return GREEN
    if duty < 40: return GREEN
    if duty < 70: return YELLOW
    return RED


def open_device():
    devs = list(MpgCooler.find_supported_devices())
    if not devs:
        return None
    d = devs[0]
    d.connect()
    return d


# ── curve canvas helper ───────────────────────────────────────────────────────
def temp_to_x(t):
    return PAD + (t - TMIN) / (TMAX - TMIN) * (CW - PAD * 2)

def x_to_temp(x):
    return round(TMIN + (x - PAD) / (CW - PAD * 2) * (TMAX - TMIN))

def duty_to_y(d):
    return CH - PAD - d / (DMAX - DMIN) * (CH - PAD * 2)

def y_to_duty(y):
    return round(DMAX - (y - PAD) / (CH - PAD * 2) * (DMAX - DMIN))


# ── main app ─────────────────────────────────────────────────────────────────
class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("MSI MPG Coreliquid Control")
        self.configure(bg=BG)
        self.resizable(False, False)

        self._device     = None
        self._lock       = threading.Lock()
        self._running    = True
        self._err_streak    = 0
        self._sliders_synced = False  # sync sliders from device on first poll

        # curve state
        self._curve_active  = False    # True when a curve is applied
        self._curve_channel = "fan1"
        self._curve_cpu_temp = 50      # °C fed to device
        self._points = [               # default S-curve [(temp, duty), ...]
            [30, 20], [45, 35], [55, 50], [65, 70], [75, 90], [85, 100]
        ]

        self._build_ui()
        self._load_settings()
        self._connect()
        self.protocol("WM_DELETE_WINDOW", self._hide_window)
        self._start_tray()

    # ── build ─────────────────────────────────────────────────────────────────
    def _build_ui(self):
        tk.Frame(self, bg=ACCENT, height=4).pack(fill="x")

        # header
        hdr = tk.Frame(self, bg=BG, pady=10)
        hdr.pack(fill="x", padx=20)
        tk.Label(hdr, text="MSI MPG CORELIQUID", font=("Segoe UI", 18, "bold"),
                 bg=BG, fg=TEXT).pack(side="left")
        self._status_dot = tk.Label(hdr, text="●", font=("Segoe UI", 17), bg=BG, fg=SUBTEXT)
        self._status_dot.pack(side="right")
        self._status_lbl = tk.Label(hdr, text="connecting…", font=("Segoe UI", 12),
                                    bg=BG, fg=SUBTEXT)
        self._status_lbl.pack(side="right", padx=6)

        # tabs
        style = ttk.Style()
        style.theme_use("default")
        style.configure("TNotebook",        background=BG, borderwidth=0)
        style.configure("TNotebook.Tab",    background=CARD, foreground=SUBTEXT,
                        padding=[18, 8], font=("Segoe UI", 12, "bold"))
        style.map("TNotebook.Tab",
                  background=[("selected", BG)],
                  foreground=[("selected", TEXT)])

        nb = ttk.Notebook(self)
        nb.pack(fill="both", expand=True, padx=10, pady=(0, 0))

        tab1 = tk.Frame(nb, bg=BG)
        tab3 = tk.Frame(nb, bg=BG)
        nb.add(tab1, text="  Monitor  ")
        nb.add(tab3, text=" Fan Curve ")

        self._build_monitor_tab(tab1)
        self._build_curve_tab(tab3)

        # bottom bar
        bot = tk.Frame(self, bg=CARD, pady=7, padx=16)
        bot.pack(fill="x")
        self._refresh_lbl = tk.Label(bot, text="", font=("Segoe UI", 11), bg=CARD, fg=SUBTEXT)
        self._refresh_lbl.pack(side="left")
        tk.Button(bot, text="⟳  Refresh", font=("Segoe UI", 11), bg=CARD, fg=TEXT,
                  relief="flat", bd=0, cursor="hand2",
                  command=self._poll_status).pack(side="right")

    # ── Monitor tab ───────────────────────────────────────────────────────────
    def _build_monitor_tab(self, parent):
        # status cards
        cards_f = tk.Frame(parent, bg=BG, padx=10, pady=8)
        cards_f.pack(fill="x")

        self._fan_cards = {}
        for i, (key, label, _, spd_key, duty_key, _, is_pump) in enumerate(CHANNELS):
            row, col = divmod(i, 2)
            card = self._make_status_card(cards_f, label, is_pump)
            card["frame"].grid(row=row, column=col, padx=5, pady=5,
                               sticky="nsew", ipadx=8, ipady=4)
            self._fan_cards[key] = {**card, "spd_key": spd_key,
                                    "duty_key": duty_key, "is_pump": is_pump}
        cards_f.columnconfigure(0, weight=1)
        cards_f.columnconfigure(1, weight=1)

        # mode buttons
        tk.Frame(parent, bg=SUBTEXT, height=1).pack(fill="x", padx=14, pady=(6, 0))
        mf = tk.Frame(parent, bg=BG, padx=14, pady=8)
        mf.pack(fill="x")
        tk.Label(mf, text="FAN MODE", font=("Segoe UI", 11, "bold"),
                 bg=BG, fg=SUBTEXT).pack(anchor="w")
        br = tk.Frame(mf, bg=BG)
        br.pack(fill="x", pady=(6, 0))

        self._mode_btns = {}
        for mode in MODES:
            b = tk.Button(br, text=mode.upper(), font=("Segoe UI", 12, "bold"),
                          bg=CARD, fg=SUBTEXT,
                          activebackground=MODE_COLORS[mode], activeforeground=BG,
                          relief="flat", bd=0, padx=10, pady=8, cursor="hand2",
                          command=lambda m=mode: self._set_mode(m))
            b.pack(side="left", expand=True, fill="x", padx=3)
            self._mode_btns[mode] = b

        # sliders
        tk.Frame(parent, bg=SUBTEXT, height=1).pack(fill="x", padx=14, pady=(10, 0))
        sf = tk.Frame(parent, bg=BG, padx=14, pady=8)
        sf.pack(fill="x")
        tk.Label(sf, text="SPEED CONTROL", font=("Segoe UI", 11, "bold"),
                 bg=BG, fg=SUBTEXT).pack(anchor="w", pady=(0, 6))

        self._sliders    = {}
        self._slider_val = {}

        for key, label, lctl_ch, _, _, min_duty, is_pump in CHANNELS:
            row_f = tk.Frame(sf, bg=BG)
            row_f.pack(fill="x", pady=4)

            tk.Label(row_f, text=label, font=("Segoe UI", 12, "bold"),
                     bg=BG, fg=TEXT, width=12, anchor="w").pack(side="left")

            slider = Slider(row_f, from_=min_duty, to=100, value=50,
                            color=GREEN,
                            command=lambda v, k=key: self._on_slider_move(k, v))
            slider.pack(side="left", padx=(4, 0))

            val_lbl = tk.Label(row_f, text="50%", font=("Segoe UI", 12),
                               bg=BG, fg=SUBTEXT, width=4, anchor="e")
            val_lbl.pack(side="left", padx=(4, 0))

            tk.Button(
                row_f, text="Set", font=("Segoe UI", 11),
                bg=CARD, fg=TEXT, activebackground=GREEN, activeforeground=BG,
                relief="flat", bd=0, padx=8, pady=3, cursor="hand2",
                command=lambda lc=lctl_ch, s=slider: self._apply_fixed_speed(lc, s.get())
            ).pack(side="left", padx=(6, 0))

            if is_pump:
                tk.Label(row_f, text="⚠ min 30%", font=("Segoe UI", 10),
                         bg=BG, fg=SUBTEXT).pack(side="left", padx=(6, 0))

            self._sliders[key]    = slider
            self._slider_val[key] = val_lbl

    def _make_status_card(self, parent, title, is_pump):
        frame = tk.Frame(parent, bg=CARD)
        tk.Label(frame, text=title.upper(), font=("Segoe UI", 10, "bold"),
                 bg=CARD, fg=SUBTEXT).pack(anchor="w", padx=10, pady=(6, 1))
        rpm_lbl = tk.Label(frame, text="— rpm", font=("Segoe UI", 20, "bold"),
                           bg=CARD, fg=TEXT)
        rpm_lbl.pack(anchor="w", padx=10)
        duty_lbl = tk.Label(frame, text="—%", font=("Segoe UI", 12), bg=CARD, fg=SUBTEXT)
        duty_lbl.pack(anchor="w", padx=10)
        bar_bg = tk.Frame(frame, bg=CARD2, height=3)
        bar_bg.pack(fill="x", padx=10, pady=(3, 6))
        bar_bg.pack_propagate(False)
        bar_fg = tk.Frame(bar_bg, bg=GREEN, height=3)
        bar_fg.place(x=0, y=0, relheight=1, relwidth=0)
        return {"frame": frame, "rpm": rpm_lbl, "duty": duty_lbl,
                "bar": bar_fg, "bar_bg": bar_bg}

    # ── Curve tab ─────────────────────────────────────────────────────────────
    def _build_curve_tab(self, parent):
        # channel selector + controls row
        ctrl = tk.Frame(parent, bg=BG, padx=14, pady=8)
        ctrl.pack(fill="x")

        tk.Label(ctrl, text="Channel:", font=("Segoe UI", 12, "bold"),
                 bg=BG, fg=TEXT).pack(side="left")

        self._curve_ch_var = tk.StringVar(value="fan1")
        ch_options = [(label, lctl_ch) for _, label, lctl_ch, *_ in CHANNELS]
        ch_menu = ttk.Combobox(ctrl, textvariable=self._curve_ch_var,
                               values=[lbl for lbl, _ in ch_options],
                               state="readonly", width=13,
                               font=("Segoe UI", 12))
        ch_menu.pack(side="left", padx=(6, 0))
        self._ch_label_to_lctl = {lbl: lc for lbl, lc in ch_options}

        tk.Button(ctrl, text="Apply Curve", font=("Segoe UI", 12, "bold"),
                  bg=GREEN, fg=BG, activebackground=GREEN,
                  relief="flat", bd=0, padx=12, pady=5, cursor="hand2",
                  command=self._apply_curve).pack(side="right", padx=4)

        tk.Button(ctrl, text="Reset Points", font=("Segoe UI", 11),
                  bg=CARD, fg=TEXT, relief="flat", bd=0, padx=10, pady=5,
                  cursor="hand2", command=self._reset_curve).pack(side="right", padx=4)

        # canvas
        canvas_f = tk.Frame(parent, bg=CARD, padx=6, pady=6)
        canvas_f.pack(padx=14, pady=(0, 6))

        self._canvas = tk.Canvas(canvas_f, width=CW, height=CH,
                                 bg=CARD2, highlightthickness=0)
        self._canvas.pack()
        self._draw_curve()

        self._canvas.bind("<ButtonPress-1>",   self._curve_press)
        self._canvas.bind("<B1-Motion>",       self._curve_drag)
        self._canvas.bind("<ButtonRelease-1>", self._curve_release)
        self._drag_idx = None

        # hint
        tk.Label(parent,
                 text="Drag points to set the curve  •  Right-click a point to remove  •  Click empty area to add",
                 font=("Segoe UI", 10), bg=BG, fg=SUBTEXT).pack()
        self._canvas.bind("<ButtonPress-3>", self._curve_right_click)

        # CPU temp feed section
        sep = tk.Frame(parent, bg=SUBTEXT, height=1)
        sep.pack(fill="x", padx=14, pady=(8, 0))

        tf = tk.Frame(parent, bg=BG, padx=14, pady=8)
        tf.pack(fill="x")

        tk.Label(tf, text="CPU Temp (°C):", font=("Segoe UI", 12, "bold"),
                 bg=BG, fg=TEXT).pack(side="left")

        self._cpu_temp_var = tk.IntVar(value=50)
        self._cpu_temp_lbl = tk.Label(tf, text="50°C", font=("Segoe UI", 12),
                                      bg=BG, fg=BLUE, width=5)

        def _on_temp(v):
            self._cpu_temp_var.set(v)
            self._cpu_temp_lbl.config(text=f"{v}°C")

        temp_slider = Slider(tf, from_=20, to=100, value=50, color=BLUE, command=_on_temp)
        temp_slider.pack(side="left", padx=(6, 0))
        self._cpu_temp_lbl.pack(side="left", padx=(4, 0))

        self._feed_btn = tk.Button(
            tf, text="Start Feeding", font=("Segoe UI", 11, "bold"),
            bg=CARD, fg=TEXT, activebackground=BLUE, activeforeground=BG,
            relief="flat", bd=0, padx=10, pady=4, cursor="hand2",
            command=self._toggle_temp_feed
        )
        self._feed_btn.pack(side="right")

        self._feed_lbl = tk.Label(tf, text="ℹ Curves need CPU temp to be sent continuously",
                                  font=("Segoe UI", 10), bg=BG, fg=SUBTEXT)
        self._feed_lbl.pack(side="right", padx=(0, 8))

        self._temp_feeding = False

    # ── canvas draw ───────────────────────────────────────────────────────────
    def _draw_curve(self):
        c = self._canvas
        c.delete("all")

        # grid
        for t in range(TMIN, TMAX + 1, 10):
            x = temp_to_x(t)
            c.create_line(x, PAD, x, CH - PAD, fill="#1a2a40", width=1)
            c.create_text(x, CH - PAD + 10, text=f"{t}", fill=SUBTEXT,
                          font=("Segoe UI", 10))

        for d in range(0, 101, 20):
            y = duty_to_y(d)
            c.create_line(PAD, y, CW - PAD, y, fill="#1a2a40", width=1)
            c.create_text(PAD - 10, y, text=f"{d}%", fill=SUBTEXT,
                          font=("Segoe UI", 10))

        # axis labels
        c.create_text(CW // 2, CH - 4, text="Temperature (°C)",
                      fill=SUBTEXT, font=("Segoe UI", 10))
        c.create_text(10, CH // 2, text="Duty %", fill=SUBTEXT,
                      font=("Segoe UI", 10), angle=90)

        # axes
        c.create_line(PAD, PAD, PAD, CH - PAD, fill=SUBTEXT, width=1)
        c.create_line(PAD, CH - PAD, CW - PAD, CH - PAD, fill=SUBTEXT, width=1)

        # curve line
        sorted_pts = sorted(self._points)
        if len(sorted_pts) >= 2:
            coords = []
            for t, d in sorted_pts:
                coords += [temp_to_x(t), duty_to_y(d)]
            c.create_line(*coords, fill=GREEN, width=2, smooth=False, tags="line")

        # points
        for i, (t, d) in enumerate(self._points):
            x, y = temp_to_x(t), duty_to_y(d)
            c.create_oval(x - PT_R, y - PT_R, x + PT_R, y + PT_R,
                          fill=BLUE, outline=TEXT, width=1, tags=f"pt{i}")
            c.create_text(x, y - PT_R - 6, text=f"{t}°/{d}%",
                          fill=TEXT, font=("Segoe UI", 12))

    def _nearest_point(self, x, y):
        best_i, best_d = None, PT_R * 2
        for i, (t, d) in enumerate(self._points):
            px, py = temp_to_x(t), duty_to_y(d)
            dist = ((x - px) ** 2 + (y - py) ** 2) ** 0.5
            if dist < best_d:
                best_d, best_i = dist, i
        return best_i

    def _curve_press(self, ev):
        self._drag_idx = self._nearest_point(ev.x, ev.y)
        if self._drag_idx is None:
            # add new point
            t = max(TMIN, min(TMAX, x_to_temp(ev.x)))
            d = max(DMIN, min(DMAX, y_to_duty(ev.y)))
            if len(self._points) < 7:
                self._points.append([t, d])
                self._draw_curve()

    def _curve_drag(self, ev):
        if self._drag_idx is None:
            return
        t = max(TMIN, min(TMAX, x_to_temp(ev.x)))
        d = max(DMIN, min(DMAX, y_to_duty(ev.y)))
        self._points[self._drag_idx] = [t, d]
        self._draw_curve()

    def _curve_release(self, ev):
        self._drag_idx = None

    def _curve_right_click(self, ev):
        i = self._nearest_point(ev.x, ev.y)
        if i is not None and len(self._points) > 2:
            self._points.pop(i)
            self._draw_curve()

    def _reset_curve(self):
        self._points = [[30, 20], [45, 35], [55, 50], [65, 70], [75, 90], [85, 100]]
        self._draw_curve()

    # ── apply curve ───────────────────────────────────────────────────────────
    def _apply_curve(self):
        label = self._curve_ch_var.get()
        lctl_ch = self._ch_label_to_lctl.get(label, "fan1")
        profile = sorted([(d, t) for t, d in self._points])  # (duty, temp) pairs

        def _do():
            with self._lock:
                dev = self._device
                if not dev:
                    return
                try:
                    dev.set_speed_profile(lctl_ch, profile)
                    self._curve_active  = True
                    self._curve_channel = lctl_ch
                    self.after(0, lambda: self._status_lbl.config(text=f"curve applied → {label}"))
                    self.after(0, lambda: self._clear_mode_highlight())
                except Exception as e:
                    msg = str(e)
                    self.after(0, lambda m=msg: messagebox.showerror("Error", m))
        threading.Thread(target=_do, daemon=True).start()

    # ── CPU temp feed ─────────────────────────────────────────────────────────
    def _toggle_temp_feed(self):
        if self._temp_feeding:
            self._temp_feeding = False
            self._feed_btn.config(text="Start Feeding", bg=CARD, fg=TEXT)
        else:
            self._temp_feeding = True
            self._feed_btn.config(text="Stop Feeding", bg=BLUE, fg=BG)
            self._send_cpu_temp()

    def _send_cpu_temp(self):
        if not self._temp_feeding or not self._running:
            return
        temp = self._cpu_temp_var.get()

        def _do():
            with self._lock:
                dev = self._device
                if dev:
                    try:
                        dev.set_oled_show_cpu_status(temp, 100)
                    except Exception:
                        pass
        threading.Thread(target=_do, daemon=True).start()
        self.after(TEMP_FEED_MS, self._send_cpu_temp)

    # ── device ────────────────────────────────────────────────────────────────
    def _connect(self):
        def _do():
            dev = open_device()
            with self._lock:
                self._device = dev
            self.after(0, self._on_connected if dev else self._on_disconnected)
        threading.Thread(target=_do, daemon=True).start()

    def _on_connected(self):
        self._status_dot.config(fg=GREEN)
        self._status_lbl.config(text="connected")
        self._poll_status()

    def _on_disconnected(self):
        self._status_dot.config(fg=RED)
        self._status_lbl.config(text="device not found")

    # ── poll ──────────────────────────────────────────────────────────────────
    def _poll_status(self):
        def _do():
            with self._lock:
                dev = self._device
                if not dev:
                    return
                try:
                    raw  = dev.get_status()
                    data = {k.strip(): v for k, v, *_ in raw}
                    self._err_streak = 0
                    self.after(0, lambda d=data: self._update_cards(d))
                    self.after(0, lambda: self._status_lbl.config(text="connected"))
                    if not self._sliders_synced:
                        self._sliders_synced = True
                        self.after(0, lambda d=data: self._sync_sliders(d))
                except Exception as e:
                    self._err_streak += 1
                    if self._err_streak >= 3:
                        msg = str(e)
                        self.after(0, lambda m=msg: self._status_lbl.config(text=f"error: {m}"))
        threading.Thread(target=_do, daemon=True).start()

        if self._running:
            self._poll_job = self.after(REFRESH_MS, self._poll_status)
            self._refresh_lbl.config(
                text=f"auto-refresh every {REFRESH_MS // 1000}s  •  {time.strftime('%H:%M:%S')}"
            )

    def _update_cards(self, data):
        for key, card in self._fan_cards.items():
            if key == "radfans":
                # average the three radiator fans, skip stopped ones
                rpms  = [data.get(sk, 0) for sk, _ in RAD_FAN_KEYS]
                duties = [data.get(dk, 0) for _, dk in RAD_FAN_KEYS]
                active = [d for d in duties if d > 0] or [0]
                rpm  = round(sum(r for r in rpms if r > 0) / max(len([r for r in rpms if r > 0]), 1))
                duty = round(sum(active) / len(active))
            else:
                rpm  = data.get(card["spd_key"], 0)
                duty = data.get(card["duty_key"], 0)
            color = rpm_color(duty, card["is_pump"])
            card["rpm"].config(text=f"{rpm:,} rpm", fg=color)
            card["duty"].config(text=f"{duty}%")
            card["bar"].config(bg=color)
            card["bar_bg"].update_idletasks()
            card["bar"].place(relwidth=min(duty / 100, 1.0), width=0)

    # ── manual speed ──────────────────────────────────────────────────────────
    def _on_slider_move(self, key, value):
        self._slider_val[key].config(text=f"{value}%")

    def _sync_sliders(self, data: dict):
        """On first connect, set sliders to match current device duty%."""
        duty_map = {
            "radfans":    data.get("Fan 1 duty", 50),
            "waterblock": data.get("Water block duty", 50),
            "pump":       data.get("Pump duty", 100),
        }
        for key, duty in duty_map.items():
            if key in self._sliders:
                self._sliders[key].set(duty)
                self._slider_val[key].config(text=f"{duty}%")

    def _apply_fixed_speed(self, lctl_ch, duty):
        def _do():
            with self._lock:
                dev = self._device
                if not dev:
                    return
                try:
                    dev.set_fixed_speed(lctl_ch, duty)
                    self.after(0, self._clear_mode_highlight)
                    self.after(0, self._save_settings)
                except Exception as e:
                    msg = str(e)
                    self.after(0, lambda m=msg: messagebox.showerror("Error", m))
        threading.Thread(target=_do, daemon=True).start()

    # ── modes ─────────────────────────────────────────────────────────────────
    def _set_mode(self, mode):
        preset = MODE_PRESETS[mode]

        def _do():
            with self._lock:
                dev = self._device
                if not dev:
                    return
                try:
                    for key, duty in preset.items():
                        lctl_ch = next(lc for k, _, lc, *_ in CHANNELS if k == key)
                        dev.set_fixed_speed(lctl_ch, duty)
                    self._curve_active = False
                    self.after(0, lambda m=mode, p=preset: self._apply_mode_ui(m, p))
                except Exception as e:
                    msg = str(e)
                    self.after(0, lambda m=msg: messagebox.showerror("Error", m))
        threading.Thread(target=_do, daemon=True).start()

    def _apply_mode_ui(self, mode, preset):
        """Update sliders and highlight button after a mode is applied."""
        for key, duty in preset.items():
            if key in self._sliders:
                self._sliders[key].set(duty)
                self._slider_val[key].config(text=f"{duty}%")
        self._highlight_mode(mode)

    def _highlight_mode(self, active):
        self._last_mode = active
        for mode, btn in self._mode_btns.items():
            btn.config(bg=MODE_COLORS[mode] if mode == active else CARD,
                       fg=BG if mode == active else SUBTEXT)
        self._save_settings()

    def _clear_mode_highlight(self):
        for btn in self._mode_btns.values():
            btn.config(bg=CARD, fg=SUBTEXT)

    # ── settings persist ──────────────────────────────────────────────────────
    def _save_settings(self):
        data = {
            "mode": getattr(self, "_last_mode", None),
            "sliders": {k: s.get() for k, s in self._sliders.items()},
            "cpu_temp": self._cpu_temp_var.get(),
        }
        try:
            with open(SETTINGS_FILE, "w") as f:
                json.dump(data, f, indent=2)
        except Exception:
            pass

    def _load_settings(self):
        try:
            with open(SETTINGS_FILE) as f:
                data = json.load(f)
        except Exception:
            return
        # restore slider positions
        for key, val in data.get("sliders", {}).items():
            if key in self._sliders:
                self._sliders[key].set(val)
                self._slider_val[key].config(text=f"{val}%")
        # restore CPU temp slider
        if "cpu_temp" in data:
            self._cpu_temp_var.set(data["cpu_temp"])
            self._cpu_temp_lbl.config(text=f"{data['cpu_temp']}°C")
        # restore mode highlight (visual only — don't re-send to device)
        if data.get("mode"):
            self._last_mode = data["mode"]
            self.after(100, lambda m=data["mode"]: self._highlight_mode(m))

    # ── system tray ───────────────────────────────────────────────────────────
    def _make_tray_image(self):
        size = 64
        img = Image.new("RGBA", (size, size), (0, 0, 0, 0))
        d = ImageDraw.Draw(img)
        d.ellipse([2, 2, 62, 62], fill=(233, 69, 96, 255))   # red circle
        # simple fan blades
        d.ellipse([22, 10, 42, 30], fill=(255, 255, 255, 180))
        d.ellipse([10, 36, 30, 56], fill=(255, 255, 255, 180))
        d.ellipse([36, 36, 56, 56], fill=(255, 255, 255, 180))
        d.ellipse([24, 24, 40, 40], fill=(233, 69, 96, 255))  # hub
        return img

    def _start_tray(self):
        if not _TRAY_AVAILABLE:
            return
        menu = pystray.Menu(
            pystray.MenuItem("Open", self._show_window, default=True),
            pystray.MenuItem("Exit", self._quit_app),
        )
        self._tray = pystray.Icon(
            "MSI Coreliquid", self._make_tray_image(),
            "MSI MPG Coreliquid Control", menu
        )
        self._tray.run_detached()

    def _hide_window(self):
        self.withdraw()

    def _show_window(self, icon=None, item=None):
        self.after(0, self.deiconify)
        self.after(0, self.lift)

    # ── cleanup ───────────────────────────────────────────────────────────────
    def _quit_app(self, icon=None, item=None):
        # Called from pystray thread — must not touch tkinter directly
        self.after(0, self._do_quit)

    def _do_quit(self):
        self._save_settings()
        self._running = self._temp_feeding = False
        if hasattr(self, "_poll_job"):
            self.after_cancel(self._poll_job)
        with self._lock:
            dev = self._device
        if dev:
            try:
                dev.disconnect()
            except Exception:
                pass
        if hasattr(self, "_tray"):
            self._tray.stop()
        self.destroy()


if __name__ == "__main__":
    app = App()
    app.mainloop()
