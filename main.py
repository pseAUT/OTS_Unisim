import os
import sys
import time
import tkinter as tk
from tkinter import ttk, messagebox
from collections import deque
from PIL import Image, ImageTk
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from matplotlib.figure import Figure
import numpy as np

# -------------- CONFIG ----------------
# Determine where we are running from:
if getattr(sys, "frozen", False):
    # running in a PyInstaller bundle
    BASE_PATH = sys._MEIPASS
else:
    # running “normally”
    BASE_PATH = os.path.dirname(os.path.abspath(__file__))

SIM_PATH   = os.path.join(BASE_PATH, "assets", "Arbitrary_Unit.usc")
IMAGE_PATH = os.path.join(BASE_PATH, "assets", "compressor_map.png")



INLET_STREAM  = "Inlet"
OUTLET_STREAM = "Outlet"

MAX_POINTS    = 1200           # keep at most this many samples in memory
UPDATE_MS     = 500            # default refresh interval
# --------------------------------------

try:
    import win32com.client
except ImportError:
    win32com = None
    print("pywin32 not installed – running in demo mode.")

# ───────────────── UniSim connector ─────────────────
class UniSimConnector:
    def __init__(self):
        self.connected = False
        if win32com is None:
            return

        try:
            print("Connecting to UniSim …")
            self.app = win32com.client.Dispatch("UniSimDesign.Application")
            self.app.Visible = True
            print("Opening:", SIM_PATH)
            self.sim = self.app.SimulationCases.Open(SIM_PATH)
            self.sim.Visible = 1                # requested
            fs = self.sim.Flowsheet
            self.inlet  = fs.MaterialStreams.Item(INLET_STREAM)
            self.outlet = fs.MaterialStreams.Item(OUTLET_STREAM)
            self.connected = True
            print("UniSim connected.")
        except Exception as e:
            print("UniSim connection failed:", e)

    def inlet_pressure(self):
        return float(self.inlet.PressureValue)

    def outlet_pressure(self):
        return float(self.outlet.PressureValue)

    def sim_time(self):
        return float(self.sim.Solver.Integrator.GetTime())

    def start(self):
        if self.connected:
            self.sim.Solver.Integrator.IsRunning = 1

    def stop(self):
        if self.connected:
            self.sim.Solver.Integrator.IsRunning = 0

# ───────────────── Live-plot window ─────────────────
class LivePlot(tk.Toplevel):
    def __init__(self, title, y_getter, t_getter):
        super().__init__()
        self.title(title)
        self.geometry("900x520")
        self.configure(bg="#F5F5F5")

        self.y_get = y_getter
        self.t_get = t_getter
        self.interval = tk.IntVar(value=UPDATE_MS)
        self.x = deque(maxlen=MAX_POINTS)
        self.y = deque(maxlen=MAX_POINTS)

        fig = Figure(figsize=(5, 4), dpi=100, facecolor="#FFF")
        self.ax = fig.add_subplot(111, facecolor="#FAFAFA")
        self.ax.set_title(title)
        self.ax.set_xlabel("Simulation Time (s)")
        self.ax.set_ylabel("Pressure")

        self.line, = self.ax.plot([], [], color="#007ACC", lw=2, antialiased=False)
        self.canvas = FigureCanvasTkAgg(fig, master=self)
        self.canvas.get_tk_widget().pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        # Interval slider
        ctrl = ttk.Frame(self, padding=8)
        ctrl.pack(fill="x")
        ttk.Label(ctrl, text="Refresh (ms)").pack(side="left", padx=(0, 5))
        ttk.Scale(ctrl, from_=200, to=3000, variable=self.interval,
                  orient="horizontal", length=240).pack(side="left")

        # Start loop
        self._after_id = None
        self.protocol("WM_DELETE_WINDOW", self._on_close)
        self.current_ymin = None
        self.current_ymax = None
        self._update()

    def _update(self):
        try:
            t_val = self.t_get()
            y_val = self.y_get()
        except Exception as e:
            messagebox.showerror("Data error", str(e))
            return

        self.x.append(t_val)
        self.y.append(y_val)

        x_data = np.array(self.x, dtype=np.float32)
        y_data = np.array(self.y, dtype=np.float32)

        # Update x-axis limits
        if len(x_data) > 0:
            xmin = max(0, x_data[-1] - 60)
            self.ax.set_xlim(xmin, xmin + 60)

        # Update y-axis limits conditionally
        if len(y_data) > 0:
            new_ymin = np.min(y_data)
            new_ymax = np.max(y_data)
            if new_ymin == new_ymax:
                new_ymin -= 1e-6
                new_ymax += 1e-6

            update_ylim = False
            if self.current_ymin is None or self.current_ymax is None:
                update_ylim = True
            else:
                if new_ymin < self.current_ymin * 0.99 or new_ymax > self.current_ymax * 1.01:
                    update_ylim = True

            if update_ylim:
                self.ax.set_ylim(new_ymin * 0.98, new_ymax * 1.02)
                self.current_ymin, self.current_ymax = self.ax.get_ylim()

        # Update plot data
        self.line.set_data(x_data, y_data)
        self.canvas.draw_idle()

        self._after_id = self.after(self.interval.get(), self._update)

    def _on_close(self):
        if self._after_id:
            self.after_cancel(self._after_id)
        self.destroy()

# ─────────────────────── Main window ───────────────────────
class MainApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("UniSim Live Pressure Viewer")
        self.geometry("880x520")
        self.configure(bg="#F5F5F5")

        self.us = UniSimConnector()

        # ttk styling
        sty = ttk.Style(self)
        sty.theme_use("clam")
        sty.configure(".", background="#F5F5F5", foreground="#333", font=("Segoe UI", 11))
        sty.configure("Accent.TButton", background="#007ACC", foreground="#FFF", padding=10)
        sty.map("Accent.TButton",
                background=[("active", "#005A9E"), ("pressed", "#004578")])

        # Canvas + background
        self.canvas = tk.Canvas(self, highlightthickness=0)
        self.canvas.pack(fill=tk.BOTH, expand=True)
        self.bg_id = self.canvas.create_image(0, 0, anchor="nw")
        self.bg_img = None
        self.bind("<Configure>", self._layout)

        # Buttons
        self.btn_inlet  = ttk.Button(self.canvas, text="Input Flow Pressure",
                                     style="Accent.TButton",
                                     command=self._plot_inlet)
        self.btn_outlet = ttk.Button(self.canvas, text="Output Flow Pressure",
                                     style="Accent.TButton",
                                     command=self._plot_outlet)
        self.btn_start  = ttk.Button(self.canvas, text="Start Sim",
                                     style="Accent.TButton",
                                     command=self._start_sim)
        self.btn_stop   = ttk.Button(self.canvas, text="Stop Sim",
                                     style="Accent.TButton",
                                     command=self._stop_sim)

        # Create window items
        self.id_inlet  = self.canvas.create_window(0, 0, window=self.btn_inlet)
        self.id_outlet = self.canvas.create_window(0, 0, window=self.btn_outlet)
        self.id_start  = self.canvas.create_window(0, 0, window=self.btn_start)
        self.id_stop   = self.canvas.create_window(0, 0, window=self.btn_stop)

        self._layout()  # Initial placement

    # ---------- callbacks ----------
    def _start_sim(self):
        if self.us.connected:
            self.us.start()
        else:
            messagebox.showinfo("Demo", "UniSim not connected – demo only.")

    def _stop_sim(self):
        if self.us.connected:
            self.us.stop()
        else:
            messagebox.showinfo("Demo", "UniSim not connected – demo only.")

    def _plot_inlet(self):
        
        t_get = self.us.sim_time
        y_get = self.us.inlet_pressure

        LivePlot("Inlet Pressure(kPa)", y_get, t_get)

    def _plot_outlet(self):

        t_get = self.us.sim_time
        y_get = self.us.outlet_pressure

        LivePlot("Outlet Pressure (kPa)", y_get, t_get)

    # ---------- layout ----------
    def _layout(self, event=None):
        w, h = self.canvas.winfo_width(), self.canvas.winfo_height()
        if w < 10 or h < 10:
            return

        # Background
        if os.path.exists(IMAGE_PATH):
            img = Image.open(IMAGE_PATH).resize((w, h), Image.LANCZOS)
        else:
            img = Image.new("RGB", (w, h), "#D0D0D0")
        self.bg_img = ImageTk.PhotoImage(img)
        self.canvas.itemconfig(self.bg_id, image=self.bg_img)
        self.canvas.tag_lower(self.bg_id)

        # Buttons
        self.canvas.coords(self.id_inlet,  0.15*w, 0.25*h)
        self.canvas.coords(self.id_outlet, 0.75*w, 0.65*h)
        self.canvas.coords(self.id_start,  0.05*w, 0.90*h)
        self.canvas.coords(self.id_stop,   0.20*w, 0.90*h)

        for it in (self.id_inlet, self.id_outlet, self.id_start, self.id_stop):
            self.canvas.tag_raise(it)

# ──────────────────────────── main ────────────────────────────
if __name__ == "__main__":
    if sys.platform != "win32":
        messagebox.showerror("Unsupported OS", "This application runs only on Windows.")
        sys.exit(1)

    MainApp().mainloop()