import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
import math
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg

# This program is Made by Chuljoon Yoo. Do not distribute the program without permission


# ─── Core calculation (no UI dependency) ───────────────────────────────

def calculate_max_distance_from_points(points):
    """
    Given a list of numeric points (each a list/tuple of floats),
    return (max_distance, point1, point2).
    """
    if len(points) < 2:
        return None, None, None

    max_distance_sq = 0
    max_point1 = None
    max_point2 = None

    for i in range(len(points)):
        for j in range(i + 1, len(points)):
            diffs = [(p1 - p2) ** 2 for p1, p2 in zip(points[i], points[j])]
            distance_sq = sum(diffs)
            if distance_sq > max_distance_sq:
                max_distance_sq = distance_sq
                max_point1 = tuple(points[i])
                max_point2 = tuple(points[j])

    return round(math.sqrt(max_distance_sq), 2), max_point1, max_point2


def calculate_max_distance_from_file(file_path, dimension_mode="3D"):
    """Read an Excel file and return (max_distance, point1, point2, points, axis_labels)."""
    try:
        df = pd.read_excel(file_path, header=None)

        expected_cols = 2 if dimension_mode == "2D" else 3
        if df.shape[1] != expected_cols:
            return f"Error: Excel file must have {expected_cols} columns for {dimension_mode}.", None, None, None, None

        axis_labels = None
        start_row = 0
        first_row = df.iloc[0].tolist()
        if all(isinstance(item, str) for item in first_row):
            axis_labels = first_row
            start_row = 1

        points_df = df.iloc[start_row:]
        if points_df.empty or points_df.shape[0] < 2:
            return "Error: Need at least two data rows.", None, None, None, axis_labels

        try:
            points = points_df.astype(float).values.tolist()
        except ValueError:
            return "Error: Data rows must contain numeric values.", None, None, None, axis_labels

        dist, p1, p2 = calculate_max_distance_from_points(points)
        return dist, p1, p2, points, axis_labels

    except FileNotFoundError:
        return f"Error: File '{file_path}' not found.", None, None, None, None
    except ModuleNotFoundError:
        return "Error: Missing 'openpyxl'. pip install openpyxl", None, None, None, None
    except Exception as e:
        return f"An unexpected error occurred: {e}", None, None, None, None


# ─── Global state ──────────────────────────────────────────────────────

current_plot_data = {
    'points': None, 'point1': None, 'point2': None,
    'axis_labels': None, 'max_distance': None, 'dimension_mode': None
}


# ─── Plotting helpers ──────────────────────────────────────────────────

def recreate_plot(dimension_mode="3D"):
    global ax, fig, canvas, canvas_widget
    if canvas_widget is not None:
        canvas_widget.destroy()
    try:
        plt.close(fig)
    except:
        pass
    fig = plt.Figure(figsize=(6, 4), dpi=100)
    if dimension_mode == "2D":
        ax = fig.add_subplot(111)
    else:
        ax = fig.add_subplot(111, projection='3d')
    canvas = FigureCanvasTkAgg(fig, master=plot_frame)
    canvas_widget = canvas.get_tk_widget()
    canvas_widget.pack(fill=tk.BOTH, expand=True)
    return fig, ax, canvas_widget, canvas


def plot_points(points, point1, point2, axis_labels=None, dimension_mode="3D"):
    if dimension_mode == "2D":
        xs = [p[0] for p in points]
        ys = [p[1] for p in points]
        ax.scatter(xs, ys, s=2)
        ax.scatter(point1[0], point1[1], color='red', s=20, label='Point 1')
        ax.scatter(point2[0], point2[1], color='blue', s=20, label='Point 2')
        ax.set_xlabel(x_name_entry.get())
        ax.set_ylabel(y_name_entry.get())
        ax.set_title('2D Scatter Plot of Data Points')
        _apply_limits_2d()
    else:
        ls = [p[0] for p in points]
        as_ = [p[1] for p in points]
        bs = [p[2] for p in points]
        ax.scatter(as_, bs, ls, s=5)
        ax.scatter(point1[1], point1[2], point1[0], color='red', s=30, label='Point 1')
        ax.scatter(point2[1], point2[2], point2[0], color='blue', s=30, label='Point 2')
        ax.set_xlabel(x_name_entry.get())
        ax.set_ylabel(y_name_entry.get())
        ax.set_zlabel(z_name_entry.get())
        ax.set_title('3D Scatter Plot of Data Points')
        _apply_limits_3d()
    ax.legend()
    canvas.draw()
    update_scale_entries(dimension_mode)


def _try_set(setter, entry_getter, **kwargs):
    try:
        v = entry_getter()
        if v:
            setter(**{list(kwargs.keys())[0]: float(v)})
    except ValueError:
        pass


def _apply_limits_2d():
    _try_set(lambda left=None: ax.set_xlim(left=left), x_min_entry.get, left=None)
    _try_set(lambda right=None: ax.set_xlim(right=right), x_max_entry.get, right=None)
    _try_set(lambda bottom=None: ax.set_ylim(bottom=bottom), y_min_entry.get, bottom=None)
    _try_set(lambda top=None: ax.set_ylim(top=top), y_max_entry.get, top=None)


def _apply_limits_3d():
    _apply_limits_2d()
    _try_set(lambda bottom=None: ax.set_zlim(bottom=bottom), z_min_entry.get, bottom=None)
    _try_set(lambda top=None: ax.set_zlim(top=top), z_max_entry.get, top=None)


def update_scale_entries(dimension_mode="3D"):
    xlim = ax.get_xlim()
    ylim = ax.get_ylim()
    for entry, val in [(x_min_entry, xlim[0]), (x_max_entry, xlim[1]),
                       (y_min_entry, ylim[0]), (y_max_entry, ylim[1])]:
        if not entry.get():
            entry.insert(0, f"{val:.2f}")
    if dimension_mode == "3D":
        zlim = ax.get_zlim()
        for entry, val in [(z_min_entry, zlim[0]), (z_max_entry, zlim[1])]:
            if not entry.get():
                entry.insert(0, f"{val:.2f}")


# ─── Shared result display ────────────────────────────────────────────

def display_result(max_distance, point1, point2, points, axis_labels, dimension_mode):
    global fig, ax, canvas, canvas_widget, current_plot_data

    current_plot_data.update({
        'points': points, 'point1': point1, 'point2': point2,
        'axis_labels': axis_labels, 'max_distance': max_distance,
        'dimension_mode': dimension_mode
    })

    for entry in [x_min_entry, x_max_entry, y_min_entry, y_max_entry, z_min_entry, z_max_entry]:
        entry.delete(0, tk.END)

    fig, ax, canvas_widget, canvas = recreate_plot(dimension_mode)

    result_text.delete(1.0, tk.END)
    result_text.insert(tk.END, f"The maximum distance is: {max_distance:.2f}\n")

    if dimension_mode == "2D":
        point1_label.config(text=f"Point 1: ({point1[0]:.2f}, {point1[1]:.2f})")
        point2_label.config(text=f"Point 2: ({point2[0]:.2f}, {point2[1]:.2f})")
    else:
        point1_label.config(text=f"Point 1: ({point1[0]:.2f}, {point1[1]:.2f}, {point1[2]:.2f})")
        point2_label.config(text=f"Point 2: ({point2[0]:.2f}, {point2[1]:.2f}, {point2[2]:.2f})")

    plot_points(points, point1, point2, axis_labels, dimension_mode)


# ─── Tab 1: Excel file ────────────────────────────────────────────────

def browse_file():
    fp = filedialog.askopenfilename(title="Select an Excel file", filetypes=[("Excel files", "*.xlsx")])
    if fp:
        file_path_entry.delete(0, tk.END)
        file_path_entry.insert(0, fp)


def calculate_from_excel():
    fp = file_path_entry.get()
    if not fp:
        messagebox.showerror("Error", "Please select an Excel file first.")
        return
    dim = dimension_var.get()
    result = calculate_max_distance_from_file(fp, dim)
    if isinstance(result[0], str):
        result_text.delete(1.0, tk.END)
        result_text.insert(tk.END, result[0])
        point1_label.config(text="")
        point2_label.config(text="")
        return
    max_dist, p1, p2, pts, labels = result
    display_result(max_dist, p1, p2, pts, labels, dim)


# ─── Tab 2: Manual spreadsheet ────────────────────────────────────────

DEFAULT_ROWS = 20


class SpreadsheetFrame(tk.Frame):
    """Excel-like grid for manual data entry with copy/paste support."""

    def __init__(self, master, dimension_var, **kwargs):
        super().__init__(master, **kwargs)
        self.dimension_var = dimension_var
        self.header_entries = []
        self.data_entries = []  # list of row-lists
        self._z_cache = {}       # {row_index: z_value} saved when switching 3D→2D
        self._z_header_cache = ""  # saved Z header name

        # ── Header row (axis names) ──
        self.header_frame = tk.Frame(self)
        self.header_frame.pack(fill=tk.X, pady=(0, 2))

        tk.Label(self.header_frame, text="Row", width=5, relief="raised", bg="#d0d0d0").grid(row=0, column=0, sticky="nsew")
        self._build_headers()

        # ── Scrollable data area ──
        container = tk.Frame(self)
        container.pack(fill=tk.BOTH, expand=True)

        self.data_canvas = tk.Canvas(container, highlightthickness=0, height=150)
        scrollbar = ttk.Scrollbar(container, orient="vertical", command=self.data_canvas.yview)
        self.data_inner = tk.Frame(self.data_canvas)

        self.data_inner.bind("<Configure>", lambda e: self.data_canvas.configure(scrollregion=self.data_canvas.bbox("all")))
        self.data_canvas.create_window((0, 0), window=self.data_inner, anchor="nw")
        self.data_canvas.configure(yscrollcommand=scrollbar.set)

        self.data_canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        # Mouse scroll
        self.data_canvas.bind_all("<MouseWheel>", lambda e: self.data_canvas.yview_scroll(-1 * (e.delta // 120), "units"))

        # ── Buttons ──
        btn_frame = tk.Frame(self)
        btn_frame.pack(fill=tk.X, pady=5)

        tk.Button(btn_frame, text="+ Add Row", command=self._add_row).pack(side=tk.LEFT, padx=5)
        tk.Button(btn_frame, text="- Remove Last Row", command=self._remove_row).pack(side=tk.LEFT, padx=5)
        tk.Button(btn_frame, text="Clear All", command=self._clear_all).pack(side=tk.LEFT, padx=5)

        # ── Initial rows ──
        for _ in range(DEFAULT_ROWS):
            self._add_row()

    def _build_headers(self):
        dim = self.dimension_var.get()
        cols = 3 if dim == "3D" else 2

        # Cache Z header before clearing (3D→2D)
        if len(self.header_entries) == 3 and cols == 2:
            self._z_header_cache = self.header_entries[0].get()

        # Clear old headers
        for e in self.header_entries:
            e.destroy()
        self.header_entries.clear()

        if cols == 3:
            z_name = self._z_header_cache if self._z_header_cache else "L*"
            defaults = [z_name, "a*", "b*"]
        else:
            defaults = ["a*", "b*"]
        for c in range(cols):
            e = tk.Entry(self.header_frame, width=12, justify="center", bg="#ffffc0",
                         font=("Arial", 9, "bold"))
            e.insert(0, defaults[c])
            e.grid(row=0, column=c + 1, sticky="nsew", padx=1)
            self.header_entries.append(e)

    def rebuild_for_dimension(self):
        """Rebuild headers and adjust data columns when dimension changes.
        3D columns: [Z(L*), X(a*), Y(b*)]  →  2D columns: [X(a*), Y(b*)]
        When switching, Z is added/removed while X and Y values are preserved.
        """
        self._build_headers()
        dim = self.dimension_var.get()
        cols = 3 if dim == "3D" else 2

        for ri, row_entries in enumerate(self.data_entries):
            current_cols = len(row_entries)
            if current_cols == cols:
                continue

            # Save current values
            old_values = [e.get() for e in row_entries]

            if current_cols == 3 and cols == 2:
                # 3D→2D: cache Z (col0), keep X (col1) and Y (col2)
                self._z_cache[ri] = old_values[0]
                new_values = [old_values[1], old_values[2]]
            elif current_cols == 2 and cols == 3:
                # 2D→3D: restore cached Z (col0), X→col1, Y→col2
                cached_z = self._z_cache.get(ri, "")
                new_values = [cached_z, old_values[0], old_values[1]]
            else:
                new_values = old_values[:cols]

            # Destroy old entries
            for e in row_entries:
                e.destroy()
            row_entries.clear()

            # Recreate entries with correct mapping
            for c in range(cols):
                e = tk.Entry(self.data_inner, width=12, justify="center")
                e.grid(row=ri, column=c + 1, sticky="nsew", padx=1, pady=1)
                e.bind("<Control-v>", self._on_paste)
                if c < len(new_values) and new_values[c]:
                    e.insert(0, new_values[c])
                row_entries.append(e)

    def _add_row(self):
        dim = self.dimension_var.get()
        cols = 3 if dim == "3D" else 2
        r = len(self.data_entries)
        row_entries = []
        lbl = tk.Label(self.data_inner, text=str(r + 1), width=5, relief="ridge", bg="#f0f0f0")
        lbl.grid(row=r, column=0, sticky="nsew")
        for c in range(cols):
            e = tk.Entry(self.data_inner, width=12, justify="center")
            e.grid(row=r, column=c + 1, sticky="nsew", padx=1, pady=1)
            e.bind("<Control-v>", self._on_paste)
            row_entries.append(e)
        self.data_entries.append(row_entries)

    def _remove_row(self):
        if not self.data_entries:
            return
        row = self.data_entries.pop()
        r = len(self.data_entries)
        # Remove row label
        for w in self.data_inner.grid_slaves(row=r):
            w.destroy()

    def _clear_all(self):
        for row in self.data_entries:
            for e in row:
                e.delete(0, tk.END)
        for e in self.header_entries:
            e.delete(0, tk.END)
        self._z_cache.clear()
        self._z_header_cache = ""
        dim = self.dimension_var.get()
        defaults = ["L*", "a*", "b*"] if dim == "3D" else ["a*", "b*"]
        for i, e in enumerate(self.header_entries):
            e.insert(0, defaults[i])

    def _on_paste(self, event):
        """Handle paste from clipboard (Excel copies as tab-separated lines)."""
        try:
            clipboard = self.winfo_toplevel().clipboard_get()
        except tk.TclError:
            return

        lines = clipboard.strip().split('\n')
        # Determine which cell triggered paste
        widget = event.widget
        start_row = None
        start_col = None
        for ri, row in enumerate(self.data_entries):
            for ci, e in enumerate(row):
                if e is widget:
                    start_row, start_col = ri, ci
                    break
            if start_row is not None:
                break

        if start_row is None:
            return

        dim = self.dimension_var.get()
        cols = 3 if dim == "3D" else 2

        for li, line in enumerate(lines):
            ri = start_row + li
            # Add rows if needed
            while ri >= len(self.data_entries):
                self._add_row()
            values = line.split('\t')
            for vi, val in enumerate(values):
                ci = start_col + vi
                if ci < cols:
                    self.data_entries[ri][ci].delete(0, tk.END)
                    self.data_entries[ri][ci].insert(0, val.strip())

        return "break"  # prevent default paste

    def get_data(self):
        """Return (points, axis_labels) from the grid."""
        axis_labels = [e.get().strip() for e in self.header_entries]
        if not any(axis_labels):
            axis_labels = None

        points = []
        for row in self.data_entries:
            vals = [e.get().strip() for e in row]
            if all(v == "" for v in vals):
                continue
            try:
                points.append([float(v) for v in vals])
            except ValueError:
                return None, axis_labels, "Error: All data cells must contain numeric values."
        return points, axis_labels, None


def calculate_from_manual():
    dim = dimension_var.get()
    points, axis_labels, error = spreadsheet.get_data()
    if error:
        messagebox.showerror("Error", error)
        return
    if points is None or len(points) < 2:
        messagebox.showerror("Error", "Need at least two data rows to calculate distance.")
        return
    expected = 2 if dim == "2D" else 3
    for i, p in enumerate(points):
        if len(p) != expected:
            messagebox.showerror("Error", f"Row {i+1} has {len(p)} values, expected {expected}.")
            return
    max_dist, p1, p2 = calculate_max_distance_from_points(points)
    display_result(max_dist, p1, p2, points, axis_labels, dim)


# ─── Shared actions ───────────────────────────────────────────────────

def reset_plot():
    global fig, ax, canvas, canvas_widget
    for entry in [x_min_entry, x_max_entry, y_min_entry, y_max_entry, z_min_entry, z_max_entry]:
        entry.delete(0, tk.END)
    if current_plot_data['points'] is None:
        return
    dim = dimension_var.get()
    try:
        fig, ax, canvas_widget, canvas = recreate_plot(dim)
        plot_points(current_plot_data['points'], current_plot_data['point1'],
                    current_plot_data['point2'], current_plot_data['axis_labels'], dim)
    except Exception as e:
        messagebox.showerror("Error", f"Error resetting plot: {e}")


def refresh_plot():
    global fig, ax, canvas, canvas_widget
    if current_plot_data['points'] is None:
        return
    dim = dimension_var.get()
    try:
        fig, ax, canvas_widget, canvas = recreate_plot(dim)
        plot_points(current_plot_data['points'], current_plot_data['point1'],
                    current_plot_data['point2'], current_plot_data['axis_labels'], dim)
    except Exception as e:
        messagebox.showerror("Error", f"Error refreshing plot: {e}")


# ═══════════════════════════════════════════════════════════════════════
#  GUI LAYOUT
# ═══════════════════════════════════════════════════════════════════════

window = tk.Tk()
window.title("Max Distance Calculator v1.3")
window.geometry("750x850")
window.minsize(700, 600)

# ── Dimension selection ───────────────────────────────────────────────
dim_frame = tk.Frame(window)
dim_frame.pack(pady=10, padx=10, fill=tk.X)

tk.Label(dim_frame, text="Select Dimension:").pack(side=tk.LEFT, padx=5)
dimension_var = tk.StringVar(value="2D")
tk.Radiobutton(dim_frame, text="2D", variable=dimension_var, value="2D").pack(side=tk.LEFT, padx=5)
tk.Radiobutton(dim_frame, text="3D", variable=dimension_var, value="3D").pack(side=tk.LEFT, padx=5)

# ── Input mode tabs (compact, no content height) ─────────────────────
notebook = ttk.Notebook(window)
notebook.pack(pady=5, padx=10, fill=tk.X)

# --- Tab 1: From Excel ---
tab_excel = ttk.Frame(notebook)
notebook.add(tab_excel, text="  📂 From Excel  ")

file_frame = tk.Frame(tab_excel)
file_frame.pack(pady=10, padx=10, fill=tk.X)

tk.Label(file_frame, text="Excel File:").pack(side=tk.LEFT)
file_path_entry = tk.Entry(file_frame, width=50)
file_path_entry.pack(side=tk.LEFT, expand=True, fill=tk.X, padx=5)
tk.Button(file_frame, text="Browse", command=browse_file).pack(side=tk.LEFT, padx=5)

tk.Button(tab_excel, text="Calculate", command=calculate_from_excel, padx=10, pady=5).pack(pady=10)

# --- Tab 2: Manual Input (just a Calculate button; spreadsheet is on the side) ---
tab_manual = ttk.Frame(notebook)
notebook.add(tab_manual, text="  ✏️ Manual Input  ")

tk.Button(tab_manual, text="Calculate", command=calculate_from_manual, padx=10, pady=5).pack(pady=10)

# ── Main content area ─────────────────────────────────────────────────
main_hbox = tk.Frame(window)
main_hbox.pack(fill=tk.BOTH, expand=True, padx=10)

# -- Left column: axis scale, result, plot --
left_panel = tk.Frame(main_hbox)
left_panel.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

# ── Axis Scale ────────────────────────────────────────────────────────
axis_frame = tk.Frame(left_panel)
axis_frame.pack(pady=5, fill=tk.X)

axis_title = tk.Frame(axis_frame)
axis_title.pack(fill=tk.X)
tk.Label(axis_title, text="Axis Scale (Leave blank for auto):", font=("Arial", 9, "bold")).pack(side=tk.LEFT)
tk.Button(axis_title, text="↩ Reset", command=reset_plot, padx=5).pack(side=tk.LEFT, padx=10)

# Z axis (shown only for 3D, before X)
z_frame = tk.Frame(axis_frame)
tk.Label(z_frame, text="Z axis:").pack(side=tk.LEFT, padx=5)
z_name_entry = tk.Entry(z_frame, width=6); z_name_entry.insert(0, "L*"); z_name_entry.pack(side=tk.LEFT, padx=5)
z_name_entry.bind("<Return>", lambda e: refresh_plot())
tk.Label(z_frame, text="Min:").pack(side=tk.LEFT, padx=5)
z_min_entry = tk.Entry(z_frame, width=10); z_min_entry.pack(side=tk.LEFT, padx=5)
z_min_entry.bind("<Return>", lambda e: refresh_plot())
tk.Label(z_frame, text="Max:").pack(side=tk.LEFT, padx=5)
z_max_entry = tk.Entry(z_frame, width=10); z_max_entry.pack(side=tk.LEFT, padx=5)
z_max_entry.bind("<Return>", lambda e: refresh_plot())

# X axis
x_frame = tk.Frame(axis_frame); x_frame.pack(fill=tk.X, pady=5)
tk.Label(x_frame, text="X axis:").pack(side=tk.LEFT, padx=5)
x_name_entry = tk.Entry(x_frame, width=6); x_name_entry.insert(0, "a*"); x_name_entry.pack(side=tk.LEFT, padx=5)
x_name_entry.bind("<Return>", lambda e: refresh_plot())
tk.Label(x_frame, text="Min:").pack(side=tk.LEFT, padx=5)
x_min_entry = tk.Entry(x_frame, width=10); x_min_entry.pack(side=tk.LEFT, padx=5)
x_min_entry.bind("<Return>", lambda e: refresh_plot())
tk.Label(x_frame, text="Max:").pack(side=tk.LEFT, padx=5)
x_max_entry = tk.Entry(x_frame, width=10); x_max_entry.pack(side=tk.LEFT, padx=5)
x_max_entry.bind("<Return>", lambda e: refresh_plot())

# Y axis
y_frame = tk.Frame(axis_frame); y_frame.pack(fill=tk.X, pady=5)
tk.Label(y_frame, text="Y axis:").pack(side=tk.LEFT, padx=5)
y_name_entry = tk.Entry(y_frame, width=6); y_name_entry.insert(0, "b*"); y_name_entry.pack(side=tk.LEFT, padx=5)
y_name_entry.bind("<Return>", lambda e: refresh_plot())
tk.Label(y_frame, text="Min:").pack(side=tk.LEFT, padx=5)
y_min_entry = tk.Entry(y_frame, width=10); y_min_entry.pack(side=tk.LEFT, padx=5)
y_min_entry.bind("<Return>", lambda e: refresh_plot())
tk.Label(y_frame, text="Max:").pack(side=tk.LEFT, padx=5)
y_max_entry = tk.Entry(y_frame, width=10); y_max_entry.pack(side=tk.LEFT, padx=5)
y_max_entry.bind("<Return>", lambda e: refresh_plot())

# ── Result display ────────────────────────────────────────────────────
result_text = tk.Text(left_panel, height=3, width=60)
result_text.pack(pady=5, fill=tk.X, expand=False)

point_info = tk.Frame(left_panel)
point_info.pack(pady=5, fill=tk.X)
point1_label = tk.Label(point_info, text="", anchor=tk.W); point1_label.pack(fill=tk.X)
point2_label = tk.Label(point_info, text="", anchor=tk.W); point2_label.pack(fill=tk.X)

# ── Plot area ─────────────────────────────────────────────────────────
plot_frame = tk.Frame(left_panel)
plot_frame.pack(pady=5, fill=tk.BOTH, expand=True)

fig = plt.Figure(figsize=(6, 4), dpi=100)
ax = fig.add_subplot(111)
canvas = FigureCanvasTkAgg(fig, master=plot_frame)
canvas_widget = canvas.get_tk_widget()
canvas_widget.pack(fill=tk.BOTH, expand=True)

# -- Right column: spreadsheet (only visible when Manual Input tab is selected) --
# Placed directly in window so it spans the full height of the GUI
right_panel = tk.Frame(window, width=280)

tk.Label(right_panel, text="📋 Data Input", font=("Arial", 11, "bold"),
         anchor=tk.W).pack(fill=tk.X, padx=5, pady=(5, 2))
ttk.Separator(right_panel, orient="horizontal").pack(fill=tk.X, padx=5)

spreadsheet = SpreadsheetFrame(right_panel, dimension_var)
spreadsheet.pack(fill=tk.BOTH, expand=True, pady=5)

# ── Copyright ─────────────────────────────────────────────────────────
copyright_label = tk.Label(window, text="Made by Peter Yoo", anchor=tk.E)
copyright_label.pack(side=tk.BOTTOM, fill=tk.X, padx=10)


# ── Tab change: show/hide right panel ─────────────────────────────────

def on_tab_changed(event):
    tab_id = notebook.index(notebook.select())
    if tab_id == 1:  # Manual Input
        # Pack right_panel at the right side of window, spanning full height
        # Re-pack main_hbox and copyright to ensure correct order
        copyright_label.pack_forget()
        main_hbox.pack_forget()
        main_hbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(10, 5))
        right_panel.pack(side=tk.RIGHT, fill=tk.BOTH, padx=(0, 10))
        copyright_label.pack(side=tk.BOTTOM, fill=tk.X, padx=10)
        window.geometry("1050x850")
    else:  # From Excel
        right_panel.pack_forget()
        main_hbox.pack_forget()
        main_hbox.pack(fill=tk.BOTH, expand=True, padx=10)
        copyright_label.pack(side=tk.BOTTOM, fill=tk.X, padx=10)
        window.geometry("750x850")


notebook.bind("<<NotebookTabChanged>>", on_tab_changed)


# ── Dimension change handler ─────────────────────────────────────────

def on_dimension_change(*args):
    global fig, ax, canvas, canvas_widget
    dim = dimension_var.get()
    if dim == "2D":
        z_frame.pack_forget()
    else:
        z_frame.pack(fill=tk.X, pady=5, before=x_frame)
    # Rebuild spreadsheet columns
    spreadsheet.rebuild_for_dimension()
    # Recreate plot
    try:
        fig, ax, canvas_widget, canvas = recreate_plot(dim)
    except Exception as e:
        print(f"Error updating plot: {e}")
        return
    if current_plot_data['points'] is not None:
        try:
            plot_points(current_plot_data['points'], current_plot_data['point1'],
                        current_plot_data['point2'], current_plot_data['axis_labels'], dim)
        except Exception as e:
            print(f"Error drawing points: {e}")

dimension_var.trace_add("write", on_dimension_change)

window.mainloop()
