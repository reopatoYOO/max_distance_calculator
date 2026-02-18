import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
import math
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg

# This program is Made by Chuljoon Yoo. Do not distribute the program without permission


def calculate_max_distance(file_path, dimension_mode="3D"):
    """
    Calculates the maximum Euclidean distance between any two points in an Excel file.
    Supports both 2D (x, y) and 3D (L*, a*, b*) coordinates.

    Args:
        file_path (str): The path to the Excel file.
        dimension_mode (str): Either "2D" or "3D" to specify the dimension of data.

    Returns:
        tuple: A tuple containing:
            - float: The maximum distance, rounded to 2 decimal places.
            - tuple: The coordinates of the first point with the maximum distance.
            - tuple: The coordinates of the second point with the maximum distance.
            - list: A list of all the points read from the file.
            - list or None: A list of axis labels from the first row if they are strings, otherwise None.
        str: An error message if any issue occurs.
    """
    try:
        df = pd.read_excel(file_path, header=None)  # Read without assuming a header

        expected_cols = 2 if dimension_mode == "2D" else 3
        if df.shape[1] != expected_cols:
            messagebox.showerror(title="Error", message=f"Error: Excel file must have {expected_cols} columns for {dimension_mode}.")
            return f"Error: Excel file must have {expected_cols} columns for {dimension_mode}.", None, None, None, None

        axis_labels = None
        start_row = 0
        first_row = df.iloc[0].tolist()
        if all(isinstance(item, str) for item in first_row):
            axis_labels = first_row
            start_row = 1

        points_df = df.iloc[start_row:]
        if points_df.empty or points_df.shape[0] < 2:
            messagebox.showerror(title="Error", message="Need at least two data rows to calculate distance.")
            return "Error: Need at least two data rows to calculate distance.", None, None, None, axis_labels

        try:
            points = points_df.astype(float).values.tolist()
        except ValueError:
            messagebox.showerror(title="Error", message="Error: Data rows must contain numeric values.")
            return "Error: Data rows must contain numeric values.", None, None, None, axis_labels

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

        return round(math.sqrt(max_distance_sq), 2), max_point1, max_point2, points, axis_labels

    except FileNotFoundError:
        return f"Error: File '{file_path}' not found.", None, None, None, None
    except ModuleNotFoundError:
        return "Error: Missing optional dependency 'openpyxl'. Use 'pip install openpyxl' or 'conda install openpyxl'.", None, None, None, None
    except Exception as e:
        return f"An unexpected error occurred: {e}", None, None, None, None

def browse_file():
    """Opens a file dialog to select an Excel file and updates the file path entry."""
    file_path = filedialog.askopenfilename(
        title="Select an Excel file",
        filetypes=[("Excel files", "*.xlsx")]
    )
    if file_path:
        file_path_entry.delete(0, tk.END)
        file_path_entry.insert(0, file_path)

# Global variables to store current plot data
current_plot_data = {
    'points': None,
    'point1': None,
    'point2': None,
    'axis_labels': None,
    'max_distance': None,
    'dimension_mode': None
}

def reset_plot():
    """Resets axis scale entries and redraws the plot with autoscale."""
    global fig, ax, canvas, canvas_widget
    
    # Clear all scale entries
    for entry in [x_min_entry, x_max_entry, y_min_entry, y_max_entry, z_min_entry, z_max_entry]:
        entry.delete(0, tk.END)
    
    if current_plot_data['points'] is None:
        return
    
    dimension_mode = dimension_var.get()
    
    try:
        fig, ax, canvas_widget, canvas = recreate_plot(dimension_mode)
        plot_points(current_plot_data['points'],
                   current_plot_data['point1'],
                   current_plot_data['point2'],
                   current_plot_data['axis_labels'],
                   dimension_mode)
    except Exception as e:
        messagebox.showerror(title="Error", message=f"Error resetting plot: {e}")

def refresh_plot():
    """Refreshes the plot with current axis scale values."""
    global fig, ax, canvas, canvas_widget
    
    if current_plot_data['points'] is None:
        return
    
    dimension_mode = dimension_var.get()
    
    try:
        fig, ax, canvas_widget, canvas = recreate_plot(dimension_mode)
        plot_points(current_plot_data['points'],
                   current_plot_data['point1'],
                   current_plot_data['point2'],
                   current_plot_data['axis_labels'],
                   dimension_mode)
    except Exception as e:
        messagebox.showerror(title="Error", message=f"Error refreshing plot: {e}")

def calculate_and_display():
    """
    Calculates the maximum distance and displays the result, including a 2D or 3D scatter plot.
    Handles potential errors from the calculation function and uses axis labels if provided.
    """
    global fig, ax, canvas, canvas_widget, current_plot_data
    
    file_path = file_path_entry.get()
    if not file_path:
        messagebox.showerror(title="Error", message="Please select an Excel file first.")
        return
    
    dimension_mode = dimension_var.get()
    
    # Clear scale entries for fresh autoscale
    for entry in [x_min_entry, x_max_entry, y_min_entry, y_max_entry, z_min_entry, z_max_entry]:
        entry.delete(0, tk.END)
    
    # Recreate plot for the selected dimension
    try:
        fig, ax, canvas_widget, canvas = recreate_plot(dimension_mode)
    except Exception as e:
        messagebox.showerror(title="Error", message=f"Error creating plot: {e}")
        return
    
    result = calculate_max_distance(file_path, dimension_mode)

    if isinstance(result[0], str):
        result_text.delete(1.0, tk.END)
        result_text.insert(tk.END, result[0])
        point1_label.config(text="")
        point2_label.config(text="")
        return

    max_distance, point1, point2, points, axis_labels = result

    # Store the current plot data
    current_plot_data['points'] = points
    current_plot_data['point1'] = point1
    current_plot_data['point2'] = point2
    current_plot_data['axis_labels'] = axis_labels
    current_plot_data['max_distance'] = max_distance
    current_plot_data['dimension_mode'] = dimension_mode

    result_text.delete(1.0, tk.END)
    result_text.insert(tk.END, f"The maximum distance is: {max_distance:.2f}\n")

    if dimension_mode == "2D":
        point1_label.config(text=f"Point 1: ({point1[0]:.2f}, {point1[1]:.2f})")
        point2_label.config(text=f"Point 2: ({point2[0]:.2f}, {point2[1]:.2f})")
    else:  # 3D
        point1_label.config(text=f"Point 1: ({point1[0]:.2f}, {point1[1]:.2f}, {point1[2]:.2f})")
        point2_label.config(text=f"Point 2: ({point2[0]:.2f}, {point2[1]:.2f}, {point2[2]:.2f})")

    plot_points(points, point1, point2, axis_labels, dimension_mode)

def plot_points(points, point1, point2, axis_labels=None, dimension_mode="3D"):
    """Generates and displays a 2D or 3D scatter plot of the points, highlighting the two furthest points."""
    
    if dimension_mode == "2D":
        x_values = [p[0] for p in points]
        y_values = [p[1] for p in points]
        
        ax.scatter(x_values, y_values, s=50)
        ax.scatter(point1[0], point1[1], color='red', s=100, label='Point 1')
        ax.scatter(point2[0], point2[1], color='blue', s=100, label='Point 2')
        
        ax.set_xlabel(x_name_entry.get())
        ax.set_ylabel(y_name_entry.get())
        
        ax.set_title('2D Scatter Plot of Data Points')
        
        # Apply axis limits if provided
        try:
            if x_min_entry.get():
                x_min = float(x_min_entry.get())
                ax.set_xlim(left=x_min)
        except ValueError:
            pass
        
        try:
            if x_max_entry.get():
                x_max = float(x_max_entry.get())
                ax.set_xlim(right=x_max)
        except ValueError:
            pass
        
        try:
            if y_min_entry.get():
                y_min = float(y_min_entry.get())
                ax.set_ylim(bottom=y_min)
        except ValueError:
            pass
        
        try:
            if y_max_entry.get():
                y_max = float(y_max_entry.get())
                ax.set_ylim(top=y_max)
        except ValueError:
            pass
            
    else:  # 3D
        l_values = [p[0] for p in points]
        a_values = [p[1] for p in points]
        b_values = [p[2] for p in points]

        ax.scatter(a_values, b_values, l_values)
        ax.scatter(point1[1], point1[2], point1[0], color='red', s=100, label=f'Point 1')
        ax.scatter(point2[1], point2[2], point2[0], color='blue', s=100, label=f'Point 2')

        ax.set_xlabel(x_name_entry.get())
        ax.set_ylabel(y_name_entry.get())
        ax.set_zlabel(z_name_entry.get())

        ax.set_title('3D Scatter Plot of Data Points')
        
        # Apply axis limits if provided for 3D
        try:
            if x_min_entry.get():
                x_min = float(x_min_entry.get())
                ax.set_xlim(left=x_min)
        except ValueError:
            pass
        
        try:
            if x_max_entry.get():
                x_max = float(x_max_entry.get())
                ax.set_xlim(right=x_max)
        except ValueError:
            pass
        
        try:
            if y_min_entry.get():
                y_min = float(y_min_entry.get())
                ax.set_ylim(bottom=y_min)
        except ValueError:
            pass
        
        try:
            if y_max_entry.get():
                y_max = float(y_max_entry.get())
                ax.set_ylim(top=y_max)
        except ValueError:
            pass
        
        try:
            if z_min_entry.get():
                z_min = float(z_min_entry.get())
                ax.set_zlim(bottom=z_min)
        except ValueError:
            pass
        
        try:
            if z_max_entry.get():
                z_max = float(z_max_entry.get())
                ax.set_zlim(top=z_max)
        except ValueError:
            pass
    
    ax.legend()
    canvas.draw()
    
    # Update scale entries with current axis limits
    update_scale_entries(dimension_mode)

def update_scale_entries(dimension_mode="3D"):
    """Updates the scale entry fields with the current axis limits."""
    xlim = ax.get_xlim()
    ylim = ax.get_ylim()
    
    # Only update empty entries (don't overwrite user input)
    for entry, val in [(x_min_entry, xlim[0]), (x_max_entry, xlim[1]),
                       (y_min_entry, ylim[0]), (y_max_entry, ylim[1])]:
        if not entry.get():
            entry.insert(0, f"{val:.2f}")
    
    if dimension_mode == "3D":
        zlim = ax.get_zlim()
        for entry, val in [(z_min_entry, zlim[0]), (z_max_entry, zlim[1])]:
            if not entry.get():
                entry.insert(0, f"{val:.2f}")

def clear_plot():
    """Clears the current plot."""
    ax.cla()
    canvas.draw()

def recreate_plot(dimension_mode="3D"):
    """Recreates the plot with appropriate 2D or 3D projection."""
    global ax, fig, canvas, canvas_widget
    
    # Destroy old canvas widget completely
    if canvas_widget is not None:
        canvas_widget.destroy()
    
    # Close old figure
    try:
        plt.close(fig)
    except:
        pass
    
    # Create new figure
    fig = plt.Figure(figsize=(6, 4), dpi=100)
    
    # Add appropriate subplot
    if dimension_mode == "2D":
        ax = fig.add_subplot(111)  # 2D subplot
    else:
        ax = fig.add_subplot(111, projection='3d')  # 3D subplot
    
    # Create new canvas inside the plot_frame
    canvas = FigureCanvasTkAgg(fig, master=plot_frame)
    canvas_widget = canvas.get_tk_widget()
    canvas_widget.pack(fill=tk.BOTH, expand=True)
    
    return fig, ax, canvas_widget, canvas

# GUI setup
window = tk.Tk()
window.title("Max Distance Calculator")

# Dimension selection frame
dimension_frame = tk.Frame(window)
dimension_frame.pack(pady=10, padx=10, fill=tk.X)

dimension_label = tk.Label(dimension_frame, text="Select Dimension:")
dimension_label.pack(side=tk.LEFT, padx=5)

dimension_var = tk.StringVar(value="2D")
radio_2d = tk.Radiobutton(dimension_frame, text="2D", variable=dimension_var, value="2D")
radio_2d.pack(side=tk.LEFT, padx=5)

radio_3d = tk.Radiobutton(dimension_frame, text="3D", variable=dimension_var, value="3D")
radio_3d.pack(side=tk.LEFT, padx=5)

def on_dimension_change(*args):
    """Update visibility of Z-axis controls and redraw plot based on dimension selection."""
    global fig, ax, canvas, canvas_widget
    
    dimension_mode = dimension_var.get()
    
    if dimension_mode == "2D":
        z_frame.pack_forget()
    else:
        z_frame.pack(fill=tk.X, pady=5, before=x_frame)
    
    # Always recreate the plot to reflect dimension change
    try:
        fig, ax, canvas_widget, canvas = recreate_plot(dimension_mode)
    except Exception as e:
        print(f"Error updating plot: {e}")
        return
    
    # If there's already plot data, redraw with new dimension
    if current_plot_data['points'] is not None:
        try:
            plot_points(current_plot_data['points'], 
                       current_plot_data['point1'], 
                       current_plot_data['point2'], 
                       current_plot_data['axis_labels'], 
                       dimension_mode)
        except Exception as e:
            print(f"Error drawing points: {e}")

dimension_var.trace_add("write", on_dimension_change)

# File selection frame
file_frame = tk.Frame(window)
file_frame.pack(pady=10, padx=10, fill=tk.X)

file_label = tk.Label(file_frame, text="Excel File:")
file_label.pack(side=tk.LEFT)

file_path_entry = tk.Entry(file_frame, width=50)
file_path_entry.pack(side=tk.LEFT, expand=True, fill=tk.X)

browse_button = tk.Button(file_frame, text="Browse", command=browse_file)
browse_button.pack(side=tk.LEFT, padx=5)

# Calculate button
calculate_button = tk.Button(window, text="Calculate", command=calculate_and_display, padx=10, pady=5)
calculate_button.pack(pady=10)

# Axis scale adjustment frame
axis_frame = tk.Frame(window)
axis_frame.pack(pady=10, padx=10, fill=tk.X)

axis_title_frame = tk.Frame(axis_frame)
axis_title_frame.pack(fill=tk.X)

axis_label = tk.Label(axis_title_frame, text="Axis Scale (Leave blank for auto):", font=("Arial", 9, "bold"))
axis_label.pack(side=tk.LEFT)

refresh_button = tk.Button(axis_title_frame, text="↩ Reset", command=reset_plot, padx=5)
refresh_button.pack(side=tk.LEFT, padx=10)

# Z-axis frame (for 3D) - displayed first in order Z, X, Y
z_frame = tk.Frame(axis_frame)
# Initially hidden since default is 2D

tk.Label(z_frame, text="Z axis:").pack(side=tk.LEFT, padx=5)
z_name_entry = tk.Entry(z_frame, width=6)
z_name_entry.insert(0, "L*")
z_name_entry.pack(side=tk.LEFT, padx=5)
z_name_entry.bind("<Return>", lambda e: refresh_plot())

tk.Label(z_frame, text="Min:").pack(side=tk.LEFT, padx=5)
z_min_entry = tk.Entry(z_frame, width=10)
z_min_entry.pack(side=tk.LEFT, padx=5)
z_min_entry.bind("<Return>", lambda e: refresh_plot())

tk.Label(z_frame, text="Max:").pack(side=tk.LEFT, padx=5)
z_max_entry = tk.Entry(z_frame, width=10)
z_max_entry.pack(side=tk.LEFT, padx=5)
z_max_entry.bind("<Return>", lambda e: refresh_plot())

# X-axis frame
x_frame = tk.Frame(axis_frame)
x_frame.pack(fill=tk.X, pady=5)

tk.Label(x_frame, text="X axis:").pack(side=tk.LEFT, padx=5)
x_name_entry = tk.Entry(x_frame, width=6)
x_name_entry.insert(0, "a*")
x_name_entry.pack(side=tk.LEFT, padx=5)
x_name_entry.bind("<Return>", lambda e: refresh_plot())

tk.Label(x_frame, text="Min:").pack(side=tk.LEFT, padx=5)
x_min_entry = tk.Entry(x_frame, width=10)
x_min_entry.pack(side=tk.LEFT, padx=5)
x_min_entry.bind("<Return>", lambda e: refresh_plot())

tk.Label(x_frame, text="Max:").pack(side=tk.LEFT, padx=5)
x_max_entry = tk.Entry(x_frame, width=10)
x_max_entry.pack(side=tk.LEFT, padx=5)
x_max_entry.bind("<Return>", lambda e: refresh_plot())

# Y-axis frame
y_frame = tk.Frame(axis_frame)
y_frame.pack(fill=tk.X, pady=5)

tk.Label(y_frame, text="Y axis:").pack(side=tk.LEFT, padx=5)
y_name_entry = tk.Entry(y_frame, width=6)
y_name_entry.insert(0, "b*")
y_name_entry.pack(side=tk.LEFT, padx=5)
y_name_entry.bind("<Return>", lambda e: refresh_plot())

tk.Label(y_frame, text="Min:").pack(side=tk.LEFT, padx=5)
y_min_entry = tk.Entry(y_frame, width=10)
y_min_entry.pack(side=tk.LEFT, padx=5)
y_min_entry.bind("<Return>", lambda e: refresh_plot())

tk.Label(y_frame, text="Max:").pack(side=tk.LEFT, padx=5)
y_max_entry = tk.Entry(y_frame, width=10)
y_max_entry.pack(side=tk.LEFT, padx=5)
y_max_entry.bind("<Return>", lambda e: refresh_plot())

# Result display area
result_text = tk.Text(window, height=3, width=60)
result_text.pack(pady=5, padx=10, fill=tk.BOTH, expand=False)

# Point information labels
point_info_frame = tk.Frame(window)
point_info_frame.pack(pady=5, padx=10, fill=tk.X)

point1_label = tk.Label(point_info_frame, text="", anchor=tk.W)
point1_label.pack(fill=tk.X)

point2_label = tk.Label(point_info_frame, text="", anchor=tk.W)
point2_label.pack(fill=tk.X)

# Plot area container
plot_frame = tk.Frame(window)
plot_frame.pack(pady=10, padx=10, fill=tk.BOTH, expand=True)

fig = plt.Figure(figsize=(6, 4), dpi=100)
ax = fig.add_subplot(111)  # 2D subplot as default
canvas = FigureCanvasTkAgg(fig, master=plot_frame)
canvas_widget = canvas.get_tk_widget()
canvas_widget.pack(fill=tk.BOTH, expand=True)

# Copyright label
copyright_label = tk.Label(window, text="Made by Peter Yoo", anchor=tk.E)
copyright_label.pack(side=tk.BOTTOM, fill=tk.X, padx=10)

window.mainloop()