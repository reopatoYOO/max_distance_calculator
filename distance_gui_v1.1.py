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

def calculate_and_display():
    """
    Calculates the maximum distance and displays the result, including a 2D or 3D scatter plot.
    Handles potential errors from the calculation function and uses axis labels if provided.
    """
    file_path = file_path_entry.get()
    dimension_mode = dimension_var.get()
    result = calculate_max_distance(file_path, dimension_mode)

    if isinstance(result[0], str):
        result_text.delete(1.0, tk.END)
        result_text.insert(tk.END, result[0])
        clear_plot()
        point1_label.config(text="")
        point2_label.config(text="")
        return

    max_distance, point1, point2, points, axis_labels = result

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
    clear_plot()
    
    if dimension_mode == "2D":
        x_values = [p[0] for p in points]
        y_values = [p[1] for p in points]
        
        ax.scatter(x_values, y_values, s=50)
        ax.scatter(point1[0], point1[1], color='red', s=100, label='Point 1')
        ax.scatter(point2[0], point2[1], color='blue', s=100, label='Point 2')
        
        if axis_labels:
            ax.set_xlabel(axis_labels[0])
            ax.set_ylabel(axis_labels[1])
        else:
            ax.set_xlabel('X')
            ax.set_ylabel('Y')
        
        ax.set_title('2D Scatter Plot of Data Points')
    else:  # 3D
        l_values = [p[0] for p in points]
        a_values = [p[1] for p in points]
        b_values = [p[2] for p in points]

        ax.scatter(a_values, b_values, l_values)
        ax.scatter(point1[1], point1[2], point1[0], color='red', s=100, label=f'Point 1')
        ax.scatter(point2[1], point2[2], point2[0], color='blue', s=100, label=f'Point 2')

        if axis_labels:
            ax.set_xlabel(axis_labels[1])
            ax.set_ylabel(axis_labels[2])
            ax.set_zlabel(axis_labels[0])
        else:
            ax.set_xlabel('a*')
            ax.set_ylabel('b*')
            ax.set_zlabel('L*')

        ax.set_title('3D Scatter Plot of Data Points')
    
    ax.legend()
    canvas.draw()

def clear_plot():
    """Clears the current plot."""
    ax.cla()
    canvas.draw()

# GUI setup
window = tk.Tk()
window.title("Max Distance Calculator")

# Dimension selection frame
dimension_frame = tk.Frame(window)
dimension_frame.pack(pady=10, padx=10, fill=tk.X)

dimension_label = tk.Label(dimension_frame, text="Select Dimension:")
dimension_label.pack(side=tk.LEFT, padx=5)

dimension_var = tk.StringVar(value="3D")
radio_2d = tk.Radiobutton(dimension_frame, text="2D", variable=dimension_var, value="2D")
radio_2d.pack(side=tk.LEFT, padx=5)

radio_3d = tk.Radiobutton(dimension_frame, text="3D", variable=dimension_var, value="3D")
radio_3d.pack(side=tk.LEFT, padx=5)

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

# 3D Plot area
fig = plt.Figure(figsize=(6, 4), dpi=100)
ax = fig.add_subplot(111, projection='3d')
canvas = FigureCanvasTkAgg(fig, master=window)
canvas_widget = canvas.get_tk_widget()
canvas_widget.pack(pady=10, padx=10, fill=tk.BOTH, expand=True)

# Copyright label
copyright_label = tk.Label(window, text="Made by Peter Yoo", anchor=tk.E)
copyright_label.pack(side=tk.BOTTOM, fill=tk.X, padx=10)

window.mainloop()