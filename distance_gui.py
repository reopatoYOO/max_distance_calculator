import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
import math
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg

## the program is built by Chuljoon Yoo.
def calculate_max_distance(file_path):
    try:
        df = pd.read_excel(file_path)
        if df.shape[1] != 3:
            return "Error: Excel file must have 3 columns (x, y, z)."

        points = list(zip(df.iloc[:, 0], df.iloc[:, 1], df.iloc[:, 2]))

        if len(points) < 2:
            return "Error: Need at least two valid points to calculate distance."

        max_distance = 0
        max_point1 = None
        max_point2 = None

        for i in range(len(points)):
            for j in range(i + 1, len(points)):
                distance = math.sqrt(sum((p1 - p2) ** 2 for p1, p2 in zip(points[i], points[j])))
                if distance > max_distance:
                    max_distance = distance
                    max_point1 = points[i]
                    max_point2 = points[j]

        return round(max_distance, 2), max_point1, max_point2, points

    except FileNotFoundError:
        return f"Error: File '{file_path}' not found.", None, None, None
    except ModuleNotFoundError:
        return "Error: Missing optional dependency 'openpyxl'. Use pip or conda to install openpyxl.", None, None, None
    except Exception as e:
        return f"An unexpected error occurred: {e}", None, None, None

def browse_file():
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
    if file_path:
        file_path_entry.delete(0, tk.END)
        file_path_entry.insert(0, file_path)

def calculate_and_display():
    file_path = file_path_entry.get()
    result = calculate_max_distance(file_path)

    if isinstance(result[0], str):
        result_text.delete(1.0, tk.END)
        result_text.insert(tk.END, result[0])
        clear_plot()
        return

    max_distance, point1, point2, points = result

    result_text.delete(1.0, tk.END)
    result_text.insert(tk.END, f"The maximum distance is: {max_distance:.2f}\n\nBetween points:\n Point1: ({point1[0]:.2f}, {point1[1]:.2f}, {point1[2]:.2f})\n Point2: ({point2[0]:.2f}, {point2[1]:.2f}, {point2[2]:.2f})")

    plot_points(points, point1, point2)

def plot_points(points, point1, point2):
    clear_plot()
    a = [p[1] for p in points]
    b = [p[2] for p in points]
    l = [p[0] for p in points]

    ax.scatter(a, b, l)
    ax.scatter(point1[1], point1[2], point1[0], color='red', s=100)
    ax.scatter(point2[1], point2[2], point2[0], color='blue', s=100)
    ax.set_xlabel('a*')
    ax.set_ylabel('b*')
    ax.set_zlabel('L*')

    canvas.draw()

def clear_plot():
    ax.cla()
    canvas.draw()

# GUI setup
window = tk.Tk()
window.title("Max Distance Calculator")

file_frame = tk.Frame(window)
file_frame.pack(pady=10)

file_label = tk.Label(file_frame, text="Excel File:")
file_label.pack(side=tk.LEFT)

file_path_entry = tk.Entry(file_frame, width=50)
file_path_entry.pack(side=tk.LEFT)

browse_button = tk.Button(file_frame, text="Browse", command=browse_file)
browse_button.pack(side=tk.LEFT)

calculate_button = tk.Button(window, text="Calculate", command=calculate_and_display)
calculate_button.pack(pady=10)

result_text = tk.Text(window, height=5, width=60)
result_text.pack(pady=10)

fig = plt.Figure(figsize=(6, 4), dpi=100)
ax = fig.add_subplot(projection='3d')
canvas = FigureCanvasTkAgg(fig, master=window)
canvas.get_tk_widget().pack()

window.mainloop()