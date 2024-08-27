import pandas as pd
from math import sqrt
import tkinter as tk
from tkinter import filedialog, messagebox

def color_distance(rgb1, rgb2):
    return sqrt(sum((a - b) ** 2 for a, b in zip(rgb1, rgb2)))

def find_closest_color(input_rgb, df):
    min_distance = float('inf')
    closest_row = None
    
    for _, row in df.iterrows():
        rgb = eval(row['RVB'])
        distance = color_distance(input_rgb, rgb)
        if distance < min_distance:
            min_distance = distance
            closest_row = row
            
    return closest_row

def load_file():
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
    if file_path:
        global df
        df = pd.read_excel(file_path, engine='openpyxl')
        messagebox.showinfo("Success", "File loaded successfully!")

def check_color():
    if df is None:
        messagebox.showerror("Error", "Please load an Excel file first.")
        return

    input_html_code = entry_color.get().strip().upper()
    if not input_html_code.startswith("#") or len(input_html_code) != 7:
        messagebox.showerror("Error", "Invalid HTML color code. Use format like #FF5733.")
        return

    exact_match = df[df['CODE HTML'] == input_html_code]
    if not exact_match.empty and exact_match['stock'].values[0] > 0:
        messagebox.showinfo("In Stock", f"Color {input_html_code} is in stock under the name '{exact_match['COULEUR'].values[0]}'.")
    else:
        input_rgb = tuple(int(input_html_code[i:i+2], 16) for i in (1, 3, 5))
        closest_match = find_closest_color(input_rgb, df)
        
        if closest_match is not None:
            messagebox.showinfo("Closest Match", f"The exact color is not in stock.\nThe closest available color is {closest_match['CODE HTML']} "
                                                f"(Name: {closest_match['COULEUR']}) with a stock of {closest_match['stock']}.")
        else:
            messagebox.showinfo("No Match", "No available colors close to the requested one.")

# Create the main window
root = tk.Tk()
root.title("Color Stock Checker")

# Define the layout
frame = tk.Frame(root)
frame.pack(pady=100, padx=100)

btn_load = tk.Button(frame, text="Load Excel File", command=load_file)
btn_load.pack(pady=10, padx=10)  # Add padding around the button

# Create and place the label for the color code entry
label_color = tk.Label(frame, text="Enter HTML Color Code:")
label_color.pack(pady=5, padx=10)  # Add padding around the label

# Create and place the entry widget for color code
entry_color = tk.Entry(frame, width=15)
entry_color.pack(pady=5)  # Add padding above and below the entry widget

# Create and place the "Check Color" button with padding and centering
btn_check = tk.Button(frame, text="Check Color", command=check_color)
btn_check.pack(pady=10)  # Add padding around the button


credit_label = tk.Label(root, text="Credit by: Chaima Belazreg")
credit_label.pack(side="bottom", pady=5)  # Position at the bottom with padding
# Initialize the DataFrame variable
df = None

# Run the application
root.mainloop()
