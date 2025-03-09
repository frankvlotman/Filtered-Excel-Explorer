import os
import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox, ttk, simpledialog
from PIL import Image

# Define the base desktop path from an environment variable.
# If the environment variable is not set, default to 'C:\Users\Frank\Desktop'
BASE_PATH = os.environ.get("MY_WORK_DESKTOP_PATH", r"C:\Users\Frank\Desktop")

# Define the path for the blank icon using the BASE_PATH
icon_path = os.path.join(BASE_PATH, "blank.ico")

# Create a blank (transparent) ICO file if it doesn't exist
def create_blank_ico(path):
    size = (16, 16)  # Size of the icon
    image = Image.new("RGBA", size, (255, 255, 255, 0))  # Transparent image
    image.save(path, format="ICO")

create_blank_ico(icon_path)

# ------------------------------
# Functions for Data Loading & Display
# ------------------------------

def load_and_process_file():
    file_path = filedialog.askopenfilename(
        title="Select CSV or Excel file",
        initialdir=BASE_PATH,
        filetypes=[("CSV Files", "*.csv"), ("Excel Files", "*.xlsx *.xls")]
    )
    if not file_path:
        messagebox.showinfo("Error", "No file selected!")
        return None

    try:
        try:
            skip_rows = int(skip_rows_entry.get())
        except ValueError:
            skip_rows = 0

        ext = os.path.splitext(file_path)[1].lower()
        if ext == ".csv":
            df = pd.read_csv(file_path, dtype=object, skiprows=skip_rows)
        else:
            df = pd.read_excel(file_path, dtype=object, keep_default_na=False, skiprows=skip_rows)

        for col in df.columns:
            df[col] = df[col].apply(
                lambda x: x.strftime("%d/%m/%Y") if pd.notna(x) and hasattr(x, "strftime") else x
            )
        df = df.fillna("")
        return df
    except Exception as e:
        messagebox.showerror("Error", f"An error occurred while processing the file: {e}")
        return None

def auto_resize_columns(tree, df):
    for col in df.columns:
        tree.heading(col, text=col, anchor="center")
        max_content_length = max(len(str(item)) for item in df[col].tolist() + [col])
        column_width = min(max_content_length * 10, 150)
        tree.column(col, width=column_width, anchor="center", stretch=False)

# For the main table, use a canvas header for column numbers.
def update_numbers_header(columns):
    numbers_canvas.delete("all")
    total_width = 0
    for i, col in enumerate(columns):
        col_width = tree.column(col, option="width")
        if not col_width:
            col_width = 100
        x0 = total_width
        x_center = x0 + col_width / 2
        numbers_canvas.create_text(x_center, 10, text=str(i+1), font=("TkDefaultFont", 8), fill="gray")
        total_width += col_width
    numbers_canvas.config(scrollregion=(0, 0, total_width, 20))

def update_treeview(tree, df):
    tree.delete(*tree.get_children())
    if df.empty:
        return
    tree["columns"] = list(df.columns)
    tree["show"] = "headings"
    auto_resize_columns(tree, df)
    update_numbers_header(list(df.columns))
    for index, row in df.iterrows():
        row_values = [row[col] for col in df.columns]
        tag = 'grey' if index % 2 == 0 else 'white'
        tree.insert("", "end", values=row_values, tags=(tag,))
    tree.tag_configure('grey', background='#f0f0f0')
    tree.tag_configure('white', background='#ffffff')
    main_display_frame.update_idletasks()

def copy_cell_value(event):
    region = tree.identify_region(event.x, event.y)
    if region != "cell":
        return
    row_id = tree.identify_row(event.y)
    column_id = tree.identify_column(event.x)
    if row_id and column_id:
        cell_value = tree.set(row_id, column_id)
        root.clipboard_clear()
        root.clipboard_append(cell_value)

def load_file():
    global df
    df = load_and_process_file()
    if df is not None:
        update_treeview(tree, df)

# ------------------------------
# Filtering Functions
# ------------------------------

def filter_and_save():
    global df
    if df is None:
        messagebox.showwarning("Warning", "No data loaded!")
        return
    filter_text = filter_entry.get()
    if filter_text:
        df_filtered = df[df.apply(lambda row: row.astype(str).str.contains(filter_text, case=False, na=False).any(), axis=1)]
    else:
        df_filtered = df.copy()
        
    if df_filtered.empty:
        messagebox.showinfo("Info", f"No rows found containing '{filter_text}'.")
        return
    save_path = filedialog.asksaveasfilename(
        defaultextension=".xlsx",
        initialdir=BASE_PATH,
        filetypes=[("Excel files", "*.xlsx")]
    )
    if save_path:
        try:
            with pd.ExcelWriter(save_path, engine='openpyxl') as writer:
                df.to_excel(writer, sheet_name='Original', index=False)
                df_filtered.to_excel(writer, sheet_name='Filtered', index=False)
            messagebox.showinfo("Success", f"File saved successfully at {save_path}")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to save file: {e}")

def filter_and_show():
    global df
    if df is None:
        messagebox.showwarning("Warning", "No data loaded!")
        return
    filter_text = filter_entry.get()
    if filter_text:
        df_filtered = df[df.apply(lambda row: row.astype(str).str.contains(filter_text, case=False, na=False).any(), axis=1)]
    else:
        df_filtered = df.copy()
        
    if df_filtered.empty:
        messagebox.showinfo("Info", f"No rows found containing '{filter_text}'.")
        return

    show_window = tk.Toplevel(root)
    show_window.title("Filtered Results")
    show_window.iconbitmap(icon_path)
    show_window.geometry("800x600")
    
    tree_frame_filtered = tk.Frame(show_window)
    tree_frame_filtered.pack(expand=True, fill=tk.BOTH)
    
    tree_filtered = ttk.Treeview(tree_frame_filtered)
    tree_filtered.grid(row=0, column=0, sticky='nsew')
    
    x_scroll_filtered = ttk.Scrollbar(tree_frame_filtered, orient=tk.HORIZONTAL, command=tree_filtered.xview)
    x_scroll_filtered.grid(row=1, column=0, sticky='ew')
    y_scroll_filtered = ttk.Scrollbar(tree_frame_filtered, orient=tk.VERTICAL, command=tree_filtered.yview)
    y_scroll_filtered.grid(row=0, column=1, sticky='ns')
    
    tree_filtered.configure(yscrollcommand=y_scroll_filtered.set, xscrollcommand=x_scroll_filtered.set)
    tree_frame_filtered.rowconfigure(0, weight=1)
    tree_frame_filtered.columnconfigure(0, weight=1)
    
    tree_filtered["columns"] = list(df_filtered.columns)
    tree_filtered["show"] = "headings"
    auto_resize_columns(tree_filtered, df_filtered)
    
    for index, row in df_filtered.iterrows():
        row_values = [row[col] for col in df_filtered.columns]
        tag = 'grey' if index % 2 == 0 else 'white'
        tree_filtered.insert("", "end", values=row_values, tags=(tag,))
    
    tree_filtered.tag_configure('grey', background='#f0f0f0')
    tree_filtered.tag_configure('white', background='#ffffff')

def filter_and_custom_show():
    global df
    if df is None:
        messagebox.showwarning("Warning", "No data loaded!")
        return
    # If filter input is blank, use all rows.
    filter_text = filter_entry.get()
    if filter_text:
        df_filtered = df[df.apply(lambda row: row.astype(str).str.contains(filter_text, case=False, na=False).any(), axis=1)]
    else:
        df_filtered = df.copy()

    if df_filtered.empty:
        messagebox.showinfo("Info", f"No rows found.")
        return

    col_input = simpledialog.askstring("Custom Columns", "Enter column numbers (1-indexed) separated by commas:")
    if not col_input:
        messagebox.showwarning("Warning", "No column input provided!")
        return

    try:
        col_indices = [int(num.strip()) - 1 for num in col_input.split(',')]
    except ValueError:
        messagebox.showerror("Error", "Invalid column numbers provided. Please enter comma-separated numbers.")
        return

    if any(i < 0 or i >= len(df.columns) for i in col_indices):
        messagebox.showerror("Error", "One or more column numbers are out of range.")
        return

    cols = list(df.columns)
    selected_cols = [cols[i] for i in col_indices]
    df_custom = df_filtered[selected_cols]

    show_window = tk.Toplevel(root)
    show_window.title("Custom Filtered Results")
    show_window.iconbitmap(icon_path)
    show_window.geometry("800x600")

    display_frame = tk.Frame(show_window)
    display_frame.pack(expand=True, fill=tk.BOTH)

    tree_container = tk.Frame(display_frame)
    tree_container.pack(expand=True, fill=tk.BOTH)

    tree_custom = ttk.Treeview(tree_container)
    tree_custom.grid(row=0, column=0, sticky='nsew')

    x_scroll_custom = ttk.Scrollbar(tree_container, orient=tk.HORIZONTAL, command=tree_custom.xview)
    x_scroll_custom.grid(row=1, column=0, sticky='ew')
    y_scroll_custom = ttk.Scrollbar(tree_container, orient=tk.VERTICAL, command=tree_custom.yview)
    y_scroll_custom.grid(row=0, column=1, sticky='ns')

    tree_custom.configure(yscrollcommand=y_scroll_custom.set, xscrollcommand=x_scroll_custom.set)
    tree_container.rowconfigure(0, weight=1)
    tree_container.columnconfigure(0, weight=1)

    tree_custom["columns"] = selected_cols
    tree_custom["show"] = "headings"
    auto_resize_columns(tree_custom, df_custom)

    for index, row in df_custom.iterrows():
        row_values = [row[col] for col in selected_cols]
        tag = 'grey' if index % 2 == 0 else 'white'
        tree_custom.insert("", "end", values=row_values, tags=(tag,))

    tree_custom.tag_configure('grey', background='#f0f0f0')
    tree_custom.tag_configure('white', background='#ffffff')

    # Download button to save custom filtered DataFrame as an XLSX.
    def download_custom():
        save_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            initialdir=BASE_PATH,
            filetypes=[("Excel files", "*.xlsx")]
        )
        if save_path:
            try:
                with pd.ExcelWriter(save_path, engine='openpyxl') as writer:
                    df_custom.to_excel(writer, sheet_name='Custom Filtered', index=False)
                messagebox.showinfo("Success", f"File saved successfully at {save_path}")
            except Exception as e:
                messagebox.showerror("Error", f"Failed to save file: {e}")

    download_button = ttk.Button(display_frame, text="Download", command=download_custom, style='Custom.TButton')
    download_button.pack(pady=5)

# ------------------------------
# Horizontal Scroll Sync for Main Table
# ------------------------------

def on_tree_xscroll(*args):
    x_scroll.set(*args)
    numbers_canvas.xview_moveto(args[0])

# ------------------------------
# GUI Setup
# ------------------------------
root = tk.Tk()
root.title("Excel Viewer With Filter Input")
root.iconbitmap(icon_path)

style = ttk.Style()
style.theme_use('clam')
style.configure('Custom.TButton', background='#d0e8f1', foreground='black')
style.map('Custom.TButton', background=[('active', '#87CEFA')], foreground=[('active', 'black')])

df = None

# Main display frame holds the numbers header and the main table.
main_display_frame = tk.Frame(root)
main_display_frame.pack(expand=True, fill=tk.BOTH)

numbers_canvas = tk.Canvas(main_display_frame, height=20, bg="lightgrey", highlightthickness=0)
numbers_canvas.pack(fill=tk.X, side=tk.TOP)

tree_frame = tk.Frame(main_display_frame)
tree_frame.pack(expand=True, fill=tk.BOTH)

tree = ttk.Treeview(tree_frame)
tree.grid(row=0, column=0, sticky='nsew')
tree.bind("<Button-3>", copy_cell_value)

x_scroll = ttk.Scrollbar(tree_frame, orient=tk.HORIZONTAL, command=lambda *args: (tree.xview(*args), numbers_canvas.xview(*args)))
x_scroll.grid(row=1, column=0, sticky='ew')
y_scroll = ttk.Scrollbar(tree_frame, orient=tk.VERTICAL, command=tree.yview)
y_scroll.grid(row=0, column=1, sticky='ns')

tree.configure(yscrollcommand=y_scroll.set, xscrollcommand=on_tree_xscroll)
tree_frame.rowconfigure(0, weight=1)
tree_frame.columnconfigure(0, weight=1)

button_frame = tk.Frame(root)
button_frame.pack(pady=10)

load_button = ttk.Button(button_frame, text="Load CSV/XLSX File", command=load_file, style='Custom.TButton')
load_button.pack(side=tk.LEFT, padx=10)

filter_label = tk.Label(button_frame, text="Filter Text:")
filter_label.pack(side=tk.LEFT, padx=5)

filter_entry = tk.Entry(button_frame)
filter_entry.pack(side=tk.LEFT, padx=5)

skip_rows_label = tk.Label(button_frame, text="Skip Rows:")
skip_rows_label.pack(side=tk.LEFT, padx=5)

skip_rows_entry = tk.Entry(button_frame, width=5)
skip_rows_entry.pack(side=tk.LEFT, padx=5)
skip_rows_entry.insert(0, "0")

filter_save_button = ttk.Button(button_frame, text="Filter & Save", command=filter_and_save, style='Custom.TButton')
filter_save_button.pack(side=tk.LEFT, padx=10)

filter_show_button = ttk.Button(button_frame, text="Filter & Show", command=filter_and_show, style='Custom.TButton')
filter_show_button.pack(side=tk.LEFT, padx=10)

custom_show_button = ttk.Button(button_frame, text="Filter & Custom Show", command=filter_and_custom_show, style='Custom.TButton')
custom_show_button.pack(side=tk.LEFT, padx=10)

root.geometry("800x600")
root.mainloop()
