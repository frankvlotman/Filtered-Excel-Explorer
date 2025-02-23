import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox, ttk, simpledialog
from PIL import Image

# Define the path for the blank icon
icon_path = 'C:\\Users\\Frank\\Desktop\\blank.ico'

# Create a blank (transparent) ICO file if it doesn't exist
def create_blank_ico(path):
    size = (16, 16)  # Size of the icon
    image = Image.new("RGBA", size, (255, 255, 255, 0))  # Transparent image
    image.save(path, format="ICO")

# Create the blank ICO file
create_blank_ico(icon_path)

# Function to load the Excel file while preserving the displayed values exactly
def load_and_process_file():
    file_path = filedialog.askopenfilename(
        title="Select Excel file",
        filetypes=[("Excel Files", "*.xlsx *.xls")]
    )
    if not file_path:
        messagebox.showinfo("Error", "No file selected!")
        return None
    try:
        # Read the file without forcing a type conversion so that we get the raw objects.
        # Also disable default NA conversion to keep empty cells as empty strings.
        df = pd.read_excel(file_path, dtype=object, keep_default_na=False)
        # Convert any datetime values to strings in the format dd/mm/yyyy.
        for col in df.columns:
            df[col] = df[col].apply(lambda x: x.strftime("%d/%m/%Y") if hasattr(x, "strftime") else x)
        df = df.fillna("")
        return df
    except Exception as e:
        messagebox.showerror("Error", f"An error occurred while processing the file: {e}")
        return None

# Function to auto-adjust the column width based on content
def auto_resize_columns(tree, df):
    for col in df.columns:
        tree.heading(col, text=col, anchor="center")
        max_content_length = max(len(str(item)) for item in df[col].tolist() + [col])
        column_width = min(max_content_length * 10, 150)
        tree.column(col, width=column_width, anchor="center", stretch=False)

# Function to update the TreeView with DataFrame data
def update_treeview(tree, df):
    tree.delete(*tree.get_children())
    if df.empty:
        return
    tree["columns"] = list(df.columns)
    tree["show"] = "headings"
    auto_resize_columns(tree, df)
    for index, row in df.iterrows():
        row_values = [row[col] for col in df.columns]
        tag = 'grey' if index % 2 == 0 else 'white'
        tree.insert("", "end", values=row_values, tags=(tag,))
    tree.tag_configure('grey', background='#f0f0f0')
    tree.tag_configure('white', background='#ffffff')
    tree_frame.update_idletasks()

# Function to copy cell value to clipboard on right-click
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

# Function to load file and display the data in the TreeView
def load_file():
    global df
    df = load_and_process_file()
    if df is not None:
        update_treeview(tree, df)

# Function to filter rows that contain the entered text (case-insensitive)
# and then save the original data along with the filtered rows to a new Excel file.
def filter_and_save():
    global df
    if df is None:
        messagebox.showwarning("Warning", "No data loaded!")
        return
    filter_text = filter_entry.get()
    if not filter_text:
        messagebox.showwarning("Warning", "Please enter text to filter!")
        return
    # Filter rows: check if any cell in the row contains the filter text.
    df_filtered = df[df.apply(lambda row: row.astype(str).str.contains(filter_text, case=False, na=False).any(), axis=1)]
    if df_filtered.empty:
        messagebox.showinfo("Info", f"No rows found containing '{filter_text}'.")
        return
    save_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
    if save_path:
        try:
            with pd.ExcelWriter(save_path, engine='openpyxl') as writer:
                df.to_excel(writer, sheet_name='Original', index=False)
                df_filtered.to_excel(writer, sheet_name='Filtered', index=False)
            messagebox.showinfo("Success", f"File saved successfully at {save_path}")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to save file: {e}")

# Function to filter rows and show them in a separate window
def filter_and_show():
    global df
    if df is None:
        messagebox.showwarning("Warning", "No data loaded!")
        return
    filter_text = filter_entry.get()
    if not filter_text:
        messagebox.showwarning("Warning", "Please enter text to filter!")
        return
    # Filter rows: check if any cell in the row contains the filter text.
    df_filtered = df[df.apply(lambda row: row.astype(str).str.contains(filter_text, case=False, na=False).any(), axis=1)]
    if df_filtered.empty:
        messagebox.showinfo("Info", f"No rows found containing '{filter_text}'.")
        return
    
    # Create a new window to display the filtered results
    show_window = tk.Toplevel(root)
    show_window.title("Filtered Results")
    show_window.iconbitmap(icon_path)  # Set the same blank icon for this window
    show_window.geometry("800x600")
    
    # Create a frame for the TreeView and scrollbars in the new window
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
    
    # Setup the TreeView with filtered data
    tree_filtered["columns"] = list(df_filtered.columns)
    tree_filtered["show"] = "headings"
    auto_resize_columns(tree_filtered, df_filtered)
    
    for index, row in df_filtered.iterrows():
        row_values = [row[col] for col in df_filtered.columns]
        tag = 'grey' if index % 2 == 0 else 'white'
        tree_filtered.insert("", "end", values=row_values, tags=(tag,))
    
    tree_filtered.tag_configure('grey', background='#f0f0f0')
    tree_filtered.tag_configure('white', background='#ffffff')

# Function to filter rows, prompt for custom column selection, and show results in a separate window.
def filter_and_custom_show():
    global df
    if df is None:
        messagebox.showwarning("Warning", "No data loaded!")
        return
    # Apply the filter text as in the other functions.
    filter_text = filter_entry.get()
    if not filter_text:
        messagebox.showwarning("Warning", "Please enter text to filter!")
        return
    df_filtered = df[df.apply(lambda row: row.astype(str).str.contains(filter_text, case=False, na=False).any(), axis=1)]
    if df_filtered.empty:
        messagebox.showinfo("Info", f"No rows found containing '{filter_text}'.")
        return

    # Prompt the user to input custom column numbers (1-indexed, e.g., "4,2,7,1").
    col_input = simpledialog.askstring("Custom Columns", "Enter column numbers (1-indexed) separated by commas:")
    if not col_input:
        messagebox.showwarning("Warning", "No column input provided!")
        return

    try:
        # Parse input and convert to 0-indexed integers.
        col_indices = [int(num.strip()) - 1 for num in col_input.split(',')]
    except ValueError:
        messagebox.showerror("Error", "Invalid column numbers provided. Please enter comma-separated numbers.")
        return

    # Validate indices are within the DataFrame's column range.
    if any(i < 0 or i >= len(df.columns) for i in col_indices):
        messagebox.showerror("Error", "One or more column numbers are out of range.")
        return

    # Create a new DataFrame with only the selected columns in the user-specified order.
    # Note: df.columns is an Index; convert to a list to allow indexing.
    cols = list(df.columns)
    selected_cols = [cols[i] for i in col_indices]
    df_custom = df_filtered[selected_cols]

    # Create a new window to display the custom results
    show_window = tk.Toplevel(root)
    show_window.title("Custom Filtered Results")
    show_window.iconbitmap(icon_path)  # Set the same blank icon for this window
    show_window.geometry("800x600")

    # Create a frame for the TreeView and scrollbars in the new window
    tree_frame_custom = tk.Frame(show_window)
    tree_frame_custom.pack(expand=True, fill=tk.BOTH)

    tree_custom = ttk.Treeview(tree_frame_custom)
    tree_custom.grid(row=0, column=0, sticky='nsew')

    x_scroll_custom = ttk.Scrollbar(tree_frame_custom, orient=tk.HORIZONTAL, command=tree_custom.xview)
    x_scroll_custom.grid(row=1, column=0, sticky='ew')
    y_scroll_custom = ttk.Scrollbar(tree_frame_custom, orient=tk.VERTICAL, command=tree_custom.yview)
    y_scroll_custom.grid(row=0, column=1, sticky='ns')

    tree_custom.configure(yscrollcommand=y_scroll_custom.set, xscrollcommand=x_scroll_custom.set)
    tree_frame_custom.rowconfigure(0, weight=1)
    tree_frame_custom.columnconfigure(0, weight=1)

    # Setup the TreeView with custom data
    tree_custom["columns"] = selected_cols
    tree_custom["show"] = "headings"
    auto_resize_columns(tree_custom, df_custom)

    for index, row in df_custom.iterrows():
        row_values = [row[col] for col in selected_cols]
        tag = 'grey' if index % 2 == 0 else 'white'
        tree_custom.insert("", "end", values=row_values, tags=(tag,))

    tree_custom.tag_configure('grey', background='#f0f0f0')
    tree_custom.tag_configure('white', background='#ffffff')

# GUI setup
root = tk.Tk()
root.title("Filtered Excel Explorer")
root.iconbitmap(icon_path)

style = ttk.Style()
style.theme_use('clam')
style.configure('Custom.TButton', background='#d0e8f1', foreground='black')
style.map('Custom.TButton', background=[('active', '#87CEFA')], foreground=[('active', 'black')])

df = None

# Create a frame for the TreeView and scrollbars
tree_frame = tk.Frame(root)
tree_frame.pack(expand=True, fill=tk.BOTH)

tree = ttk.Treeview(tree_frame)
tree.grid(row=0, column=0, sticky='nsew')
tree.bind("<Button-3>", copy_cell_value)

x_scroll = ttk.Scrollbar(tree_frame, orient=tk.HORIZONTAL, command=tree.xview)
x_scroll.grid(row=1, column=0, sticky='ew')
y_scroll = ttk.Scrollbar(tree_frame, orient=tk.VERTICAL, command=tree.yview)
y_scroll.grid(row=0, column=1, sticky='ns')

tree.configure(yscrollcommand=y_scroll.set, xscrollcommand=x_scroll.set)
tree_frame.rowconfigure(0, weight=1)
tree_frame.columnconfigure(0, weight=1)

# Frame for buttons and filter input
button_frame = tk.Frame(root)
button_frame.pack(pady=10)

load_button = ttk.Button(button_frame, text="Load Excel File", command=load_file, style='Custom.TButton')
load_button.pack(side=tk.LEFT, padx=10)

filter_label = tk.Label(button_frame, text="Filter Text:")
filter_label.pack(side=tk.LEFT, padx=5)

filter_entry = tk.Entry(button_frame)
filter_entry.pack(side=tk.LEFT, padx=5)

filter_save_button = ttk.Button(button_frame, text="Filter & Save", command=filter_and_save, style='Custom.TButton')
filter_save_button.pack(side=tk.LEFT, padx=10)

filter_show_button = ttk.Button(button_frame, text="Filter & Show", command=filter_and_show, style='Custom.TButton')
filter_show_button.pack(side=tk.LEFT, padx=10)

# New "Filter & Custom Show" button
custom_show_button = ttk.Button(button_frame, text="Filter & Custom Show", command=filter_and_custom_show, style='Custom.TButton')
custom_show_button.pack(side=tk.LEFT, padx=10)

root.geometry("800x600")
root.mainloop()
