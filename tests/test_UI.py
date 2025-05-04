import tkinter as tk
from tkinter import filedialog, ttk, messagebox
import json
import os
import time

def save_last_settings():
    settings = {
        "dwg_folder": dwg_folder_var.get(),
        "pdf_folder": pdf_folder_var.get(),
        "export_dwg": dwg_var.get(),
        "export_pdf": pdf_var.get(),
        "drawings": [drawings_list.item(item, "values") for item in drawings_list.get_children()],
        "step_folder": step_folder_var.get(),
        "parts": [parts_list.item(item, "values") for item in parts_list.get_children()]
    }
    with open("last_settings.json", "w") as file:
        json.dump(settings, file, indent=4)

def select_dwg_folder():
    folder = filedialog.askdirectory()
    if folder:
        dwg_folder_var.set(folder)
        status_bar.config(text="DWG folder selected")

def select_pdf_folder():
    folder = filedialog.askdirectory()
    if folder:
        pdf_folder_var.set(folder)
        status_bar.config(text="PDF folder selected")

def select_step_folder():
    folder = filedialog.askdirectory()
    if folder:
        step_folder_var.set(folder)
        status_bar.config(text="STEP export folder selected")

def select_drawings():
    files = filedialog.askopenfilenames(filetypes=[("SolidWorks Drawings", "*.SLDDRW")])
    for file in files:
        filename = os.path.basename(file)
        drawings_list.insert("", "end", values=(filename, file))
    status_bar.config(text="Drawings selected")

def select_parts():
    files = filedialog.askopenfilenames(filetypes=[("SolidWorks Parts", "*.SLDPRT")])
    for file in files:
        filename = os.path.basename(file)
        parts_list.insert("", "end", values=(filename, file))
    status_bar.config(text="Parts selected")

def delete_selected(table):
    selected_items = table.selection()
    for item in selected_items:
        table.delete(item)
    status_bar.config(text="Selected items deleted")

def export_drawings():
    if not drawings_list.get_children():
        messagebox.showwarning("Export Warning", "No drawings selected for export.")
        return
    
    progress_bar.grid()
    progress_bar["maximum"] = len(drawings_list.get_children())
    status_bar.config(text="Exporting drawings...")
    root.update_idletasks()
    
    for i, item in enumerate(drawings_list.get_children(), start=1):
        print(f"Exporting: {drawings_list.item(item, 'values')[0]}")
        time.sleep(0.5)
        progress_bar["value"] = i
        root.update_idletasks()
    
    progress_bar.grid_remove()
    status_bar.config(text="Export completed successfully")
    messagebox.showinfo("Export Completed", "Drawing export process has finished.")

def export_parts():
    if not parts_list.get_children():
        messagebox.showwarning("Export Warning", "No parts selected for STEP export.")
        return
    
    progress_bar.grid()
    progress_bar["maximum"] = len(parts_list.get_children())
    status_bar.config(text="Exporting parts to STEP...")
    root.update_idletasks()
    
    for i, item in enumerate(parts_list.get_children(), start=1):
        print(f"Exporting: {parts_list.item(item, 'values')[0]}")
        time.sleep(0.5)
        progress_bar["value"] = i
        root.update_idletasks()
    
    progress_bar.grid_remove()
    status_bar.config(text="STEP export completed successfully")
    messagebox.showinfo("Export Completed", "STEP export process has finished.")

def save_settings():
    file_path = filedialog.asksaveasfilename(defaultextension=".json", filetypes=[("JSON Files", "*.json")])
    if file_path:
        save_last_settings()
        os.rename("last_settings.json", file_path)
        status_bar.config(text="Settings saved")

def load_settings():
    file_path = filedialog.askopenfilename(filetypes=[("JSON Files", "*.json")])
    if file_path:
        with open(file_path, "r") as file:
            settings = json.load(file)
            dwg_folder_var.set(settings.get("dwg_folder", ""))
            pdf_folder_var.set(settings.get("pdf_folder", ""))
            step_folder_var.set(settings.get("step_folder", ""))
            drawings_list.delete(*drawings_list.get_children())
            parts_list.delete(*parts_list.get_children())
            for drawing in settings.get("drawings", []):
                drawings_list.insert("", "end", values=drawing)
            for part in settings.get("parts", []):
                parts_list.insert("", "end", values=part)
        status_bar.config(text="Settings loaded")
root = tk.Tk()
root.title("SolidWorks Export Manager")

notebook = ttk.Notebook(root)
notebook.pack(expand=True, fill="both")

dwg_pdf_tab = ttk.Frame(notebook)
step_tab = ttk.Frame(notebook)
notebook.add(dwg_pdf_tab, text="Drawing Export")
notebook.add(step_tab, text="STEP Export")

dwg_folder_var = tk.StringVar()
pdf_folder_var = tk.StringVar()
step_folder_var = tk.StringVar()

dwg_var = tk.BooleanVar()
pdf_var = tk.BooleanVar()

tk.Label(dwg_pdf_tab, text="DWG Export Folder:").pack()
tk.Entry(dwg_pdf_tab, textvariable=dwg_folder_var, width=50).pack()
tk.Button(dwg_pdf_tab, text="Browse", command=select_dwg_folder).pack()

tk.Label(dwg_pdf_tab, text="PDF Export Folder:").pack()
tk.Entry(dwg_pdf_tab, textvariable=pdf_folder_var, width=50).pack()
tk.Button(dwg_pdf_tab, text="Browse", command=select_pdf_folder).pack()

tk.Button(dwg_pdf_tab, text="Select Drawings", command=select_drawings).pack()
drawings_list = ttk.Treeview(dwg_pdf_tab, columns=("File Name", "File Path"), show="headings")
drawings_list.heading("File Name", text="File Name")
drawings_list.heading("File Path", text="File Path")
drawings_list.pack()
tk.Button(dwg_pdf_tab, text="Export Drawings", command=export_drawings).pack()

tk.Label(step_tab, text="STEP Export Folder:").pack()
tk.Entry(step_tab, textvariable=step_folder_var, width=50).pack()
tk.Button(step_tab, text="Browse", command=select_step_folder).pack()

tk.Button(step_tab, text="Select Parts", command=select_parts).pack()
parts_list = ttk.Treeview(step_tab, columns=("File Name", "File Path"), show="headings")
parts_list.heading("File Name", text="File Name")
parts_list.heading("File Path", text="File Path")
parts_list.pack()
tk.Button(step_tab, text="Export Parts", command=export_parts).pack()

status_bar = tk.Label(root, text="Ready", bd=1, relief=tk.SUNKEN, anchor="w")
status_bar.pack(fill="x")

root.mainloop()
