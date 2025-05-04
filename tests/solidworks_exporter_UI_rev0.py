import tkinter as tk
from tkinter import filedialog, ttk
import json
import os
import win32com.client
import pythoncom
import time
import re

def list_slddrw_files(folder_path):
    """Lists all SLDDRW files in the given folder."""
    slddrw_files = [f for f in os.listdir(folder_path) if f.lower().endswith('.slddrw')]
    return slddrw_files

def list_sldprt_files(folder_path):
    """Lists all SLDPRT files in the given folder."""
    sldprt_files = [f for f in os.listdir(folder_path) if f.lower().endswith('.sldprt')]
    return sldprt_files

errors = win32com.client.VARIANT(pythoncom.VT_BYREF | pythoncom.VT_I4, 0)
warnings = win32com.client.VARIANT(pythoncom.VT_BYREF | pythoncom.VT_I4, 0)

def open_and_rebuild_drawing(sw_app, drawing_path):
    try:
        # Open the drawing file
        drawing = sw_app.OpenDoc6(drawing_path, 3, 0, "", errors, warnings)  # 3 indicates drawing document type

        # Rebuild/Refresh the drawing
        drawing.ForceRebuild3(True)

        return drawing
    except Exception as e:
        print(f"An error occurred while opening and rebuilding the drawing: {e}")
        return None

# SolidWorks Interaction - Export to PDF
def export_drawing_to_pdf(sw_app,drawing, pdf_export_path, export_individual_sheets=False):
    try:
        # Get sheet names if exporting individual sheets
        sheet_names = list(drawing.GetSheetNames)
        file_name = os.path.splitext(os.path.basename(pdf_export_path))[0]
        if export_individual_sheets:
            for index, sheet_name in enumerate(sheet_names, start=1):
                # Activate individual sheet
                drawing.ActivateSheet(sheet_name)
                
                pdf_export_dir = os.path.dirname(pdf_export_path)
                # Define export path
                sheet_pdf_export_path = os.path.join(pdf_export_dir, f"{file_name}_sheet{index}.pdf")

                # Save individual sheet as PDF
                export_pdf_data = sw_app.GetExportFileData(1)  # 1 indicates PDF
                export_pdf_data.SetSheets(2, win32com.client.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_BSTR, [sheet_name]))  # 2 = Export current sheet
                export_pdf_data.ViewPdfAfterSaving = False
                success_pdf = drawing.Extension.SaveAs(sheet_pdf_export_path, 0, 0, export_pdf_data, errors, warnings)
                if not success_pdf:
                    print(f"Failed to save sheet {sheet_name} as PDF.")
                else:
                    print(f"Exported sheet {sheet_name} as PDF: {sheet_pdf_export_path}")
        else:
            # Save as PDF (including all sheets if present)
            success_pdf = drawing.SaveAs3(pdf_export_path, 0, 1)
            if success_pdf != 0:
                print("Failed to save the drawing as PDF.")
            else:
                print(f"Exported PDF: {pdf_export_path}")

    except Exception as e:
        print(f"An error occurred: {e}")

# SolidWorks Interaction - Export to DWG
def export_drawing_to_dwg(sw_app,drawing, dwg_export_path, export_individual_sheets=False):
    export_folder_dwg = dwg_export_path
    try:
        # Get sheet names if exporting individual sheets
        sheet_names = list(drawing.GetSheetNames)
        
        file_name = os.path.splitext(os.path.basename(dwg_export_path))[0]
        dwg_export_dir = os.path.dirname(dwg_export_path)
        
        if export_individual_sheets:
            for index, sheet_name in enumerate(sheet_names, start=1):
                # Activate individual sheet
                drawing.ActivateSheet(sheet_name)
                
                
                # Define export path
                sheet_dwg_export_path = os.path.join(dwg_export_dir, f"{file_name}_sheet{index}.dwg")
               # new_path = os.path.join(os.path.dirname(old_path), new_filename)
                
                # Save individual sheet as DWG using SaveAs3
                success_dwg = drawing.SaveAs3(sheet_dwg_export_path, 0, 2)  # 2 = Save only the active sheet
                if success_dwg != 0:
                    print(f"Failed to save sheet {sheet_name} as DWG.")
                else:
                    print(f"Exported sheet {sheet_name} as DWG: {sheet_dwg_export_path}")
        else:
            # Save as DWG (including all sheets if present)
            success_dwg = drawing.SaveAs3(dwg_export_path, 0, 1)
            if success_dwg != 0:
                print("Failed to save the drawing as DWG.")
            else:
                print(f"Exported DWG: {dwg_export_path}")

    except Exception as e:
        print(f"An error occurred: {e}")


def rename_dwg_files(dwg_folder, file_name):
    try:
        # List all DWG files in the folder
        dwg_files = [f for f in os.listdir(dwg_folder) if f.lower().endswith('.dwg')]
        pattern = re.compile(r"^(\d{2})_" + re.escape(file_name) + r"\.dwg$")

        # Filter files matching the pattern like '00_filename.dwg', '01_filename.dwg', etc.
        matching_files = [f for f in dwg_files if pattern.match(f)]
        
        # Rename each file to include the sheet number in a readable format
        for dwg_file in matching_files:
            sheet_index = int(pattern.match(dwg_file).group(1)) + 1
            new_file_name = f"{file_name}_sheet{sheet_index}.dwg"
            old_file_path = os.path.join(dwg_folder, dwg_file)
            new_file_path = os.path.join(dwg_folder, new_file_name)
            os.rename(old_file_path, new_file_path)
            print(f"Renamed {dwg_file} to {new_file_name}")
    except Exception as e:
        print(f"An error occurred while renaming DWG files: {e}")

# New function to open a part or assembly and export it as STEP
def export_part_or_assembly_configurations_to_step(sw_app, part_path, export_folder, selected_configs=None):
    try:
        # Open the part or assembly file
        errors = win32com.client.VARIANT(pythoncom.VT_BYREF | pythoncom.VT_I4, 0)
        warnings = win32com.client.VARIANT(pythoncom.VT_BYREF | pythoncom.VT_I4, 0)
        model = sw_app.OpenDoc6(part_path, 1 if part_path.endswith('.SLDPRT') else 2, 0, "", errors, warnings)  # 1 for part, 2 for assembly

        # Get the configuration names
        configs = model.GetConfigurationNames
        for config_name in configs:
            # If selected_configs is provided, only export those configurations
            if selected_configs and config_name not in selected_configs:
                continue
            # Activate each configuration
            model.ShowConfiguration2(config_name)

            # Define the export path for each configuration
            step_export_path = os.path.join(export_folder, f"{os.path.splitext(os.path.basename(part_path))[0]}_{config_name}.step")

            # Save as STEP
            success_step = model.SaveAs(step_export_path)
            if not success_step:
                print(f"Failed to save configuration '{config_name}' as STEP: {part_path}")
            else:
                print(f"Exported configuration '{config_name}' as STEP: {step_export_path}")

        # Close the part or assembly
        sw_app.CloseDoc(model.GetTitle)
    except Exception as e:
        print(f"An error occurred while exporting part/assembly configurations to STEP: {e}")



def select_dwg_folder():
    folder = filedialog.askdirectory()
    if folder:
        dwg_folder_var.set(folder)

def select_pdf_folder():
    folder = filedialog.askdirectory()
    if folder:
        pdf_folder_var.set(folder)

def select_drawings():
    files = filedialog.askopenfilenames(filetypes=[("SolidWorks Drawings", "*.SLDDRW")])
    for file in files:
        filename = file.split("/")[-1]  # Extract only the file name
        drawings_list.insert("", "end", values=(filename, file))

def delete_selected():
    selected_items = drawings_list.selection()
    for item in selected_items:
        drawings_list.delete(item)



def save_settings():
    settings = {
        "dwg_folder": dwg_folder_var.get(),
        "pdf_folder": pdf_folder_var.get(),
        "export_dwg": dwg_var.get(),
        "export_pdf": pdf_var.get(),
        "drawings": [drawings_list.item(item, "values") for item in drawings_list.get_children()]
    }
    file_path = filedialog.asksaveasfilename(defaultextension=".json", filetypes=[("JSON Files", "*.json")])
    if file_path:
        with open(file_path, "w") as file:
            json.dump(settings, file, indent=4)

def load_settings():
    file_path = filedialog.askopenfilename(filetypes=[("JSON Files", "*.json")])
    if file_path:
        with open(file_path, "r") as file:
            settings = json.load(file)
            dwg_folder_var.set(settings.get("dwg_folder", ""))
            pdf_folder_var.set(settings.get("pdf_folder", ""))
            dwg_var.set(settings.get("export_dwg", False))
            pdf_var.set(settings.get("export_pdf", False))
            drawings_list.delete(*drawings_list.get_children())
            for drawing in settings.get("drawings", []):
                drawings_list.insert("", "end", values=drawing)
                

def export_DRW_Solidworks(drawings_list,export_folder_dwg,export_folder_pdf,
                          export_individual_sheets_pdf,export_individual_sheets_dwg):
    # Connect to SolidWorks
    sw_app = win32com.client.Dispatch('SldWorks.Application')
    sw_app.Visible = True
    
    for drawing_SLDDRW in drawings_list:
        #print(drawing_SLDDRW)
        # drawing_path = drawing_folder + '\\' + drawing_SLDDRW
        drawing_path = drawing_SLDDRW
        
        # Ensure export folder exists
        os.makedirs(export_folder_dwg, exist_ok=True)
        os.makedirs(export_folder_pdf, exist_ok=True)
        
        # Open and rebuild the drawing
        drawing = open_and_rebuild_drawing(sw_app, drawing_path)
        if not drawing:
            continue
        
        # Export file paths
        file_name = os.path.splitext(os.path.basename(drawing_path))[0]
        pdf_export_path = os.path.join(export_folder_pdf, file_name + '.pdf')
        dwg_export_path = os.path.join(export_folder_dwg, file_name + '.dwg')
        
        # Export the drawing to DWG and PDF
        export_drawing_to_pdf(sw_app, drawing, pdf_export_path, export_individual_sheets=export_individual_sheets_pdf)
        export_drawing_to_dwg(sw_app, drawing, dwg_export_path, export_individual_sheets=export_individual_sheets_dwg)
    
        # Rename DWG files if multiple sheets are present
        #rename_dwg_files(export_folder_dwg, file_name)
        
        # Close the drawing
        sw_app.CloseDoc(drawing.GetTitle)
        
def export_drawings():
    export_folder_dwg = dwg_folder_var.get()
    export_folder_pdf = pdf_folder_var.get()
    export_individual_sheets_dwg = dwg_var.get()
    export_individual_sheets_pdf = pdf_var.get()
    drawings = [drawings_list.item(item, "values") for item in drawings_list.get_children()]
    drawings_list2 = [d[1] for d in drawings]
    
    print("Exporting with options:")
    print(f"DWG Folder: {export_folder_dwg}")
    print(f"PDF Folder: {export_folder_pdf}")
    # print(f"Export DWG: {export_individual_sheets_dwg}")
    # print(f"Export PDF: {export_individual_sheets_pdf}")
    # print(f"Drawings: {drawings}")
    
    export_DRW_Solidworks(drawings_list2,export_folder_dwg,export_folder_pdf,
                               export_individual_sheets_pdf,export_individual_sheets_dwg)
    # Here, call the SolidWorks API to process the export
    
root = tk.Tk()
root.title("SolidWorks Drawing Exporter")

# Folder selection
dwg_folder_var = tk.StringVar()
pdf_folder_var = tk.StringVar()

tk.Label(root, text="DWG Export Folder:").grid(row=0, column=0, sticky="w")
tk.Entry(root, textvariable=dwg_folder_var, width=50).grid(row=0, column=1)
tk.Button(root, text="Browse", command=select_dwg_folder).grid(row=0, column=2)

tk.Label(root, text="PDF Export Folder:").grid(row=1, column=0, sticky="w")
tk.Entry(root, textvariable=pdf_folder_var, width=50).grid(row=1, column=1)
tk.Button(root, text="Browse", command=select_pdf_folder).grid(row=1, column=2)

# Checkboxes for individual sheet export
dwg_var = tk.BooleanVar()
pdf_var = tk.BooleanVar()
tk.Checkbutton(root, text="Export individual DWG sheets", variable=dwg_var).grid(row=2, column=0, columnspan=2, sticky="w")
tk.Checkbutton(root, text="Export individual PDF sheets", variable=pdf_var).grid(row=3, column=0, columnspan=2, sticky="w")

# File selection
tk.Button(root, text="Select Drawings", command=select_drawings).grid(row=4, column=0, columnspan=3, pady=5)

# Table with two columns
drawings_list = ttk.Treeview(root, columns=("File Name", "File Path"), show="headings")
drawings_list.heading("File Name", text="File Name")
drawings_list.heading("File Path", text="File Path")
drawings_list.column("File Name", width=150)
drawings_list.column("File Path", width=350)
drawings_list.grid(row=5, column=0, columnspan=3, padx=5, pady=5, sticky="nsew")

# Add vertical and horizontal scrollbars
scrollbar_y = ttk.Scrollbar(root, orient="vertical", command=drawings_list.yview)
drawings_list.configure(yscrollcommand=scrollbar_y.set)
scrollbar_y.grid(row=5, column=3, sticky="ns")

scrollbar_x = ttk.Scrollbar(root, orient="horizontal", command=drawings_list.xview)
drawings_list.configure(xscrollcommand=scrollbar_x.set)
scrollbar_x.grid(row=6, column=0, columnspan=3, sticky="ew")

# Delete button
tk.Button(root, text="Delete Selected", command=delete_selected).grid(row=7, column=0, columnspan=3, pady=5)

# Export button
tk.Button(root, text="Export", command=export_drawings).grid(row=8, column=0, columnspan=3, pady=10)

# Save and Load buttons
tk.Button(root, text="Save Settings", command=save_settings).grid(row=9, column=0, pady=5)
tk.Button(root, text="Load Settings", command=load_settings).grid(row=9, column=1, pady=5)

root.mainloop()
