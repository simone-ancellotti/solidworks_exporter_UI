# -*- coding: utf-8 -*-
"""
Created on Tue Jun  3 15:39:34 2025

@author: user
"""

import os
import win32com.client
import pythoncom
import time
import re

swUserPreferenceIntegerValue_e = {'swDxfMultiSheetOption':253}
swDxfMultisheet_e = {'swDxfActiveSheetOnly':0,'swDxfMultiSheet':2,'swDxfSeparateSheets':1}


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

def open_and_rebuild_drawing(sw_app, drawing_path,flagForceRebuild = True):
    try:
        # Open the drawing file
        drawing = sw_app.OpenDoc6(drawing_path, 3, 0, "", errors, warnings)  # 3 indicates drawing document type

        # Rebuild/Refresh the drawing
        drawing.ForceRebuild3(flagForceRebuild)

        return drawing
    except Exception as e:
        print(f"An error occurred while opening and rebuilding the drawing: {e}")
        return None
    
def open_and_rebuild_SLDWRKS_file(sw_app, file_path,flagForceRebuild = True):
    try:
        type_of_file = 0 
        if file_path.upper().endswith('.SLDDRW'):
            type_of_file= 3
        else:
            if file_path.upper().endswith('.SLDPRT'):
                type_of_file =  1
            else: 
                type_of_file = 2
        # Open the drawing file
        drawing = sw_app.OpenDoc6(file_path, type_of_file, 0, "", errors, warnings)  # 3 indicates drawing document type

        # Rebuild/Refresh the drawing
        drawing.ForceRebuild3(flagForceRebuild)

        return drawing
    except Exception as e:
        print(f"An error occurred while opening and rebuilding the file: {e}")
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
                #name_out_file = f"{file_name}_sheet{index}.pdf"
                name_out_file = f"{file_name}_{sheet_name}.pdf"
                sheet_pdf_export_path = os.path.join(pdf_export_dir, name_out_file )

                # Save individual sheet as PDF
                export_pdf_data = sw_app.GetExportFileData(1)  # 1 indicates PDF
                export_pdf_data.SetSheets(2, win32com.client.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_BSTR, [sheet_name]))  # 2 = Export current sheet
                export_pdf_data.ViewPdfAfterSaving = False
                success_pdf = drawing.Extension.SaveAs(sheet_pdf_export_path, 0, 0, export_pdf_data, errors, warnings)
                if not success_pdf or not(os.path.isfile(sheet_pdf_export_path)):
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
    try:
        # Get sheet names if exporting individual sheets
        sheet_names = list(drawing.GetSheetNames)
        
        file_name = os.path.splitext(os.path.basename(dwg_export_path))[0]
        dwg_export_dir = os.path.dirname(dwg_export_path)
        
        if export_individual_sheets:
            boolstatus = sw_app.SetUserPreferenceIntegerValue(
                swUserPreferenceIntegerValue_e['swDxfMultiSheetOption'],
                swDxfMultisheet_e['swDxfActiveSheetOnly']
                )
            for index, sheet_name in enumerate(sheet_names, start=1):
                # Activate individual sheet
                drawing.ActivateSheet(sheet_name)
                
                
                # Define export path
                # name_out_file = f"{file_name}_sheet{index}.dwg"
                name_out_file = f"{file_name}_{sheet_name}.dwg"
                sheet_dwg_export_path = os.path.join(dwg_export_dir, name_out_file )
               # new_path = os.path.join(os.path.dirname(old_path), new_filename)
                
                # Save individual sheet as DWG using SaveAs3
                success_dwg = drawing.SaveAs3(sheet_dwg_export_path, 0, 1)  # 2 = Save only the active sheet
                if success_dwg != 0 or not(os.path.isfile(sheet_dwg_export_path)):
                    print(f"Failed to save sheet {sheet_name} as DWG.")
                else:
                    print(f"Exported sheet {sheet_name} as DWG: {sheet_dwg_export_path}")
        else:
            # Save as DWG (including all sheets if present)
            boolstatus = sw_app.SetUserPreferenceIntegerValue(
                swUserPreferenceIntegerValue_e['swDxfMultiSheetOption'],
                swDxfMultisheet_e['swDxfMultiSheet']
                )
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
        model = sw_app.OpenDoc6(part_path, 1 if part_path.upper().endswith('.SLDPRT') else 2, 0, "", errors, warnings)  # 1 for part, 2 for assembly

        # Get the configuration names
        configs = model.GetConfigurationNames
        
        found_config_flag = False
        for config_name in configs:
            # If selected_configs is provided, only export those configurations
            if selected_configs and config_name not in selected_configs:
                
                continue
            # Activate each configuration
            model.ShowConfiguration2(config_name)
            found_config_flag = True
            
            config_name_epurated = config_name.replace('<','_')
            config_name_epurated = config_name_epurated.replace('>','')
            config_name_epurated = config_name_epurated.replace(' ','_')
            # Define the export path for each configuration
            step_export_path = os.path.join(export_folder, f"{os.path.splitext(os.path.basename(part_path))[0]}_{config_name_epurated}.step")

            # Save as STEP
            success_step = model.SaveAs(step_export_path)
            if not success_step or not(os.path.isfile(step_export_path)):
                print(f"Failed to save configuration '{config_name}' as STEP: {part_path}")
            else:
                print(f"Exported configuration '{config_name}' as STEP: {step_export_path}")
        
        if not(found_config_flag):
            print('Error: Selected Configuartion not found')
            
        # Close the part or assembly
        sw_app.CloseDoc(model.GetTitle)
    except Exception as e:
        print(f"An error occurred while exporting part/assembly configurations to STEP: {e}")

def get_model_GetConfigurationNames(sw_app, part_path):
    try: 
        model = sw_app.OpenDoc6(part_path, 1 if part_path.upper().endswith('.SLDPRT') else 2, 0, "", errors, warnings)  # 1 for part, 2 for assembly
        # Get the configuration names
        configs = model.GetConfigurationNames
        # Close the part or assembly
        sw_app.CloseDoc(model.GetTitle)
        
        return  configs
    except Exception as e:
        print(f"An error occurred while exporting part/assembly configurations : {e}")
        





