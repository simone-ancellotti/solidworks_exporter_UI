# -*- coding: utf-8 -*-
"""
Created on Mon Jun  2 11:29:38 2025

@author: user
"""

import win32com.client
import pythoncom
import os
from datetime import date, timedelta





# List your SLDDRW files here (absolute or relative paths)
slddrw_files = [
    r"G:/My Drive/ULIX tecnico/JOB/JOB 7 Isopor/CAD/JOB-7_Battery_HeatExchanger/drawings/Job-7_cover_battery_extended_ULIX.SLDDRW",
    #r"G:/My Drive/ULIX tecnico/JOB/JOB 7 Isopor/CAD/JOB-7_Battery_HeatExchanger/drawings/Job-7_extention_C_ULIX_rev2.SLDDRW",
    # Add more paths
]

swSummInfoField_e = {
    "swSumInfoAuthor":	2,
    "swSumInfoComment":	4,
    "swSumInfoCreateDate":	6,
    "swSumInfoCreateDate2":	8,
    "swSumInfoKeywords":	3,
    "swSumInfoSaveDate":	7,
    "swSumInfoSaveDate2":	9,
    "swSumInfoSavedBy":	5,
    "swSumInfoSubject":	1,
    "swSumInfoTitle":	0,
    }


def get_properties_from_drawing(sw_app, drawing_path):
    errors = win32com.client.VARIANT(pythoncom.VT_BYREF | pythoncom.VT_I4, 0)
    warnings = win32com.client.VARIANT(pythoncom.VT_BYREF | pythoncom.VT_I4, 0)
    # Open drawing
    drawing = sw_app.OpenDoc6(drawing_path, 3, 0, "", errors, warnings)  # 3 = drawing
    if not drawing:
        print(f"Failed to open {drawing_path}")
        return

    # Get Custom Property Manager (for the active configuration, usually "Sheet1" or use "")
    config_name = ""  # For custom properties not tied to a specific config
    prop_mgr = drawing.Extension.CustomPropertyManager(config_name)
    prop_mgr.GetNames

    # List of properties you want
    prop_names = [
        "Revision", "machining1", "machining2", "machining3", "machining4",
        "machining5", "all_chamfers", "all_fillets", "Checked by",
        "Approved by", "Revision Comment", "Revision Date", "Revision Done by",
        "Revision Approved by", "Customer", "Project", "Description", "Drawing Date"
    ]
    
    # Get each property value
    properties = {}
    for prop_name in prop_names:
        #val, resolved, was_found = prop_mgr.Get(name, False, "", "")
        #properties[name] = resolved if was_found else ""
        val = prop_mgr.Get(prop_name)
        properties[prop_name] = val

    # Summary Info (Author, Title, etc.)
    summary_info = drawing.SummaryInfo
    author = summary_info(swSummInfoField_e["swSumInfoAuthor"])
    title = summary_info(swSummInfoField_e["swSumInfoTitle"])

    properties["Author"] = author
    properties["Title"] = title

    # Close the drawing
    sw_app.CloseDoc(drawing.GetTitle)
    
    return properties

def update_drawing_property(sw_app, drawing_path, prop_name, prop_value, is_summary=False, summary_index=None):
    errors = win32com.client.VARIANT(pythoncom.VT_BYREF | pythoncom.VT_I4, 0)
    warnings = win32com.client.VARIANT(pythoncom.VT_BYREF | pythoncom.VT_I4, 0)
    drawing = sw_app.OpenDoc6(drawing_path, 3, 0, "", errors, warnings)
    if not drawing:
        print(f"Failed to open {drawing_path}")
        return False

    if is_summary and summary_index is not None:
        drawing.SummaryInfo.SetCustomInfo(summary_index, prop_value)
    else:
        prop_mgr = drawing.Extension.CustomPropertyManager("")
        #prop_mgr.Add3(prop_name, 64,prop_value ,1 )
        prop_mgr.Set2(prop_name,prop_value)
    #sw_app.CloseDoc(drawing.GetTitle)
    return True

#def main():


if __name__ == "__main__":
    sw_app = win32com.client.Dispatch("SldWorks.Application")
    sw_app.Visible = True
    
    today = date.today()+ timedelta(days=7)
    formatted_date = today.strftime('%B %d, %Y')
    print(formatted_date)
    all_results = {}

    for drawing_path in slddrw_files:
        print(f"Processing: {drawing_path}")
        props = get_properties_from_drawing(sw_app, drawing_path)
        all_results[os.path.basename(drawing_path)] = props
    
    update_drawing_property(sw_app, drawing_path, 'Revision Date', formatted_date, is_summary=False, summary_index=None)

    # Print results
    for filename, props in all_results.items():
        print(f"\nProperties for {filename}:")
        for k, v in props.items():
            print(f"  {k}: {v}")
