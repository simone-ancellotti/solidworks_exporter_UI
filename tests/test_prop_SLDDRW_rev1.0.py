# -*- coding: utf-8 -*-
"""
Created on Mon Jun  2 11:29:38 2025

@author: user
"""

import win32com.client
import pythoncom
import os
import sys
from datetime import date, timedelta

sys.path.append(os.path.join(os.path.dirname(__file__), '../solidworks_functions'))
from solidworks_porp_manager import (
    get_properties_from_drawing,
    update_drawing_property,
)



# List your SLDDRW files here (absolute or relative paths)
slddrw_files = [
    r"G:/My Drive/ULIX tecnico/Varie/Pentano/Pentane_Distillator/Distillator_pentane_CAD/drawings/Mesh_Support.SLDDRW",
    #r"G:/My Drive/ULIX tecnico/JOB/JOB 7 Isopor/CAD/JOB-7_Battery_HeatExchanger/drawings/Job-7_extention_C_ULIX_rev2.SLDDRW",
    # Add more paths
]







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
    
    props=all_results[os.path.basename(drawing_path)]
    props['Revision Date']['value'] = formatted_date
    props['Title']['value'] = 'Strirato'
    
    update_drawing_property( sw_app, drawing_path, props)
    
    props = get_properties_from_drawing(sw_app, drawing_path)
    all_results[os.path.basename(drawing_path)] = props

    # Print results
    for filename, props in all_results.items():
        print(f"\nProperties for {filename}:")
        for k, v in props.items():
            print(f"  {k}: {v['value']}")
    
    
