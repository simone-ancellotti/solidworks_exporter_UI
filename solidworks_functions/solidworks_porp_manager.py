# -*- coding: utf-8 -*-
"""
Created on Wed Jun  4 11:54:24 2025

@author: user
"""
import win32com.client
import pythoncom

# List of properties you want

prop_names = [
    "Revision",  "Checked by",
    "Approved by", "Revision Comment", "Revision Date", "Revision Done by",
    "Revision Approved by", "Customer", "Project", "Description", "Drawing Date",
    "machining1", "machining2", "machining3", "machining4",
    "machining5", "all_chamfers", "all_fillets"
]

properties_columns_all = ["Author","Title"] + prop_names 

errors = win32com.client.VARIANT(pythoncom.VT_BYREF | pythoncom.VT_I4, 0)
warnings = win32com.client.VARIANT(pythoncom.VT_BYREF | pythoncom.VT_I4, 0)

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
prop2SumnInfo = {
    "Author": "swSumInfoAuthor",
    "Title" : "swSumInfoTitle",
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

    # Get each property value
    properties = {}
    for prop_name in prop_names:
        #val, resolved, was_found = prop_mgr.Get(name, False, "", "")
        #properties[name] = resolved if was_found else ""
        val = prop_mgr.Get(prop_name)
        properties[prop_name] = {'value':val,'type':"CustomProperty"}

    # Summary Info (Author, Title, etc.)
    summary_info = drawing.SummaryInfo
    author = summary_info(swSummInfoField_e["swSumInfoAuthor"])
    title = summary_info(swSummInfoField_e["swSumInfoTitle"])

    properties["Author"] = {'value':author,'type':"SummInfoField"}
    properties["Title"] = {'value':title,'type':"SummInfoField"}

    # Close the drawing
    sw_app.CloseDoc(drawing.GetTitle)
    
    return properties

def update_drawing_property(sw_app, drawing_path, properties, flagRebuild = False ):
    
    
    drawing = sw_app.OpenDoc6(drawing_path, 3, 0, "", errors, warnings)
    drawing.ForceRebuild3(flagRebuild)
    if not drawing:
        print(f"Failed to open {drawing_path}")
        return False
    
    prop_mgr = drawing.Extension.CustomPropertyManager("")
    for prop_name, prop_content in properties.items():
        #print(prop_content)
        prop_value = prop_content['value']
        prop_type = prop_content['type']
        is_summary = prop_type == "SummInfoField"
        if is_summary :
            summary_index = swSummInfoField_e[prop2SumnInfo[prop_name]]
            if drawing.SummaryInfo(summary_index) != prop_value:
                drawing.SummaryInfo(summary_index, prop_value)
        else:
            if prop_name in prop_mgr.GetNames:
                if prop_mgr.Get(prop_name)!= prop_value:
                    #prop_mgr.Add3(prop_name, 64,prop_value ,1 )
                    prop_mgr.Set2(prop_name,prop_value)
    drawing.Save()
    sw_app.CloseDoc(drawing.GetTitle)
    return True