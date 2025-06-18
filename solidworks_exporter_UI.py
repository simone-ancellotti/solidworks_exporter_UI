import win32com.client
import pythoncom
import sys
import os
import json
import time
import pandas as pd
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QWidget, QLabel, QPushButton, QVBoxLayout,
    QHBoxLayout, QFileDialog, QLineEdit, QTabWidget, QCheckBox, QProgressBar,
    QTableWidget, QTableWidgetItem, QMessageBox, QHeaderView,QAction,QFileDialog,
    QStyledItemDelegate, QComboBox
)
from PyQt5.QtCore import Qt, QTimer



sys.path.append(os.path.join(os.path.dirname(__file__), './solidworks_functions'))
from solidworks_export import (
    open_and_rebuild_drawing,
    export_drawing_to_pdf,
    export_drawing_to_dwg,
    export_part_or_assembly_configurations_to_step,
    get_model_GetConfigurationNames,
)

from solidworks_porp_manager import (
    get_properties_from_drawing,
    prop_names,
    properties_columns_all,
    update_drawing_property,
    )



class SolidWorksExportManager(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("SolidWorks Export Manager")
        self.setMinimumWidth(800)

        self.dwg_folder = ""
        self.pdf_folder = ""
        self.step_folder = ""
        
        self.slddrw_properties={}
        
        self.sw_app = None

        self.tabs = QTabWidget()
        self.setCentralWidget(self.tabs)

        self.status = QLabel("Ready")
        self.statusBar().addWidget(self.status)
        
        self.init_dwg_pdf_tab()
        self.title_blocks_tab()
        self.init_step_tab()
        
        self.progress = QProgressBar()
        self.statusBar().addPermanentWidget(self.progress)
        self.progress.setVisible(False)
        self.create_menu()
        
    def create_menu(self):
        # Create the menubar
        menubar = self.menuBar()
        
        # Create the File menu
        file_menu = menubar.addMenu("File")
        
        # Add actions to the File menu
        load_action = QAction("Load settings", self)
        load_action.triggered.connect(self.load_settings)
        file_menu.addAction(load_action)
        
        save_action = QAction("Save settings", self)
        save_action.triggered.connect(self.save_settings)
        file_menu.addAction(save_action)
        
        file_menu.addSeparator()

    def init_dwg_pdf_tab(self):
        tab = QWidget()
        layout = QVBoxLayout()

        # DWG Folder
        self.dwg_folder_edit = QLineEdit()
        btn_dwg_folder = QPushButton("Browse")
        btn_dwg_folder.clicked.connect(self.select_dwg_folder)
        layout.addWidget(QLabel("DWG Export Folder:"))
        layout.addLayout(self.make_folder_layout(self.dwg_folder_edit, btn_dwg_folder))

        # PDF Folder
        self.pdf_folder_edit = QLineEdit()
        btn_pdf_folder = QPushButton("Browse")
        btn_pdf_folder.clicked.connect(self.select_pdf_folder)
        layout.addWidget(QLabel("PDF Export Folder:"))
        layout.addLayout(self.make_folder_layout(self.pdf_folder_edit, btn_pdf_folder))
        
        
        

       # self.pdf_check_boxes = QVBoxLayout()
        self.SLDWRK_visible_checkbox = QCheckBox("Run Solidworks visible")
        self.SLDWRK_visible_checkbox.setChecked(True)  # Default True
        self.dwg_checkbox = QCheckBox("Export DWG")
        self.dwg_checkbox.setChecked(True)  # Default True
        
        self.SLDWRK_close_expDRW_checkbox = QCheckBox("Close exported drws")
        self.SLDWRK_close_expDRW_checkbox.setChecked(True)  
        
        self.pdf_checkbox = QCheckBox("Export PDF")
        self.pdf_checkbox.setChecked(True)
        self.pdf_indiv_checkbox = QCheckBox("Export individual PDF sheets")
        self.dwg_indiv_checkbox = QCheckBox("Export individual DWG sheets")
        layout.addLayout(self.make_check_buttons_layout([
            self.SLDWRK_visible_checkbox,self.pdf_checkbox, self.pdf_indiv_checkbox]))
        layout.addLayout(self.make_check_buttons_layout([
            self.SLDWRK_close_expDRW_checkbox,self.dwg_checkbox, self.dwg_indiv_checkbox]))
        # layout.addLayout(self.make_check_2buttons_layout(
        #     self.pdf_checkbox, self.pdf_indiv_checkbox
        #     ))
        # layout.addLayout(self.make_check_2buttons_layout(
        #     self.dwg_checkbox, self.dwg_indiv_checkbox
        #     ))
        # layout.addLayout(self.make_check_buttons_layout([pdf_checkbox, pdf_indiv_checkbox,dwg_checkbox, dwg_indiv_checkbox]))



        # Drawing List
        btn_drawings = QPushButton("Select Drawings")
        btn_drawings.clicked.connect(self.select_drawings)
        layout.addWidget(btn_drawings)
        
        self.drawings_table = QTableWidget(0, 2)
        self.drawings_table.setHorizontalHeaderLabels(["File Name", "File Path"])
        self.drawings_table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        layout.addWidget(self.drawings_table)

        btn_delete = QPushButton("Delete Selected")
        btn_delete.setFixedWidth(150)
        btn_delete.clicked.connect(self.delete_selected_drawings)
        #layout.addWidget(btn_delete, alignment=Qt.AlignLeft)
        
        btn_export = QPushButton("Export Drawings")
        btn_export.setFixedWidth(150)
        btn_export.clicked.connect(self.export_drawings)
        #layout.addWidget(btn_export)
        layout.addLayout(self.make_folder_layout(btn_delete, btn_export))

        tab.setLayout(layout)
        self.tabs.addTab(tab, "Drawing Export")


    def title_blocks_tab(self):
        tab = QWidget()
        layout = QVBoxLayout()

        # Drawing List
        btn_drawings = QPushButton("Retrive title blocks")
        btn_drawings.clicked.connect(self.retrive_SLDDRW_properties)
        btn_updateSLWDRW = QPushButton("Sync. into SLWDRW")
        btn_updateSLWDRW.clicked.connect(self.update_into_SLDDRW_files)
        #layout.addWidget(btn_drawings)
        layout.addLayout(self.make_check_buttons_layout([
            btn_drawings,btn_updateSLWDRW]))
        
        
        columns = properties_columns_all
        self.drawings_prop_table = QTableWidget()
        self.drawings_prop_table.setColumnCount(len(columns) + 1)
        self.drawings_prop_table.setHorizontalHeaderLabels(["File name"] + columns)
        self.drawings_prop_table.setRowCount(len(self.slddrw_properties))
        self.drawings_prop_table.cellChanged.connect(self.on_cell_changed_drawings_prop_table)
        layout.addWidget(self.drawings_prop_table)
        
        self.btn_export_excel = QPushButton("Export to Excel")
        self.btn_export_excel.clicked.connect(self.export_table_to_excel)
        
        self.btn_import_excel = QPushButton("Import from Excel")
        self.btn_import_excel.clicked.connect(self.import_table_from_excel)
        layout.addLayout(self.make_check_buttons_layout([
            self.btn_export_excel,self.btn_import_excel]))


        tab.setLayout(layout)
        self.tabs.addTab(tab, "Title blocks")
        
    def init_step_tab(self):
        tab = QWidget()
        layout = QVBoxLayout()

        # STEP Folder
        self.step_folder_edit = QLineEdit()
        btn_step_folder = QPushButton("Browse")
        btn_step_folder.clicked.connect(self.select_step_folder)
        layout.addWidget(QLabel("STEP Export Folder:"))
        layout.addLayout(self.make_folder_layout(self.step_folder_edit, btn_step_folder))

        # Parts List
        btn_parts = QPushButton("Select Parts")
        btn_parts.clicked.connect(self.select_parts)
        btn_parts_config = QPushButton("Extract Config.")
        btn_parts_config.clicked.connect(self.extract_parts_config)
        #layout.addWidget(btn_parts)
        layout.addLayout(self.make_check_buttons_layout([
            btn_parts ,btn_parts_config]))

        self.parts_table_HeaderLabels = ["File Name", "File Path","config"]
        self.parts_table = QTableWidget(0, 3)
        self.parts_table.setHorizontalHeaderLabels(self.parts_table_HeaderLabels)
        self.parts_table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        layout.addWidget(self.parts_table)

        btn_export_parts = QPushButton("Export Parts")
        btn_export_parts.clicked.connect(self.export_parts)
        btn_export_parts.setFixedWidth(150)
        # layout.addWidget(btn_export_parts)
        btn_delete_prt = QPushButton("Delete Selected")
        btn_delete_prt.setFixedWidth(150)
        btn_delete_prt.clicked.connect(self.delete_selected_parts)
        layout.addLayout(self.make_check_buttons_layout([
            btn_delete_prt ,btn_export_parts]))

        tab.setLayout(layout)
        self.tabs.addTab(tab, "STEP Export")
        
        
    def make_folder_layout(self, line_edit, button):
        layout = QHBoxLayout()
        layout.addWidget(line_edit)
        layout.addWidget(button)
        return layout
    
    def make_check_2buttons_layout(self, button1, button2):
        layout = QHBoxLayout()
        layout.addWidget(button1)
        layout.addWidget(button2)
        return layout
    def make_check_buttons_layout(self, button_list):
        layout = QHBoxLayout()
        for button in button_list:
            layout.addWidget(button)
        return layout

    def select_dwg_folder(self):
        folder = QFileDialog.getExistingDirectory(self, "Select DWG Folder")
        if folder:
            self.dwg_folder_edit.setText(folder)
            self.status.setText("DWG folder selected")

    def select_pdf_folder(self):
        folder = QFileDialog.getExistingDirectory(self, "Select PDF Folder")
        if folder:
            self.pdf_folder_edit.setText(folder)
            self.status.setText("PDF folder selected")

    def select_step_folder(self):
        folder = QFileDialog.getExistingDirectory(self, "Select STEP Folder")
        if folder:
            self.step_folder_edit.setText(folder)
            self.status.setText("STEP export folder selected")

    def select_drawings(self):
        files, _ = QFileDialog.getOpenFileNames(self, "Select Drawings", "", "SolidWorks Drawings (*.SLDDRW)")
        for file in files:
            filename = os.path.basename(file)
            self.drawings_table.insertRow(self.drawings_table.rowCount())
            self.drawings_table.setItem(self.drawings_table.rowCount()-1, 0, QTableWidgetItem(filename))
            self.drawings_table.setItem(self.drawings_table.rowCount()-1, 1, QTableWidgetItem(file))
            
        self.status.setText("Drawings selected")
        
        self.sync_drawings_table1_and_2()
        

    def select_parts(self):
        files, _ = QFileDialog.getOpenFileNames(self, "Select Parts", "", "SolidWorks Files (*.SLDPRT *.SLDASM)" )
        for file in files:
            filename = os.path.basename(file)
            self.parts_table.insertRow(self.parts_table.rowCount())
            i = self.parts_table_HeaderLabels.index("File Name")
            self.parts_table.setItem(self.parts_table.rowCount()-1, i, QTableWidgetItem(filename))
            i = self.parts_table_HeaderLabels.index("File Path")
            self.parts_table.setItem(self.parts_table.rowCount()-1, i, QTableWidgetItem(file))
            
            
        i = self.parts_table_HeaderLabels.index("config")
        self.add_QComboBox_column(self.parts_table,i,configs = ["All"])
        self.status.setText("Parts selected")

    def add_QComboBox_column(self,qtTable, target_column,configs = ["All"]):
        for row in range(qtTable.rowCount()):
            self.add_QComboBox_cell(qtTable, target_column, row,configs)
            
    def add_QComboBox_cell(self,qtTable, target_column,target_row,configs = ["All"],value = None):
        combo = QComboBox()
        # Get list of configs for this row (maybe from your data)
        combo.addItems(configs)
        if value:
            idx = combo.findText(value)
            if idx >= 0:
                combo.setCurrentIndex(idx)
        qtTable.setCellWidget(target_row, target_column, combo)
        return combo
        
    def extract_parts_config(self):
        rows = self.parts_table.rowCount()
        if rows == 0:
            QMessageBox.warning(self, "Warning", "No parts selected for config. extractions.")
            return
        # Connect to SolidWorks
        self.start_SLDWRK()
        self.sw_app.Visible = self.SLDWRK_visible_checkbox.isChecked()
        
        count = rows
        self.progress.setMaximum(count)
        self.progress.setValue(0)
        self.progress.setVisible(True)
        self.status.setText(f"Extraction config...")
        self.current_index = 0
        self.total_items = count
        data_table_parts = self.get_table_data(self.parts_table)
        row = -1
        for row_table in data_table_parts:
            i = self.parts_table_HeaderLabels.index("File Name")
            part_name = row_table[i]
            i = self.parts_table_HeaderLabels.index("File Path")
            part_path = row_table[i]
            row +=1
            configs = get_model_GetConfigurationNames(
                sw_app = self.sw_app,
                part_path = part_path,
                )
            configs_all = ("All",)+tuple(configs)
            config_column = self.parts_table_HeaderLabels.index("config")
            self.add_QComboBox_cell(self.parts_table, 
                                    target_column = config_column,
                                    target_row = row,
                                    configs = configs_all )
            print(f"Extracting config: {self.current_index}/{self.total_items}")
            self.current_index += 1
            self.progress.setValue(self.current_index-1)
            
        self.progress.setVisible(False)
        self.status.setText("Extraction config completed")
        QMessageBox.information(self, "Export Completed", "Extraction config process has finished.")
        
        return 
    
    def export_drawings(self):
        rows = self.drawings_table.rowCount()
        if rows == 0:
            QMessageBox.warning(self, "Export Warning", "No drawings selected for export.")
            return
        
        # Connect to SolidWorks
        self.start_SLDWRK()
        self.sw_app.Visible = self.SLDWRK_visible_checkbox.isChecked()
        
        self.export_type = "drawings"
        #self.run_export_simulation(rows, "drawings")
        export_folder_dwg = self.dwg_folder_edit.text()
        
        
        export_folder_pdf = self.pdf_folder_edit.text()
        flag_export_individual_sheets_dwg = self.dwg_indiv_checkbox.isChecked()
        flag_export_individual_sheets_pdf = self.pdf_indiv_checkbox.isChecked()
        flag_export_dwg_ = self.dwg_checkbox.isChecked()
        flag_export_pdf_ = self.pdf_checkbox.isChecked()
        drawings = self.get_drawings_table1_data()
        
        num_SLWDRW = len(drawings) 
        progress_bar_length = 0
        if flag_export_dwg_:
            progress_bar_length += num_SLWDRW
        if flag_export_pdf_:
            progress_bar_length += num_SLWDRW
        
        self.progress.setMaximum(progress_bar_length)
        self.progress.setValue(0)
        self.progress.setVisible(True)
        self.status.setText("Exporting drawings...")
        self.current_index = 0
        
        drawings_path_list2 = [d[1] for d in drawings]
        
        print("Exporting with options:")
        print(f"DWG Folder: {export_folder_dwg}")
        print(f"PDF Folder: {export_folder_pdf}")
        print(f"Export DWG: {flag_export_dwg_}")
        print(f"Export PDF: {flag_export_pdf_}")
        print(f"Export individual DWG: {flag_export_individual_sheets_dwg}")
        print(f"Export individual PDF: {flag_export_individual_sheets_pdf}")
        print("")
        
        self.export_DRW_Solidworks(self.sw_app,drawings_path_list2,export_folder_dwg,export_folder_pdf,
                              flag_export_dwg_, flag_export_pdf_,
                              flag_export_individual_sheets_pdf,flag_export_individual_sheets_dwg)
        
        self.progress.setVisible(False)
        self.status.setText(f"{self.export_type.capitalize()} export completed")
        QMessageBox.information(self, "Export Completed", f"{self.export_type.capitalize()} export process has finished.")

        
        
            

    def export_DRW_Solidworks(self,sw_app,drawings_list,export_folder_dwg,export_folder_pdf,
                              flag_export_dwg, flag_export_pdf,
                              export_individual_sheets_pdf,export_individual_sheets_dwg):

        
        
        for i ,drawing_SLDDRW in enumerate(drawings_list, start=1):
            #print(drawing_SLDDRW)
            # drawing_path = drawing_folder + '\\' + drawing_SLDDRW
            drawing_path = drawing_SLDDRW
            
            # Ensure export folder exists
            os.makedirs(export_folder_dwg, exist_ok=True)
            os.makedirs(export_folder_pdf, exist_ok=True)
            
            # Open and rebuild the drawing
            drawing = open_and_rebuild_drawing(sw_app, drawing_path, flagForceRebuild = True)
            if not drawing:
                continue
            
            # Export file paths
            file_name = os.path.splitext(os.path.basename(drawing_path))[0]
            pdf_export_path = os.path.join(export_folder_pdf, file_name + '.pdf')
            dwg_export_path = os.path.join(export_folder_dwg, file_name + '.dwg')
            
            # Export the drawing to DWG and PDF
            if flag_export_pdf:
                export_drawing_to_pdf(sw_app, drawing, pdf_export_path, export_individual_sheets=export_individual_sheets_pdf)
                self.current_index += 1
            if flag_export_dwg:
                export_drawing_to_dwg(sw_app, drawing, dwg_export_path, export_individual_sheets=export_individual_sheets_dwg)
                self.current_index += 1
            self.progress.setValue(self.current_index)
            # Close the drawing
            if self.SLDWRK_close_expDRW_checkbox.isChecked():
                sw_app.CloseDoc(drawing.GetTitle)
    
    def delete_selected_rows_of_table(self,table):
        selected = table.selectionModel().selectedRows()
        for index in sorted(selected, reverse=True):
            table.removeRow(index.row())
            
    def delete_selected_drawings(self):
        # selected = self.drawings_table.selectionModel().selectedRows()
        # for index in sorted(selected, reverse=True):
        #     self.drawings_table.removeRow(index.row())
        self.delete_selected_rows_of_table(self.drawings_table)
        
        self.sync_drawings_table1_and_2()
        self.status.setText("Selected drawings deleted")
        
    def delete_selected_parts(self):
        self.delete_selected_rows_of_table(self.parts_table)
        self.status.setText("Selected parts deleted")
     

    def export_parts(self):
        rows = self.parts_table.rowCount()
        if rows == 0:
            QMessageBox.warning(self, "Export Warning", "No parts selected for STEP export.")
            return
        #self.run_export_simulation(rows, "parts")
        # Connect to SolidWorks
        self.start_SLDWRK()
        self.sw_app.Visible = self.SLDWRK_visible_checkbox.isChecked()
        
        
        self.export_type = "drawings"
        export_folder_parts = self.step_folder_edit.text()
        count = rows
        self.progress.setMaximum(count)
        self.progress.setValue(0)
        self.progress.setVisible(True)
        self.status.setText(f"Exporting {self.export_type}...")
        self.current_index = 0
        self.total_items = count
        data_table_parts = self.get_table_data(self.parts_table)
        for row_table in data_table_parts:
            i = self.parts_table_HeaderLabels.index("File Name")
            part_name = row_table[i]
            i = self.parts_table_HeaderLabels.index("File Path")
            part_path = row_table[i]
            i = self.parts_table_HeaderLabels.index("config")
            part_config = row_table[i]
            
            if part_config == "All":
                selected_configs = None
            else:
                selected_configs = [part_config]
                
            export_part_or_assembly_configurations_to_step(
                sw_app = self.sw_app,
                part_path = part_path,
                export_folder = export_folder_parts,
                selected_configs=selected_configs
                )
            self.current_index += 1
            self.progress.setValue(self.current_index-1)
            print(f"Exporting {self.export_type}: {self.current_index}/{self.total_items}")
        self.progress.setVisible(False)
        self.status.setText(f"{self.export_type.capitalize()} export completed")
        QMessageBox.information(self, "Export Completed", f"{self.export_type.capitalize()} export process has finished.")
    
    
        
        
    # def run_export_simulation(self, count, export_type):
    #     self.progress.setMaximum(count)
    #     self.progress.setValue(0)
    #     self.progress.setVisible(True)
    #     self.status.setText(f"Exporting {export_type}...")

    #     self.current_index = 0
    #     self.total_items = count
    #     self.export_type = export_type

    #     self.timer = QTimer()
    #     self.timer.timeout.connect(self.perform_export_step)
    #     self.timer.start(500)

    # def perform_export_step(self):
    #     self.current_index += 1
    #     if self.current_index > self.total_items:
    #         self.timer.stop()
    #         self.progress.setVisible(False)
    #         self.status.setText(f"{self.export_type.capitalize()} export completed")
    #         QMessageBox.information(self, "Export Completed", f"{self.export_type.capitalize()} export process has finished.")
    #     else:
    #         self.progress.setValue(self.current_index)
    #         #print(f"Exporting {self.export_type}: {self.current_index}/{self.total_items}")
            
    
    def start_SLDWRK(self):
        if not(self.sw_app):
            self.status.setText("Starting SolidWorks...")
            self.sw_app = win32com.client.Dispatch('SldWorks.Application')
    
    def get_drawings_table1_data(self):
        data = []
        for row in range(self.drawings_table.rowCount()):
            row_data = []
            for col in range(self.drawings_table.columnCount()):
                item = self.drawings_table.item(row, col)
                row_data.append(item.text() if item else "")
            data.append(row_data)
        return data
    
    def get_table_data(self,qtTable):
        data = []
        for row in range(qtTable.rowCount()):
            row_data = []
            for col in range(qtTable.columnCount()):
                combo = qtTable.cellWidget(row, col)  # Get the widget
                if combo is not None and isinstance(combo, QComboBox):
                    options = [combo.itemText(c) for c in range(combo.count())]
                    value = {'value':combo.currentText(),'options':options}
                else:
                    item = qtTable.item(row, col)
                    value = item.text()

                row_data.append(value if item else "")
                    
            data.append(row_data)
        return data
    def get_TableWidget_horizontalHeaderItem(self,qtTable):
        headers = []
        for col in range(qtTable.columnCount()):
            header_item = qtTable.horizontalHeaderItem(col)
            if header_item is not None:
                headers.append(header_item.text())
        return headers
    
    def get_table_data_JSON(self,qtTable):
        headers = self.get_TableWidget_horizontalHeaderItem(qtTable)
        data = []
        for row in range(qtTable.rowCount()):
            row_data = {}
            for col in range(qtTable.columnCount()):
                header_text = headers[col]
                combo = qtTable.cellWidget(row, col)  # Get the widget
                if combo is not None and isinstance(combo, QComboBox):
                    options = [combo.itemText(c) for c in range(combo.count())]
                    value = combo.currentText()
                    type_value = str(type(QComboBox))
                else:
                    item = qtTable.item(row, col)
                    value = item.text()
                    options = []
                    type_value = type('')
                row_data.update( { header_text : {'value':value,'options':options,'type':str(type_value) } })
            data.append(row_data)
        return data
    
    def load_table_from_JSON_setting(self,settingsJSON,qtTable,field_name):
        qtTable.setRowCount(0)
        for row_data in settingsJSON.get(field_name, []):
            row = qtTable.rowCount()
            qtTable.insertRow(row)
            for col, value in enumerate(row_data):
                qtTable.setItem(row, col, QTableWidgetItem(value))
                
    def load_table_from_JSON_setting2(self,settingsJSON,qtTable,field_name):
        headers = self.get_TableWidget_horizontalHeaderItem(qtTable)
        qtTable.setRowCount(0)
        for row_data in settingsJSON.get(field_name, []):
            if isinstance(row_data,dict):
                row = qtTable.rowCount()
                qtTable.insertRow(row)
                for header, cell_content in row_data.items():
                    if header in headers:
                        col = headers.index(header)
                        value = cell_content.get('value','')
                        
                        type_value = cell_content.get('type','')
                        
                        if str(type(QComboBox)) == type_value:
                            options = cell_content.get('options',[])
                            self.add_QComboBox_cell(qtTable, col,row,configs = options,value = value)
                        else:
                            qtTable.setItem(row, col, QTableWidgetItem(value))
    
    def retrive_SLDDRW_properties(self):
        rows = self.drawings_table.rowCount()
        if rows == 0:
            QMessageBox.warning(self, "Warning", "No drawings selected retriving properties.")
            return
        
        self.slddrw_properties = {}
        slddrw_files = self.get_drawings_table1_data()
        #slddrw_files_path = [drw[1] for drw in slddrw_files ]
        #self.sw_app = win32com.client.Dispatch('SldWorks.Application')
        if len(slddrw_files)>0:
            self.start_SLDWRK()
        self.sw_app.Visible = self.SLDWRK_visible_checkbox.isChecked()
        
        progress_bar_length = self.drawings_table.rowCount()
        self.progress.setMaximum(progress_bar_length)
        self.progress.setValue(0)
        self.progress.setVisible(True)
        self.status.setText("Retriving properties from drawings...")
        self.current_index = 0

        
        for drawing_name,drawing_path in slddrw_files:
            print(f"Processing: {drawing_path}")
            props = get_properties_from_drawing(self.sw_app, drawing_path)
            self.slddrw_properties[os.path.basename(drawing_path)] = props
            
            self.update_SLDDRW_prop_table_from_Dict()
            
            self.current_index +=1
            self.progress.setValue(self.current_index)
            self.status.setText(f"Properties drawing {drawing_name} retriving completed")
            
        self.progress.setVisible(False)
        self.status.setText("Properties drawings retriving completed")
        QMessageBox.information(self, "Properties retriving Completed", "Properties drawings retriving process has finished.")
        
        
    def update_SLDDRW_prop_table_from_Dict(self):
        # for filename, props in self.slddrw_properties.items():
            
        #     print(f"\nProperties for {filename}:")
        #     for k, v in props.items():
        #         print(f"  {k}: {v['value']}")
        
        columns = properties_columns_all
        self.drawings_prop_table.setRowCount(len(self.slddrw_properties))
        for row, (filename, props) in enumerate(self.slddrw_properties.items()):
            self.drawings_prop_table.setItem(row, 0, QTableWidgetItem(filename))
            for col, prop in enumerate(columns, start=1):
                value = props.get(prop, {}).get("value", "")
                self.drawings_prop_table.setItem(row, col, QTableWidgetItem(value))

    def sync_drawings_prop_table_to_dict(self):
        """
        Updates self.slddrw_properties with the latest values from self.drawings_prop_table.
        Only 'value' field is updated; 'type' remains unchanged.
        """
        row_count = self.drawings_prop_table.rowCount()
        col_count = self.drawings_prop_table.columnCount()
        columns = properties_columns_all  # already imported
        for row in range(row_count):
            file_name_item = self.drawings_prop_table.item(row, 0)
            if not file_name_item:
                continue
            file_name = file_name_item.text()
            # Ensure the file exists in dict
            if file_name not in self.slddrw_properties:
                self.slddrw_properties[file_name] = {}
            for col in range(1, col_count):
                prop_name = columns[col-1]
                item = self.drawings_prop_table.item(row, col)
                value = item.text() if item else ""
                # Keep type if exists, else use CustomProperty
                old_type = self.slddrw_properties[file_name].get(prop_name, {}).get("type", "CustomProperty")
                self.slddrw_properties[file_name][prop_name] = {"value": value, "type": old_type}
                
    def sync_drawings_table1_and_2(self):
        slddrw_files = self.get_drawings_table1_data()
        
        slddrw_files_table1_set = {drw[0] for drw in slddrw_files }
        slddrw_propertiesJSON_set = set(self.slddrw_properties.keys())
        slddrw_files_deleted_set= slddrw_propertiesJSON_set - slddrw_files_table1_set
        
        slddrw_files_newAdded_set=  slddrw_files_table1_set - slddrw_propertiesJSON_set
        
        flag_some_row_deleted = False
        for drawing_name in list(slddrw_files_deleted_set):
                del self.slddrw_properties[drawing_name]
                flag_some_row_deleted = True
        
        flag_some_row_newAdded = False
        for drawing_name in list(slddrw_files_newAdded_set):
            flag_some_row_newAdded = True
            self.slddrw_properties.update({drawing_name:{}})
            
            
        if flag_some_row_deleted or flag_some_row_newAdded:
            self.update_SLDDRW_prop_table_from_Dict()
        return
    
    def export_table_to_excel(self):
        # Let user choose where to save
        file_path, _ = QFileDialog.getSaveFileName(self, "Save Excel File", "drawings_properties", "Excel Files (*.xlsx)")
        if not file_path:
            return
        headers = [self.drawings_prop_table.horizontalHeaderItem(i).text() for i in range(self.drawings_prop_table.columnCount())]
        data = []
        for row in range(self.drawings_prop_table.rowCount()):
            row_data = []
            for col in range(self.drawings_prop_table.columnCount()):
                item = self.drawings_prop_table.item(row, col)
                row_data.append(item.text() if item else "")
            data.append(row_data)
        df = pd.DataFrame(data, columns=headers)
        df.to_excel(file_path, index=False)
        
    def import_table_from_excel(self):
        file_path, _ = QFileDialog.getOpenFileName(self, "Open Excel File", "", "Excel Files (*.xlsx *.xls)")
        if not file_path:
            return
        df = pd.read_excel(file_path)
        df = df.fillna("")  # Replace NaN with empty string
        # Clear existing table
        self.drawings_prop_table.setRowCount(0)
        # Optionally reset columns if headers change
        self.drawings_prop_table.setColumnCount(len(df.columns))
        self.drawings_prop_table.setHorizontalHeaderLabels(list(df.columns))
        # Fill table
        for i, row in df.iterrows():
            self.drawings_prop_table.insertRow(i)
            for j, val in enumerate(row):
                item = QTableWidgetItem(str(val))
                self.drawings_prop_table.setItem(i, j, item)
        
        self.sync_drawings_prop_table_to_dict()
        self.sync_drawings_table1_and_2()



    def save_settings(self):
        settings = {
            "dwg_folder": self.dwg_folder_edit.text(),
            "pdf_folder": self.pdf_folder_edit.text(),
            "step_folder": self.step_folder_edit.text(),
            "flag_SLDWRK_visible" : self.SLDWRK_visible_checkbox.isChecked(),
            "flag_SLDWRK__close_expDRW" : self.SLDWRK_close_expDRW_checkbox.isChecked(),
            "export_dwg": self.dwg_checkbox.isChecked(),
            "export_pdf": self.pdf_checkbox.isChecked(),
            "flag_indiv_dwg": self.dwg_indiv_checkbox.isChecked(),
            "flag_indiv_pdf": self.pdf_indiv_checkbox.isChecked(),
            "drawings": self.get_drawings_table1_data(),
            "drawings_properties": self.slddrw_properties,
            "parts": self.get_table_data_JSON(self.parts_table),
        }
        file_path, _ = QFileDialog.getSaveFileName(
            self,                             # parent widget (can be None if not in a class)
            "Save Settings",                  # dialog title
            "",                               # default directory or file
            "JSON Files (*.json);;All Files (*)"  # filter for file types
        )
        if file_path:
            with open(file_path, "w") as file:
                json.dump(settings, file, indent=4)
    
    def load_settings(self):
        file_path, _ = QFileDialog.getOpenFileName(
            self,                                # parent widget (use None if not in a class)
            "Open Settings",                     # dialog title
            "",                                  # initial directory (empty means current)
            "JSON Files (*.json);;All Files (*)" # filter for file types
        )
        if file_path:
            with open(file_path, "r") as file:
                settings = json.load(file)
                self.dwg_folder_edit.setText(settings.get("dwg_folder", ""))
                self.pdf_folder_edit.setText(settings.get("pdf_folder", ""))
                self.step_folder_edit.setText(settings.get("step_folder", ""))
                self.dwg_checkbox.setChecked(settings.get("export_dwg" , False))
                self.pdf_checkbox.setChecked(settings.get("export_pdf" , False)) 
                self.SLDWRK_visible_checkbox.setChecked(settings.get("flag_SLDWRK_visible" , False)) 
                self.SLDWRK_close_expDRW_checkbox.setChecked(settings.get("flag_SLDWRK__close_expDRW" , True)) 
                self.dwg_indiv_checkbox.setChecked(settings.get("flag_indiv_dwg", False))
                self.pdf_indiv_checkbox.setChecked(settings.get("flag_indiv_pdf", False))
                # self.drawings_table.setRowCount(0)
                # for row_data in settings.get("drawings", []):
                #     row = self.drawings_table.rowCount()
                #     self.drawings_table.insertRow(row)
                #     for col, value in enumerate(row_data):
                #         self.drawings_table.setItem(row, col, QTableWidgetItem(value))
                self.load_table_from_JSON_setting(settings,self.drawings_table,"drawings")
                self.slddrw_properties = settings.get("drawings_properties",{})
                self.update_SLDDRW_prop_table_from_Dict()
                self.sync_drawings_table1_and_2()
                
                self.load_table_from_JSON_setting2(settings,self.parts_table,"parts")
                
                # i = self.parts_table_HeaderLabels.index("config")
                # self.add_QComboBox_column(self.parts_table,i,configs = ["All"])

                

        
    def on_cell_changed_drawings_prop_table(self):
        #print('tony')
        #self.sync_drawings_prop_table_to_dict()
        return                    
    
    def update_into_SLDDRW_files(self):
        self.sync_drawings_prop_table_to_dict()
        rows = self.drawings_prop_table.rowCount()
        if rows == 0:
            QMessageBox.warning(self, "Warning", "No drawings selected for sync. properties.")
            return
        #print(self.slddrw_properties)
        slddrw_files = self.get_drawings_table1_data()
        if len(slddrw_files)>0:
            self.start_SLDWRK()
            self.status.setText("Preparing sync. properties...")
            
        progress_bar_length = self.drawings_table.rowCount()
        self.progress.setMaximum(progress_bar_length)
        self.progress.setValue(0)
        self.progress.setVisible(True)
        self.status.setText("Sync. properties from table to drawings...")
        self.current_index = 0
        for drawing_name, drawing_path in slddrw_files:
            props = self.slddrw_properties[drawing_name]
            update_drawing_property( self.sw_app, drawing_path, props ,flagRebuild = False)

            self.current_index +=1
            self.progress.setValue(self.current_index)
            self.status.setText(f"Properties drawing {drawing_name} retriving completed")
            
        self.progress.setVisible(False)
        self.status.setText("Sync. properties drawings completed")
        QMessageBox.information(self, "Sync. properties drawings completed", "Sync. properties drawings completed process has finished.")
        
        return 0
    
    
    
if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = SolidWorksExportManager()
    window.show()
    sys.exit(app.exec())
