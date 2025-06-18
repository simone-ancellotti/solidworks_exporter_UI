import sys
import os
import json
import time
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QWidget, QLabel, QPushButton, QVBoxLayout,
    QHBoxLayout, QFileDialog, QLineEdit, QTabWidget, QCheckBox, QProgressBar,
    QTableWidget, QTableWidgetItem, QMessageBox, QHeaderView,QAction
)
from PyQt5.QtCore import Qt, QTimer
from ../solidworks_export import (
    open_and_rebuild_drawing,
    export_drawing_to_pdf,
    export_drawing_to_dwg
)

class SolidWorksExportManager(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("SolidWorks Export Manager")
        self.setMinimumWidth(800)

        self.dwg_folder = ""
        self.pdf_folder = ""
        self.step_folder = ""

        self.tabs = QTabWidget()
        self.setCentralWidget(self.tabs)

        self.status = QLabel("Ready")
        self.statusBar().addWidget(self.status)
        
        self.init_dwg_pdf_tab()
        self.init_step_tab()
        self.title_blocks_tab()
        
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
        
        
        

        self.pdf_check_boxes = QVBoxLayout()
        self.dwg_checkbox = QCheckBox("Export DWG")
        self.dwg_checkbox.setChecked(True)  # Default True
        
        self.pdf_checkbox = QCheckBox("Export PDF")
        self.pdf_checkbox.setChecked(True)
        self.pdf_indiv_checkbox = QCheckBox("Export individual PDF sheets")
        self.dwg_indiv_checkbox = QCheckBox("Export individual DWG sheets")
        layout.addLayout(self.make_check_2buttons_layout(
            self.pdf_checkbox, self.pdf_indiv_checkbox
            ))
        layout.addLayout(self.make_check_2buttons_layout(
            self.dwg_checkbox, self.dwg_indiv_checkbox
            ))
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

    def init_step_tab(self):
        tab = QWidget()
        layout = QVBoxLayout()

        # Drawing List
        btn_drawings = QPushButton("Upload title blocks")
        btn_drawings.clicked.connect(self.select_drawings)
        layout.addWidget(btn_drawings)

        tab.setLayout(layout)
        self.tabs.addTab(tab, "Title blocks")
        
    def title_blocks_tab(self):
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
        layout.addWidget(btn_parts)

        self.parts_table = QTableWidget(0, 2)
        self.parts_table.setHorizontalHeaderLabels(["File Name", "File Path"])
        self.parts_table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        layout.addWidget(self.parts_table)

        btn_export_parts = QPushButton("Export Parts")
        btn_export_parts.clicked.connect(self.export_parts)
        layout.addWidget(btn_export_parts)

        tab.setLayout(layout)
        self.tabs.addTab(tab, "STEP Export")

    def load_settings(self):
        print("Load settings clicked")

    def save_settings(self):
        print("Save settings clicked")
        
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

    def select_parts(self):
        files, _ = QFileDialog.getOpenFileNames(self, "Select Parts", "", "SolidWorks Parts (*.SLDPRT)")
        for file in files:
            filename = os.path.basename(file)
            self.parts_table.insertRow(self.parts_table.rowCount())
            self.parts_table.setItem(self.parts_table.rowCount()-1, 0, QTableWidgetItem(filename))
            self.parts_table.setItem(self.parts_table.rowCount()-1, 1, QTableWidgetItem(file))
        self.status.setText("Parts selected")

    def export_drawings(self):
        rows = self.drawings_table.rowCount()
        if rows == 0:
            QMessageBox.warning(self, "Export Warning", "No drawings selected for export.")
            return
        self.run_export_simulation(rows, "drawings")

    def delete_selected_drawings(self):
        selected = self.drawings_table.selectionModel().selectedRows()
        for index in sorted(selected, reverse=True):
            self.drawings_table.removeRow(index.row())
        self.status.setText("Selected drawings deleted")

    def export_parts(self):
        rows = self.parts_table.rowCount()
        if rows == 0:
            QMessageBox.warning(self, "Export Warning", "No parts selected for STEP export.")
            return
        self.run_export_simulation(rows, "parts")
        
        

    def run_export_simulation(self, count, export_type):
        self.progress.setMaximum(count)
        self.progress.setValue(0)
        self.progress.setVisible(True)
        self.status.setText(f"Exporting {export_type}...")

        self.current_index = 0
        self.total_items = count
        self.export_type = export_type

        self.timer = QTimer()
        self.timer.timeout.connect(self.perform_export_step)
        self.timer.start(500)

    def perform_export_step(self):
        self.current_index += 1
        if self.current_index > self.total_items:
            self.timer.stop()
            self.progress.setVisible(False)
            self.status.setText(f"{self.export_type.capitalize()} export completed")
            QMessageBox.information(self, "Export Completed", f"{self.export_type.capitalize()} export process has finished.")
        else:
            self.progress.setValue(self.current_index)
            print(f"Exporting {self.export_type}: {self.current_index}/{self.total_items}")


    # def save_settings(self):
    #     settings = {
    #         "dwg_folder": self.dwg_folder.get(),
    #         "pdf_folder": self.pdf_folder.get(),
    #         "export_dwg": dwg_var.get(),
    #         "export_pdf": pdf_var.get(),
    #         "flag_export_dwg": flag_export_dwg.get(),
    #         "flag_export_pdf": flag_export_pdf.get(),
    #         "drawings": [drawings_list.item(item, "values") for item in drawings_list.get_children()]
    #     }
    #     file_path = filedialog.asksaveasfilename(defaultextension=".json", filetypes=[("JSON Files", "*.json")])
    #     if file_path:
    #         with open(file_path, "w") as file:
    #             json.dump(settings, file, indent=4)
    
    # def load_settings():
    #     file_path = filedialog.askopenfilename(filetypes=[("JSON Files", "*.json")])
    #     if file_path:
    #         with open(file_path, "r") as file:
    #             settings = json.load(file)
    #             dwg_folder_var.set(settings.get("dwg_folder", ""))
    #             pdf_folder_var.set(settings.get("pdf_folder", ""))
    #             dwg_var.set(settings.get("export_dwg", False))
    #             pdf_var.set(settings.get("export_pdf", False))
    #             flag_export_dwg.set(settings.get("flag_export_dwg", True))
    #             flag_export_pdf.set(settings.get("flag_export_pdf", True))
    #             drawings_list.delete(*drawings_list.get_children())
    #             for drawing in settings.get("drawings", []):
    #                 drawings_list.insert("", "end", values=drawing)
if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = SolidWorksExportManager()
    window.show()
    sys.exit(app.exec())
