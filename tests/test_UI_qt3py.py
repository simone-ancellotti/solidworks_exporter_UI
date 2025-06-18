from PyQt5 import QtWidgets, QtCore
import sys
import win32com.client
import pythoncom
import os
import json

class SolidWorksExporterUI(QtWidgets.QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("SolidWorks Exporter")
        self.resize(700, 500)
        self.tabs = QtWidgets.QTabWidget()
        self.setCentralWidget(self.tabs)

        self.tab_drawings = QtWidgets.QWidget()
        self.tab_parts = QtWidgets.QWidget()
        self.tabs.addTab(self.tab_drawings, "Export Drawings")
        self.tabs.addTab(self.tab_parts, "Export STEP")

        self.init_drawings_tab()
        self.init_parts_tab()
        self.statusBar().showMessage("Ready.")

    def init_drawings_tab(self):
        layout = QtWidgets.QVBoxLayout()
        # --- Add widgets (folders, checkboxes, table, buttons, etc) ---
        # For example:
        self.dwg_folder_btn = QtWidgets.QPushButton("Select DWG Folder")
        self.pdf_folder_btn = QtWidgets.QPushButton("Select PDF Folder")
        # ...other widgets...
        self.export_btn = QtWidgets.QPushButton("Export")
        self.export_btn.clicked.connect(self.export_drawings)
        layout.addWidget(self.dwg_folder_btn)
        layout.addWidget(self.pdf_folder_btn)
        # ...etc...
        self.tab_drawings.setLayout(layout)

    def init_parts_tab(self):
        layout = QtWidgets.QVBoxLayout()
        # Add widgets for selecting parts, export folder, config selection, etc.
        self.tab_parts.setLayout(layout)

    def export_drawings(self):
        # Call your export logic
        pass

if __name__ == "__main__":
    app = QtWidgets.QApplication(sys.argv)
    window = SolidWorksExporterUI()
    window.show()
    sys.exit(app.exec_())
