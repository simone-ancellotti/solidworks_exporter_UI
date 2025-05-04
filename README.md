# SolidWorks Drawing Exporter UI

A user-friendly Python GUI built with `tkinter` that allows you to batch export SolidWorks drawings (`.SLDDRW`) to both PDF and DWG formats. It also supports exporting part/assembly files to STEP format, making it an efficient tool for engineers and technical teams working with multiple SolidWorks outputs.

## 🚀 Features

- Export multiple `.SLDDRW` drawings to:
  - PDF (single file or individual sheets)
  - DWG (single file or individual sheets)
- Export parts or assemblies (`.SLDPRT`/`.SLDASM`) to STEP format
- Select and manage files through an intuitive interface
- Save and load project settings to JSON
- Visual progress bar and status updates
- Built-in error messages and user confirmations

## 📦 Requirements

- Python 3.8+
- SolidWorks installed on the same machine
- Python packages:
  - `pywin32`
  - `tkinter` (usually comes with Python)
  
To install `pywin32`:
```bash
pip install pywin32

## 🚀 Features

- Export multiple `.SLDDRW` drawings to:
  - PDF (single file or individual sheets)
  - DWG (single file or individual sheets)
- Export parts or assemblies (`.SLDPRT`/`.SLDASM`) to STEP format
- Select and manage files through an intuitive interface
- Save and load project settings to JSON
- Visual progress bar and status updates
- Built-in error messages and user confirmations

## 📦 Requirements

- Python 3.8+
- SolidWorks installed on the same machine
- Python packages:
  - `pywin32`
  - `tkinter` (usually comes with Python)
  
To install `pywin32`:
```bash
pip install pywin32
🛠 How to Use
Run the script:

bash
Copy
python solidworks_exporter_UI.py
Choose export folders for DWG and PDF

Select the .SLDDRW drawings to export

Configure sheet export options and flags

Click Export to batch-process all files

Optionally save your setup for future reuse

💡 Example Use Case
Ideal for mechanical design teams needing to deliver:

Manufacturing drawings in DWG

Documentation in PDF

CAD exchange files in STEP

🗂 Folder Structure
bash
Copy
solidworks_exporter_UI/
│
├── solidworks_exporter_UI.py   # Main script
├── README.md                   # This file
├── requirements.txt            # Optional
├── saved_settings.json         # User-defined settings (optional)
