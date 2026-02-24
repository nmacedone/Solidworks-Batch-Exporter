# NeuralFab Batch Exporter

A powerful, standalone desktop application built with Python and PySide6 that automates the process of generating and exporting multiple geometric configurations of a SolidWorks part. 

Instead of manually tweaking dimensions, rebuilding, and saving out individual STEP, IGES, or STL files, this tool allows you to define a table of dimensions and autonomous handles the batch modification and export process in the background.

## Features
* **Dynamic Dimension Extraction**: Uses a native VBA macro to reliably traverse the SolidWorks feature tree and pull every configurable dimension into a user-friendly dropdown.
* **Batch Processing Grid**: Build a spreadsheet-like configuration table directly in the app. Define custom output filenames and dimension sets for each configuration.
* **Background Execution**: SolidWorks runs minimized in your taskbar. The application uses a multithreaded background worker so the UI remains completely responsive during heavy CAD operations.
* **Auto-Organized Outputs**: Automatically creates a clean `[partname]_batch_exports` subfolder in your chosen directory, and backs up the unmodified original part before processing your custom configurations.
* **Live Console Logging**: Built-in terminal readout to monitor exactly what the SolidWorks API is doing in real-time, making error tracking incredibly simple.

---

## Prerequisites
* **OS:** Windows 10 or 11
* **CAD:** SolidWorks (Installed and licensed on the host machine)
* **Python:** Python 3.9 or newer (Python 3.13 supported)

## Installation

1. **Clone or Download the Repository** Ensure all three primary files are located in the exact same directory:
   * `main.py`
   * `sw_controller.py`
   * `GetDimensions.swp` *(Critical: This macro must be present for dimension extraction)*

2. **Install Python Dependencies**
   Open your terminal/command prompt in the project directory and install the required packages:
   ```bash
   pip install -r requirements.txt
Note: This will install PySide6 for the GUI and pywin32 for the SolidWorks COM interface.

How to Use
1. Launch the Application
Run the main Python script from your terminal:

Bash
python main.py
2. Load Your Part
Click Load Part (.SLDPRT) and select your SolidWorks file.

The tool will briefly communicate with SolidWorks and run the GetDimensions.swp macro.

Once successful, the console will confirm how many dimensions were extracted.

3. Build Your Configuration Table
Click + Add Dimension Column for every dimension parameter you want to change.

In the newly added column(s), click the dropdown header (Row 0) and select the specific dimension (e.g., D1@Boss-Extrude1@Part1.Part).

Click + Add Config Row for every unique file you want to generate.

For each row:

Type a custom suffix in the Filename column (e.g., config_1, thick_base).

Enter the desired numerical values (in millimeters) for each dimension column.

4. Select Output & Format
Choose your desired export format (STEP, IGES, or STL) from the dropdown at the top.

(Optional) Change the Output Folder. By default, it auto-selects the folder where your .SLDPRT file lives.

5. Calculate Configurations
Click the green Calculate configurations button.

SolidWorks will open minimized in your taskbar.

Watch the live console output as the tool autonomously changes the dimensions, rebuilds the model, and exports each file into a newly created [PartName]_batch_exports folder.

Troubleshooting
"Failed to load dimensions. Check SolidWorks connection."

Ensure SolidWorks is not currently displaying a blocking dialog box (like a missing font warning or a rebuild error).

Verify that GetDimensions.swp is in the exact same folder as main.py.

"Write access denied" or "Part is being used"

If the script crashes or is force-closed midway through a batch, a "zombie" version of SolidWorks might get stuck running invisibly in the background holding onto your file locks.

Fix: Press Ctrl + Shift + Esc to open Windows Task Manager, find SLDWORKS.exe, and click End Task.

Dimensions aren't changing the model

Ensure your input values are in millimeters.

Ensure the dimension you selected in the dropdown actually drives the geometry you want to change (avoid selecting read-only reference dimensions).