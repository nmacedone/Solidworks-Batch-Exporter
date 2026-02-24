# Vesrion 2.0.0.0 Improves upon previous version by:
# - Adding a dedicated subfolder for exports based on the part name

import sys
import os
from pathlib import Path
from PySide6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, 
    QPushButton, QLabel, QLineEdit, QFileDialog, QTableWidget, 
    QTableWidgetItem, QComboBox, QMessageBox, QHeaderView, QPlainTextEdit
)
from PySide6.QtCore import Qt, QThread, Signal
from sw_controller import SolidWorksController

class BatchExportWorker(QThread):
    """Background worker to handle SW operations without freezing the UI."""
    progress_update = Signal(int, str)  # row_index, status_text
    log_update = Signal(str)            # verbose logging
    finished = Signal()

    def __init__(self, model_path, output_dir, export_format, configurations):
        super().__init__()
        self.model_path = model_path
        self.output_dir = output_dir
        self.export_format = export_format
        self.configurations = configurations  # List of dicts: {'row': int, 'filename': str, 'dims': {name: val}}

    def run(self):
        self.log_update.emit("Initializing SolidWorks Controller...")
        sw = SolidWorksController()
        
        if not sw.connect(log_callback=self.log_update.emit):
            self.progress_update.emit(-1, "Failed to connect to SolidWorks.")
            self.finished.emit()
            return
            
        self.log_update.emit(f"Opening document: {self.model_path}")
        if not sw.open_document(self.model_path, log_callback=self.log_update.emit):
            self.progress_update.emit(-1, "Failed to open document.")
            self.log_update.emit("ERROR: Failed to open document.")
            sw.close()
            self.finished.emit()
            return

        try:
            self.log_update.emit("SolidWorks ready. Starting batch export...")

            # --- Create dedicated subfolder based on part name ---
            part_name = Path(self.model_path).stem
            batch_folder = Path(self.output_dir) / f"{part_name}_batch_exports"
            batch_folder.mkdir(parents=True, exist_ok=True)
            self.log_update.emit(f"Created export directory: {batch_folder}")

            # --- Export the original unmodified part first ---
            orig_filename = f"{part_name}_original.{self.export_format}"
            orig_out_file = batch_folder / orig_filename
            self.log_update.emit(f"\n--- Exporting original state to: {orig_out_file} ---")
            
            success_orig = sw.export_file(orig_out_file, log_callback=self.log_update.emit)
            if success_orig:
                self.log_update.emit(f"SUCCESS: Saved {orig_filename}")
            else:
                self.log_update.emit(f"ERROR: Failed to save {orig_filename}")

            for config in self.configurations:
                row = config['row']
                filename = config['filename']
                dims = config['dims']
                
                self.progress_update.emit(row, "Processing...")
                
                # Combine part name and user-defined config name
                full_filename = f"{part_name}_{filename}"
                self.log_update.emit(f"\n--- Processing Row {row} | File: {full_filename} ---")
                
                # Apply dimensions
                for dim_name, val in dims.items():
                    self.log_update.emit(f"Setting {dim_name} = {val}")
                    sw.modify_dimension(dim_name, val)
                    
                self.log_update.emit("Rebuilding model...")
                sw.rebuild()
                
                # Construct output path inside the new subfolder
                out_file = batch_folder / f"{full_filename}.{self.export_format}"
                
                # Export
                self.log_update.emit(f"Exporting to: {out_file}")
                success = sw.export_file(out_file, log_callback=self.log_update.emit)
                
                if success:
                    self.progress_update.emit(row, "✓ Saved")
                    self.log_update.emit(f"SUCCESS: Saved {full_filename}.{self.export_format}")
                else:
                    self.progress_update.emit(row, "✗ Error")
                    self.log_update.emit(f"ERROR: Failed to save {full_filename}.{self.export_format}")

            self.log_update.emit("\nBatch export completed. Closing document...")
            
        except Exception as e:
            self.log_update.emit(f"\nFATAL THREAD ERROR: {e}")
            self.progress_update.emit(-1, "Worker crashed.")
        finally:
            # THIS GUARANTEES the invisible SolidWorks drops the file lock no matter what happens
            sw.close()
            self.finished.emit()

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("NeuralFab Batch Exporter")
        self.resize(1000, 600)

        self.available_dimensions = []
        self.dimension_columns = []  # Tracks which column index holds a dimension

        self.init_ui()

    def init_ui(self):
        main_widget = QWidget()
        self.setCentralWidget(main_widget)
        layout = QVBoxLayout(main_widget)

        # --- Top Controls ---
        top_layout = QHBoxLayout()
        
        # File Selection
        self.btn_load_part = QPushButton("Load Part (.SLDPRT)")
        self.btn_load_part.clicked.connect(self.load_part)
        self.line_part_path = QLineEdit()
        self.line_part_path.setReadOnly(True)
        
        top_layout.addWidget(self.btn_load_part)
        top_layout.addWidget(self.line_part_path)
        
        # Output Folder Selection
        self.btn_out_folder = QPushButton("Output Folder")
        self.btn_out_folder.clicked.connect(self.select_output_folder)
        self.line_out_folder = QLineEdit()
        self.line_out_folder.setReadOnly(True)
        
        top_layout.addWidget(self.btn_out_folder)
        top_layout.addWidget(self.line_out_folder)
        
        # Format Selection
        self.combo_format = QComboBox()
        self.combo_format.addItems(["STEP", "IGES", "STL"])
        top_layout.addWidget(QLabel("Format:"))
        top_layout.addWidget(self.combo_format)

        layout.addLayout(top_layout)

        # --- Table Controls ---
        table_ctrl_layout = QHBoxLayout()
        
        self.btn_add_dim = QPushButton("+ Add Dimension Column")
        self.btn_add_dim.clicked.connect(self.add_dimension_column)
        self.btn_add_dim.setEnabled(False)  # Disabled until part is loaded
        
        self.btn_add_row = QPushButton("+ Add Config Row")
        self.btn_add_row.clicked.connect(self.add_config_row)
        
        table_ctrl_layout.addWidget(self.btn_add_dim)
        table_ctrl_layout.addWidget(self.btn_add_row)
        table_ctrl_layout.addStretch()
        
        layout.addLayout(table_ctrl_layout)

        # --- Table ---
        self.table = QTableWidget(1, 2)  # 1 row (header), 2 cols (Status, Filename)
        self.table.horizontalHeader().setVisible(False)  # Hide default header, we use Row 0
        
        # Setup our Row 0 "Header"
        status_item = QTableWidgetItem("Status")
        status_item.setFlags(Qt.ItemIsEnabled)
        filename_item = QTableWidgetItem("Filename")
        filename_item.setFlags(Qt.ItemIsEnabled)
        
        self.table.setItem(0, 0, status_item)
        self.table.setItem(0, 1, filename_item)
        
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        layout.addWidget(self.table)

        # Add initial blank config row
        self.add_config_row()

        # --- Bottom Controls ---
        bottom_layout = QHBoxLayout()
        self.btn_calculate = QPushButton("Calculate configurations")
        self.btn_calculate.setStyleSheet("background-color: #2e7d32; color: white; font-weight: bold; padding: 10px;")
        self.btn_calculate.clicked.connect(self.start_calculation)
        bottom_layout.addStretch()
        bottom_layout.addWidget(self.btn_calculate)
        
        layout.addLayout(bottom_layout)

        # --- Console Output ---
        self.console_output = QPlainTextEdit()
        self.console_output.setReadOnly(True)
        self.console_output.setStyleSheet("background-color: #1e1e1e; color: #4af626; font-family: Consolas, monospace;")
        self.console_output.setMaximumHeight(150)
        layout.addWidget(self.console_output)
        self.append_log("System initialized. Ready.")

    def load_part(self):
        file_name, _ = QFileDialog.getOpenFileName(self, "Select SolidWorks Part", "", "SolidWorks Parts (*.SLDPRT)")
        if file_name:
            self.line_part_path.setText(file_name)
            
            # Auto-set output folder to same directory if empty
            if not self.line_out_folder.text():
                self.line_out_folder.setText(str(Path(file_name).parent))
                
            self.fetch_dimensions(file_name)

    def select_output_folder(self):
        folder = QFileDialog.getExistingDirectory(self, "Select Output Directory")
        if folder:
            self.line_out_folder.setText(folder)

    def append_log(self, text):
        """Appends text to the console output and scrolls to the bottom."""
        self.console_output.appendPlainText(text)
        scrollbar = self.console_output.verticalScrollBar()
        scrollbar.setValue(scrollbar.maximum())

    def fetch_dimensions(self, file_path):
        """Temporarily opens SolidWorks to grab available dimensions for the comboboxes."""
        self.btn_load_part.setText("Loading dimensions...")
        self.append_log(f"Connecting to SolidWorks to load part: {file_path}")
        QApplication.processEvents()
        
        # A quick local function to force the UI to repaint on every log message
        def live_log(msg):
            self.append_log(msg)
            QApplication.processEvents()
            
        sw = SolidWorksController()
        if sw.connect(log_callback=live_log) and sw.open_document(file_path, log_callback=live_log):
            try:
                self.append_log("Part loaded successfully. Extracting dimensions...")
                self.available_dimensions = sw.get_all_dimensions(log_callback=live_log)
                self.btn_add_dim.setEnabled(True)
                self.btn_load_part.setText("Load Part (.SLDPRT)")
                self.append_log(f"Extracted {len(self.available_dimensions)} dimensions.")
                QMessageBox.information(self, "Success", f"Loaded {len(self.available_dimensions)} dimensions.")
            except Exception as e:
                self.append_log(f"ERROR: Extraction crashed: {e}")
            finally:
                # Guarantee document closure even if the macro crashes
                sw.close()
        else:
            self.append_log("ERROR: Failed to load dimensions. Check SolidWorks connection.")
            self.btn_load_part.setText("Load Part (.SLDPRT)")
            QMessageBox.critical(self, "Error", "Failed to load dimensions from SolidWorks.")

    def add_dimension_column(self):
        if not self.available_dimensions:
            return
            
        col_idx = self.table.columnCount()
        self.table.insertColumn(col_idx)
        self.dimension_columns.append(col_idx)
        
        # Add Combobox to Row 0
        combo = QComboBox()
        combo.addItems(["-- Select Dimension --"] + self.available_dimensions)
        self.table.setCellWidget(0, col_idx, combo)
        
        # Add empty items to remaining rows
        for row in range(1, self.table.rowCount()):
            self.table.setItem(row, col_idx, QTableWidgetItem(""))

    def add_config_row(self):
        row_idx = self.table.rowCount()
        self.table.insertRow(row_idx)
        
        # Status Cell
        status_item = QTableWidgetItem("-")
        status_item.setFlags(Qt.ItemIsEnabled)
        status_item.setTextAlignment(Qt.AlignCenter)
        self.table.setItem(row_idx, 0, status_item)
        
        # Filename cell
        self.table.setItem(row_idx, 1, QTableWidgetItem(f"config_{row_idx}"))
        
        # Blank cells for dimensions
        for col in self.dimension_columns:
            self.table.setItem(row_idx, col, QTableWidgetItem(""))

    def start_calculation(self):
        if not self.line_part_path.text() or not self.line_out_folder.text():
            QMessageBox.warning(self, "Missing Info", "Please select a part file and output folder.")
            return

        if len(self.dimension_columns) == 0:
            QMessageBox.warning(self, "Missing Info", "Please add at least one dimension column.")
            return

        # Parse configurations from table
        configs = []
        for row in range(1, self.table.rowCount()):
            filename_item = self.table.item(row, 1)
            filename = filename_item.text() if filename_item else f"config_{row}"
            
            dims = {}
            valid_row = True
            for col in self.dimension_columns:
                combo = self.table.cellWidget(0, col)
                dim_name = combo.currentText()
                
                if dim_name == "-- Select Dimension --":
                    continue
                    
                val_item = self.table.item(row, col)
                if val_item and val_item.text().strip():
                    try:
                        dims[dim_name] = float(val_item.text().strip())
                    except ValueError:
                        self.table.item(row, 0).setText("✗ Invalid Num")
                        valid_row = False
                        break
            
            if valid_row and dims:
                configs.append({
                    'row': row,
                    'filename': filename,
                    'dims': dims
                })

        if not configs:
            QMessageBox.warning(self, "No valid data", "No valid configurations found to calculate.")
            return

        # Disable UI during calculation
        self.btn_calculate.setEnabled(False)
        self.btn_calculate.setText("Running...")

        # Start Worker
        fmt = self.combo_format.currentText().lower()
        self.append_log(f"\nStarting batch export for {len(configs)} configurations...")
        self.worker = BatchExportWorker(self.line_part_path.text(), self.line_out_folder.text(), fmt, configs)
        self.worker.progress_update.connect(self.update_status)
        self.worker.log_update.connect(self.append_log)
        self.worker.finished.connect(self.calculation_finished)
        self.worker.start()

    def update_status(self, row, status):
        if row == -1:
            QMessageBox.critical(self, "Error", status)
            return
            
        item = self.table.item(row, 0)
        if item:
            item.setText(status)

    def calculation_finished(self):
        self.append_log("Worker finished execution.")
        self.btn_calculate.setEnabled(True)
        self.btn_calculate.setText("Calculate configurations")
        QMessageBox.information(self, "Done", "Batch export completed!")

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec())