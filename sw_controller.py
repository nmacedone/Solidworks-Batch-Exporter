import win32com.client as win32
import pythoncom
import os
import time
from pathlib import Path

class SolidWorksController:
    def __init__(self):
        self.sw_app = None
        self.sw_model = None

    def connect(self, log_callback=None):
        """Initializes the COM connection to SolidWorks."""
        pythoncom.CoInitialize()
        try:
            if log_callback: log_callback("Attempting to connect to SolidWorks...")
            self.sw_app = win32.Dispatch("SldWorks.Application")
            
            # CRITICAL EXPORT FIX: SolidWorks MUST be visible and UserControl=True 
            # for export translators (STEP/STL) to load. If hidden, exports crash.
            self.sw_app.UserControl = True
            self.sw_app.Visible = True
            
            # To keep it out of the user's way, we instantly minimize the window to the taskbar.
            try:
                self.sw_app.FrameState = 1 # 1 = swWindowMinimized
            except:
                pass
            
            if log_callback: log_callback("Connected successfully to SolidWorks (Minimized Mode).")
            return True
        except Exception as e:
            if log_callback: log_callback(f"ERROR: Connection failure: {e}")
            return False

    def open_document(self, path, log_callback=None):
        """Opens a SolidWorks part document."""
        if not os.path.exists(path):
            if log_callback: log_callback(f"ERROR: File does not exist at {path}")
            return False
            
        try:
            # We open the document normally. (Hiding the document itself causes STEP 
            # exports to crash because the exporter cannot read the graphics body).
            self.sw_model = self.sw_app.OpenDoc(str(path), 1) # 1 = swDocPART
                
            if self.sw_model is None:
                if log_callback: log_callback("ERROR: OpenDoc failed. File may be corrupt or already open.")
                return False
            return True
        except Exception as e:
            if log_callback: log_callback(f"ERROR: Exception during open_document: {e}")
            return False

    def get_all_dimensions(self, log_callback=None):
        """Uses a VBA macro to robustly extract dimensions bypassing Python COM limits."""
        if not self.sw_model:
            return []
        
        macro_path = Path(__file__).parent / "GetDimensions.swp"
        temp_file = Path(r"C:\temp\sw_dimensions.txt")
        target_file = Path(r"C:\temp\sw_target_file.txt")
        
        if not macro_path.exists():
            if log_callback: log_callback(f"ERROR: Macro not found at {macro_path}")
            return []
            
        if log_callback: log_callback("Executing GetDimensions.swp to extract dimensions natively...")
        
        # Ensure temp dir exists and clear old files
        temp_file.parent.mkdir(parents=True, exist_ok=True)
        for file in [temp_file, target_file]:
            if file.exists():
                try:
                    file.unlink()
                except:
                    pass
                    
        # Write the target file path so VBA knows which document to process
        try:
            with open(target_file, 'w') as f:
                path_attr = self.sw_model.GetPathName
                model_path = path_attr() if callable(path_attr) else path_attr
                f.write(model_path)
        except Exception as e:
            if log_callback: log_callback(f"ERROR: Failed to write target file: {e}")
            return []
            
        # Run macro
        try:
            errs = win32.VARIANT(pythoncom.VT_BYREF | pythoncom.VT_I4, 0)
            success = self.sw_app.RunMacro2(str(macro_path), "GetDimensions1", "main", 0, errs)
            
            if not success:
                if log_callback: log_callback(f"ERROR: Macro failed to run. Code: {errs.value}")
                return []
        except Exception as e:
            if log_callback: log_callback(f"ERROR: Exception running macro: {e}")
            return []
            
        # Wait for the macro to finish writing the file
        max_wait = 50 # 5 seconds max
        while not temp_file.exists() and max_wait > 0:
            time.sleep(0.1)
            max_wait -= 1
            
        if not temp_file.exists():
            if log_callback: log_callback("ERROR: Macro ran but sw_dimensions.txt was not created.")
            return []
            
        # Read dimensions from the text file
        dims = set()
        try:
            with open(temp_file, 'r') as f:
                for line in f:
                    dim = line.strip()
                    if dim:
                        dims.add(dim)
                        if log_callback: log_callback(f"  -> Found dimension: {dim}")
        except Exception as e:
            if log_callback: log_callback(f"ERROR: Failed to read output file: {e}")
            
        if log_callback: log_callback(f"Successfully extracted {len(dims)} unique dimensions.")
        
        return sorted(list(dims))

    def modify_dimension(self, dim_name, new_value):
        """Modifies a specific dimension. Assumes input is in millimeters."""
        try:
            param = self.sw_model.Parameter(dim_name)
            if param:
                # SystemValue is strictly in meters, so we convert mm to meters
                param.SystemValue = float(new_value) / 1000.0
                return True
            return False
        except Exception as e:
            print(f"Error modifying {dim_name}: {e}")
            return False

    def rebuild(self):
        """Forces a model rebuild to apply dimensional changes."""
        if self.sw_model:
            # ForceRebuild3(False) = Rebuild only top level
            return self.sw_model.ForceRebuild3(False)
        return False

    def export_file(self, output_path, log_callback=None):
        """Exports the active model to the specified path (STEP, IGES, STL)."""
        if not self.sw_model:
            return False
            
        try:
            # API exports explicitly REQUIRE an absolute path
            abs_path = os.path.abspath(str(output_path))
            
            # The 'Type Mismatch' error (Code 4) is caused by passing None to the 4th argument
            # of Extension.SaveAs. To bypass this COM limitation in Python, we will use the 
            # strictly typed SaveAs3 method which only takes basic data types (String, Int, Int).
            
            # Options: 1 (Silent) + 2 (SaveAsCopy) = 3
            # SaveAsCopy is CRITICAL so it doesn't rename the active memory file!
            err_code = self.sw_model.SaveAs3(abs_path, 0, 3)
            success = (err_code == 0)
            
            if not success and log_callback:
                log_callback(f"  -> SaveAs3 failed with error code: {err_code}")
                    
            return success
        except Exception as e:
            if log_callback: log_callback(f"  -> Exception during export: {e}")
            return False

    def close(self):
        """Closes the active document cleanly to release file locks."""
        if self.sw_app and self.sw_model:
            try:
                title_attr = self.sw_model.GetTitle
                doc_name = title_attr() if callable(title_attr) else title_attr
                self.sw_app.CloseDoc(doc_name)
            except:
                pass
        
        # Explicitly release COM objects to allow the process to terminate
        self.sw_model = None
        self.sw_app = None