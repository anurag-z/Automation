import time
import os
import ctypes
import pandas as pd
import pytest
from pywinauto import Application, Desktop
from PIL import ImageGrab, ImageOps, Image, ImageEnhance
import pytesseract
import win32gui

# ==========================================
# 1. CONSTANTS & SETUP
# ==========================================
WORKING_DIR = r"C:\1040ta5"
LAUNCH_CMD = r'cmd.exe /k FADS'
TESSERACT_PATH = r'C:\Users\C302461\AppData\Local\Programs\Tesseract-OCR\tesseract.exe'
DEBUG_DIR = r"C:\Temp\OCR_Debug"
FADS_DEBUG_DIR = r"C:\FADS_Debug"
EXCEL_PATH = r"C:\Test\FieldAnalysis_Combined_1040.xlsx"
EXCEL_SHEET = 'Demo'

# --- DPI Awareness ---
try:
    ctypes.windll.shcore.SetProcessDpiAwareness(1)
    ctypes.windll.user32.SetProcessDPIAware()
    pytesseract.pytesseract.tesseract_cmd = TESSERACT_PATH
except Exception as e:
    print(f"[DPI/Tesseract Setup Error] {e}")

# ==========================================
# 2. UTILITY & NAVIGATION METHODS
# ==========================================
def ensure_dir_exists(directory):
    if not os.path.exists(directory):
        os.makedirs(directory)

def launch_fads():
    app = Application(backend="uia").start(LAUNCH_CMD, work_dir=WORKING_DIR)
    time.sleep(2)
    return app

def safe_type(window, keystrokes, wait_time=0.2):
    """Acts as a Pause/Play button. Waits until FADS is active before typing."""
    try:
        real_window = window.wrapper_object()
        fads_handle = real_window.handle
        
        attempts = 0
        while win32gui.GetForegroundWindow() != fads_handle:
            print(f"⚠️ Paused: Waiting for FADS to be active to type '{keystrokes}'...")
            try:
                real_window.set_focus()
            except:
                pass
            time.sleep(1.0)
            attempts += 1
            if attempts > 15:
                raise Exception("Focus timeout.")

        window.type_keys(keystrokes)
        if wait_time > 0:
            time.sleep(wait_time)
            
    except Exception as e:
        print(f"❌ CRITICAL: Could not type '{keystrokes}'. Error: {e}")
        raise e

def navigate_field(area, fieldname, row, column, window):
    """Navigates to the specific screen for the given area and field."""
    try:
        window.set_focus()
        window.maximize()
        safe_type(window, "{F3}{ESC}")
        
        for key in "MAS":
            safe_type(window, key, wait_time=0.2)
            
        safe_type(window, area + "{ENTER}", wait_time=0.5)
        safe_type(window, "{ESC}{ESC}")
        
        for key in "FFS":
            safe_type(window, key, wait_time=0.2)
            
        safe_type(window, fieldname + "{ENTER}", wait_time=0.8)
        safe_type(window, "{F9}AA", wait_time=1.2)
        
        safe_type(window, "{ESC}")
    except Exception as e:
        print(f"❌ [navigate_field] Error navigating to {area} -> {fieldname}: {e}")

# ==========================================
# 3. OCR & EXTRACTION METHODS
# ==========================================
def capture_and_read(window):
    """Placeholder for your robust image processing and OCR logic."""
    # Insert your exact capture_and_read image enhancement logic here
    # Make sure it returns the raw text string.
    pass 

def safe_get_field_values(window):
    """Waits for UI to settle, captures screen, and returns the raw list."""
    try:
        real_window = window.wrapper_object()
        fads_handle = real_window.handle
        
        attempts = 0
        while win32gui.GetForegroundWindow() != fads_handle:
            try: real_window.set_focus()
            except: pass
            time.sleep(1.0)
            attempts += 1
            if attempts > 15: raise Exception("Focus timeout.")

        time.sleep(0.5) # Critical pause for Windows animation
        
        # Assume capture_and_read is defined and returns the raw string
        screen_data = capture_and_read(window) 
        if not screen_data or screen_data.strip() == "":
            return []
            
        lines = screen_data.strip().split('\n')
        last_line = lines[-1]
        ls = last_line.split()
        return ls

    except Exception as e:
        print(f"❌ Error extracting field values: {e}")
        return []

# ==========================================
# 4. PYTEST DATA PROVIDER
# ==========================================
def get_excel_tasks():
    """Reads Excel and provides a list of tuples for Pytest."""
    df = pd.read_excel(EXCEL_PATH, sheet_name=EXCEL_SHEET)
    # Drop completely empty rows just in case
    df = df.dropna(how='all') 
    
    areas = df['Area'].tolist()
    field_names = df['Field Name'].tolist()
    rows = df['Row'].tolist()
    lengths = df['Length'].tolist()
    
    return list(zip(areas, field_names, rows, lengths))

# ==========================================
# 5. PYTEST FIXTURE & TEST LOGIC
# ==========================================
@pytest.fixture(scope="session")
def fads_window():
    """Launches the app once for the entire test session."""
    ensure_dir_exists(DEBUG_DIR)
    app = launch_fads()
    
    window = Desktop(backend="uia").window(title_re=".*FADS PRIME.*")
    window.wait('ready', timeout=10)
    print("\n[+] FADS Application Launched.")
    
    yield window 

@pytest.mark.parametrize("area, expected_fieldname, expected_row, expected_length", get_excel_tasks())
def test_fads_field_validation(fads_window, area, expected_fieldname, expected_row, expected_length):
    """The main assertion test that runs for every row in Excel."""
    print(f"\n--- Testing: {area} -> {expected_fieldname} ---")
    
    try:
        # 1. Navigate
        navigate_field(area, expected_fieldname, expected_row, expected_length, fads_window)
        time.sleep(0.5)
        
        # 2. Extract Data (Returns list 'ls')
        ls = safe_get_field_values(fads_window)
        assert ls and len(ls) >= 8, f"❌ OCR returned incomplete list: {ls}"
        
        # 3. Transform the List based on your index rules
        actual_fieldname = ls[1]
        actual_length = ls[3]
        raw_coords = ls[7]
        
        assert "," in raw_coords, f"❌ Expected comma in index 7, but got: '{raw_coords}'"
        actual_row, actual_col = raw_coords.split(",")
        
        print(f"Captured -> Field: {actual_fieldname}, Row: {actual_row}, Col: {actual_col}, Len: {actual_length}")
        
        # 4. Assertions
        assert actual_fieldname == expected_fieldname, f"Fieldname Mismatch! Excel: {expected_fieldname}, Screen: {actual_fieldname}"
        assert str(actual_row) == str(expected_row), f"Row Mismatch! Excel: {expected_row}, Screen: {actual_row}"
        assert str(actual_length) == str(expected_length), f"Length Mismatch! Excel: {expected_length}, Screen: {actual_length}"
        
        print(f"✅ Validation Passed for {expected_fieldname}")

    finally:
        # 5. Safe Backout (Ensures menu reset even if test fails)
        print(f"Backing out to Main Menu...")
        safe_type(fads_window, "{F9}", wait_time=0.2)
        safe_type(fads_window, "{F9}", wait_time=0.2)
        safe_type(fads_window, "Y", wait_time=0.5)



from PIL import ImageChops, ImageStat, Image

BASELINE_IMG_PATH = r"C:\Temp\fads_main_menu_baseline.png"

def is_main_menu_active(window, tolerance=3.0):
    """
    Takes a live screenshot of the FADS window and compares it to the baseline Main Menu.
    Returns True if the screens match within the allowed tolerance.
    """
    try:
        if not os.path.exists(BASELINE_IMG_PATH):
            print(f"❌ Baseline image missing at {BASELINE_IMG_PATH}")
            return False
            
        # 1. Load baseline and convert to Grayscale ('L') to ignore minor color shifts
        baseline = Image.open(BASELINE_IMG_PATH).convert('L')
        
        # 2. Capture the current FADS window
        rect = window.rectangle()
        current_screen = ImageGrab.grab(bbox=(rect.left, rect.top, rect.right, rect.bottom)).convert('L')
        
        # Ensure sizes match (in case the window was resized slightly)
        if baseline.size != current_screen.size:
            current_screen = current_screen.resize(baseline.size)
            
        # 3. Calculate the difference
        diff = ImageChops.difference(baseline, current_screen)
        stat = ImageStat.Stat(diff)
        mean_diff = stat.mean[0] # Get average difference of the grayscale band
        
        # 4. Check against tolerance
        # If mean_diff is 0, they are identical. If it's < 3.0, it's a 99% match (ignoring cursor blinks).
        if mean_diff < tolerance:
            return True
        else:
            print(f"   [Visual Check] Not at menu. Difference score: {mean_diff:.2f}")
            return False
            
    except Exception as e:
        print(f"❌ [is_main_menu_active] Error: {e}")
        return False


except Exception as e:
        # --- PANIC RECOVERY WITH VISUAL VERIFICATION ---
        print(f"\n⚠️ TEST FAILED: Initiating visual recovery to Main Menu...")
        
        max_attempts = 8
        recovered = False
        
        for attempt in range(1, max_attempts + 1):
            print(f"Attempt {attempt}/{max_attempts} to reach Main Menu...")
            
            # 1. Look at the screen
            if is_main_menu_active(fads_window):
                print(f"✅ Visual Confirmation: Reached Main Menu successfully!")
                recovered = True
                break
                
            # 2. If not at the menu, press ESC to back out one level
            safe_type(fads_window, "{ESC}", wait_time=0.5)
            
        if not recovered:
            print("❌ CRITICAL: App is stuck. Could not reach Main Menu after 8 attempts.")
            # Optional: Add logic here to restart the FADS process entirely
            
        # Re-raise the error so Pytest properly marks this row as FAILED in your HTML/Excel report
        raise e
import pytest

# Your exact filter dictionary
FILTERS = {
    "Safe to Extend": ["Yes"],
    "Action Required": ["Move and Extend (Non-Group)", "Extend Only (Non-Group)"]
}

# ==========================================
# 1. HELPER FUNCTIONS
# ==========================================
def is_blank_field(val):
    """Safely checks if an Excel cell is empty, None, or a Pandas NaN."""
    if val is None:
        return True
    str_val = str(val).strip().lower()
    if str_val == "" or str_val == "nan":
        return True
    return False

def passes_business_filters(task):
    """Checks if a row meets all criteria in the FILTERS dictionary."""
    for column_name, allowed_values in FILTERS.items():
        # Get the value from the Excel row (default to empty string if missing)
        raw_val = task.get(column_name, "")
        
        # Convert to string and strip spaces (Protects against Excel typos like "Yes ")
        clean_val = str(raw_val).strip() if raw_val is not None else ""
        
        # If the cleaned Excel value is NOT in our list of allowed values, reject the row
        if clean_val not in allowed_values:
            return False
            
    # If it survived the loop, it matches all filters!
    return True

# ==========================================
# 2. PYTEST COLLECTION PHASE
# ==========================================
# Step A: Load ALL data
all_tasks = get_excel_tasks()

# Step B: Apply the Business Filters first
# This throws away any rows that aren't "Yes" and "Move/Extend"
filtered_tasks = [task for task in all_tasks if passes_business_filters(task)]

# Step C: Split the surviving rows into Flow 1 (Has Field Name)
flow1_tasks = [
    task for task in filtered_tasks 
    if not is_blank_field(task.get("Field Name"))
]

# Step D: Split the surviving rows into Flow 2 (Blank Field Name)
flow2_tasks = [
    task for task in filtered_tasks 
    if is_blank_field(task.get("Field Name"))
]

# ==========================================
# 3. TEST DEFINITIONS 
# ==========================================
@pytest.mark.parametrize("task", flow1_tasks, ids=lambda t: f"Flow1_Row{t.get('Row')}")
def test_fads_field_validation(fads_window, task):
    # Your standard validation flow...
    pass

@pytest.mark.parametrize("task", flow2_tasks, ids=lambda t: f"Flow2_Row{t.get('Row')}")
def test_fads_blank_field_flow(fads_window, task):
    # Your alternate blank field flow...
    pass
