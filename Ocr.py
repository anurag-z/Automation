import time
from pywinauto import Application
from PIL import ImageGrab, ImageOps
import pytesseract

# --- CONFIGURATION ---
# Path to your Tesseract engine
pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'

# Launch settings from your C# code
WORKING_DIR = r"C:\1040ta5"
LAUNCH_CMD = 'cmd.exe /k FADS'

def navigate_field(area, fieldname, window):
    print(f"Processing Screen: {area}")
    
    # Block 1: Navigate to Form (Matches your C# image)
    window.type_keys("{F3}")
    window.type_keys("{ESC}")
    window.type_keys("M")
    window.type_keys("A")
    window.type_keys("S")
    window.type_keys(area) # Form name
    window.type_keys("{ENTER}")
    time.sleep(0.5)
    
    # Block 2: Navigate to Field
    window.type_keys("{ESC}")
    window.type_keys("{ESC}")
    window.type_keys("F")
    window.type_keys("F")
    window.type_keys("S")
    window.type_keys(fieldname) # Field Name
    window.type_keys("{ENTER}")
    time.sleep(0.5)
    
    # Block 3: Final trigger and OCR
    window.type_keys("{F9}")
    window.type_keys("A")
    window.type_keys("A")
    time.sleep(1) # Wait for DOS screen to refresh
    
    # OCR Section (Captures via Window Handle/HWND)
    rect = window.rectangle()
    screenshot = ImageGrab.grab(bbox=(rect.left, rect.top, rect.right, rect.bottom))
    
    # Invert for DOS (Blue background -> White) for better OCR
    processed_img = ImageOps.invert(screenshot.convert('RGB'))
    result = pytesseract.image_to_string(processed_img, config='--psm 6')
    
    print(f"Result for {area}: {result.strip()}")
    window.type_keys("{ESC}")

def main():
    # 1. Launch Process (Matches your Process.Start)
    print("Launching FADS...")
    app = Application(backend="win32").start(LAUNCH_CMD, work_dir=WORKING_DIR)
    time.sleep(3) # Matches your Thread.Sleep(3000)
    
    # 2. Connect to the specific FADS window
    window = app.window(title_re=".*FADS PRIME.*")
    window.set_focus()

    # 3. Data Loop (Matches your List<Dictionary> foreach)
    data = [
        {"AreaName": "Basis", "FieldName": "KL72"},
        {"AreaName": "F4797", "FieldName": "POST"},
        {"AreaName": "FSCHDAMT", "FieldName": "STAA"},
        {"AreaName": "FCOMB", "FieldName": "NITEMSTD"}
    ]

    for row in data:
        navigate_field(row["AreaName"], row["FieldName"], window)
        time.sleep(0.1) # Matches your Thread.Sleep(100)

if __name__ == "__main__":
    main()


import os
import shutil

# 1. Check if Tesseract is found automatically
auto_path = shutil.which("tesseract")
print(f"Tesseract in System PATH: {auto_path}")

# 2. Check common manual paths
common_paths = [
    r'C:\Program Files\Tesseract-OCR\tesseract.exe',
    r'C:\Program Files (x86)\Tesseract-OCR\tesseract.exe'
]

print("\nChecking common install locations:")
for path in common_paths:
    if os.path.exists(path):
        print(f"[FOUND] {path}")
    else:
        print(f"[MISSING] {path}")

# 3. Check if Pytesseract library is actually ready
try:
    import pytesseract
    print("\n[OK] Pytesseract library is installed.")
except ImportError:
    print("\n[ERROR] Pytesseract library not found. Use 'python -m pip install pytesseract'")
