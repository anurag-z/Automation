import time
import subprocess
from pywinauto import Application
from PIL import ImageGrab, ImageOps
import pytesseract

# --- CONFIGURATION ---
# Path to your Tesseract engine (Verified in your earlier check)
pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'

# Launch settings from your C# image_1d23cf.png
WORKING_DIR = r"C:\1040ta5"
LAUNCH_CMD = 'cmd.exe /k FADS'

def navigate_field(area, fieldname, window):
    print(f"\n[+] Navigating: {area} -> {fieldname}")
    
    # Block 1: Form Selection (Matches your C# image_1d2446.png)
    window.type_keys("{F3}")
    window.type_keys("{ESC}")
    for key in "MAS": 
        window.type_keys(key)
        time.sleep(0.1)
    window.type_keys(area + "{ENTER}")
    time.sleep(0.5)
    
    # Block 2: Field Selection
    window.type_keys("{ESC}{ESC}")
    for key in "FFS":
        window.type_keys(key)
        time.sleep(0.1)
    window.type_keys(fieldname + "{ENTER}")
    time.sleep(0.5)
    
    # Block 3: Final trigger (F9 -> A -> A)
    window.type_keys("{F9}AA")
    time.sleep(1.5) 
    
    # Block 4: Capture via Window Handle (HWND)
    rect = window.rectangle()
    screenshot = ImageGrab.grab(bbox=(rect.left, rect.top, rect.right, rect.bottom))
    
    # Invert Blue to White for accurate OCR
    processed_img = ImageOps.invert(screenshot.convert('RGB'))
    text_result = pytesseract.image_to_string(processed_img, config='--psm 6')
    
    print("--- OCR RESULT ---")
    print(text_result.strip())
    print("------------------")
    
    window.type_keys("{ESC}")

def main():
    try:
        # FIX: Launch using subprocess to avoid 'Not a GUI process' error
        print("Launching FADS...")
        subprocess.Popen(LAUNCH_CMD, cwd=WORKING_DIR, shell=True)
        
        # Give it 3 seconds to load (Matches your C# Thread.Sleep(3000))
        time.sleep(3) 
        
        # Connect to the window by title (as seen in image_1cbe8d.png)
        app = Application(backend="win32").connect(title_re=".*FADS PRIME.*")
        window = app.window(title_re=".*FADS PRIME.*")
        window.set_focus()

        # Data list from your C# image_1d23cf.png
        tasks = [
            {"Area": "Basis", "Field": "KL72"},
            {"Area": "F4797", "Field": "POST"},
            {"Area": "FSCHDAMT", "Field": "STAA"},
            {"Area": "FCOMB", "Field": "NITEMSTD"}
        ]

        for t in tasks:
            navigate_field(t["Area"], t["Field"], window)
            time.sleep(0.2)

    except Exception as e:
        print(f"Error: {e}")

if __name__ == "__main__":
    main()
