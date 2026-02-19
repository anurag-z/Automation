import time
import os
from pywinauto import Application
from PIL import ImageGrab, ImageOps
import pytesseract

# --- CONFIGURATION ---
pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'
WORKING_DIR = r"C:\1040ta5"
LAUNCH_CMD = r'cmd.exe /k FADS'

def navigate_field(area, fieldname, window):
    print(f"\n[+] Processing: {area} -> {fieldname}")
    window.set_focus()
    
    # Navigation Sequence (Matches your C# logic)
    window.type_keys("{F3}{ESC}")
    for key in "MAS": 
        window.type_keys(key)
        time.sleep(0.1)
    window.type_keys(area + "{ENTER}")
    time.sleep(0.5)
    
    window.type_keys("{ESC}{ESC}")
    for key in "FFS":
        window.type_keys(key)
        time.sleep(0.1)
    window.type_keys(fieldname + "{ENTER}")
    time.sleep(0.5)
    
    window.type_keys("{F9}AA")
    time.sleep(1.2) # Delay for FADS screen to draw
    
    # OCR Logic targeting the Window HWND
    rect = window.rectangle()
    screenshot = ImageGrab.grab(bbox=(rect.left, rect.top, rect.right, rect.bottom))
    processed_img = ImageOps.invert(screenshot.convert('RGB'))
    result = pytesseract.image_to_string(processed_img, config='--psm 6')
    
    print(f"--- DATA READ ---\n{result.strip()}\n-----------------")
    window.type_keys("{ESC}")

def main():
    try:
        print("Launching FADS in a new external window...")
        # create_new_console=True forces it out of the VS Code terminal
        # wait_for_idle=False prevents the 'Not a GUI process' error
        app = Application(backend="win32").start(
            LAUNCH_CMD, 
            work_dir=WORKING_DIR, 
            create_new_console=True, 
            wait_for_idle=False
        )
        
        # Give it time to load the new window
        time.sleep(3) 
        
        # Connect to the window using title_re for regex matching
        window = app.window(title_re=".*FADS PRIME.*")
        window.wait('ready', timeout=10) # Safe wait for console window
        
        tasks = [
            {"Area": "Basis", "Field": "KL72"},
            {"Area": "F4797", "Field": "POST"}
        ]

        for t in tasks:
            navigate_field(t["Area"], t["Field"], window)
            time.sleep(0.2)

    except Exception as e:
        print(f"FAILED: {e}")

if __name__ == "__main__":
    main()



import pytesseract
from PIL import Image, ImageOps

# --- CONFIGURATION ---
# Update this path to where your Tesseract engine is installed
pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'

def read_fads_screen(image_path):
    try:
        # 1. Load the image from your local drive
        img = Image.open(image_path)
        
        # 2. Convert to RGB and Invert colors
        # This turns the blue background white and the white text black
        img_rgb = img.convert('RGB')
        inverted_img = ImageOps.invert(img_rgb)
        
        # 3. Convert to Grayscale (L) for cleaner OCR processing
        final_img = inverted_img.convert('L')
        
        # Optional: Save the "fixed" image to see what the OCR is looking at
        final_img.save("processed_for_ocr.png")

        # 4. Run OCR with Page Segmentation Mode (PSM) 6
        # PSM 6 is best for uniform blocks of text like DOS/Console screens
        text = pytesseract.image_to_string(final_img, config='--psm 6')

        return text

    except Exception as e:
        return f"Error processing image: {e}"

# --- EXECUTION ---
if __name__ == "__main__":
    # Replace this with the path to the image you want to read
    test_image = "image_1e902a.png" 
    
    extracted_text = read_fads_screen(test_image)
    
    print("\n--- EXTRACTED TEXT ---")
    print(extracted_text)
    print("----------------------")
