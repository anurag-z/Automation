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
from PIL import Image, ImageOps, ImageEnhance

# --- CONFIGURATION ---
# Ensure this points to your Tesseract EXE
pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'

def read_fads_clean(image_path):
    try:
        # 1. Load the original blue image
        img = Image.open(image_path).convert('RGB')
        
        # 2. Invert colors: Blue background becomes white, white text becomes black
        inverted_img = ImageOps.invert(img)
        
        # 3. Convert to Grayscale (L) for better contrast
        gray_img = inverted_img.convert('L')
        
        # 4. Enhance Contrast: Sharpen the text edges
        enhancer = ImageEnhance.Contrast(gray_img)
        high_contrast = enhancer.enhance(2.0)
        
        # 5. Thresholding: Force pixels to be either pure black or pure white
        # This removes "ghosting" or shadows that cause errors like '1@' instead of '10'
        final_img = high_contrast.point(lambda x: 0 if x < 140 else 255, '1')
        
        # Save for debugging to see the improved quality
        final_img.save("ocr_debug_clean.png")

        # 6. Run OCR with PSM 6 (Uniform block of text)
        # Added whitelist for common characters found in your FADS screens
        custom_config = r'--psm 6 -c tessedit_char_whitelist=0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz,: '
        text = pytesseract.image_to_string(final_img, config=custom_config)

        return text.strip()

    except Exception as e:
        return f"Error: {e}"

if __name__ == "__main__":
    # Path to your captured FADS screen snippet
    image_to_read = "image_1e902a.png" 
    print("Reading screen...")
    print("-" * 30)
    print(read_fads_clean(image_to_read))
    print("-" * 30)
import pytesseract
from PIL import Image, ImageOps, ImageEnhance

# --- CONFIGURATION ---
pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'

def fix_ocr_numbers(image_path):
    try:
        # 1. Load the image
        img = Image.open(image_path).convert('RGB')
        
        # 2. Invert colors (Crucial: Blue -> White background)
        inverted = ImageOps.invert(img)
        
        # 3. Convert to Grayscale
        gray = inverted.convert('L')
        
        # 4. Sharpen and Increase Contrast
        # This makes the "0" hole larger so it doesn't look like a "6"
        enhancer = ImageEnhance.Contrast(gray)
        high_contrast = enhancer.enhance(3.0) # Aggressive contrast
        
        # 5. Thresholding
        # Adjust 160 higher if the '0' still looks like '6'
        # Adjust 160 lower if the characters start disappearing
        bw_img = high_contrast.point(lambda x: 0 if x < 160 else 255, '1')
        
        # 6. Save for inspection (Check if '0' looks like '0' here)
        bw_img.save("number_check.png")

        # 7. OCR Configuration
        # --psm 6: Uniform block of text
        # -c tessedit_char_whitelist: Only look for numbers, letters, and spaces
        custom_config = r'--psm 6 -c preserve_interword_spaces=1 -c tessedit_char_whitelist=0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz, '
        
        text = pytesseract.image_to_string(bw_img, config=custom_config)
        return text

    except Exception as e:
        return f"Error: {e}"

if __name__ == "__main__":
    # Test on your snippet
    print("Reading with high-precision settings...")
    result = fix_ocr_numbers("image_1e902a.png")
    print(f"\nExtracted Text:\n{result}")
