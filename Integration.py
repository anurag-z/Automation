import time
from pywinauto import Application
from PIL import ImageGrab, ImageOps, ImageResampling
import pytesseract

# --- CONFIG ---
pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'

def capture_and_read(window):
    """Captures the window at runtime and applies the working OCR logic."""
    try:
        # 1. Get the exact coordinates of the FADS window at this moment
        rect = window.rectangle() 
        
        # 2. Take the screenshot (Left, Top, Right, Bottom)
        screenshot = ImageGrab.grab(bbox=(rect.left, rect.top, rect.right, rect.bottom))
        
        # 3. Apply the High-Res logic that worked for '10' and '30'
        width, height = screenshot.size
        img = screenshot.resize((width * 3, height * 3), resample=ImageResampling.LANCZOS)
        
        # 4. Pre-process (Invert blue to white and Threshold)
        inverted = ImageOps.invert(img.convert('RGB'))
        bw_img = inverted.convert('L').point(lambda x: 0 if x < 180 else 255, '1')
        
        # 5. Extract Text
        # Using PSM 6 for the uniform console block and preserving spaces
        config = r'--psm 6 -c preserve_interword_spaces=1'
        text = pytesseract.image_to_string(bw_img, config=config)
        
        return text.strip()
    except Exception as e:
        return f"Capture Error: {e}"

def main():
    # Connect to your already running FADS
    app = Application(backend="win32").connect(title_re=".*FADS PRIME.*")
    window = app.window(title_re=".*FADS PRIME.*")
    
    # Example: Fire your keys, then capture
    window.set_focus()
    
    # ... your navigateField logic here ...
    
    # Run the OCR at runtime
    print("Reading screen data...")
    screen_data = capture_and_read(window)
    print(f"OCR Output:\n{screen_data}")

if __name__ == "__main__":
    main()
import os
import time
from pywinauto import Application
from PIL import ImageGrab, ImageOps, Image
import pytesseract

# --- CONFIG ---
pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'
DEBUG_DIR = r"C:\Temp\OCR_Debug"

# Ensure the debug folder exists
if not os.path.exists(DEBUG_DIR):
    os.makedirs(DEBUG_DIR)

def get_bottom_bar_text_with_debug(window, area_name):
    """Captures the bottom bar and saves images for debugging coordinates."""
    try:
        # 1. Get window handle coordinates
        rect = window.rectangle()
        
        # 2. Define the Bottom ROI
        # We target the very bottom where the ADD ATTR line sits
        left = rect.left
        right = rect.right
        bottom = rect.bottom
        top_of_bar = rect.bottom - 70  # Try 70 pixels for just the status line
        
        # 3. Capture Raw Image
        raw_capture = ImageGrab.grab(bbox=(left, top_of_bar, right, bottom))
        raw_capture.save(os.path.join(DEBUG_DIR, f"1_Raw_{area_name}.png"))
        
        # 4. Processing (The High-Res fix you confirmed works)
        width, height = raw_capture.size
        img = raw_capture.resize((width * 3, height * 3), resample=Image.LANCZOS)
        
        # Invert and Threshold
        inverted = ImageOps.invert(img.convert('RGB'))
        # Using 180 threshold to keep '0' from becoming '6'
        bw_img = inverted.convert('L').point(lambda x: 0 if x < 180 else 255, '1')
        
        # Save the processed image - IF THIS IS BLANK, THE THRESHOLD IS TOO HIGH
        bw_img.save(os.path.join(DEBUG_DIR, f"2_Processed_{area_name}.png"))
        
        # 5. OCR
        config = r'--psm 6 -c preserve_interword_spaces=1'
        text = pytesseract.image_to_string(bw_img, config=config)
        
        return text.strip()

    except Exception as e:
        return f"Debug Error: {e}"

# Usage inside your loop:
# result = get_bottom_bar_text_with_debug(window, row["AreaName"])
