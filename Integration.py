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
   """Captures only the bottom two lines of the FADS window."""
    try:
        # 1. Force the window to the front before calculating coordinates
        window.set_focus()
        time.sleep(0.5) # Give Windows time to draw the pixels
        
        # 2. Get window coordinates
        rect = window.rectangle()
        
        # 3. Define the Bottom ROI (Region of Interest)
        # We target the bottom 60 pixels for the status line
        # Use rect.bottom - 60 to avoid capturing the taskbar or main form
        left = rect.left + 5    # Small offset to avoid window borders
        top = rect.bottom - 65  # The 'ADD ATTR' line starts roughly 60px from bottom
        right = rect.right - 5
        bottom = rect.bottom - 5
        
        # 4. Capture and Save Raw for verification
        raw = ImageGrab.grab(bbox=(left, top, right, bottom))
        raw.save(os.path.join(DEBUG_DIR, f"Raw_{label}.png"))
        
        # 5. Pre-process for OCR (High-Res Fix)
        # Resize 3x to ensure '10' isn't read as '16'
        width, height = raw.size
        img = raw.resize((width * 3, height * 3), resample=Image.LANCZOS)
        
        # Invert and Binary Threshold
        inverted = ImageOps.invert(img.convert('RGB'))
        # Using 170 as threshold based on your FADS blue color
        bw_img = inverted.convert('L').point(lambda x: 0 if x < 170 else 255, '1')
        bw_img.save(os.path.join(DEBUG_DIR, f"Processed_{label}.png"))
        
        # 6. OCR
        return pytesseract.image_to_string(bw_img, config=r'--psm 6 -c preserve_interword_spaces=1')

    except Exception as e:
        return f"Error: {e}"
        # Usage inside your loop:
# result = get_bottom_bar_text_with_debug(window, row["AreaName"])
