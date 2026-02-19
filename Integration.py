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
import time
from pywinauto import Application
from PIL import ImageGrab, ImageOps, Image
import pytesseract

# --- CONFIG ---
pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'

def get_bottom_bar_text(window):
    """Captures and reads ONLY the bottom section of the FADS window."""
    try:
        # 1. Get total window coordinates
        rect = window.rectangle()
        
        # 2. Calculate ROI (Region of Interest)
        # We only want the bottom part (approx. last 80 pixels)
        left = rect.left
        right = rect.right
        bottom = rect.bottom
        top_of_bar = rect.bottom - 80  # Adjust this number to crop higher or lower
        
        # 3. Capture the cropped area
        screenshot = ImageGrab.grab(bbox=(left, top_of_bar, right, bottom))
        
        # 4. Apply the working High-Res fix
        width, height = screenshot.size
        img = screenshot.resize((width * 3, height * 3), resample=Image.LANCZOS)
        
        # 5. Pre-process (Invert & Threshold)
        inverted = ImageOps.invert(img.convert('RGB'))
        bw_img = inverted.convert('L').point(lambda x: 0 if x < 180 else 255, '1')
        
        # 6. OCR with preserved spaces
        config = r'--psm 6 -c preserve_interword_spaces=1'
        return pytesseract.image_to_string(bw_img, config=config)

    except Exception as e:
        return f"ROI Error: {e}"

# --- Usage Example ---
# app = Application(backend="win32").connect(title_re=".*FADS PRIME.*")
# window = app.window(title_re=".*FADS PRIME.*")
# text = get_bottom_bar_text(window)
# print(text)
