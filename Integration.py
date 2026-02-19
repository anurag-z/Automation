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
  """
    Python version of your C# CaptureAndRead logic
    1. Grabs full window.
    2. Scans for the white line.
    3. Crops below it.
    """
    try:
        window.set_focus()
        time.sleep(0.5)
        
        # 1. GET COORDINATES
        rect = window.rectangle()
        full_content = ImageGrab.grab(bbox=(rect.left, rect.top, rect.right, rect.bottom))
        width, height = full_content.size
        
        # 2. FIND THE WHITE BORDER LINE
        # Scans from bottom upwards to find the first bright horizontal line
        white_line_y = -1
        pixels = full_content.load()
        middle_x = width // 2
        
        # Scan bottom 25% of the window for the white line
        for y in range(height - 1, height // 2, -1):
            r, g, b = pixels[middle_x, y]
            # 0.85 brightness equivalent in RGB (approx 215)
            if r > 215 and g > 215 and b > 215:
                white_line_y = y
                break
        
        # 3. CROP BELOW THE BORDER
        # Start 3 pixels below to avoid noise, exactly like your C# code
        start_y = white_line_y + 3 if white_line_y != -1 else height - 35
        
        # Define the region to crop
        crop_region = (0, start_y, width, height)
        raw_crop = full_content.crop(crop_region)
        
        # Save for verification
        raw_crop.save(os.path.join(DEBUG_DIR, f"{label}_1_Target_Crop.png"))
        
        # 4. PRE-PROCESS (High-Res Fix for '10' vs '16')
        # Blow up the small crop by 3x to sharpen the numbers
        new_w, new_h = raw_crop.size
        upscaled = raw_crop.resize((new_w * 3, new_h * 3), resample=Image.LANCZOS)
        
        # Invert (Blue -> White) and Threshold
        inverted = ImageOps.invert(upscaled.convert('RGB'))
        bw_img = inverted.convert('L').point(lambda x: 0 if x < 180 else 255, '1')
        bw_img.save(os.path.join(DEBUG_DIR, f"{label}_2_Final_OCR.png"))
        
        # 5. RUN OCR
        return pytesseract.image_to_string(bw_img, config=r'--psm 6 -c preserve_interword_spaces=1')

    except Exception as e:
        return f"Logic Error: {e}"
def capture_and_read_approved(window, label):
    """
    Implements your approved C# logic:
    - Scan for white border line (>0.85 brightness)
    - Crop 3 pixels below that line
    - Apply verified 3x Resize + 180 Threshold
    """
    try:
        # Ensure window is active to get valid coordinates
        window.set_focus()
        time.sleep(0.5) 
        
        # A. GET COORDINATES
        rect = window.rectangle()
        full_content = ImageGrab.grab(bbox=(rect.left, rect.top, rect.right, rect.bottom))
        width, height = full_content.size
        
        # B. FIND WHITE BORDER LINE (Approved C# logic)
        white_line_y = -1
        pixels = full_content.load()
        middle_x = width // 2
        
        # Scan from bottom to middle for bright white line
        for y in range(height - 1, height // 2, -1):
            r, g, b = pixels[middle_x, y]
            # Equivalent to your C# .GetBrightness() > 0.85f
            if r > 215 and g > 215 and b > 215:
                white_line_y = y
                break
        
        # C. CROP BELOW BORDER (Approved C# offsets)
        # Use your +3 pixel offset to avoid line noise
        start_y = white_line_y + 3 if white_line_y != -1 else height - 35
        raw_crop = full_content.crop((0, start_y, width, height))
        
        # D. APPLY THE VERIFIED OCR FIX
        # Resize 3x to ensure '10' is read correctly (not '16')
        w, h = raw_crop.size
        upscaled = raw_crop.resize((w * 3, h * 3), resample=Image.LANCZOS)
        
        # Invert blue-to-white and Threshold at 180
        inverted = ImageOps.invert(upscaled.convert('RGB'))
        bw_img = inverted.convert('L').point(lambda x: 0 if x < 180 else 255, '1')
        
        # Save for your debugging
        bw_img.save(os.path.join(DEBUG_DIR, f"{label}_Target_Crop.png"))
        
        # E. READ TEXT
        config = r'--psm 6 -c preserve_interword_spaces=1'
        return pytesseract.image_to_string(bw_img, config=config).strip()

    except Exception as e:
        return f"Error: {e}"
