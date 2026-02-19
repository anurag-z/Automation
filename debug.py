from PIL import Image, ImageGrab, ImageOps, ImageEnhance
import pytesseract
import os

# IMPORTANT: Point this to where Tesseract is installed on your machine
# pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'

def capture_and_read(window):
    try:
        # 1. Get the exact coordinates of the FADS window at this moment
        rect = window.rectangle()

        # 2. Take the screenshot (Left, Top, Right, Bottom)
        screenshot = ImageGrab.grab(bbox=(rect.left, rect.top, rect.right, rect.bottom))
        img = screenshot.convert('RGB')
        
        width, height = img.size

        # 3. FIND THE WHITE BORDER LINE
        white_line_y = -1
        mid_x = width // 2

        # Scan the middle column from bottom up to the middle
        for y in range(height - 1, height // 2, -1):
            r, g, b = img.getpixel((mid_x, y))
            
            # Brightness check to find the white line
            if r > 200 and g > 200 and b > 200:
                white_line_y = y
                break

        # 4. CROP BELOW THE BORDER
        # Start 3 pixels below the white line, or fallback to height - 35
        start_y = (white_line_y + 3) if white_line_y != -1 else (height - 35)
        
        # SAFETY FIX: Prevent the "lower is less than upper" crash
        if start_y >= height:
            start_y = height - 30 
        
        # Crop the image: (left, top, right, bottom)
        cropped_img = img.crop((0, start_y, width, height))

        # 5. APPLY YOUR WORKING IMAGE PROCESSING (on the cropped area only)
        crop_width, crop_height = cropped_img.size
        img_resized = cropped_img.resize((crop_width * 3, crop_height * 3), resample=Image.Resampling.LANCZOS)
        
        gray = img_resized.convert('L')
        inverted = ImageOps.invert(gray)
        
        enhancer = ImageEnhance.Sharpness(inverted)
        sharpened = enhancer.enhance(2.0)
        
        bw_img = sharpened.point(lambda x: 0 if x < 180 else 255, '1')

        # 6. SAVE DEBUG IMAGE (Using Raw String for C: Drive path)
        # Make sure the folder "C:\FADS_Debug" actually exists on your computer!
        debug_folder = r"C:\FADS_Debug"
        if not os.path.exists(debug_folder):
            os.makedirs(debug_folder)
            
        debug_path = os.path.join(debug_folder, "debug_bottom_line_crop.png")
        bw_img.save(debug_path)

        # 7. PRECISE OCR CONFIG & READ
        custom_config = r'--psm 6 -c preserve_interword_spaces=1'
        text = pytesseract.image_to_string(bw_img, config=custom_config)
        
        # 8. POST-PROCESSING CLEANUP
        # Fix the Q and @ issues without messing with Tesseract settings
        text = text.replace("Q ", "0 ").replace("@", "0")
        
        return text.strip()

    except Exception as e:
        return f"Capture Error: {e}"
from PIL import ImageGrab
import os
import ctypes

# This ensures Windows scaling (like 125% zoom) doesn't mess up your coordinates
ctypes.windll.user32.SetProcessDPIAware()

def capture_hardcoded_region():
    # --- ADJUST THESE COORDINATES ---
    # Look at your screen as an X/Y grid starting at (0,0) in the top-left corner.
    # These numbers are estimates based on your screenshot. 
    
    LEFT = 0       # Starts at the far left edge
    TOP = 640      # Starts roughly where the blue box ends and the black begins
    RIGHT = 1100   # Extends right to catch all the text
    BOTTOM = 690   # Ends right underneath the text
    
    # Create the bounding box
    bbox = (LEFT, TOP, RIGHT, BOTTOM)
    
    try:
        # Take the screenshot
        screenshot = ImageGrab.grab(bbox=bbox)
        
        # Save it to the C: drive to check your alignment
        save_dir = r"C:\FADS_Debug"
        if not os.path.exists(save_dir):
            os.makedirs(save_dir)
            
        save_path = os.path.join(save_dir, "hardcoded_crop.png")
        screenshot.save(save_path)
        
        print(f"SUCCESS! Screenshot saved to: {save_path}")
        print(f"Used coordinates: {bbox}")
        print("-> Go look at the image. If it's too high or low, adjust the TOP and BOTTOM numbers!")
        
    except Exception as e:
        print(f"Capture Error: {e}")

if __name__ == "__main__":
    capture_hardcoded_region()
