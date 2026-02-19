import os
import ctypes
import win32gui
import pytesseract
from PIL import ImageGrab

# 1. FIX WINDOWS SCALING ISSUES (CRITICAL)
# This tells Windows to give Python the exact pixel coordinates on your screen.
# Without this, Windows Display Scaling (like 125%) will capture the wrong area.
ctypes.windll.user32.SetProcessDPIAware()

# IMPORTANT: Point this to where Tesseract is installed on your machine
pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'

def find_fads_window():
    """Finds the window handle (hWnd) by checking if 'FADS PRIME' is in the title."""
    found_hwnd = None
    
    def callback(hwnd, extra):
        nonlocal found_hwnd
        if win32gui.IsWindowVisible(hwnd):
            title = win32gui.GetWindowText(hwnd)
            # Match the title from your screenshot
            if "FADS PRIME" in title:
                found_hwnd = hwnd
        return True
        
    win32gui.EnumWindows(callback, None)
    return found_hwnd

def capture_and_read():
    hwnd = find_fads_window()
    if not hwnd:
        print("Could not find the 'FADS PRIME' window. Make sure the app is open.")
        return

    # 2. GET INNER WINDOW COORDINATES
    # ClientToScreen ensures we ignore the top white title bar and window borders
    left, top = win32gui.ClientToScreen(hwnd, (0, 0))
    _, _, right, bottom = win32gui.GetClientRect(hwnd)
    
    width = right
    height = bottom

    # 3. CAPTURE THE SCREEN INTO MEMORY
    bbox = (left, top, left + width, top + height)
    img = ImageGrab.grab(bbox).convert('RGB')

    # 4. SCAN FOR THE WHITE LINE
    # We scan the middle of the screen, moving from the bottom upwards
    white_line_y = -1
    mid_x = width // 2

    for y in range(height - 1, height // 2, -1):
        r, g, b = img.getpixel((mid_x, y))
        
        # Check if the pixel is bright white/light gray
        if r > 200 and g > 200 and b > 200:
            white_line_y = y
            break

    # 5. CROP BELOW THE LINE
    if white_line_y != -1:
        start_y = white_line_y + 3  # Start just below the white line
    else:
        print("Warning: White line not found. Using default fallback.")
        start_y = height - 40     # Fallback if screen is empty

    capture_height = height - start_y
    if capture_height <= 0:
        capture_height = 30

    # Crop format: (left, upper, right, lower)
    crop_box = (0, start_y, width, start_y + capture_height)
    cropped_img = img.crop(crop_box)

    # 6. SAVE DEBUG SCREENSHOT
    debug_dir = "debug_output"
    os.makedirs(debug_dir, exist_ok=True)
    debug_path = os.path.join(debug_dir, "runtime_target_crop.png")
    cropped_img.save(debug_path)
    print(f"[Debug] Cropped image saved to: {os.path.abspath(debug_path)}")

    # 7. READ TEXT USING OCR
    # --psm 6 tells Tesseract it's looking at a uniform block of text
    text = pytesseract.image_to_string(cropped_img, config='--psm 6')
    
    print("\n--- EXTRACTED TEXT ---")
    print(text.strip())
    print("----------------------")

if __name__ == "__main__":
    capture_and_read()
