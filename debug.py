import os
import win32gui
import pytesseract
from PIL import ImageGrab

# IMPORTANT: Point this to where Tesseract is installed on your Windows machine
pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'

def capture_and_read(hwnd, debug_dir, save_debug_image=True):
    """
    Captures a specific window, finds the bottom white line, 
    crops below it, runs OCR, and optionally saves a debug image.
    """
    if save_debug_image and not os.path.exists(debug_dir):
        os.makedirs(debug_dir)

    # 1. GET COORDINATES (Replicating C#: GetClientRect + ClientToScreen)
    # This ensures we only capture the inner window, excluding title bars
    left_top = win32gui.ClientToScreen(hwnd, (0, 0))
    _, _, client_right, client_bottom = win32gui.GetClientRect(hwnd)
    
    width = client_right
    height = client_bottom
    screen_x, screen_y = left_top

    # Capture the screen region into memory
    bbox = (screen_x, screen_y, screen_x + width, screen_y + height)
    full_content = ImageGrab.grab(bbox=bbox).convert('RGB')

    # 2. FIND THE WHITE BORDER LINE
    white_line_y = -1
    mid_x = width // 2

    # Loop from bottom up to the middle of the window
    for y in range(height - 1, height // 2, -1):
        r, g, b = full_content.getpixel((mid_x, y))
        
        # Calculate brightness (Luminance formula)
        brightness = (0.299 * r + 0.587 * g + 0.114 * b) / 255.0
        
        if brightness > 0.85:
            white_line_y = y
            break

    # 3. CROP BELOW THE BORDER
    # Start 3 pixels below the white line to avoid OCR noise
    start_y = (white_line_y + 3) if white_line_y != -1 else (height - 35)
    capture_height = height - start_y
    
    if capture_height <= 0:
        capture_height = 30

    crop_region = (0, start_y, width, start_y + capture_height)
    raw_crop = full_content.crop(crop_region)

    # 4. OPTIONAL: SAVE DEBUG SCREENSHOT
    if save_debug_image:
        save_path = os.path.join(debug_dir, "1_Target_Crop.png")
        raw_crop.save(save_path)
        print(f"[Debug] Cropped image saved to: {save_path}")

    # 5. RUN OCR IN RUNTIME (In-Memory)
    # --psm 6 tells Tesseract to assume a single uniform block of text
    extracted_text = pytesseract.image_to_string(raw_crop, config='--psm 6')
    
    return extracted_text.strip()

# --- HOW TO TEST IT ---
if __name__ == "__main__":
    # Replace with the exact Title of your target window
    target_window_title = "Your Terminal Window Title"
    
    hwnd = win32gui.FindWindow(None, target_window_title)
    
    if hwnd:
        print(f"Window found! Handle ID: {hwnd}")
        text = capture_and_read(hwnd, debug_dir="debug_output", save_debug_image=True)
        print("\n--- EXTRACTED TEXT ---")
        print(text)
    else:
        print(f"Could not find a window titled '{target_window_title}'")
