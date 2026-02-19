from PIL import Image, ImageGrab, ImageOps, ImageEnhance
import pytesseract

def capture_and_read(window):
    try:
        # 1. Get the exact coordinates of the window
        rect = window.rectangle()
        screenshot = ImageGrab.grab(bbox=(rect.left, rect.top, rect.right, rect.bottom))
        img = screenshot.convert('RGB')
        
        width, height = img.size

        # 2. FIND THE WHITE BORDER LINE
        white_line_y = -1
        mid_x = width // 2

        # Scan the middle column from bottom up to the middle
        for y in range(height - 1, height // 2, -1):
            r, g, b = img.getpixel((mid_x, y))
            
            # Brightness check to find the white line
            if r > 200 and g > 200 and b > 200:
                white_line_y = y
                break

        # 3. CROP BELOW THE BORDER
        # Start 3 pixels below the white line, or fallback to height - 35
        start_y = (white_line_y + 3) if white_line_y != -1 else (height - 35)
        
        # Crop the image: (left, top, right, bottom)
        cropped_img = img.crop((0, start_y, width, height))

        # 4. APPLY YOUR WORKING IMAGE PROCESSING (on the cropped area only)
        crop_width, crop_height = cropped_img.size
        img_resized = cropped_img.resize((crop_width * 3, crop_height * 3), resample=Image.Resampling.LANCZOS)
        
        gray = img_resized.convert('L')
        inverted = ImageOps.invert(gray)
        
        enhancer = ImageEnhance.Sharpness(inverted)
        sharpened = enhancer.enhance(2.0)
        
        bw_img = sharpened.point(lambda x: 0 if x < 180 else 255, '1')

        # SAVE THIS TO CHECK: It should now only show the bottom row!
        bw_img.save("debug_bottom_line_crop.png")

        # 5. Precise OCR Config
        custom_config = r'--psm 6 -c preserve_interword_spaces=1 -c tessedit_char_blacklist=@Q'
        text = pytesseract.image_to_string(bw_img, config=custom_config)
        
        # Extra safety cleanup just in case
        text = text.replace("Q ", "0 ").replace("@", "0")
        
        return text.strip()

    except Exception as e:
        return f"Capture Error: {e}"
