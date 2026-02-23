@pytest.fixture(scope="session")
def fads_window():
    """Launches the app once, and cleanly closes it after all tests are done."""
    # ==========================================
    # ⬆️ SESSION SETUP (Runs once at the beginning)
    # ==========================================
    app = launch_fads() 
    window = Desktop(backend="uia").window(title_re=".*FADS PRIME.*")
    window.wait('ready', timeout=10)
    
    # Pass the window to your 100+ tests
    yield window 
    
    # ==========================================
    # ⬇️ SESSION TEARDOWN (Runs once at the very end)
    # ==========================================
    print("\n[🏁 Bulk Execution Complete] Closing FADS and Command Prompt...")
    
    try:
        # Option A: Graceful Exit (Highly Recommended for DOS apps)
        # Ensure we are at the main menu first
        for _ in range(5):
            safe_type(window, "{ESC}", wait_time=0.2)
        
        # Press the specific key to Exit FADS (e.g., F6, or ESC and Y)
        # UPDATE THESE KEYS to match exactly how you manually quit FADS!
        safe_type(window, "{ESC}Y", wait_time=1.0) 
        
        # If quitting FADS drops you into a standard "C:\>" DOS prompt, 
        # type 'exit' to close the black CMD window completely.
        safe_type(window, "exit{ENTER}", wait_time=0.5)
        
    except Exception as e:
        print(f"⚠️ Graceful exit failed: {e}. Forcing application to close...")
        
    finally:
        # Option B: The "Nuke" Option
        # Just in case the keystrokes failed, we force-kill the process
        try:
            app.kill()
            print("✅ FADS process killed.")
        except:
            pass




from PIL import ImageGrab, Image
import os
import pytesseract

def read_bottom_status_bar(window):
    """
    Grabs the bottom left corner and uses a color threshold to erase 
    white text and blue backgrounds, leaving only the black F9=Exit text.
    """
    try:
        rect = window.wrapper_object().rectangle()
        img = ImageGrab.grab(bbox=(rect.left, rect.top, rect.right, rect.bottom))
        width, height = img.size
        
        # 1. MASSIVE CROP: Grab the bottom 250 pixels, and the left half of the screen.
        # This guarantees we capture both the F6 line and the F9 line.
        crop_box = (0, height - 250, width // 2, height)
        cropped_img = img.crop(crop_box)
        
        # 2. Resize 3x for Tesseract clarity
        img_resized = cropped_img.resize((cropped_img.width * 3, cropped_img.height * 3), resample=Image.Resampling.LANCZOS)
        
        # 3. Convert to Grayscale
        gray = img_resized.convert('L')
        
        # 4. THE MAGIC THRESHOLD:
        # White text = ~255. Blue background = ~130. Black text = ~0.
        # This rule says: If it's darker than 80, make it pure black (0). Otherwise, pure white (255).
        # This completely erases the white "Create a print file" text!
        bw_img = gray.point(lambda x: 0 if x < 80 else 255, '1')
        
        # 5. Save the debug image so you can see the magic trick
        ensure_dir_exists(r"C:\FADS_Debug")
        bw_img.save(os.path.join(r"C:\FADS_Debug", "debug_erased_text.png"))
        
        # 6. OCR (Using PSM 6 since it's a block of text, even if most of it is invisible now)
        custom_config = r'--psm 6 -c preserve_interword_spaces=1'
        text = pytesseract.image_to_string(bw_img, config=custom_config)
        
        clean_text = text.strip().replace("‘", "").replace("'", "")
        print(f"   [Scanner] Status Bar reads: '{clean_text}'")
        
        return clean_text

    except Exception as e:
        print(f"❌ [read_bottom_status_bar] Failed: {e}")
        return ""
